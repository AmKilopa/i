use crate::collector::{TelemetryCollector, warm_up_collector};
use crate::protocol::{
    AgentCommand, AgentEvent, HealthResponse, PROTOCOL_VERSION, TelemetrySnapshot,
};
use anyhow::{Context, Result};
use axum::Json;
use axum::Router;
use axum::extract::State;
use axum::extract::ws::{Message, WebSocket, WebSocketUpgrade};
use axum::http::{HeaderMap, StatusCode};
use axum::response::{IntoResponse, Response};
use axum::routing::get;
use futures_util::{SinkExt, StreamExt};
use serde::Serialize;
use std::net::SocketAddr;
use std::sync::Arc;
use std::time::Duration;
use sysinfo::{Pid, System};
use tokio::sync::watch;

#[derive(Clone)]
struct AppState {
    token: Arc<str>,
    telemetry: watch::Receiver<TelemetrySnapshot>,
}

pub async fn serve(address: SocketAddr, token: String, interval: Duration) -> Result<()> {
    let mut collector = TelemetryCollector::new();
    warm_up_collector(&mut collector).await;
    let initial = collector.snapshot();
    let (telemetry_tx, telemetry_rx) = watch::channel(initial);

    tokio::spawn(async move {
        let mut ticker = tokio::time::interval(interval);
        ticker.set_missed_tick_behavior(tokio::time::MissedTickBehavior::Skip);
        loop {
            ticker.tick().await;
            if telemetry_tx.send(collector.snapshot()).is_err() {
                break;
            }
        }
    });

    let state = AppState {
        token: Arc::from(token),
        telemetry: telemetry_rx,
    };
    let app = Router::new()
        .route("/health", get(health))
        .route("/api/v1/snapshot", get(snapshot))
        .route("/api/v1/stream", get(stream))
        .with_state(state);
    let listener = tokio::net::TcpListener::bind(address)
        .await
        .with_context(|| format!("cannot listen on {address}"))?;
    tracing::info!(%address, "computer agent is listening");
    axum::serve(listener, app)
        .with_graceful_shutdown(shutdown_signal())
        .await
        .context("agent server stopped unexpectedly")
}

async fn health() -> Json<HealthResponse> {
    Json(HealthResponse {
        status: "ok".to_owned(),
        service: "Computer Agent".to_owned(),
        protocol_version: PROTOCOL_VERSION,
    })
}

async fn snapshot(
    State(state): State<AppState>,
    headers: HeaderMap,
) -> Result<Json<TelemetrySnapshot>, StatusCode> {
    if !is_authorized(&headers, &state.token) {
        return Err(StatusCode::UNAUTHORIZED);
    }
    Ok(Json(state.telemetry.borrow().clone()))
}

async fn stream(
    State(state): State<AppState>,
    headers: HeaderMap,
    websocket: WebSocketUpgrade,
) -> Response {
    if !is_authorized(&headers, &state.token) {
        return StatusCode::UNAUTHORIZED.into_response();
    }
    websocket
        .on_upgrade(move |socket| stream_telemetry(socket, state.telemetry))
        .into_response()
}

async fn stream_telemetry(socket: WebSocket, mut telemetry: watch::Receiver<TelemetrySnapshot>) {
    let (mut sender, mut receiver) = socket.split();
    let first = telemetry.borrow().clone();
    if send_json(&mut sender, &first).await.is_err() {
        return;
    }
    loop {
        tokio::select! {
            changed = telemetry.changed() => {
                if changed.is_err() {
                    return;
                }
                let snapshot = telemetry.borrow_and_update().clone();
                if send_json(&mut sender, &snapshot).await.is_err() {
                    return;
                }
            }
            message = receiver.next() => {
                let Some(Ok(message)) = message else {
                    return;
                };
                match message {
                    Message::Text(text) => {
                        let result = match serde_json::from_str::<AgentCommand>(&text) {
                            Ok(command) => execute_command(command).await,
                            Err(error) => AgentEvent::ActionResult {
                                request_id: 0,
                                success: false,
                                message: format!("Некорректная команда: {error}"),
                            },
                        };
                        if send_json(&mut sender, &result).await.is_err() {
                            return;
                        }
                    }
                    Message::Ping(payload) => {
                        if sender.send(Message::Pong(payload)).await.is_err() {
                            return;
                        }
                    }
                    Message::Close(_) => return,
                    _ => {}
                }
            }
        }
    }
}

async fn send_json<T: Serialize>(
    sender: &mut futures_util::stream::SplitSink<WebSocket, Message>,
    value: &T,
) -> Result<()> {
    let payload = serde_json::to_string(value)?;
    sender.send(Message::Text(payload.into())).await?;
    Ok(())
}

async fn execute_command(command: AgentCommand) -> AgentEvent {
    match command {
        AgentCommand::TerminateProcess { request_id, pid } => {
            let result = tokio::task::spawn_blocking(move || terminate_process(pid)).await;
            match result {
                Ok(Ok(message)) => AgentEvent::ActionResult {
                    request_id,
                    success: true,
                    message,
                },
                Ok(Err(message)) => AgentEvent::ActionResult {
                    request_id,
                    success: false,
                    message,
                },
                Err(error) => AgentEvent::ActionResult {
                    request_id,
                    success: false,
                    message: format!("Команда завершилась с ошибкой: {error}"),
                },
            }
        }
    }
}

fn terminate_process(pid: u32) -> std::result::Result<String, String> {
    if pid <= 4 {
        return Err("Системный процесс защищён".to_owned());
    }
    if sysinfo::get_current_pid()
        .ok()
        .is_some_and(|current| current.as_u32() == pid)
    {
        return Err("Агент нельзя завершить из панели".to_owned());
    }
    let system = System::new_all();
    let process = system
        .process(Pid::from_u32(pid))
        .ok_or_else(|| format!("Процесс PID {pid} уже не существует"))?;
    let name = process.name().to_string_lossy().into_owned();
    if is_protected_process(&name) {
        return Err(format!("Системный процесс {name} защищён"));
    }
    process
        .kill_and_wait()
        .map_err(|error| format!("Не удалось завершить {name}: {error}"))?;
    Ok(format!("{name} завершён"))
}

fn is_protected_process(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "system"
            | "registry"
            | "smss.exe"
            | "csrss.exe"
            | "wininit.exe"
            | "services.exe"
            | "lsass.exe"
            | "winlogon.exe"
            | "computer-agent.exe"
            | "computer.agent.exe"
            | "klp-pulse-agent.exe"
            | "klppulse.agent.exe"
    )
}

fn is_authorized(headers: &HeaderMap, token: &str) -> bool {
    if token.is_empty() {
        return true;
    }
    let expected = format!("Bearer {token}");
    headers
        .get(axum::http::header::AUTHORIZATION)
        .and_then(|value| value.to_str().ok())
        .is_some_and(|value| value == expected)
}

async fn shutdown_signal() {
    let _ = tokio::signal::ctrl_c().await;
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn protects_critical_windows_processes() {
        assert!(is_protected_process("lsass.exe"));
        assert!(is_protected_process("Computer.Agent.exe"));
        assert!(!is_protected_process("notepad.exe"));
    }
}
