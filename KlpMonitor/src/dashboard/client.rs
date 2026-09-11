use crate::protocol::{AgentCommand, AgentEvent, PROTOCOL_VERSION, TelemetrySnapshot};
use futures_util::{SinkExt, StreamExt};
use std::time::Duration;
use tokio::sync::{mpsc, watch};
use tokio_tungstenite::connect_async;
use tokio_tungstenite::tungstenite::Message;
use tokio_tungstenite::tungstenite::client::IntoClientRequest;
use tokio_tungstenite::tungstenite::http::HeaderValue;

#[derive(Debug)]
pub enum ClientEvent {
    Connecting,
    Connected,
    Snapshot(Box<TelemetrySnapshot>),
    ActionResult {
        request_id: u64,
        success: bool,
        message: String,
    },
    Disconnected(String),
}

pub fn spawn_connector(
    endpoint: String,
    token: String,
) -> (
    mpsc::UnboundedReceiver<ClientEvent>,
    mpsc::UnboundedSender<AgentCommand>,
    watch::Sender<bool>,
) {
    let (events_tx, events_rx) = mpsc::unbounded_channel();
    let (commands_tx, mut commands_rx) = mpsc::unbounded_channel::<AgentCommand>();
    let (stop_tx, mut stop_rx) = watch::channel(false);
    tokio::spawn(async move {
        let mut delay = 1_u64;
        loop {
            if *stop_rx.borrow() {
                break;
            }
            let _ = events_tx.send(ClientEvent::Connecting);
            while let Ok(command) = commands_rx.try_recv() {
                let request_id = command.request_id();
                let _ = events_tx.send(ClientEvent::ActionResult {
                    request_id,
                    success: false,
                    message: "Компьютер сейчас не подключён".to_owned(),
                });
            }
            match connect_once(
                &endpoint,
                &token,
                &events_tx,
                &mut commands_rx,
                &mut stop_rx,
            )
            .await
            {
                Ok(()) if *stop_rx.borrow() => break,
                Ok(()) => {
                    let _ =
                        events_tx.send(ClientEvent::Disconnected("Connection closed".to_owned()));
                }
                Err(error) => {
                    let _ = events_tx.send(ClientEvent::Disconnected(short_error(&error)));
                }
            }
            tokio::select! {
                _ = tokio::time::sleep(Duration::from_secs(delay)) => {}
                _ = stop_rx.changed() => {
                    if *stop_rx.borrow() {
                        break;
                    }
                }
            }
            delay = (delay * 2).min(5);
        }
    });
    (events_rx, commands_tx, stop_tx)
}

async fn connect_once(
    endpoint: &str,
    token: &str,
    events: &mpsc::UnboundedSender<ClientEvent>,
    commands: &mut mpsc::UnboundedReceiver<AgentCommand>,
    stop: &mut watch::Receiver<bool>,
) -> Result<(), String> {
    let mut request = endpoint
        .into_client_request()
        .map_err(|error| error.to_string())?;
    if !token.is_empty() {
        let value =
            HeaderValue::from_str(&format!("Bearer {token}")).map_err(|error| error.to_string())?;
        request.headers_mut().insert("authorization", value);
    }
    let (mut stream, _) = connect_async(request)
        .await
        .map_err(|error| error.to_string())?;
    let _ = events.send(ClientEvent::Connected);

    loop {
        tokio::select! {
            _ = stop.changed() => {
                if *stop.borrow() {
                    let _ = stream.close(None).await;
                    return Ok(());
                }
            }
            message = stream.next() => {
                match message {
                    Some(Ok(Message::Text(text))) => {
                        if let Ok(AgentEvent::ActionResult { request_id, success, message }) =
                            serde_json::from_str::<AgentEvent>(&text)
                        {
                            let _ = events.send(ClientEvent::ActionResult {
                                request_id,
                                success,
                                message,
                            });
                            continue;
                        }
                        let snapshot: TelemetrySnapshot = serde_json::from_str(&text)
                            .map_err(|error| format!("Invalid telemetry: {error}"))?;
                        if snapshot.protocol_version != PROTOCOL_VERSION {
                            return Err(format!(
                                "Protocol mismatch: agent {}, dashboard {}",
                                snapshot.protocol_version,
                                PROTOCOL_VERSION
                            ));
                        }
                        let _ = events.send(ClientEvent::Snapshot(Box::new(snapshot)));
                    }
                    Some(Ok(Message::Ping(payload))) => {
                        stream.send(Message::Pong(payload)).await.map_err(|error| error.to_string())?;
                    }
                    Some(Ok(Message::Close(_))) | None => return Ok(()),
                    Some(Err(error)) => return Err(error.to_string()),
                    _ => {}
                }
            }
            command = commands.recv() => {
                let Some(command) = command else {
                    return Ok(());
                };
                let payload = serde_json::to_string(&command).map_err(|error| error.to_string())?;
                stream
                    .send(Message::Text(payload.into()))
                    .await
                    .map_err(|error| error.to_string())?;
            }
        }
    }
}

trait CommandRequestId {
    fn request_id(&self) -> u64;
}

impl CommandRequestId for AgentCommand {
    fn request_id(&self) -> u64 {
        match self {
            AgentCommand::TerminateProcess { request_id, .. } => *request_id,
        }
    }
}

pub fn normalize_endpoint(value: &str) -> String {
    let mut endpoint = value.trim().trim_end_matches('/').to_owned();
    if let Some(rest) = endpoint.strip_prefix("http://") {
        endpoint = format!("ws://{rest}");
    } else if let Some(rest) = endpoint.strip_prefix("https://") {
        endpoint = format!("wss://{rest}");
    } else if !endpoint.starts_with("ws://") && !endpoint.starts_with("wss://") {
        endpoint = format!("ws://{endpoint}");
    }
    let authority = endpoint
        .split_once("://")
        .map(|(_, rest)| rest)
        .unwrap_or(&endpoint);
    if !authority.contains('/') {
        endpoint.push_str("/api/v1/stream");
    }
    endpoint
}

fn short_error(error: &str) -> String {
    let line = error.lines().next().unwrap_or(error).trim();
    if line.len() > 110 {
        format!("{}...", &line[..107])
    } else {
        line.to_owned()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn normalizes_plain_host() {
        assert_eq!(
            normalize_endpoint("192.168.1.10:47821"),
            "ws://192.168.1.10:47821/api/v1/stream"
        );
    }

    #[test]
    fn preserves_explicit_path() {
        assert_eq!(
            normalize_endpoint("https://pulse.local/custom"),
            "wss://pulse.local/custom"
        );
    }
}
