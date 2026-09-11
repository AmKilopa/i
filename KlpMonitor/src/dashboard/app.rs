use super::client::ClientEvent;
use crate::protocol::{AgentCommand, TelemetrySnapshot};
use crossterm::event::{KeyCode, MouseEvent};
use std::collections::VecDeque;
use std::time::{Duration, Instant};
use tokio::sync::mpsc;

#[derive(Clone, Debug)]
pub enum ConnectionState {
    Connecting,
    Online,
    Offline(String),
}

pub struct Toast {
    pub message: String,
    pub success: bool,
    pub created_at: Instant,
}

pub struct DashboardApp {
    pub connection: ConnectionState,
    pub snapshot: Option<TelemetrySnapshot>,
    pub last_snapshot_at: Option<Instant>,
    pub should_quit: bool,
    pub cpu_history: VecDeque<u64>,
    pub memory_history: VecDeque<u64>,
    pub download_history: VecDeque<u64>,
    pub upload_history: VecDeque<u64>,
    pub toast: Option<Toast>,
    _commands: mpsc::UnboundedSender<AgentCommand>,
}

impl DashboardApp {
    pub fn new(commands: mpsc::UnboundedSender<AgentCommand>) -> Self {
        Self {
            connection: ConnectionState::Connecting,
            snapshot: None,
            last_snapshot_at: None,
            should_quit: false,
            cpu_history: VecDeque::with_capacity(120),
            memory_history: VecDeque::with_capacity(120),
            download_history: VecDeque::with_capacity(120),
            upload_history: VecDeque::with_capacity(120),
            toast: None,
            _commands: commands,
        }
    }

    pub fn handle_client_event(&mut self, event: ClientEvent) {
        match event {
            ClientEvent::Connecting => self.connection = ConnectionState::Connecting,
            ClientEvent::Connected => self.connection = ConnectionState::Online,
            ClientEvent::Disconnected(error) => self.connection = ConnectionState::Offline(error),
            ClientEvent::Snapshot(snapshot) => {
                self.push_history(&snapshot);
                self.snapshot = Some(*snapshot);
                self.last_snapshot_at = Some(Instant::now());
                self.connection = ConnectionState::Online;
            }
            ClientEvent::ActionResult {
                request_id,
                success,
                message,
            } => {
                let _ = request_id;
                self.toast = Some(Toast {
                    message,
                    success,
                    created_at: Instant::now(),
                });
            }
        }
    }

    pub fn handle_key(&mut self, key: KeyCode) {
        if matches!(key, KeyCode::Char('q') | KeyCode::Esc) {
            self.should_quit = true;
        }
    }

    pub fn handle_mouse(&mut self, _event: MouseEvent) {}

    pub fn clear_expired_toast(&mut self) {
        if self
            .toast
            .as_ref()
            .is_some_and(|toast| toast.created_at.elapsed() > Duration::from_secs(4))
        {
            self.toast = None;
        }
    }

    fn push_history(&mut self, snapshot: &TelemetrySnapshot) {
        let memory = percent(snapshot.memory.used_bytes, snapshot.memory.total_bytes) as u64;
        push_limited(
            &mut self.cpu_history,
            snapshot.cpu.usage_percent.round() as u64,
        );
        push_limited(&mut self.memory_history, memory);
        push_limited(
            &mut self.download_history,
            snapshot.network.received_bytes_per_second,
        );
        push_limited(
            &mut self.upload_history,
            snapshot.network.transmitted_bytes_per_second,
        );
    }
}

fn push_limited(history: &mut VecDeque<u64>, value: u64) {
    if history.len() >= 120 {
        history.pop_front();
    }
    history.push_back(value);
}

fn percent(used: u64, total: u64) -> f64 {
    if total == 0 {
        0.0
    } else {
        used as f64 / total as f64 * 100.0
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::protocol::{CpuStats, MemoryStats};

    #[test]
    fn only_exit_keys_close_dashboard() {
        let (commands, _) = mpsc::unbounded_channel();
        let mut app = DashboardApp::new(commands);
        app.handle_key(KeyCode::Tab);
        assert!(!app.should_quit);
        app.handle_key(KeyCode::Char('q'));
        assert!(app.should_quit);
    }

    #[test]
    fn history_starts_with_real_samples() {
        let (commands, _) = mpsc::unbounded_channel();
        let mut app = DashboardApp::new(commands);
        app.handle_client_event(ClientEvent::Snapshot(Box::new(TelemetrySnapshot {
            cpu: CpuStats {
                usage_percent: 37.0,
                ..CpuStats::default()
            },
            memory: MemoryStats {
                total_bytes: 100,
                used_bytes: 62,
                ..MemoryStats::default()
            },
            ..TelemetrySnapshot::default()
        })));
        assert_eq!(app.cpu_history.iter().copied().collect::<Vec<_>>(), [37]);
        assert_eq!(app.memory_history.iter().copied().collect::<Vec<_>>(), [62]);
    }
}
