mod app;
mod client;
mod ui;

use anyhow::Result;
use app::DashboardApp;
use client::spawn_connector;
use crossterm::cursor::{Hide, Show};
use crossterm::event::{
    DisableMouseCapture, EnableMouseCapture, Event, KeyCode, KeyEventKind, KeyModifiers,
};
use crossterm::execute;
use crossterm::terminal::{
    EnterAlternateScreen, LeaveAlternateScreen, disable_raw_mode, enable_raw_mode,
};
use ratatui::Terminal;
use ratatui::backend::CrosstermBackend;
use std::io::{Stdout, stdout};
use std::time::Duration;

pub async fn run(endpoint: String, token: String) -> Result<()> {
    let endpoint = client::normalize_endpoint(&endpoint);
    let (mut events, commands, stop) = spawn_connector(endpoint, token);
    let mut session = TerminalSession::new()?;
    let mut app = DashboardApp::new(commands);

    while !app.should_quit {
        while let Ok(event) = events.try_recv() {
            app.handle_client_event(event);
        }
        session.terminal.draw(|frame| ui::draw(frame, &mut app))?;

        if crossterm::event::poll(Duration::from_millis(50))? {
            match crossterm::event::read()? {
                Event::Key(key) if key.kind == KeyEventKind::Press => {
                    if key.modifiers.contains(KeyModifiers::CONTROL)
                        && matches!(key.code, KeyCode::Char('c'))
                    {
                        app.should_quit = true;
                    } else {
                        app.handle_key(key.code);
                    }
                }
                Event::Mouse(mouse) => app.handle_mouse(mouse),
                Event::Resize(_, _) => {}
                _ => {}
            }
        }
    }

    let _ = stop.send(true);
    Ok(())
}

struct TerminalSession {
    terminal: Terminal<CrosstermBackend<Stdout>>,
}

impl TerminalSession {
    fn new() -> Result<Self> {
        enable_raw_mode()?;
        let mut output = stdout();
        execute!(output, EnterAlternateScreen, EnableMouseCapture, Hide)?;
        let mut terminal = Terminal::new(CrosstermBackend::new(output))?;
        terminal.clear()?;
        Ok(Self { terminal })
    }
}

impl Drop for TerminalSession {
    fn drop(&mut self) {
        let _ = disable_raw_mode();
        let _ = execute!(
            self.terminal.backend_mut(),
            Show,
            DisableMouseCapture,
            LeaveAlternateScreen
        );
        let _ = self.terminal.show_cursor();
    }
}
