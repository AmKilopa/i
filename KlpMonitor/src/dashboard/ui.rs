use super::app::{ConnectionState, DashboardApp};
use crate::protocol::TelemetrySnapshot;
use chrono::{Datelike, Local, Timelike};
use ratatui::Frame;
use ratatui::layout::{Alignment, Constraint, Direction, Layout, Margin, Rect};
use ratatui::style::{Color, Modifier, Style};
use ratatui::text::{Line, Span};
use ratatui::widgets::{Block, BorderType, Borders, Clear, Paragraph};

const BG: Color = Color::Rgb(3, 8, 18);
const PANEL: Color = Color::Rgb(5, 12, 25);
const OUTLINE: Color = Color::Rgb(78, 112, 168);
const OUTLINE_DIM: Color = Color::Rgb(34, 54, 84);
const TEXT: Color = Color::Rgb(224, 233, 246);
const MUTED: Color = Color::Rgb(125, 143, 174);
const CYAN: Color = Color::Rgb(30, 220, 244);
const BLUE: Color = Color::Rgb(75, 143, 255);
const PURPLE: Color = Color::Rgb(193, 104, 255);
const GREEN: Color = Color::Rgb(96, 224, 174);
const RED: Color = Color::Rgb(246, 104, 132);

pub fn draw(frame: &mut Frame, app: &mut DashboardApp) {
    let area = frame.area();
    frame.render_widget(Block::default().style(Style::default().bg(BG)), area);
    app.clear_expired_toast();

    if area.width < 112 || area.height < 30 {
        draw_too_small(frame, area);
        return;
    }

    let shell_area = area.inner(Margin {
        horizontal: 1,
        vertical: 1,
    });
    let shell = Block::default()
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(OUTLINE))
        .style(Style::default().bg(BG))
        .title_top(fetch_title().centered())
        .title_bottom(fetch_footer().centered());
    let content = shell.inner(shell_area).inner(Margin {
        horizontal: 1,
        vertical: 1,
    });
    frame.render_widget(shell, shell_area);

    let columns = fetch_columns(content);
    let left = split_cards(columns[0]);
    let right = split_cards(columns[2]);
    draw_clock(frame, columns[1]);

    if let Some(snapshot) = app.snapshot.as_ref() {
        draw_system(frame, left[0], snapshot);
        draw_resources(frame, left[1], snapshot);
        draw_hardware(frame, right[0], snapshot);
        draw_network(frame, right[1], snapshot);
    } else {
        draw_placeholder(frame, left[0], "SYSTEM", app);
        draw_placeholder(frame, left[1], "RESOURCES", app);
        draw_placeholder(frame, right[0], "HARDWARE", app);
        draw_placeholder(frame, right[1], "NETWORK", app);
    }

    if let Some(toast) = &app.toast {
        draw_toast(frame, area, &toast.message, toast.success);
    }
}

fn fetch_title() -> Line<'static> {
    Line::from(vec![
        Span::styled("● ", Style::default().fg(PURPLE)),
        Span::styled("● ", Style::default().fg(CYAN)),
        Span::styled("●   ", Style::default().fg(BLUE)),
        Span::styled(
            "Klpfetch",
            Style::default().fg(CYAN).add_modifier(Modifier::BOLD),
        ),
        Span::styled("  v0.0.1", Style::default().fg(TEXT)),
        Span::styled("   ● ", Style::default().fg(BLUE)),
        Span::styled("● ", Style::default().fg(CYAN)),
        Span::styled("●", Style::default().fg(PURPLE)),
    ])
}

fn fetch_footer() -> Line<'static> {
    Line::from(vec![
        Span::styled("● ", Style::default().fg(PURPLE)),
        Span::styled("● ", Style::default().fg(CYAN)),
        Span::styled("●     ", Style::default().fg(BLUE)),
        Span::styled("[", Style::default().fg(MUTED)),
        Span::styled("Q", Style::default().fg(CYAN).add_modifier(Modifier::BOLD)),
        Span::styled("]  Quit", Style::default().fg(TEXT)),
        Span::styled("     ● ", Style::default().fg(BLUE)),
        Span::styled("● ", Style::default().fg(CYAN)),
        Span::styled("●", Style::default().fg(PURPLE)),
    ])
}

fn fetch_columns(area: Rect) -> [Rect; 3] {
    let gaps = 4;
    let available = area.width.saturating_sub(gaps);
    let preferred_side = available.saturating_mul(30) / 100;
    let maximum_side = available.saturating_sub(40) / 2;
    let side = preferred_side.min(maximum_side).max(24);
    let center = available.saturating_sub(side.saturating_mul(2));
    [
        Rect::new(area.x, area.y, side, area.height),
        Rect::new(area.x + side + 2, area.y, center, area.height),
        Rect::new(area.x + side + center + 4, area.y, side, area.height),
    ]
}

fn split_cards(area: Rect) -> [Rect; 2] {
    let usable = area.height.saturating_sub(1);
    let top_height = usable.saturating_mul(45) / 100;
    [
        Rect::new(area.x, area.y, area.width, top_height),
        Rect::new(
            area.x,
            area.y + top_height + 1,
            area.width,
            usable.saturating_sub(top_height),
        ),
    ]
}

fn draw_clock(frame: &mut Frame, area: Rect) {
    let height = area.height.saturating_mul(58).div_ceil(100).clamp(17, 23);
    let card_area = centered_rect(area.width, height.min(area.height), area);
    let block = Block::default()
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(OUTLINE))
        .style(Style::default().bg(PANEL));
    let inner = block.inner(card_area).inner(Margin {
        horizontal: 1,
        vertical: 0,
    });
    frame.render_widget(block, card_area);

    let now = Local::now();
    if inner.width < 49 || inner.height < 13 {
        let content = vec![
            Line::from(Span::styled(
                format!("{:02}:{:02}:{:02}", now.hour(), now.minute(), now.second()),
                Style::default().fg(CYAN).add_modifier(Modifier::BOLD),
            )),
            Line::from(""),
            Line::from(Span::styled(fetch_date(&now), Style::default().fg(TEXT))),
        ];
        frame.render_widget(
            Paragraph::new(content).alignment(Alignment::Center),
            centered_rect(inner.width, 3, inner),
        );
        return;
    }

    let mut content = vec![
        Line::from(Span::styled("━━━━", Style::default().fg(PURPLE))),
        Line::from(""),
    ];
    content.extend(render_time(now.hour(), now.minute(), now.second()));
    content.push(Line::from(""));
    content.push(Line::from(Span::styled(
        "━━━━━━━━━━━━━━━━",
        Style::default().fg(CYAN),
    )));
    content.push(Line::from(""));
    content.push(Line::from(vec![
        Span::styled("▣  ", Style::default().fg(CYAN)),
        Span::styled(fetch_date(&now), Style::default().fg(TEXT)),
    ]));
    content.push(Line::from(""));
    content.push(Line::from(vec![
        Span::styled("───── ┆ ───── ", Style::default().fg(OUTLINE)),
        Span::styled("◇", Style::default().fg(PURPLE)),
        Span::styled(" ───── ┆ ─────", Style::default().fg(OUTLINE)),
    ]));
    frame.render_widget(
        Paragraph::new(content).alignment(Alignment::Center),
        centered_rect(inner.width, 13, inner),
    );
}

fn render_time(hour: u32, minute: u32, second: u32) -> Vec<Line<'static>> {
    let value = format!("{hour:02}:{minute:02}:{second:02}");
    (0..5)
        .map(|row| {
            let mut spans = Vec::new();
            for (index, character) in value.chars().enumerate() {
                let color = if character == ':' {
                    PURPLE
                } else {
                    gradient_color(index as f64 / 7.0)
                };
                spans.push(Span::styled(
                    time_pattern(character)[row].to_owned(),
                    Style::default().fg(color).add_modifier(Modifier::BOLD),
                ));
                if index + 1 < value.len() {
                    spans.push(Span::raw(" "));
                }
            }
            Line::from(spans)
        })
        .collect()
}

fn time_pattern(character: char) -> &'static [&'static str; 5] {
    match character {
        '0' => &["█████", "█   █", "█   █", "█   █", "█████"],
        '1' => &["  ██ ", " ███ ", "  ██ ", "  ██ ", "█████"],
        '2' => &["█████", "    █", "█████", "█    ", "█████"],
        '3' => &["█████", "    █", "█████", "    █", "█████"],
        '4' => &["█   █", "█   █", "█████", "    █", "    █"],
        '5' => &["█████", "█    ", "█████", "    █", "█████"],
        '6' => &["█████", "█    ", "█████", "█   █", "█████"],
        '7' => &["█████", "    █", "   █ ", "  █  ", "  █  "],
        '8' => &["█████", "█   █", "█████", "█   █", "█████"],
        '9' => &["█████", "█   █", "█████", "    █", "█████"],
        ':' => &["     ", "  ◆  ", "     ", "  ◆  ", "     "],
        _ => &["     ", "     ", "     ", "     ", "     "],
    }
}

fn fetch_date(now: &chrono::DateTime<Local>) -> String {
    const WEEKDAYS: [&str; 7] = [
        "MONDAY",
        "TUESDAY",
        "WEDNESDAY",
        "THURSDAY",
        "FRIDAY",
        "SATURDAY",
        "SUNDAY",
    ];
    const MONTHS: [&str; 12] = [
        "JANUARY",
        "FEBRUARY",
        "MARCH",
        "APRIL",
        "MAY",
        "JUNE",
        "JULY",
        "AUGUST",
        "SEPTEMBER",
        "OCTOBER",
        "NOVEMBER",
        "DECEMBER",
    ];
    let weekday = WEEKDAYS[now.weekday().num_days_from_monday() as usize];
    let month = MONTHS[now.month0() as usize];
    format!("{weekday}, {} {month} {}", now.day(), now.year())
}

fn draw_system(frame: &mut Frame, area: Rect, snapshot: &TelemetrySnapshot) {
    let block = fetch_block("SYSTEM");
    let inner = block.inner(area).inner(Margin {
        horizontal: 1,
        vertical: 1,
    });
    frame.render_widget(block, area);
    draw_info_rows(
        frame,
        inner,
        vec![
            InfoRow::new("♙", "User", snapshot.host.user_name.clone(), CYAN),
            InfoRow::new("▣", "Host", snapshot.host.hostname.clone(), CYAN),
            InfoRow::new(
                "⊞",
                "OS",
                format!("{} {}", snapshot.host.os_name, snapshot.host.os_version),
                CYAN,
            ),
            InfoRow::new("⌘", "Kernel", snapshot.host.kernel_version.clone(), CYAN),
            InfoRow::new(
                "◷",
                "Uptime",
                format_uptime(snapshot.host.uptime_seconds),
                CYAN,
            ),
        ],
    );
}

fn draw_hardware(frame: &mut Frame, area: Rect, snapshot: &TelemetrySnapshot) {
    let block = fetch_block("HARDWARE");
    let inner = block.inner(area).inner(Margin {
        horizontal: 1,
        vertical: 1,
    });
    frame.render_widget(block, area);
    let gpu = snapshot
        .gpu
        .as_ref()
        .map_or_else(|| "Not detected".to_owned(), |value| value.name.clone());
    let disk_total = snapshot
        .disks
        .iter()
        .map(|disk| disk.total_bytes)
        .sum::<u64>();
    let storage_kind = snapshot
        .disks
        .iter()
        .find(|disk| !disk.kind.is_empty())
        .map_or("Storage", |disk| disk.kind.as_str());
    draw_info_rows(
        frame,
        inner,
        vec![
            InfoRow::new("▦", "CPU", snapshot.host.cpu_model.clone(), CYAN),
            InfoRow::new("▣", "GPU", gpu, CYAN),
            InfoRow::new(
                "▤",
                "RAM",
                format!(
                    "{} physical memory",
                    format_bytes(snapshot.memory.total_bytes)
                ),
                CYAN,
            ),
            InfoRow::new(
                "▰",
                "Storage",
                format!("{} {storage_kind}", format_bytes(disk_total)),
                CYAN,
            ),
        ],
    );
}

fn draw_network(frame: &mut Frame, area: Rect, snapshot: &TelemetrySnapshot) {
    let block = fetch_block("NETWORK");
    let inner = block.inner(area).inner(Margin {
        horizontal: 1,
        vertical: 1,
    });
    frame.render_widget(block, area);
    let (address, interface) = local_network(snapshot);
    draw_info_rows(
        frame,
        inner,
        vec![
            InfoRow::new("⌘", "Local IP", address, CYAN),
            InfoRow::new("◇", "Interface", interface, CYAN),
            InfoRow::new(
                "↓",
                "Download",
                format_rate(snapshot.network.received_bytes_per_second),
                GREEN,
            ),
            InfoRow::new(
                "↑",
                "Upload",
                format_rate(snapshot.network.transmitted_bytes_per_second),
                BLUE,
            ),
            InfoRow::new(
                "◌",
                "Received",
                format_bytes(snapshot.network.total_received_bytes),
                CYAN,
            ),
            InfoRow::new(
                "◍",
                "Sent",
                format_bytes(snapshot.network.total_transmitted_bytes),
                PURPLE,
            ),
        ],
    );
}

fn draw_resources(frame: &mut Frame, area: Rect, snapshot: &TelemetrySnapshot) {
    let block = fetch_block("RESOURCES");
    let inner = block.inner(area).inner(Margin {
        horizontal: 1,
        vertical: 1,
    });
    frame.render_widget(block, area);
    let rows = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(4),
            Constraint::Length(4),
            Constraint::Length(4),
        ])
        .spacing(if inner.height >= 14 { 1 } else { 0 })
        .split(inner);

    draw_resource(
        frame,
        rows[0],
        ResourceRow {
            label: "CPU".to_owned(),
            percent: snapshot.cpu.usage_percent as f64 / 100.0,
            detail: format!(
                "{:.2} GHz  ·  {} threads",
                snapshot.cpu.frequency_mhz as f64 / 1000.0,
                snapshot.host.logical_cores
            ),
        },
    );
    let memory_ratio = ratio(snapshot.memory.used_bytes, snapshot.memory.total_bytes);
    draw_resource(
        frame,
        rows[1],
        ResourceRow {
            label: "Memory".to_owned(),
            percent: memory_ratio,
            detail: format!(
                "{} / {}",
                format_bytes(snapshot.memory.used_bytes),
                format_bytes(snapshot.memory.total_bytes)
            ),
        },
    );
    let disk = primary_disk(snapshot);
    let disk_total = disk.map_or(0, |value| value.total_bytes);
    let disk_free = disk.map_or(0, |value| value.available_bytes);
    let disk_ratio = ratio(disk_total.saturating_sub(disk_free), disk_total);
    let mount = disk.map_or("Disk", |value| value.mount_point.as_str());
    draw_resource(
        frame,
        rows[2],
        ResourceRow {
            label: format!("Disk ({mount})"),
            percent: disk_ratio,
            detail: format!(
                "{} / {}",
                format_bytes(disk_total.saturating_sub(disk_free)),
                format_bytes(disk_total)
            ),
        },
    );
}

struct InfoRow {
    glyph: &'static str,
    label: &'static str,
    value: String,
    color: Color,
}

impl InfoRow {
    fn new(glyph: &'static str, label: &'static str, value: String, color: Color) -> Self {
        Self {
            glyph,
            label,
            value,
            color,
        }
    }
}

fn draw_info_rows(frame: &mut Frame, area: Rect, rows: Vec<InfoRow>) {
    let count = rows.len();
    let spacing = usize::from(area.height as usize >= count.saturating_mul(2).saturating_sub(1));
    let layout = Layout::default()
        .direction(Direction::Vertical)
        .constraints(vec![Constraint::Length(1); count])
        .spacing(spacing as u16)
        .split(area);
    for (index, row) in rows.into_iter().enumerate() {
        let value_width = layout[index].width.saturating_sub(18) as usize;
        frame.render_widget(
            Paragraph::new(Line::from(vec![
                Span::styled(
                    format!("{}  ", row.glyph),
                    Style::default().fg(TEXT).add_modifier(Modifier::BOLD),
                ),
                Span::styled(format!("{:<10}", row.label), Style::default().fg(TEXT)),
                Span::styled(":  ", Style::default().fg(MUTED)),
                Span::styled(
                    compact_text(&row.value, value_width),
                    Style::default().fg(row.color),
                ),
            ])),
            layout[index],
        );
    }
}

struct ResourceRow {
    label: String,
    percent: f64,
    detail: String,
}

fn draw_resource(frame: &mut Frame, area: Rect, resource: ResourceRow) {
    let rows = Layout::default()
        .direction(Direction::Vertical)
        .constraints([
            Constraint::Length(1),
            Constraint::Length(1),
            Constraint::Length(1),
            Constraint::Length(1),
        ])
        .split(area);
    let header = Layout::default()
        .direction(Direction::Horizontal)
        .constraints([Constraint::Min(10), Constraint::Length(5)])
        .split(rows[0]);
    frame.render_widget(
        Paragraph::new(resource.label).style(Style::default().fg(TEXT)),
        header[0],
    );
    frame.render_widget(
        Paragraph::new(format!("{:>4.0}%", resource.percent * 100.0))
            .alignment(Alignment::Right)
            .style(Style::default().fg(TEXT)),
        header[1],
    );
    frame.render_widget(
        Paragraph::new(gradient_bar(
            resource.percent,
            rows[1].width.saturating_sub(1) as usize,
        )),
        rows[1],
    );
    frame.render_widget(
        Paragraph::new(resource.detail).style(Style::default().fg(MUTED)),
        rows[2],
    );
}

fn gradient_bar(value: f64, width: usize) -> Line<'static> {
    let segments = width.clamp(8, 48);
    let filled = (value.clamp(0.0, 1.0) * segments as f64).round() as usize;
    let spans = (0..segments)
        .map(|index| {
            if index < filled {
                Span::styled(
                    "▮",
                    Style::default().fg(gradient_color(index as f64 / segments as f64)),
                )
            } else {
                Span::styled("▮", Style::default().fg(OUTLINE_DIM))
            }
        })
        .collect::<Vec<_>>();
    Line::from(spans)
}

fn gradient_color(position: f64) -> Color {
    let value = position.clamp(0.0, 1.0);
    if value < 0.5 {
        mix_color((30, 220, 244), (75, 143, 255), value * 2.0)
    } else {
        mix_color((75, 143, 255), (193, 104, 255), (value - 0.5) * 2.0)
    }
}

fn mix_color(from: (u8, u8, u8), to: (u8, u8, u8), amount: f64) -> Color {
    let channel =
        |start: u8, end: u8| (start as f64 + (end as f64 - start as f64) * amount).round() as u8;
    Color::Rgb(
        channel(from.0, to.0),
        channel(from.1, to.1),
        channel(from.2, to.2),
    )
}

fn draw_placeholder(frame: &mut Frame, area: Rect, title: &str, app: &DashboardApp) {
    let (label, color) = match &app.connection {
        ConnectionState::Connecting => ("Connecting...", CYAN),
        ConnectionState::Offline(error) => (
            if error.is_empty() {
                "Waiting for computer"
            } else {
                "Reconnecting..."
            },
            RED,
        ),
        ConnectionState::Online => ("Loading data...", GREEN),
    };
    frame.render_widget(
        Paragraph::new(label)
            .alignment(Alignment::Center)
            .style(Style::default().fg(color))
            .block(fetch_block(title)),
        area,
    );
}

fn fetch_block(title: &str) -> Block<'static> {
    Block::default()
        .title_top(
            Line::from(Span::styled(
                format!(" {title} "),
                Style::default().fg(CYAN).add_modifier(Modifier::BOLD),
            ))
            .centered(),
        )
        .borders(Borders::ALL)
        .border_type(BorderType::Rounded)
        .border_style(Style::default().fg(OUTLINE))
        .style(Style::default().bg(PANEL))
}

fn local_network(snapshot: &TelemetrySnapshot) -> (String, String) {
    for interface in &snapshot.network.interfaces {
        for address in &interface.addresses {
            let plain = address.split('/').next().unwrap_or(address);
            if plain.contains('.') && !plain.starts_with("127.") && plain != "0.0.0.0" {
                return (plain.to_owned(), interface.name.clone());
            }
        }
    }
    ("Not available".to_owned(), "Not available".to_owned())
}

fn primary_disk(snapshot: &TelemetrySnapshot) -> Option<&crate::protocol::DiskStats> {
    snapshot
        .disks
        .iter()
        .find(|disk| disk.mount_point.eq_ignore_ascii_case("C:\\"))
        .or_else(|| snapshot.disks.iter().find(|disk| !disk.removable))
        .or_else(|| snapshot.disks.first())
}

fn compact_text(value: &str, maximum: usize) -> String {
    if maximum == 0 {
        return String::new();
    }
    if value.chars().count() <= maximum {
        return value.to_owned();
    }
    let keep = maximum.saturating_sub(1);
    format!("{}…", value.chars().take(keep).collect::<String>())
}

fn centered_rect(width: u16, height: u16, area: Rect) -> Rect {
    let width = width.min(area.width);
    let height = height.min(area.height);
    Rect {
        x: area.x + area.width.saturating_sub(width) / 2,
        y: area.y + area.height.saturating_sub(height) / 2,
        width,
        height,
    }
}

fn ratio(value: u64, total: u64) -> f64 {
    if total == 0 {
        0.0
    } else {
        value as f64 / total as f64
    }
}

fn format_bytes(value: u64) -> String {
    const UNITS: [&str; 5] = ["B", "KB", "MB", "GB", "TB"];
    let mut amount = value as f64;
    let mut index = 0;
    while amount >= 1024.0 && index < UNITS.len() - 1 {
        amount /= 1024.0;
        index += 1;
    }
    if index == 0 {
        format!("{} {}", value, UNITS[index])
    } else if amount >= 100.0 {
        format!("{amount:.0} {}", UNITS[index])
    } else {
        format!("{amount:.1} {}", UNITS[index])
    }
}

fn format_rate(value: u64) -> String {
    format!("{}/s", format_bytes(value))
}

fn format_uptime(seconds: u64) -> String {
    let days = seconds / 86_400;
    let hours = seconds % 86_400 / 3_600;
    let minutes = seconds % 3_600 / 60;
    if days > 0 {
        format!("{days}d {hours}h {minutes}m")
    } else if hours > 0 {
        format!("{hours}h {minutes}m")
    } else {
        format!("{minutes}m")
    }
}

fn draw_toast(frame: &mut Frame, area: Rect, message: &str, success: bool) {
    let width = (message.chars().count() as u16 + 8).clamp(28, area.width.saturating_sub(4));
    let toast = Rect {
        x: area.x + area.width.saturating_sub(width) / 2,
        y: area.y + 2,
        width,
        height: 3,
    };
    let color = if success { GREEN } else { RED };
    frame.render_widget(Clear, toast);
    frame.render_widget(
        Paragraph::new(message)
            .alignment(Alignment::Center)
            .style(Style::default().fg(TEXT).bg(PANEL))
            .block(
                Block::default()
                    .borders(Borders::ALL)
                    .border_type(BorderType::Rounded)
                    .border_style(Style::default().fg(color)),
            ),
        toast,
    );
}

fn draw_too_small(frame: &mut Frame, area: Rect) {
    frame.render_widget(
        Paragraph::new(vec![
            Line::from(Span::styled(
                "Klpfetch needs a larger terminal",
                Style::default().fg(TEXT).add_modifier(Modifier::BOLD),
            )),
            Line::from(Span::styled(
                format!(
                    "{}×{} available  ·  112×30 required",
                    area.width, area.height
                ),
                Style::default().fg(MUTED),
            )),
        ])
        .alignment(Alignment::Center),
        centered_rect(56, 3, area),
    );
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::dashboard::app::DashboardApp;
    use crate::protocol::{
        CpuStats, DiskStats, GpuStats, HostInfo, MemoryStats, NetworkInterfaceStats, NetworkStats,
        PROTOCOL_VERSION, ProcessStats,
    };
    use ratatui::Terminal;
    use ratatui::backend::TestBackend;
    use tokio::sync::mpsc;

    #[test]
    fn renders_klpfetch_reference_layout() {
        let backend = TestBackend::new(150, 42);
        let mut terminal = Terminal::new(backend).unwrap();
        let (commands, _) = mpsc::unbounded_channel();
        let mut app = DashboardApp::new(commands);
        app.snapshot = Some(sample_snapshot());
        terminal.draw(|frame| draw(frame, &mut app)).unwrap();
        let content = screen_content(&terminal);
        assert!(content.contains("Klpfetch"));
        assert!(content.contains("v0.0.1"));
        assert!(content.contains("SYSTEM"));
        assert!(content.contains("RESOURCES"));
        assert!(content.contains("HARDWARE"));
        assert!(content.contains("NETWORK"));
        assert!(content.contains("Quit"));
        assert!(!content.contains("Theme"));
        assert!(!content.contains("ДИНАМИКА"));
        assert!(!content.contains("ПРИЛОЖЕНИЯ"));
        assert!(!content.contains("Spotify"));
        assert!(!content.contains("Opera"));
    }

    #[test]
    fn clock_column_is_geometrically_centered() {
        let area = Rect::new(2, 2, 146, 36);
        let columns = fetch_columns(area);
        assert_eq!(columns[0].width, columns[2].width);
        assert_eq!(columns[1].x + columns[1].width / 2, area.x + area.width / 2);
    }

    #[test]
    fn time_contains_seconds_in_large_font() {
        let time = render_time(14, 47, 32);
        assert_eq!(time.len(), 5);
        assert!(time.iter().all(|line| line.width() == 47));
    }

    #[test]
    fn renders_small_terminal_message() {
        let backend = TestBackend::new(60, 16);
        let mut terminal = Terminal::new(backend).unwrap();
        let (commands, _) = mpsc::unbounded_channel();
        let mut app = DashboardApp::new(commands);
        terminal.draw(|frame| draw(frame, &mut app)).unwrap();
        assert!(screen_content(&terminal).contains("larger terminal"));
    }

    fn screen_content(terminal: &Terminal<ratatui::backend::TestBackend>) -> String {
        terminal
            .backend()
            .buffer()
            .content
            .iter()
            .map(|cell| cell.symbol())
            .collect::<String>()
    }

    fn sample_snapshot() -> TelemetrySnapshot {
        TelemetrySnapshot {
            protocol_version: PROTOCOL_VERSION,
            sequence: 17,
            host: HostInfo {
                user_name: "lenovo".to_owned(),
                hostname: "LEGION-5".to_owned(),
                os_name: "Windows".to_owned(),
                os_version: "11 Pro 23H2".to_owned(),
                kernel_version: "10.0.22631.3737".to_owned(),
                cpu_model: "AMD Ryzen 5 5600H".to_owned(),
                physical_cores: 6,
                logical_cores: 12,
                uptime_seconds: 187_080,
            },
            cpu: CpuStats {
                usage_percent: 23.0,
                frequency_mhz: 2_810,
                per_core_percent: vec![23.0; 12],
            },
            memory: MemoryStats {
                total_bytes: 16 * 1024 * 1024 * 1024,
                used_bytes: 9 * 1024 * 1024 * 1024,
                available_bytes: 7 * 1024 * 1024 * 1024,
                ..MemoryStats::default()
            },
            network: NetworkStats {
                received_bytes_per_second: 1_250_000,
                transmitted_bytes_per_second: 480_000,
                total_received_bytes: 82_000_000_000,
                total_transmitted_bytes: 12_000_000_000,
                interfaces: vec![NetworkInterfaceStats {
                    name: "Wi-Fi".to_owned(),
                    addresses: vec!["192.168.1.24/24".to_owned()],
                    ..NetworkInterfaceStats::default()
                }],
            },
            disks: vec![DiskStats {
                name: "System".to_owned(),
                mount_point: "C:\\".to_owned(),
                file_system: "NTFS".to_owned(),
                kind: "NVMe SSD".to_owned(),
                total_bytes: 512 * 1024 * 1024 * 1024,
                available_bytes: 202 * 1024 * 1024 * 1024,
                removable: false,
            }],
            processes: vec![
                ProcessStats {
                    name: "opera.exe".to_owned(),
                    ..ProcessStats::default()
                },
                ProcessStats {
                    name: "Spotify.exe".to_owned(),
                    ..ProcessStats::default()
                },
            ],
            gpu: Some(GpuStats {
                name: "NVIDIA GeForce RTX 3060".to_owned(),
                usage_percent: 31.0,
                memory_used_bytes: 2 * 1024 * 1024 * 1024,
                memory_total_bytes: 6 * 1024 * 1024 * 1024,
                temperature_celsius: Some(52.0),
            }),
            ..TelemetrySnapshot::default()
        }
    }
}
