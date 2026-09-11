use anyhow::{Result, bail};
use clap::Parser;
use computer_monitor::protocol::DEFAULT_PORT;
use std::net::SocketAddr;
use std::time::Duration;
use tracing_subscriber::EnvFilter;

#[derive(Parser)]
#[command(
    name = "Computer Agent",
    version,
    about = "Remote Windows telemetry agent"
)]
struct Arguments {
    #[arg(long, default_value_t = default_address())]
    listen: SocketAddr,
    #[arg(long, env = "COMPUTER_MONITOR_TOKEN", default_value = "")]
    token: String,
    #[arg(long, default_value_t = 200, value_parser = clap::value_parser!(u64).range(100..=10000))]
    interval_ms: u64,
}

#[tokio::main]
async fn main() -> Result<()> {
    tracing_subscriber::fmt()
        .with_env_filter(
            EnvFilter::try_from_default_env().unwrap_or_else(|_| EnvFilter::new("info")),
        )
        .with_target(false)
        .compact()
        .init();
    let arguments = Arguments::parse();
    if !arguments.listen.ip().is_loopback() && arguments.token.trim().len() < 12 {
        bail!("LAN mode requires a token with at least 12 characters");
    }
    computer_monitor::server::serve(
        arguments.listen,
        arguments.token,
        Duration::from_millis(arguments.interval_ms),
    )
    .await
}

fn default_address() -> SocketAddr {
    ([127, 0, 0, 1], DEFAULT_PORT).into()
}
