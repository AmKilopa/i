use anyhow::Result;
use clap::Parser;
use computer_monitor::protocol::DEFAULT_PORT;

#[derive(Parser)]
#[command(
    name = "Computer Dashboard",
    version,
    about = "Interactive remote PC telemetry dashboard"
)]
struct Arguments {
    #[arg(long, default_value_t = default_endpoint())]
    connect: String,
    #[arg(long, env = "COMPUTER_MONITOR_TOKEN", default_value = "")]
    token: String,
}

#[tokio::main]
async fn main() -> Result<()> {
    let arguments = Arguments::parse();
    computer_monitor::dashboard::run(arguments.connect, arguments.token).await
}

fn default_endpoint() -> String {
    format!("127.0.0.1:{DEFAULT_PORT}")
}
