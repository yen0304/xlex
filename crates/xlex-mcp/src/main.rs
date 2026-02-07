use clap::Parser;
use rmcp::{transport::stdio, ServiceExt};

use xlex_mcp::server::XlexMcpServer;

#[derive(Parser)]
#[command(name = "xlex-mcp", about = "MCP server for xlex Excel manipulation")]
struct Cli {
    /// Transport to use
    #[arg(long, default_value = "stdio")]
    transport: String,
}

#[tokio::main]
async fn main() -> anyhow::Result<()> {
    let cli = Cli::parse();

    match cli.transport.as_str() {
        "stdio" => {
            let service = XlexMcpServer::new().serve(stdio()).await?;
            service.waiting().await?;
        }
        other => {
            eprintln!("Unsupported transport: {other}. Only 'stdio' is currently supported.");
            std::process::exit(1);
        }
    }

    Ok(())
}
