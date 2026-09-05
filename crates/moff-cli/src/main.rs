use std::io::{self, Write};
use std::process::ExitCode;

use clap::{ArgAction, Parser, Subcommand};
use moff_core::{ErrorCode, known_source, schema};
use serde_json::{Value, json};

#[derive(Debug, Parser)]
#[allow(
    clippy::struct_excessive_bools,
    reason = "CLI flags are independent user-provided inputs"
)]
#[command(
    name = "moff",
    version,
    about = "Agent-first Microsoft 365 CLI scaffolding"
)]
struct Cli {
    /// Render JSON with indentation instead of compact JSON.
    #[arg(long, global = true)]
    pretty: bool,

    /// Wrap command results in the stable agent response envelope.
    #[arg(long, global = true)]
    envelope: bool,

    /// Explicitly request the default JSON output format.
    #[arg(long, global = true)]
    json: bool,

    /// Select a named local profile; no profile is loaded in scaffold mode.
    #[arg(long, global = true)]
    profile: Option<String>,

    /// Report planned work without making changes.
    #[arg(long, global = true)]
    dry_run: bool,

    /// Confirm a future mutating operation without prompting.
    #[arg(long, global = true)]
    yes: bool,

    /// Disable ANSI color in human diagnostics.
    #[arg(long, global = true, action = ArgAction::SetTrue)]
    no_color: bool,

    #[command(subcommand)]
    command: Option<Command>,
}

#[derive(Debug, Subcommand)]
enum Command {
    /// Authentication scaffolding; never contacts an identity provider.
    Auth,
    /// Local diagnostics scaffolding; never contacts a provider.
    Doctor,
    /// Emit the versioned machine-readable command map.
    Schema {
        /// Restrict the map to one provider source.
        #[arg(long)]
        source: Option<String>,
    },
    /// Outlook scaffolding; never contacts Microsoft Graph.
    Outlook {
        /// Optional Outlook operation name; only known sources are accepted.
        source: Option<String>,
    },
}

struct AppError {
    code: ErrorCode,
    message: String,
}

impl AppError {
    fn unknown_source(source: &str) -> Self {
        Self {
            code: ErrorCode::UnknownSource,
            message: format!("unknown source: {source}"),
        }
    }
}

fn command_data(command: Option<Command>) -> Result<Value, AppError> {
    match command {
        None => Ok(json!({"command": "root", "status": "scaffold"})),
        Some(Command::Auth) => Ok(json!({"command": "auth", "status": "scaffold"})),
        Some(Command::Doctor) => Ok(json!({"command": "doctor", "status": "scaffold"})),
        Some(Command::Schema { source }) => {
            if let Some(source) = source.as_deref()
                && !known_source(source)
            {
                return Err(AppError::unknown_source(source));
            }
            Ok(schema(source.as_deref()))
        }
        Some(Command::Outlook { source }) => {
            if let Some(source) = source.as_deref()
                && !known_source(source)
            {
                return Err(AppError::unknown_source(source));
            }
            Ok(json!({"command": "outlook", "status": "scaffold"}))
        }
    }
}

fn emit_json(value: &Value, pretty: bool) -> io::Result<()> {
    let stdout = io::stdout();
    let mut output = stdout.lock();
    if pretty {
        serde_json::to_writer_pretty(&mut output, value)?;
    } else {
        serde_json::to_writer(&mut output, value)?;
    }
    writeln!(output)
}

fn run(cli: Cli) -> Result<(), AppError> {
    if cli.json && cli.pretty {
        return Err(AppError {
            code: ErrorCode::IncompatibleOutputModes,
            message: "--json and --pretty cannot be used together".to_owned(),
        });
    }

    let data = command_data(cli.command)?;
    let response = if cli.envelope {
        json!({
            "ok": true,
            "data": data,
            "warnings": [],
            "request_id": "scaffold-00000000",
        })
    } else {
        data
    };
    emit_json(&response, cli.pretty).expect("stdout is writable");
    Ok(())
}

fn main() -> ExitCode {
    let cli = Cli::parse();
    let envelope = cli.envelope;
    let pretty = cli.pretty;

    match run(cli) {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("moff: [{}] {}", error.code.as_str(), error.message);
            if envelope {
                let response = json!({
                    "ok": false,
                    "data": {"error": {"code": error.code, "message": error.message}},
                    "warnings": [],
                    "request_id": "scaffold-00000000",
                });
                emit_json(&response, pretty).expect("stdout is writable");
            }
            ExitCode::from(2)
        }
    }
}
