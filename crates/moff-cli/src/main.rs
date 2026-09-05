use std::process::ExitCode;

use clap::{Parser, Subcommand};
use moff_outlook_local::{LocalError, doctor, visible_mail};
use serde_json::json;

#[derive(Debug, Parser)]
#[command(name = "moff", version, about = "Local macOS Outlook screen reader")]
struct Cli {
    #[command(subcommand)]
    command: Option<Command>,
}
#[derive(Debug, Subcommand)]
enum Command {
    Outlook {
        #[command(subcommand)]
        command: OutlookCommand,
    },
    Schema,
}
#[derive(Debug, Subcommand)]
enum OutlookCommand {
    Doctor,
    Mail {
        #[command(subcommand)]
        command: MailCommand,
    },
}
#[derive(Debug, Subcommand)]
enum MailCommand {
    Visible {
        #[arg(long, default_value_t = 25)]
        limit: u16,
    },
}

#[allow(clippy::needless_pass_by_value)]
fn print_json(value: serde_json::Value) {
    println!(
        "{}",
        serde_json::to_string(&value).expect("JSON values serialize")
    );
}
#[allow(clippy::needless_pass_by_value)]
fn print_error(local_error: LocalError) -> ExitCode {
    eprintln!("[{}] {}", local_error.code(), local_error);
    ExitCode::from(2)
}
fn main() -> ExitCode {
    let cli = Cli::parse();
    match cli.command {
        Some(Command::Schema) => {
            print_json(
                json!({"schema":"moff.schema/v1","commands":{"outlook":{"summary":"Local macOS Outlook screen/OCR reader","source":"outlook-screen-local"}}}),
            );
            ExitCode::SUCCESS
        }
        Some(Command::Outlook {
            command: OutlookCommand::Doctor,
        }) => match doctor() {
            Ok(value) => {
                print_json(value);
                ExitCode::SUCCESS
            }
            Err(error) => print_error(error),
        },
        Some(Command::Outlook {
            command:
                OutlookCommand::Mail {
                    command: MailCommand::Visible { limit },
                },
        }) => match visible_mail(limit) {
            Ok(value) => {
                print_json(value);
                ExitCode::SUCCESS
            }
            Err(error) => print_error(error),
        },
        None => {
            print_json(
                json!({"command":"moff","status":"local_only","provider":"outlook-screen-local"}),
            );
            ExitCode::SUCCESS
        }
    }
}
