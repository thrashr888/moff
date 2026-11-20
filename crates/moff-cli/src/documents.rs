use std::{path::PathBuf, process::ExitCode};

use clap::Subcommand;
use moff_outlook_owa::{DocumentEdit, DocumentSession};

use super::{SessionArgs, error, output, read_json, safari_write_guard};

#[derive(Debug, Subcommand)]
pub(super) enum DocumentCommand {
    /// Read Word main-body headings, text, tables and links from its signed-in `SharePoint` tab.
    Read {
        #[arg(long)]
        url: String,
    },
    /// Preview appending to a unique heading; --confirm applies the version-checked edit once.
    Edit {
        #[arg(long)]
        url: String,
        #[arg(long)]
        section: String,
        /// JSON: paragraphs [{text, `heading_level`?, link?, bold?}], optional `create_before` heading.
        #[arg(long)]
        body_file: PathBuf,
        #[arg(long, requires = "expected_etag")]
        confirm: bool,
        #[arg(long, requires = "confirm")]
        expected_etag: Option<String>,
    },
}

pub(super) fn run(args: SessionArgs, command: DocumentCommand) -> ExitCode {
    let result = (|| -> Result<(&str, serde_json::Value), String> {
        match command {
            DocumentCommand::Read { url } => {
                let doc =
                    DocumentSession::open(&url, args.safari_window.zip(args.safari_tab), false)
                        .map_err(|e| e.to_string())?;
                Ok(("docs.read", doc.read().map_err(|e| e.to_string())?))
            }
            DocumentCommand::Edit {
                url,
                section,
                body_file,
                confirm,
                expected_etag,
            } => {
                if confirm {
                    safari_write_guard(&args, true)?;
                }
                let edit: DocumentEdit = serde_json::from_value(read_json(&body_file)?).map_err(
                    |_| "invalid document edit; provide paragraphs and optional create_before",
                )?;
                let doc =
                    DocumentSession::open(&url, args.safari_window.zip(args.safari_tab), confirm)
                        .map_err(|e| e.to_string())?;
                let action = if confirm {
                    "docs.edit"
                } else {
                    "docs.edit.preview"
                };
                Ok((
                    action,
                    doc.edit(&section, &edit, expected_etag.as_deref())
                        .map_err(|e| e.to_string())?,
                ))
            }
        }
    })();
    match result {
        Ok((action, value)) => {
            output(action, &value);
            ExitCode::SUCCESS
        }
        Err(message) => error(message),
    }
}
