//! Shared command contracts for the Moff agent-first CLI.

use std::collections::BTreeMap;

use serde::Serialize;
use serde_json::{Value, json};

/// Stable schema identifier emitted by `moff schema`.
pub const SCHEMA_VERSION: &str = "moff.schema/v1";

/// A machine-readable description of a root command.
#[derive(Debug, Clone, Serialize)]
pub struct CommandDefinition {
    pub summary: &'static str,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub source: Option<&'static str>,
}

const COMMANDS: [(&str, CommandDefinition); 4] = [
    (
        "auth",
        CommandDefinition {
            summary: "Authentication scaffolding (no provider calls)",
            source: None,
        },
    ),
    (
        "doctor",
        CommandDefinition {
            summary: "Local environment diagnostics scaffolding",
            source: None,
        },
    ),
    (
        "schema",
        CommandDefinition {
            summary: "Emit the machine-readable command schema",
            source: None,
        },
    ),
    (
        "outlook",
        CommandDefinition {
            summary: "Outlook command scaffolding (no provider calls)",
            source: Some("outlook"),
        },
    ),
];

/// Returns the root commands from the authoritative registry.
pub fn command_definitions() -> impl Iterator<Item = (&'static str, &'static CommandDefinition)> {
    COMMANDS
        .iter()
        .map(|(name, definition)| (*name, definition))
}

/// True when a provider source is known to the command registry.
#[must_use]
pub fn known_source(source: &str) -> bool {
    command_definitions().any(|(_, definition)| definition.source == Some(source))
}

/// Builds the stable, versioned command schema.
#[must_use]
pub fn schema(source: Option<&str>) -> Value {
    let commands: BTreeMap<_, _> = command_definitions()
        .filter(|(_, definition)| source.is_none_or(|selected| definition.source == Some(selected)))
        .collect();

    json!({
        "schema": SCHEMA_VERSION,
        "commands": commands,
    })
}

/// A machine-readable error classification.
#[derive(Debug, Clone, Copy, Serialize)]
pub enum ErrorCode {
    #[serde(rename = "incompatible_output_modes")]
    IncompatibleOutputModes,
    #[serde(rename = "unknown_source")]
    UnknownSource,
}

impl ErrorCode {
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::IncompatibleOutputModes => "incompatible_output_modes",
            Self::UnknownSource => "unknown_source",
        }
    }
}
