//! Shared output contracts for Moff's local-only CLI.
use serde::Serialize;

pub const SCHEMA_VERSION: &str = "moff.schema/v1";

#[derive(Debug, Serialize)]
pub struct ErrorOutput<'a> {
    pub error: &'a str,
    pub message: &'a str,
}
