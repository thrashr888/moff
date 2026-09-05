//! Shared command contracts and secure delegated authentication for Moff.
#![allow(
    clippy::missing_errors_doc,
    reason = "public error conditions are represented by DeviceFlowError"
)]

use std::collections::{BTreeMap, HashMap};
use std::fmt;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

use directories::ProjectDirs;
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

pub const SCHEMA_VERSION: &str = "moff.schema/v1";

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
            summary: "Delegated Microsoft Entra authentication",
            source: None,
        },
    ),
    (
        "doctor",
        CommandDefinition {
            summary: "Local environment diagnostics",
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
pub fn command_definitions() -> impl Iterator<Item = (&'static str, &'static CommandDefinition)> {
    COMMANDS
        .iter()
        .map(|(name, definition)| (*name, definition))
}
#[must_use]
pub fn known_source(source: &str) -> bool {
    command_definitions().any(|(_, definition)| definition.source == Some(source))
}
#[must_use]
pub fn schema(source: Option<&str>) -> Value {
    let commands: BTreeMap<_, _> = command_definitions()
        .filter(|(_, definition)| source.is_none_or(|selected| definition.source == Some(selected)))
        .collect();
    json!({"schema": SCHEMA_VERSION,"commands": commands})
}

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

/// Non-secret profile configuration persisted in the platform data directory.
#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
pub struct Profile {
    pub name: String,
    pub client_id: String,
    pub tenant: String,
    pub scopes: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub account_display_name: Option<String>,
}

pub struct ProfileStore {
    root: PathBuf,
}
impl ProfileStore {
    #[must_use]
    pub fn platform_default() -> Self {
        let root = std::env::var_os("MOFF_DATA_DIR").map_or_else(
            || {
                ProjectDirs::from("com", "moff", "moff")
                    .map_or_else(|| PathBuf::from(".moff"), |dirs| dirs.data_dir().to_owned())
            },
            PathBuf::from,
        );
        Self { root }
    }
    #[must_use]
    pub fn at(root: impl Into<PathBuf>) -> Self {
        Self { root: root.into() }
    }
    pub fn validate_profile_name(name: &str) -> Result<String, DeviceFlowError> {
        let trimmed = name.trim();
        if trimmed.is_empty()
            || trimmed != name
            || !trimmed
                .chars()
                .all(|c| c.is_ascii_alphanumeric() || matches!(c, '-' | '_'))
        {
            return Err(DeviceFlowError::InvalidProfile);
        }
        Ok(trimmed.to_owned())
    }
    fn path(&self, name: &str) -> Result<PathBuf, DeviceFlowError> {
        Ok(self
            .root
            .join("profiles")
            .join(format!("{}.json", Self::validate_profile_name(name)?)))
    }
    pub fn save(&self, profile: &Profile) -> Result<(), DeviceFlowError> {
        let path = self.path(&profile.name)?;
        let parent = path.parent().ok_or(DeviceFlowError::Storage)?;
        fs::create_dir_all(parent).map_err(|_| DeviceFlowError::Storage)?;
        let data = serde_json::to_vec_pretty(profile).map_err(|_| DeviceFlowError::Storage)?;
        fs::write(path, data).map_err(|_| DeviceFlowError::Storage)
    }
    pub fn load(&self, name: &str) -> Result<Option<Profile>, DeviceFlowError> {
        let path = self.path(name)?;
        match fs::read(path) {
            Ok(data) => serde_json::from_slice(&data)
                .map(Some)
                .map_err(|_| DeviceFlowError::Storage),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(None),
            Err(_) => Err(DeviceFlowError::Storage),
        }
    }
    pub fn remove(&self, name: &str) -> Result<(), DeviceFlowError> {
        let path = self.path(name)?;
        match fs::remove_file(path) {
            Ok(()) => Ok(()),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
            Err(_) => Err(DeviceFlowError::Storage),
        }
    }
    #[must_use]
    pub fn root(&self) -> &Path {
        &self.root
    }
}

/// Token values deliberately cannot be serialized or debug printed.
pub struct AccessToken {
    value: String,
    expires_in: u64,
}
impl AccessToken {
    #[must_use]
    pub fn expires_in_seconds(&self) -> u64 {
        self.expires_in
    }
    #[must_use]
    pub fn expose_to_graph_only(&self) -> &str {
        &self.value
    }
}
impl fmt::Debug for AccessToken {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("AccessToken(REDACTED)")
    }
}

pub trait TokenVault {
    fn get(&self, profile: &str) -> Result<Option<String>, DeviceFlowError>;
    fn set(&self, profile: &str, refresh_token: &str) -> Result<(), DeviceFlowError>;
    fn delete(&self, profile: &str) -> Result<(), DeviceFlowError>;
}
pub struct KeyringVault;
impl KeyringVault {
    fn entry(profile: &str) -> Result<keyring::Entry, DeviceFlowError> {
        let name = ProfileStore::validate_profile_name(profile)?;
        keyring::Entry::new("moff", &format!("profile:{name}:refresh-token"))
            .map_err(|_| DeviceFlowError::Storage)
    }
}
impl TokenVault for KeyringVault {
    fn get(&self, profile: &str) -> Result<Option<String>, DeviceFlowError> {
        match Self::entry(profile)?.get_password() {
            Ok(secret) => Ok(Some(secret)),
            Err(keyring::Error::NoEntry) => Ok(None),
            Err(_) => Err(DeviceFlowError::Storage),
        }
    }
    fn set(&self, profile: &str, refresh_token: &str) -> Result<(), DeviceFlowError> {
        Self::entry(profile)?
            .set_password(refresh_token)
            .map_err(|_| DeviceFlowError::Storage)
    }
    fn delete(&self, profile: &str) -> Result<(), DeviceFlowError> {
        match Self::entry(profile)?.delete_credential() {
            Ok(()) | Err(keyring::Error::NoEntry) => Ok(()),
            Err(_) => Err(DeviceFlowError::Storage),
        }
    }
}
#[derive(Default)]
pub struct InMemoryVault(Mutex<HashMap<String, String>>);
impl TokenVault for InMemoryVault {
    fn get(&self, profile: &str) -> Result<Option<String>, DeviceFlowError> {
        Ok(self
            .0
            .lock()
            .map_err(|_| DeviceFlowError::Storage)?
            .get(profile)
            .cloned())
    }
    fn set(&self, profile: &str, secret: &str) -> Result<(), DeviceFlowError> {
        self.0.lock().map_err(|_| DeviceFlowError::Storage)?.insert(
            ProfileStore::validate_profile_name(profile)?,
            secret.to_owned(),
        );
        Ok(())
    }
    fn delete(&self, profile: &str) -> Result<(), DeviceFlowError> {
        self.0
            .lock()
            .map_err(|_| DeviceFlowError::Storage)?
            .remove(profile);
        Ok(())
    }
}

#[derive(Debug, Clone)]
pub struct HttpResponse {
    pub status: u16,
    pub body: String,
}
pub trait HttpClient {
    fn post_form(
        &self,
        url: &str,
        form: &BTreeMap<String, String>,
    ) -> Result<HttpResponse, DeviceFlowError>;
}
pub trait Sleeper {
    fn sleep_seconds(&self, seconds: u64);
}
pub struct ThreadSleeper;
impl Sleeper for ThreadSleeper {
    fn sleep_seconds(&self, seconds: u64) {
        std::thread::sleep(std::time::Duration::from_secs(seconds));
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DeviceFlowError {
    InvalidProfile,
    Protocol,
    Terminal(String),
    NoRefreshToken,
    Storage,
    Transport,
}
impl fmt::Display for DeviceFlowError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let text = match self {
            Self::InvalidProfile => "invalid profile name",
            Self::Protocol => "identity provider returned an invalid response",
            Self::Terminal(_) => "device authorization did not complete",
            Self::NoRefreshToken => "profile has no refresh token",
            Self::Storage => "secure storage operation failed",
            Self::Transport => "identity provider request failed",
        };
        formatter.write_str(text)
    }
}
impl std::error::Error for DeviceFlowError {}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VerificationInstructions {
    pub user_code: String,
    pub verification_uri: String,
    pub message: String,
}
pub struct AuthService<'a> {
    http: &'a dyn HttpClient,
    vault: &'a dyn TokenVault,
    sleeper: &'a dyn Sleeper,
    base_url: &'a str,
}
impl<'a> AuthService<'a> {
    #[must_use]
    pub fn new(
        http: &'a dyn HttpClient,
        vault: &'a dyn TokenVault,
        sleeper: &'a dyn Sleeper,
        base_url: &'a str,
    ) -> Self {
        Self {
            http,
            vault,
            sleeper,
            base_url,
        }
    }
    fn endpoint(&self, tenant: &str, action: &str) -> String {
        format!(
            "{}/{tenant}/oauth2/v2.0/{action}",
            self.base_url.trim_end_matches('/')
        )
    }
    fn form(entries: &[(&str, String)]) -> BTreeMap<String, String> {
        entries
            .iter()
            .map(|(key, value)| ((*key).to_owned(), value.clone()))
            .collect()
    }
    pub fn login(
        &self,
        profile: &str,
        client_id: &str,
        tenant: &str,
        scopes: &[&str],
    ) -> Result<VerificationInstructions, DeviceFlowError> {
        self.login_with_instructions(profile, client_id, tenant, scopes, |_| {})
    }

    /// Runs device authorization and presents safe instructions before polling begins.
    pub fn login_with_instructions<F: FnOnce(&VerificationInstructions)>(
        &self,
        profile: &str,
        client_id: &str,
        tenant: &str,
        scopes: &[&str],
        present: F,
    ) -> Result<VerificationInstructions, DeviceFlowError> {
        ProfileStore::validate_profile_name(profile)?;
        let scope = std::iter::once("offline_access")
            .chain(scopes.iter().copied())
            .collect::<Vec<_>>()
            .join(" ");
        let request = Self::form(&[("client_id", client_id.to_owned()), ("scope", scope)]);
        let device: Value = serde_json::from_str(
            &self
                .http
                .post_form(&self.endpoint(tenant, "devicecode"), &request)?
                .body,
        )
        .map_err(|_| DeviceFlowError::Protocol)?;
        let device_code = device["device_code"]
            .as_str()
            .ok_or(DeviceFlowError::Protocol)?
            .to_owned();
        let instructions = VerificationInstructions {
            user_code: device["user_code"]
                .as_str()
                .ok_or(DeviceFlowError::Protocol)?
                .to_owned(),
            verification_uri: device["verification_uri"]
                .as_str()
                .ok_or(DeviceFlowError::Protocol)?
                .to_owned(),
            message: device["message"]
                .as_str()
                .unwrap_or("Open the verification URL and enter the code.")
                .to_owned(),
        };
        present(&instructions);
        let mut interval = device["interval"].as_u64().unwrap_or(5);
        let expires = device["expires_in"]
            .as_u64()
            .ok_or(DeviceFlowError::Protocol)?;
        let mut elapsed = 0;
        loop {
            self.sleeper.sleep_seconds(interval);
            elapsed += interval;
            if elapsed > expires {
                return Err(DeviceFlowError::Terminal("expired_token".to_owned()));
            }
            let form = Self::form(&[
                ("client_id", client_id.to_owned()),
                (
                    "grant_type",
                    "urn:ietf:params:oauth:grant-type:device_code".to_owned(),
                ),
                ("device_code", device_code.clone()),
            ]);
            let response = self
                .http
                .post_form(&self.endpoint(tenant, "token"), &form)?;
            let body: Value =
                serde_json::from_str(&response.body).map_err(|_| DeviceFlowError::Protocol)?;
            if let Some(error) = body["error"].as_str() {
                match error {
                    "authorization_pending" => continue,
                    "slow_down" => {
                        interval += 5;
                        continue;
                    }
                    "authorization_declined" | "expired_token" => {
                        return Err(DeviceFlowError::Terminal(error.to_owned()));
                    }
                    _ => return Err(DeviceFlowError::Terminal(error.to_owned())),
                }
            }
            let refresh = body["refresh_token"]
                .as_str()
                .ok_or(DeviceFlowError::Protocol)?;
            self.vault.set(profile, refresh)?;
            return Ok(instructions);
        }
    }
    pub fn access_token(
        &self,
        profile: &str,
        client_id: &str,
        tenant: &str,
    ) -> Result<AccessToken, DeviceFlowError> {
        let refresh = self
            .vault
            .get(profile)?
            .ok_or(DeviceFlowError::NoRefreshToken)?;
        let form = Self::form(&[
            ("client_id", client_id.to_owned()),
            ("grant_type", "refresh_token".to_owned()),
            ("refresh_token", refresh),
        ]);
        let body: Value = serde_json::from_str(
            &self
                .http
                .post_form(&self.endpoint(tenant, "token"), &form)?
                .body,
        )
        .map_err(|_| DeviceFlowError::Protocol)?;
        let access = body["access_token"]
            .as_str()
            .ok_or(DeviceFlowError::Protocol)?
            .to_owned();
        if let Some(rotated) = body["refresh_token"].as_str() {
            self.vault.set(profile, rotated)?;
        }
        Ok(AccessToken {
            value: access,
            expires_in: body["expires_in"]
                .as_u64()
                .ok_or(DeviceFlowError::Protocol)?,
        })
    }
}
