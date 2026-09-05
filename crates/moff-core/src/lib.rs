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

use base64::{Engine as _, engine::general_purpose::STANDARD};
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
            summary: "Microsoft Graph v1.0 Outlook Mail provider",
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
    fn get_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
    ) -> Result<Option<String>, DeviceFlowError>;
    fn set_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
        delta_link: &str,
    ) -> Result<(), DeviceFlowError>;
    fn delete_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
    ) -> Result<(), DeviceFlowError>;
}
pub struct KeyringVault;
impl KeyringVault {
    fn entry(profile: &str) -> Result<keyring::Entry, DeviceFlowError> {
        let name = ProfileStore::validate_profile_name(profile)?;
        keyring::Entry::new("moff", &format!("profile:{name}:refresh-token"))
            .map_err(|_| DeviceFlowError::Storage)
    }
    fn delta_entry(
        profile: &str,
        graph_base: &str,
        folder: &str,
    ) -> Result<keyring::Entry, DeviceFlowError> {
        let name = ProfileStore::validate_profile_name(profile)?;
        keyring::Entry::new(
            "moff",
            &format!("profile:{name}:delta-link:{graph_base}:{folder}"),
        )
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
    fn get_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
    ) -> Result<Option<String>, DeviceFlowError> {
        match Self::delta_entry(profile, graph_base, folder)?.get_password() {
            Ok(link) => Ok(Some(link)),
            Err(keyring::Error::NoEntry) => Ok(None),
            Err(_) => Err(DeviceFlowError::Storage),
        }
    }
    fn set_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
        delta_link: &str,
    ) -> Result<(), DeviceFlowError> {
        Self::delta_entry(profile, graph_base, folder)?
            .set_password(delta_link)
            .map_err(|_| DeviceFlowError::Storage)
    }
    fn delete_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
    ) -> Result<(), DeviceFlowError> {
        match Self::delta_entry(profile, graph_base, folder)?.delete_credential() {
            Ok(()) | Err(keyring::Error::NoEntry) => Ok(()),
            Err(_) => Err(DeviceFlowError::Storage),
        }
    }
}
#[derive(Default)]
pub struct InMemoryVault {
    tokens: Mutex<HashMap<String, String>>,
    delta_links: Mutex<HashMap<(String, String, String), String>>,
}
impl TokenVault for InMemoryVault {
    fn get(&self, profile: &str) -> Result<Option<String>, DeviceFlowError> {
        Ok(self
            .tokens
            .lock()
            .map_err(|_| DeviceFlowError::Storage)?
            .get(profile)
            .cloned())
    }
    fn set(&self, profile: &str, secret: &str) -> Result<(), DeviceFlowError> {
        self.tokens
            .lock()
            .map_err(|_| DeviceFlowError::Storage)?
            .insert(
                ProfileStore::validate_profile_name(profile)?,
                secret.to_owned(),
            );
        Ok(())
    }
    fn delete(&self, profile: &str) -> Result<(), DeviceFlowError> {
        self.tokens
            .lock()
            .map_err(|_| DeviceFlowError::Storage)?
            .remove(profile);
        Ok(())
    }
    fn get_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
    ) -> Result<Option<String>, DeviceFlowError> {
        Ok(self
            .delta_links
            .lock()
            .map_err(|_| DeviceFlowError::Storage)?
            .get(&(
                ProfileStore::validate_profile_name(profile)?,
                graph_base.to_owned(),
                folder.to_owned(),
            ))
            .cloned())
    }
    fn set_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
        delta_link: &str,
    ) -> Result<(), DeviceFlowError> {
        self.delta_links
            .lock()
            .map_err(|_| DeviceFlowError::Storage)?
            .insert(
                (
                    ProfileStore::validate_profile_name(profile)?,
                    graph_base.to_owned(),
                    folder.to_owned(),
                ),
                delta_link.to_owned(),
            );
        Ok(())
    }
    fn delete_delta_link(
        &self,
        profile: &str,
        graph_base: &str,
        folder: &str,
    ) -> Result<(), DeviceFlowError> {
        self.delta_links
            .lock()
            .map_err(|_| DeviceFlowError::Storage)?
            .remove(&(
                ProfileStore::validate_profile_name(profile)?,
                graph_base.to_owned(),
                folder.to_owned(),
            ));
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

/// Deterministic transport boundary for Microsoft Graph.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GraphRequest {
    pub method: String,
    pub url: String,
    pub headers: BTreeMap<String, String>,
    pub body: Option<String>,
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct GraphResponse {
    pub status: u16,
    pub body: String,
    pub request_id: Option<String>,
    pub retry_after_seconds: Option<u64>,
}
impl GraphResponse {
    #[must_use]
    pub fn json(status: u16, body: &str) -> Self {
        Self {
            status,
            body: body.to_owned(),
            request_id: None,
            retry_after_seconds: None,
        }
    }
    #[must_use]
    pub const fn with_retry_after(mut self, seconds: u64) -> Self {
        self.retry_after_seconds = Some(seconds);
        self
    }
}
pub trait GraphHttp {
    fn request(&self, request: GraphRequest) -> Result<GraphResponse, GraphError>;
}
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum GraphError {
    InvalidBase,
    UntrustedNextLink,
    Transport,
    OutcomeUnknown {
        request_id: Option<String>,
    },
    Protocol,
    UnsupportedAttachment {
        attachment_type: String,
    },
    Http {
        status: u16,
        request_id: Option<String>,
    },
}
impl fmt::Display for GraphError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::InvalidBase => f.write_str("invalid Graph base URL"),
            Self::UntrustedNextLink => f.write_str("Graph next link has an untrusted host"),
            Self::Transport => f.write_str("Graph request failed"),
            Self::OutcomeUnknown { request_id } => write!(
                f,
                "mutation outcome is unknown (request id: {})",
                request_id.as_deref().unwrap_or("unavailable")
            ),
            Self::Protocol => f.write_str("Graph returned an invalid response"),
            Self::UnsupportedAttachment { attachment_type } => {
                write!(f, "unsupported attachment type: {attachment_type}")
            }
            Self::Http { status, .. } => write!(f, "Graph returned HTTP {status}"),
        }
    }
}
impl std::error::Error for GraphError {}
pub struct GraphClient<'a> {
    http: &'a dyn GraphHttp,
    base: String,
    token: &'a str,
    sleeper: Option<&'a dyn Sleeper>,
}
impl<'a> GraphClient<'a> {
    pub fn new(http: &'a dyn GraphHttp, base: &str, token: &'a str) -> Result<Self, GraphError> {
        let base = base.trim_end_matches('/');
        if base != "https://graph.microsoft.com/v1.0" || base.contains('?') {
            return Err(GraphError::InvalidBase);
        }
        Ok(Self {
            http,
            base: base.to_owned(),
            token,
            sleeper: None,
        })
    }
    #[cfg(feature = "test-support")]
    pub fn new_for_test(
        http: &'a dyn GraphHttp,
        base: &str,
        token: &'a str,
    ) -> Result<Self, GraphError> {
        let base = base.trim_end_matches('/');
        let local_http = (base.starts_with("http://127.0.0.1:")
            || base.starts_with("http://localhost:"))
            && base.ends_with("/v1.0");
        if !local_http || base.contains('?') {
            return Err(GraphError::InvalidBase);
        }
        Ok(Self {
            http,
            base: base.to_owned(),
            token,
            sleeper: None,
        })
    }
    pub fn new_with_sleeper(
        http: &'a dyn GraphHttp,
        base: &str,
        token: &'a str,
        sleeper: &'a dyn Sleeper,
    ) -> Result<Self, GraphError> {
        let mut client = Self::new(http, base, token)?;
        client.sleeper = Some(sleeper);
        Ok(client)
    }
    fn wait(&self, seconds: u64) {
        if let Some(sleeper) = self.sleeper {
            sleeper.sleep_seconds(seconds);
        } else {
            ThreadSleeper.sleep_seconds(seconds);
        }
    }
    fn request(&self, method: &str, path: &str, body: Option<Value>) -> Result<Value, GraphError> {
        self.request_with_headers(method, path, body, BTreeMap::new())
    }
    fn request_with_headers(
        &self,
        method: &str,
        path: &str,
        body: Option<Value>,
        extra_headers: BTreeMap<String, String>,
    ) -> Result<Value, GraphError> {
        let url = if self.is_permitted_absolute_url(path) {
            path.to_owned()
        } else {
            format!(
                "{}{}",
                self.base,
                if path.starts_with('/') {
                    path.to_owned()
                } else {
                    format!("/{path}")
                }
            )
        };
        let mut headers = BTreeMap::from([
            ("Authorization".to_owned(), format!("Bearer {}", self.token)),
            ("Accept".to_owned(), "application/json".to_owned()),
        ]);
        if url.contains("/messages") || url.contains("/attachments") {
            headers.insert("Prefer".to_owned(), "IdType=\"ImmutableId\"".to_owned());
        }
        headers.extend(extra_headers);
        let encoded = body
            .map(|value| serde_json::to_string(&value).map_err(|_| GraphError::Protocol))
            .transpose()?;
        if encoded.is_some() {
            headers.insert("Content-Type".to_owned(), "application/json".to_owned());
        }
        let req = GraphRequest {
            method: method.to_owned(),
            url,
            headers,
            body: encoded,
        };
        let mut attempts = 0;
        loop {
            attempts += 1;
            match self.http.request(req.clone()) {
                Ok(response) if (200..300).contains(&response.status) => {
                    return if response.body.is_empty() {
                        Ok(json!({}))
                    } else {
                        serde_json::from_str(&response.body).map_err(|_| GraphError::Protocol)
                    };
                }
                Ok(response)
                    if method == "GET"
                        && attempts < 3
                        && matches!(response.status, 408 | 429 | 500 | 502 | 503 | 504) =>
                {
                    self.wait(response.retry_after_seconds.unwrap_or(1).min(60));
                }
                Ok(response) => {
                    if method != "GET" && matches!(response.status, 408 | 429 | 500..=599) {
                        return Err(GraphError::OutcomeUnknown {
                            request_id: response.request_id,
                        });
                    }
                    return Err(GraphError::Http {
                        status: response.status,
                        request_id: response.request_id,
                    });
                }
                Err(GraphError::Transport) if method == "GET" && attempts < 3 => {}
                Err(GraphError::Transport) if method != "GET" => {
                    return Err(GraphError::OutcomeUnknown { request_id: None });
                }
                Err(error) => return Err(error),
            }
        }
    }
    pub fn get_json(&self, path: &str) -> Result<Value, GraphError> {
        self.request("GET", path, None)
    }
    pub fn get_json_with_headers(
        &self,
        path: &str,
        headers: BTreeMap<String, String>,
    ) -> Result<Value, GraphError> {
        self.request_with_headers("GET", path, None, headers)
    }
    pub fn post_json(&self, path: &str, body: Value) -> Result<Value, GraphError> {
        self.request("POST", path, Some(body))
    }
    pub fn post_empty(&self, path: &str) -> Result<Value, GraphError> {
        let url = format!(
            "{}{}",
            self.base,
            if path.starts_with('/') {
                path.to_owned()
            } else {
                format!("/{path}")
            }
        );
        let mut headers = BTreeMap::from([
            ("Authorization".to_owned(), format!("Bearer {}", self.token)),
            ("Accept".to_owned(), "application/json".to_owned()),
            ("Content-Length".to_owned(), "0".to_owned()),
        ]);
        if url.contains("/messages/") {
            headers.insert("Prefer".to_owned(), "IdType=\"ImmutableId\"".to_owned());
        }
        match self.http.request(GraphRequest {
            method: "POST".to_owned(),
            url,
            headers,
            body: Some(String::new()),
        }) {
            Ok(response) if (200..300).contains(&response.status) => Ok(json!({})),
            Ok(response) if matches!(response.status, 408 | 429 | 500..=599) => {
                Err(GraphError::OutcomeUnknown {
                    request_id: response.request_id,
                })
            }
            Ok(response) => Err(GraphError::Http {
                status: response.status,
                request_id: response.request_id,
            }),
            Err(GraphError::Transport) => Err(GraphError::OutcomeUnknown { request_id: None }),
            Err(error) => Err(error),
        }
    }
    pub fn mutate_json(
        &self,
        method: &str,
        path: &str,
        body: Option<Value>,
    ) -> Result<Value, GraphError> {
        self.request(method, path, body)
    }
    fn is_permitted_absolute_url(&self, url: &str) -> bool {
        url.strip_prefix(&self.base)
            .is_some_and(|suffix| suffix.starts_with('/') || suffix.starts_with('?'))
    }
    pub fn validate_next_link(&self, next: &str) -> Result<(), GraphError> {
        if !self.is_permitted_absolute_url(next) {
            return Err(GraphError::UntrustedNextLink);
        }
        Ok(())
    }
    pub fn next_json(&self, next: &str) -> Result<Value, GraphError> {
        self.validate_next_link(next)?;
        self.get_json(next)
    }
}

/// Graph Mail routes with deterministic, injectable HTTP transport.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BodyPreference {
    Html,
    Text,
    None,
}
#[derive(Debug, Clone, Copy)]
pub struct GetMessageOptions {
    pub body: BodyPreference,
    pub include_headers: bool,
}
#[derive(Debug, Clone, Copy)]
pub struct ListMessagesOptions<'a> {
    pub limit: u16,
    pub unread: bool,
    pub from: Option<&'a str>,
    pub since: Option<&'a str>,
}
pub struct MailService<'a> {
    client: GraphClient<'a>,
}
impl<'a> MailService<'a> {
    #[must_use]
    pub fn new(client: GraphClient<'a>) -> Self {
        Self { client }
    }
    fn escape(value: &str) -> String {
        value
            .bytes()
            .map(|b| match b {
                b'A'..=b'Z' | b'a'..=b'z' | b'0'..=b'9' | b'-' | b'_' | b'.' | b'~' => {
                    char::from(b).to_string()
                }
                _ => format!("%{b:02X}"),
            })
            .collect()
    }
    fn messages(result: &Value) -> Value {
        json!({"content_origin":"external_email_untrusted","value":result["value"].clone(),"next_link":result["@odata.nextLink"].clone()})
    }
    fn graph_response(mut result: Value) -> Value {
        result["content_origin"] = json!("external_graph_untrusted");
        result
    }
    pub fn folders(&self) -> Result<Value, GraphError> {
        self.client
            .get_json("/me/mailFolders?$select=id,displayName,parentFolderId,totalItemCount,unreadItemCount")
            .map(Self::graph_response)
    }
    pub fn list(&self, folder: &str, limit: u16) -> Result<Value, GraphError> {
        Ok(Self::messages(&self.client.get_json(&format!("/me/mailFolders/{}/messages?$select=id%2Csubject%2Cfrom%2CreceivedDateTime%2CisRead&$top={limit}", Self::escape(folder)))?))
    }
    pub fn list_with_options(
        &self,
        folder: &str,
        options: ListMessagesOptions<'_>,
    ) -> Result<Value, GraphError> {
        let mut predicates = Vec::new();
        if options.unread {
            predicates.push("isRead eq false".to_owned());
        }
        if let Some(address) = options.from {
            predicates.push(format!("from/emailAddress/address eq '{address}'"));
        }
        if let Some(since) = options.since {
            predicates.push(format!("receivedDateTime ge {since}"));
        }
        let mut query = format!(
            "%24select=id%2Csubject%2Cfrom%2CreceivedDateTime%2CisRead&%24top={}",
            options.limit
        );
        if !predicates.is_empty() {
            query.push_str("&%24filter=");
            query.push_str(&Self::escape(&predicates.join(" and ")));
        }
        Ok(Self::messages(&self.client.get_json(&format!(
            "/me/mailFolders/{}/messages?{query}",
            Self::escape(folder)
        ))?))
    }
    pub fn search(&self, query: &str, limit: u16) -> Result<Value, GraphError> {
        Ok(Self::messages(&self.client.get_json(&format!("/me/messages?$search=%22{}%22&$select=id%2Csubject%2Cfrom%2CreceivedDateTime%2CisRead&$top={limit}", Self::escape(query)))?))
    }
    pub fn search_in_folder(
        &self,
        query: &str,
        folder: Option<&str>,
        limit: u16,
    ) -> Result<Value, GraphError> {
        let route = folder.map_or_else(
            || "/me/messages".to_owned(),
            |folder| format!("/me/mailFolders/{}/messages", Self::escape(folder)),
        );
        Ok(Self::messages(&self.client.get_json(&format!(
            "{route}?%24search=%22{}%22&%24select=id%2Csubject%2Cfrom%2CreceivedDateTime%2CisRead&%24top={limit}",
            Self::escape(query)
        ))?))
    }
    pub fn get(&self, id: &str) -> Result<Value, GraphError> {
        let mut result = self
            .client
            .get_json(&format!("/me/messages/{}", Self::escape(id)))?;
        result["content_origin"] = json!("external_email_untrusted");
        Ok(result)
    }
    pub fn get_with_options(
        &self,
        id: &str,
        options: GetMessageOptions,
    ) -> Result<Value, GraphError> {
        let mut select = "id,subject,from,toRecipients,receivedDateTime,isRead".to_owned();
        if options.body != BodyPreference::None {
            select.push_str(",body");
        }
        if options.include_headers {
            select.push_str(",internetMessageHeaders");
        }
        let query = format!(
            "%24select={}&%24expand=attachments%28%24select%3Did%2Cname%2Csize%2CcontentType%29",
            Self::escape(&select)
        );
        let prefer = match options.body {
            BodyPreference::Text => "IdType=\"ImmutableId\", outlook.body-content-type=\"text\"",
            BodyPreference::Html | BodyPreference::None => "IdType=\"ImmutableId\"",
        };
        let mut result = self.client.get_json_with_headers(
            &format!("/me/messages/{}?{query}", Self::escape(id)),
            BTreeMap::from([("Prefer".to_owned(), prefer.to_owned())]),
        )?;
        result["content_origin"] = json!("external_email_untrusted");
        Ok(result)
    }
    pub fn attachments(&self, id: &str) -> Result<Value, GraphError> {
        let result = self.client.get_json(&format!(
            "/me/messages/{}/attachments?$select=id,name,size,contentType,@odata.type",
            Self::escape(id)
        ))?;
        Ok(
            json!({"content_origin":"external_attachment_untrusted","value":result["value"].clone()}),
        )
    }
    pub fn download_attachment(
        &self,
        message_id: &str,
        attachment_id: &str,
    ) -> Result<Vec<u8>, GraphError> {
        let attachment = self.client.get_json(&format!(
            "/me/messages/{}/attachments/{}",
            Self::escape(message_id),
            Self::escape(attachment_id)
        ))?;
        let attachment_type = attachment["@odata.type"].as_str().unwrap_or_default();
        if attachment_type != "#microsoft.graph.fileAttachment" {
            return Err(GraphError::UnsupportedAttachment {
                attachment_type: attachment_type.to_owned(),
            });
        }
        let content = attachment["contentBytes"]
            .as_str()
            .ok_or(GraphError::Protocol)?;
        STANDARD.decode(content).map_err(|_| GraphError::Protocol)
    }
    pub fn mark(&self, id: &str, read: bool) -> Result<Value, GraphError> {
        self.client
            .mutate_json(
                "PATCH",
                &format!("/me/messages/{}", Self::escape(id)),
                Some(json!({"isRead":read})),
            )
            .map(Self::graph_response)
    }
    pub fn move_message(&self, id: &str, folder: &str) -> Result<Value, GraphError> {
        self.client
            .post_json(
                &format!("/me/messages/{}/move", Self::escape(id)),
                json!({"destinationId":folder}),
            )
            .map(Self::graph_response)
    }
    pub fn copy_message(&self, id: &str, folder: &str) -> Result<Value, GraphError> {
        self.client
            .post_json(
                &format!("/me/messages/{}/copy", Self::escape(id)),
                json!({"destinationId":folder}),
            )
            .map(Self::graph_response)
    }
    pub fn soft_delete(&self, id: &str) -> Result<Value, GraphError> {
        self.client
            .mutate_json(
                "DELETE",
                &format!("/me/messages/{}", Self::escape(id)),
                None,
            )
            .map(Self::graph_response)
    }
    pub fn draft_create(&self, message: Value) -> Result<Value, GraphError> {
        self.client
            .post_json("/me/messages", message)
            .map(Self::graph_response)
    }
    pub fn draft_update(&self, id: &str, message: Value) -> Result<Value, GraphError> {
        self.client
            .mutate_json(
                "PATCH",
                &format!("/me/messages/{}", Self::escape(id)),
                Some(message),
            )
            .map(Self::graph_response)
    }
    pub fn send(&self, message: &Value) -> Result<Value, GraphError> {
        self.client
            .post_json(
                "/me/sendMail",
                json!({"message":message,"saveToSentItems":true}),
            )
            .map(Self::graph_response)
    }
    pub fn send_draft(&self, id: &str) -> Result<Value, GraphError> {
        self.client
            .post_empty(&format!("/me/messages/{}/send", Self::escape(id)))
            .map(Self::graph_response)
    }
    pub fn reply(&self, id: &str, payload: Value, all: bool) -> Result<Value, GraphError> {
        self.client
            .post_json(
                &format!(
                    "/me/messages/{}/{}",
                    Self::escape(id),
                    if all { "replyAll" } else { "reply" }
                ),
                payload,
            )
            .map(Self::graph_response)
    }
    pub fn forward(&self, id: &str, payload: Value) -> Result<Value, GraphError> {
        self.client
            .post_json(
                &format!("/me/messages/{}/forward", Self::escape(id)),
                payload,
            )
            .map(Self::graph_response)
    }
    pub fn delta(&self, folder: &str, continuation: Option<&str>) -> Result<Value, GraphError> {
        self.delta_limited(folder, continuation, None)
    }
    pub fn delta_limited(
        &self,
        folder: &str,
        continuation: Option<&str>,
        limit: Option<u16>,
    ) -> Result<Value, GraphError> {
        let value = if let Some(link) = continuation {
            self.client.next_json(link)?
        } else {
            let route = format!("/me/mailFolders/{}/messages/delta", Self::escape(folder));
            let route = match limit {
                Some(limit) => format!("{route}?%24top={limit}"),
                None => route,
            };
            self.client.get_json(&route)?
        };
        Ok(
            json!({"content_origin":"external_email_untrusted","value":value["value"].clone(),"next_link":value["@odata.nextLink"].clone(),"delta_link":value["@odata.deltaLink"].clone()}),
        )
    }
    pub fn validate_delta_link(&self, link: &str) -> Result<(), GraphError> {
        self.client.validate_next_link(link)
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
