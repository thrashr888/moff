#![allow(
    clippy::enum_variant_names,
    clippy::needless_pass_by_value,
    reason = "CLI-owned values and Microsoft scope vocabulary are clearer at call sites"
)]

use std::collections::BTreeMap;
use std::fs::{self, OpenOptions};
use std::io::{self, Read, Write};
use std::path::Path;
use std::process::ExitCode;

use clap::{ArgAction, Parser, Subcommand, ValueEnum};
use moff_core::{
    AuthService, BodyPreference, DeviceFlowError, ErrorCode, GetMessageOptions, GraphClient,
    GraphError, GraphHttp, GraphRequest, GraphResponse, HttpClient, HttpResponse, KeyringVault,
    ListMessagesOptions, MailService, Profile, ProfileStore, ThreadSleeper, TokenVault,
    known_source, schema,
};
use moff_outlook_local::{
    OsaScriptOutlookUi, doctor as local_doctor, visible_mail as local_visible_mail,
};
use serde_json::{Value, json};
use time::{OffsetDateTime, format_description::well_known::Rfc3339};
use uuid::Uuid;

#[derive(Debug, Parser)]
#[allow(clippy::struct_excessive_bools, reason = "CLI flags are independent")]
#[command(name = "moff", version, about = "Agent-first Microsoft 365 CLI")]
struct Cli {
    #[arg(long, global = true)]
    pretty: bool,
    #[arg(long, global = true)]
    envelope: bool,
    #[arg(long, global = true)]
    json: bool,
    #[arg(long, global = true)]
    profile: Option<String>,
    #[arg(long, global = true)]
    dry_run: bool,
    #[arg(long, global = true)]
    yes: bool,
    #[arg(long, global = true, action = ArgAction::SetTrue)]
    no_color: bool,
    #[command(subcommand)]
    command: Option<Command>,
}
#[derive(Debug, Subcommand)]
enum Command {
    Auth {
        #[command(subcommand)]
        command: AuthCommand,
    },
    Doctor {
        #[arg(long)]
        network: bool,
    },
    Schema {
        #[arg(long)]
        source: Option<String>,
    },
    Outlook {
        #[arg(long, value_enum, default_value_t = Provider::Graph)]
        provider: Provider,
        #[command(subcommand)]
        command: OutlookCommand,
    },
}
#[derive(Clone, Debug, ValueEnum, PartialEq, Eq)]
enum Provider {
    Graph,
    #[value(name = "local-macos")]
    LocalMacos,
}
#[derive(Debug, Subcommand)]
enum OutlookCommand {
    Doctor,
    Folders {
        #[command(subcommand)]
        command: FoldersCommand,
    },
    Mail {
        #[command(subcommand)]
        command: MailCommand,
    },
    Attachments {
        #[command(subcommand)]
        command: AttachmentsCommand,
    },
}
#[derive(Debug, Subcommand)]
enum FoldersCommand {
    List,
}
#[derive(Debug, Subcommand)]
enum MailCommand {
    Visible {
        #[arg(long, default_value_t = 25)]
        limit: u16,
    },
    List {
        #[arg(long, default_value = "inbox")]
        folder: String,
        #[arg(long, default_value_t = 25)]
        limit: u16,
        #[arg(long)]
        unread: bool,
        #[arg(long)]
        from: Option<String>,
        #[arg(long)]
        since: Option<String>,
    },
    Search {
        query: String,
        #[arg(long)]
        folder: Option<String>,
        #[arg(long, default_value_t = 25)]
        limit: u16,
    },
    Get {
        #[arg(long)]
        id: String,
        #[arg(long, value_enum, default_value_t = GetBody::Html)]
        body: GetBody,
        #[arg(long)]
        include_headers: bool,
    },
    Attachments {
        #[command(subcommand)]
        command: AttachmentsCommand,
    },
    Mark {
        id: String,
        #[arg(long)]
        read: bool,
    },
    Move {
        id: String,
        folder: String,
    },
    Copy {
        id: String,
        folder: String,
    },
    Delete {
        id: String,
    },
    DraftCreate {
        #[arg(long)]
        to: Vec<String>,
        #[arg(long)]
        subject: Option<String>,
        #[arg(long)]
        body: Option<String>,
        #[arg(long)]
        body_file: Option<String>,
        #[arg(long)]
        body_stdin: bool,
        #[arg(long, value_enum, default_value_t = BodyType::Text)]
        body_type: BodyType,
    },
    DraftUpdate {
        #[arg(long)]
        id: String,
        #[arg(long)]
        subject: Option<String>,
        #[arg(long)]
        body: Option<String>,
        #[arg(long)]
        body_file: Option<String>,
        #[arg(long)]
        body_stdin: bool,
        #[arg(long, value_enum, default_value_t = BodyType::Text)]
        body_type: BodyType,
    },
    Send {
        #[arg(long)]
        to: Vec<String>,
        #[arg(long)]
        subject: Option<String>,
        #[arg(long)]
        body: Option<String>,
        #[arg(long)]
        body_file: Option<String>,
        #[arg(long)]
        body_stdin: bool,
        #[arg(long)]
        confirm_send: bool,
        #[arg(long, value_enum, default_value_t = BodyType::Text)]
        body_type: BodyType,
    },
    SendDraft {
        id: String,
        #[arg(long)]
        confirm_send: bool,
    },
    Reply {
        #[arg(long)]
        id: String,
        #[arg(long)]
        comment: Option<String>,
        #[arg(long)]
        body: Option<String>,
        #[arg(long)]
        body_file: Option<String>,
        #[arg(long)]
        body_stdin: bool,
        #[arg(long)]
        confirm_send: bool,
        #[arg(long, value_enum, default_value_t = BodyType::Text)]
        body_type: BodyType,
    },
    ReplyAll {
        id: String,
        #[arg(long)]
        comment: Option<String>,
        #[arg(long)]
        body: Option<String>,
        #[arg(long)]
        body_file: Option<String>,
        #[arg(long)]
        body_stdin: bool,
        #[arg(long)]
        confirm_send: bool,
        #[arg(long, value_enum, default_value_t = BodyType::Text)]
        body_type: BodyType,
    },
    Forward {
        #[arg(long)]
        id: String,
        #[arg(long)]
        to: Vec<String>,
        #[arg(long)]
        comment: Option<String>,
        #[arg(long)]
        body: Option<String>,
        #[arg(long)]
        body_file: Option<String>,
        #[arg(long)]
        body_stdin: bool,
        #[arg(long)]
        confirm_send: bool,
        #[arg(long, value_enum, default_value_t = BodyType::Text)]
        body_type: BodyType,
    },
    Sync {
        #[arg(long, default_value = "inbox")]
        folder: String,
        #[arg(long)]
        reset: bool,
        #[arg(long, default_value_t = 100)]
        limit: u16,
    },
}
#[derive(Clone, Debug, ValueEnum)]
enum BodyType {
    Text,
    Html,
}
#[allow(dead_code)]
impl BodyType {
    const fn graph(&self) -> &'static str {
        match self {
            Self::Text => "Text",
            Self::Html => "HTML",
        }
    }
}
#[derive(Clone, Debug, ValueEnum)]
enum GetBody {
    Html,
    Text,
    None,
}
impl GetBody {
    const fn preference(&self) -> BodyPreference {
        match self {
            Self::Html => BodyPreference::Html,
            Self::Text => BodyPreference::Text,
            Self::None => BodyPreference::None,
        }
    }
}
#[derive(Debug, Subcommand)]
enum AttachmentsCommand {
    List {
        #[arg(long)]
        id: String,
    },
    Download {
        #[arg(long)]
        id: String,
        #[arg(long)]
        attachment_id: String,
        #[arg(long)]
        output: String,
        #[arg(long)]
        force: bool,
    },
}
#[derive(Debug, Subcommand)]
enum AuthCommand {
    Login {
        #[arg(long)]
        client_id: String,
        #[arg(long, default_value = "common")]
        tenant: String,
        #[arg(long, value_enum)]
        scope: Vec<Scope>,
    },
    Status,
    Logout,
}
#[derive(Clone, Debug, ValueEnum)]
enum Scope {
    #[value(name = "mail.read")]
    MailRead,
    #[value(name = "mail.write")]
    MailWrite,
    #[value(name = "mail.send")]
    MailSend,
}
impl Scope {
    const fn as_entra(&self) -> &'static str {
        match self {
            Self::MailRead => "Mail.Read",
            Self::MailWrite => "Mail.ReadWrite",
            Self::MailSend => "Mail.Send",
        }
    }
}

struct ReqwestHttp(reqwest::blocking::Client);
impl HttpClient for ReqwestHttp {
    fn post_form(
        &self,
        url: &str,
        form: &BTreeMap<String, String>,
    ) -> Result<HttpResponse, DeviceFlowError> {
        let response = self
            .0
            .post(url)
            .form(form)
            .send()
            .map_err(|_| DeviceFlowError::Transport)?;
        let status = response.status().as_u16();
        let body = response.text().map_err(|_| DeviceFlowError::Transport)?;
        Ok(HttpResponse { status, body })
    }
}

#[allow(dead_code)]
struct ReqwestGraph(reqwest::blocking::Client);
impl GraphHttp for ReqwestGraph {
    fn request(&self, request: GraphRequest) -> Result<GraphResponse, GraphError> {
        let method = reqwest::Method::from_bytes(request.method.as_bytes())
            .map_err(|_| GraphError::Protocol)?;
        let mut call = self.0.request(method, &request.url);
        for (name, value) in request.headers {
            call = call.header(name, value);
        }
        if let Some(body) = request.body {
            call = call.body(body);
        }
        let response = call.send().map_err(|_| GraphError::Transport)?;
        let status = response.status().as_u16();
        let request_id = response
            .headers()
            .get("request-id")
            .and_then(|v| v.to_str().ok())
            .map(str::to_owned);
        let retry_after_seconds = response
            .headers()
            .get("retry-after")
            .and_then(|value| value.to_str().ok())
            .and_then(|value| value.parse().ok());
        let body = response.text().map_err(|_| GraphError::Transport)?;
        Ok(GraphResponse {
            status,
            body,
            request_id,
            retry_after_seconds,
        })
    }
}

struct AppError {
    code: &'static str,
    message: String,
}
impl AppError {
    fn new(code: &'static str, message: impl Into<String>) -> Self {
        Self {
            code,
            message: message.into(),
        }
    }
}
fn profile_name(value: Option<String>) -> Result<String, AppError> {
    ProfileStore::validate_profile_name(value.as_deref().unwrap_or("default"))
        .map_err(|error| AppError::new("invalid_profile", error.to_string()))
}
fn valid_tenant(value: &str) -> bool {
    matches!(value, "common" | "consumers" | "organizations") || Uuid::parse_str(value).is_ok()
}
fn valid_email(value: &str) -> bool {
    let Some((local, domain)) = value.split_once('@') else {
        return false;
    };
    !local.is_empty()
        && !domain.is_empty()
        && !value.contains(char::is_whitespace)
        && !domain.starts_with('.')
        && !domain.ends_with('.')
        && domain.contains('.')
        && !domain.contains("..")
        && value.matches('@').count() == 1
}
fn valid_rfc3339(value: &str) -> bool {
    OffsetDateTime::parse(value, &Rfc3339).is_ok()
}
const MAX_BODY_BYTES: u64 = 1_048_576;

fn read_utf8_body(reader: &mut dyn Read) -> Result<String, AppError> {
    let mut bytes = Vec::new();
    reader
        .take(MAX_BODY_BYTES + 1)
        .read_to_end(&mut bytes)
        .map_err(|_| AppError::new("body_read", "unable to read body source"))?;
    if bytes.len() as u64 > MAX_BODY_BYTES {
        return Err(AppError::new("body_too_large", "body source exceeds 1 MiB"));
    }
    String::from_utf8(bytes)
        .map_err(|_| AppError::new("invalid_body_utf8", "body source must be valid UTF-8"))
}

fn resolve_body_source(
    body: &mut Option<String>,
    body_file: Option<&String>,
    body_stdin: bool,
    required: bool,
) -> Result<(), AppError> {
    if body
        .as_ref()
        .is_some_and(|value| value.len() as u64 > MAX_BODY_BYTES)
    {
        return Err(AppError::new("body_too_large", "body source exceeds 1 MiB"));
    }
    let sources =
        usize::from(body.is_some()) + usize::from(body_file.is_some()) + usize::from(body_stdin);
    if sources == 0 && required {
        return Err(AppError::new(
            "body_source_required",
            "exactly one body source is required",
        ));
    }
    if sources > 1 {
        return Err(AppError::new(
            "body_source_conflict",
            "only one body source may be used",
        ));
    }
    if let Some(path) = body_file {
        #[cfg(unix)]
        let mut file = {
            use std::os::unix::fs::OpenOptionsExt;
            let file = OpenOptions::new()
                .read(true)
                .custom_flags(libc::O_NOFOLLOW)
                .open(path)
                .map_err(|_| {
                    AppError::new("body_file", "body file must be a readable regular file")
                })?;
            if !file.metadata().is_ok_and(|metadata| metadata.is_file()) {
                return Err(AppError::new(
                    "body_file",
                    "body file must be a readable regular file",
                ));
            }
            file
        };
        #[cfg(not(unix))]
        let mut file = {
            let metadata = fs::symlink_metadata(path).map_err(|_| {
                AppError::new("body_file", "body file must be a readable regular file")
            })?;
            if !metadata.file_type().is_file() {
                return Err(AppError::new(
                    "body_file",
                    "body file must be a readable regular file",
                ));
            }
            fs::File::open(path).map_err(|_| {
                AppError::new("body_file", "body file must be a readable regular file")
            })?
        };
        *body = Some(read_utf8_body(&mut file)?);
    } else if body_stdin {
        *body = Some(read_utf8_body(&mut io::stdin().lock())?);
    }
    Ok(())
}

#[allow(
    clippy::unnested_or_patterns,
    reason = "shared compose source validation"
)]
fn resolve_compose_body(command: &mut OutlookCommand) -> Result<(), AppError> {
    match command {
        OutlookCommand::Mail {
            command:
                MailCommand::DraftCreate {
                    body,
                    body_file,
                    body_stdin,
                    ..
                },
        }
        | OutlookCommand::Mail {
            command:
                MailCommand::Send {
                    body,
                    body_file,
                    body_stdin,
                    ..
                },
        } => resolve_body_source(body, body_file.as_ref(), *body_stdin, true),
        OutlookCommand::Mail {
            command:
                MailCommand::DraftUpdate {
                    body,
                    body_file,
                    body_stdin,
                    ..
                },
        } => resolve_body_source(body, body_file.as_ref(), *body_stdin, false),
        OutlookCommand::Mail {
            command:
                MailCommand::Reply {
                    comment,
                    body,
                    body_file,
                    body_stdin,
                    ..
                },
        }
        | OutlookCommand::Mail {
            command:
                MailCommand::ReplyAll {
                    comment,
                    body,
                    body_file,
                    body_stdin,
                    ..
                },
        }
        | OutlookCommand::Mail {
            command:
                MailCommand::Forward {
                    comment,
                    body,
                    body_file,
                    body_stdin,
                    ..
                },
        } => {
            if comment.is_some() && (body.is_some() || body_file.is_some() || *body_stdin) {
                return Err(AppError::new(
                    "invalid_reply",
                    "reply requires exactly one of --comment or a body source",
                ));
            }
            resolve_body_source(body, body_file.as_ref(), *body_stdin, comment.is_none())
        }
        _ => Ok(()),
    }
}

fn validate_outlook_command(command: &OutlookCommand) -> Result<(), AppError> {
    match command {
        OutlookCommand::Mail {
            command: MailCommand::List { from, since, .. },
        } => {
            if from.as_deref().is_some_and(|value| !valid_email(value)) {
                return Err(AppError::new(
                    "invalid_email",
                    "--from must be a valid email address",
                ));
            }
            if since.as_deref().is_some_and(|value| !valid_rfc3339(value)) {
                return Err(AppError::new("invalid_rfc3339", "--since must be RFC3339"));
            }
        }
        OutlookCommand::Mail {
            command: MailCommand::Send { to, .. } | MailCommand::Forward { to, .. },
        } if to.is_empty() => {
            return Err(AppError::new(
                "recipient_required",
                "at least one recipient is required",
            ));
        }
        OutlookCommand::Mail {
            command:
                MailCommand::Send { to, .. }
                | MailCommand::Forward { to, .. }
                | MailCommand::DraftCreate { to, .. },
        } if to.iter().any(|value| !valid_email(value)) => {
            return Err(AppError::new(
                "invalid_email",
                "recipient must be a valid email address",
            ));
        }
        OutlookCommand::Mail {
            command: MailCommand::DraftUpdate { subject, body, .. },
        } if subject.is_none() && body.is_none() => {
            return Err(AppError::new(
                "draft_update_required",
                "draft update requires --subject or a body source",
            ));
        }
        _ => {}
    }
    Ok(())
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

fn write_attachment(path: &Path, bytes: &[u8], force: bool) -> Result<(), AppError> {
    if path.exists() && !force {
        return Err(AppError::new(
            "output_exists",
            "output exists; use --force with --yes",
        ));
    }
    let parent = path.parent().unwrap_or_else(|| Path::new("."));
    let file_name = path
        .file_name()
        .ok_or_else(|| AppError::new("output", "output must name a file"))?;
    let nonce = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_err(|_| AppError::new("output", "system clock is unavailable"))?
        .as_nanos();
    let temporary = parent.join(format!(
        ".{}.{}.{}.tmp",
        file_name.to_string_lossy(),
        std::process::id(),
        nonce
    ));
    let mut options = OpenOptions::new();
    options.create_new(true).write(true);
    #[cfg(unix)]
    {
        use std::os::unix::fs::OpenOptionsExt;
        options.mode(0o600);
    }
    let mut file = options
        .open(&temporary)
        .map_err(|_| AppError::new("output", "unable to create temporary attachment file"))?;
    if file
        .write_all(bytes)
        .and_then(|()| file.sync_all())
        .is_err()
    {
        let _ = fs::remove_file(&temporary);
        return Err(AppError::new("output", "unable to write attachment file"));
    }
    drop(file);
    #[cfg(unix)]
    if !force {
        return fs::hard_link(&temporary, path)
            .and_then(|()| fs::remove_file(&temporary))
            .map_err(|error| {
                let _ = fs::remove_file(&temporary);
                if error.kind() == io::ErrorKind::AlreadyExists {
                    AppError::new("output_exists", "output exists; use --force with --yes")
                } else {
                    AppError::new("output", "unable to finalize attachment file")
                }
            });
    }
    fs::rename(&temporary, path).map_err(|_| {
        let _ = fs::remove_file(&temporary);
        AppError::new("output", "unable to finalize attachment file")
    })
}

fn sync_delta_cycle(
    mail: &MailService<'_>,
    vault: &dyn TokenVault,
    profile_name: &str,
    graph_base: &str,
    folder: &str,
    limit: u16,
) -> Result<Value, AppError> {
    let mut continuation = vault
        .get_delta_link(profile_name, graph_base, folder)
        .map_err(storage_error)?;
    let mut records = Vec::new();
    loop {
        let delta = mail
            .delta_limited(folder, continuation.as_deref(), Some(limit))
            .map_err(graph_error)?;
        let values = delta["value"]
            .as_array()
            .ok_or_else(|| AppError::new("graph", "Graph returned an invalid delta page"))?;
        records.extend(values.iter().cloned());
        if let Some(next) = delta["next_link"].as_str() {
            continuation = Some(next.to_owned());
            continue;
        }
        let terminal = delta["delta_link"].as_str().ok_or_else(|| {
            AppError::new("graph", "Graph delta response omitted terminal cursor")
        })?;
        mail.validate_delta_link(terminal).map_err(graph_error)?;
        vault
            .set_delta_link(profile_name, graph_base, folder, terminal)
            .map_err(storage_error)?;
        return Ok(
            json!({"content_origin":"external_email_untrusted","value":records,"next_link":Value::Null,"delta_link":terminal}),
        );
    }
}

#[cfg(feature = "test-support")]
fn validate_test_endpoint(endpoint: &str, graph: bool) -> Result<String, AppError> {
    let parsed = reqwest::Url::parse(endpoint)
        .map_err(|_| AppError::new("test_endpoint", "test endpoint must be loopback"))?;
    let loopback = matches!(parsed.host_str(), Some("127.0.0.1" | "localhost"));
    if parsed.scheme() != "http"
        || !loopback
        || parsed.port().is_none_or(|port| port == 0)
        || !parsed.username().is_empty()
        || parsed.password().is_some()
        || parsed.query().is_some()
        || parsed.fragment().is_some()
        || (graph && parsed.path() != "/v1.0")
    {
        return Err(AppError::new(
            "test_endpoint",
            "test endpoint must be loopback",
        ));
    }
    Ok(endpoint.to_owned())
}
#[cfg(feature = "test-support")]
fn test_endpoint(name: &str, graph: bool) -> Result<String, AppError> {
    let endpoint = std::env::var(name)
        .map_err(|_| AppError::new("test_endpoint", "test endpoint is required"))?;
    validate_test_endpoint(&endpoint, graph)
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "feature-gated test endpoints can return validation errors"
)]
fn entra_base(test_endpoints: bool) -> Result<String, AppError> {
    #[cfg(feature = "test-support")]
    if test_endpoints {
        return test_endpoint("MOFF_ENTRA_BASE_URL", false);
    }
    let _ = test_endpoints;
    Ok("https://login.microsoftonline.com".to_owned())
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "feature-gated test endpoints can return validation errors"
)]
fn graph_base(test_endpoints: bool) -> Result<String, AppError> {
    #[cfg(feature = "test-support")]
    if test_endpoints {
        return test_endpoint("MOFF_GRAPH_BASE_URL", true);
    }
    let _ = test_endpoints;
    Ok("https://graph.microsoft.com/v1.0".to_owned())
}

fn auth_data(
    command: AuthCommand,
    profile_name: String,
    dry_run: bool,
    test_endpoints: bool,
) -> Result<Value, AppError> {
    let store = ProfileStore::platform_default();
    let vault = KeyringVault;
    match command {
        AuthCommand::Status => {
            let profile = store.load(&profile_name).map_err(storage_error)?;
            let token_present = vault.get(&profile_name).map_err(storage_error)?.is_some();
            Ok(
                json!({"command":"auth.status","profile":profile_name,"configured":profile.is_some(),"credentials_present":token_present,"expires_at":"unknown"}),
            )
        }
        AuthCommand::Logout => {
            vault.delete(&profile_name).map_err(storage_error)?;
            store.remove(&profile_name).map_err(storage_error)?;
            Ok(json!({"command":"auth.logout","profile":profile_name,"logged_out":true}))
        }
        AuthCommand::Login {
            client_id,
            tenant,
            scope,
        } => {
            Uuid::parse_str(&client_id)
                .map_err(|_| AppError::new("invalid_client_id", "invalid UUID client ID"))?;
            if !valid_tenant(&tenant) {
                return Err(AppError::new(
                    "invalid_tenant",
                    "tenant must be common, consumers, organizations, or a tenant UUID",
                ));
            }
            let scopes: Vec<&str> = scope.iter().map(Scope::as_entra).collect();
            if dry_run {
                return Ok(
                    json!({"command":"auth.login","profile":profile_name,"plan":"device_authorization","tenant":tenant,"scopes":scopes}),
                );
            }
            let http = ReqwestHttp(
                reqwest::blocking::Client::builder()
                    .build()
                    .map_err(|_| AppError::new("transport", "unable to create HTTPS client"))?,
            );
            let sleeper = ThreadSleeper;
            let base = entra_base(test_endpoints)?;
            let service = AuthService::new(&http, &vault, &sleeper, &base);
            service
                .login_with_instructions(
                    &profile_name,
                    &client_id,
                    &tenant,
                    &scopes,
                    |instructions| {
                        eprintln!(
                            "{}\nVisit {} and enter code {}",
                            instructions.message,
                            instructions.verification_uri,
                            instructions.user_code
                        );
                    },
                )
                .map_err(auth_error)?;
            store
                .save(&Profile {
                    name: profile_name.clone(),
                    client_id,
                    tenant,
                    scopes: scopes.into_iter().map(str::to_owned).collect(),
                    account_display_name: None,
                })
                .map_err(storage_error)?;
            Ok(json!({"command":"auth.login","profile":profile_name,"authenticated":true}))
        }
    }
}
fn storage_error(error: DeviceFlowError) -> AppError {
    AppError::new("storage", error.to_string())
}
fn auth_error(error: DeviceFlowError) -> AppError {
    AppError::new("authentication", error.to_string())
}
#[allow(
    clippy::too_many_lines,
    reason = "command policy is deliberately co-located"
)]
fn outlook_data(
    mut command: OutlookCommand,
    dry_run: bool,
    yes: bool,
    profile_name: &str,
    vault: &dyn TokenVault,
    test_endpoints: bool,
) -> Result<Value, AppError> {
    resolve_compose_body(&mut command)?;
    validate_outlook_command(&command)?;
    let (name, _mutation) = match &command {
        OutlookCommand::Doctor => ("outlook.doctor", false),
        OutlookCommand::Mail {
            command: MailCommand::Visible { .. },
        } => ("outlook.mail.visible", false),
        OutlookCommand::Folders {
            command: FoldersCommand::List,
        } => ("outlook.folders.list", false),
        OutlookCommand::Attachments {
            command: AttachmentsCommand::List { .. },
        }
        | OutlookCommand::Mail {
            command:
                MailCommand::Attachments {
                    command: AttachmentsCommand::List { .. },
                },
        } => ("outlook.attachments.list", false),
        OutlookCommand::Attachments {
            command: AttachmentsCommand::Download { force, .. },
        }
        | OutlookCommand::Mail {
            command:
                MailCommand::Attachments {
                    command: AttachmentsCommand::Download { force, .. },
                },
        } => {
            if *force && !yes {
                return Err(AppError::new(
                    "confirmation_required",
                    "attachment overwrite requires --yes",
                ));
            }
            ("outlook.attachments.download", false)
        }
        OutlookCommand::Mail {
            command: MailCommand::List { limit, .. } | MailCommand::Search { limit, .. },
        } => {
            if *limit > 1000 {
                return Err(AppError::new(
                    "invalid_limit",
                    "result limit must be <= 1000",
                ));
            }
            (
                if matches!(
                    &command,
                    OutlookCommand::Mail {
                        command: MailCommand::Search { .. }
                    }
                ) {
                    "outlook.mail.search"
                } else {
                    "outlook.mail.list"
                },
                false,
            )
        }
        OutlookCommand::Mail {
            command: MailCommand::Get { .. },
        } => ("outlook.mail.get", false),
        OutlookCommand::Mail {
            command: MailCommand::Sync { limit, reset, .. },
        } => {
            if *limit > 1000 {
                return Err(AppError::new(
                    "invalid_limit",
                    "result limit must be <= 1000",
                ));
            }
            if *reset && !yes {
                return Err(AppError::new(
                    "confirmation_required",
                    "sync reset requires --yes",
                ));
            }
            ("outlook.mail.sync", false)
        }
        OutlookCommand::Mail {
            command: MailCommand::Delete { .. },
        } => {
            if !yes {
                return Err(AppError::new(
                    "confirmation_required",
                    "soft delete requires --yes",
                ));
            }
            ("outlook.mail.delete", true)
        }
        OutlookCommand::Mail {
            command:
                MailCommand::Send { confirm_send, .. }
                | MailCommand::SendDraft { confirm_send, .. }
                | MailCommand::Reply { confirm_send, .. }
                | MailCommand::ReplyAll { confirm_send, .. }
                | MailCommand::Forward { confirm_send, .. },
        } => {
            if !confirm_send {
                return Err(AppError::new(
                    "confirmation_required",
                    "sending mail requires --confirm-send",
                ));
            }
            let name = match &command {
                OutlookCommand::Mail {
                    command: MailCommand::Send { .. },
                } => "outlook.mail.send",
                OutlookCommand::Mail {
                    command: MailCommand::SendDraft { .. },
                } => "outlook.mail.send-draft",
                OutlookCommand::Mail {
                    command: MailCommand::Reply { .. },
                } => "outlook.mail.reply",
                OutlookCommand::Mail {
                    command: MailCommand::ReplyAll { .. },
                } => "outlook.mail.reply-all",
                _ => "outlook.mail.forward",
            };
            (name, true)
        }
        OutlookCommand::Mail {
            command: MailCommand::Mark { .. },
        } => ("outlook.mail.mark", true),
        OutlookCommand::Mail {
            command: MailCommand::Move { .. },
        } => ("outlook.mail.move", true),
        OutlookCommand::Mail {
            command: MailCommand::Copy { .. },
        } => ("outlook.mail.copy", true),
        OutlookCommand::Mail {
            command: MailCommand::DraftCreate { .. },
        } => ("outlook.mail.draft-create", true),
        OutlookCommand::Mail {
            command: MailCommand::DraftUpdate { .. },
        } => ("outlook.mail.draft-update", true),
    };
    if dry_run {
        return Ok(json!({"command":name,"dry_run":true,"network_mutation":false}));
    }
    execute_outlook(command, profile_name, vault, test_endpoints)
}

#[allow(
    clippy::too_many_lines,
    reason = "command execution is deliberately co-located with CLI dispatch"
)]
fn execute_outlook(
    command: OutlookCommand,
    profile_name: &str,
    vault: &dyn TokenVault,
    test_endpoints: bool,
) -> Result<Value, AppError> {
    let store = ProfileStore::platform_default();
    let profile = store
        .load(profile_name)
        .map_err(storage_error)?
        .ok_or_else(|| AppError::new("authentication", "no credentials: run auth login first"))?;
    let required = match &command {
        OutlookCommand::Doctor
        | OutlookCommand::Folders { .. }
        | OutlookCommand::Attachments { .. }
        | OutlookCommand::Mail {
            command:
                MailCommand::List { .. }
                | MailCommand::Search { .. }
                | MailCommand::Get { .. }
                | MailCommand::Sync { .. },
        } => "Mail.Read",
        OutlookCommand::Mail {
            command:
                MailCommand::Send { .. }
                | MailCommand::SendDraft { .. }
                | MailCommand::Reply { .. }
                | MailCommand::ReplyAll { .. }
                | MailCommand::Forward { .. },
        } => "Mail.Send",
        OutlookCommand::Mail { .. } => "Mail.ReadWrite",
    };
    if !profile
        .scopes
        .iter()
        .any(|scope| scope == required || (required == "Mail.Read" && scope == "Mail.ReadWrite"))
    {
        return Err(AppError::new(
            "consent_required",
            format!("typed consent requirement: {required}"),
        ));
    }
    let http = ReqwestHttp(
        reqwest::blocking::Client::builder()
            .build()
            .map_err(|_| AppError::new("transport", "unable to create HTTPS client"))?,
    );
    let sleeper = ThreadSleeper;
    let entra = entra_base(test_endpoints)?;
    let token = AuthService::new(&http, vault, &sleeper, &entra)
        .access_token(profile_name, &profile.client_id, &profile.tenant)
        .map_err(auth_error)?;
    let graph_http = ReqwestGraph(
        reqwest::blocking::Client::builder()
            .build()
            .map_err(|_| AppError::new("transport", "unable to create HTTPS client"))?,
    );
    let graph_base = graph_base(test_endpoints)?;
    let graph = {
        #[cfg(feature = "test-support")]
        if test_endpoints {
            GraphClient::new_for_test(&graph_http, &graph_base, token.expose_to_graph_only())
        } else {
            GraphClient::new(&graph_http, &graph_base, token.expose_to_graph_only())
        }
        #[cfg(not(feature = "test-support"))]
        {
            let _ = test_endpoints;
            GraphClient::new(&graph_http, &graph_base, token.expose_to_graph_only())
        }
    }
    .map_err(graph_error)?;
    let mail = MailService::new(graph);
    let message = |to: &[String], subject: &Option<String>, body: &str, body_type: &BodyType| json!({"subject":subject,"body":{"contentType":body_type.graph(),"content":body},"toRecipients":to.iter().map(|address| json!({"emailAddress":{"address":address}})).collect::<Vec<_>>()});
    let result = match command {
        OutlookCommand::Doctor
        | OutlookCommand::Mail {
            command: MailCommand::Visible { .. },
        } => unreachable!("local-only command is dispatched before Graph"),
        OutlookCommand::Folders { .. } => mail.folders(),
        OutlookCommand::Attachments { command: AttachmentsCommand::List { id } }
        | OutlookCommand::Mail { command: MailCommand::Attachments { command: AttachmentsCommand::List { id } } } => mail.attachments(&id),
        OutlookCommand::Attachments { command: AttachmentsCommand::Download { id, attachment_id, output, force } }
        | OutlookCommand::Mail { command: MailCommand::Attachments { command: AttachmentsCommand::Download { id, attachment_id, output, force } } } => {
            let bytes = mail.download_attachment(&id, &attachment_id).map_err(graph_error)?;
            write_attachment(Path::new(&output), &bytes, force)?;
            return Ok(json!({"content_origin":"external_attachment_untrusted","output":output,"bytes":bytes.len()}));
        }
        OutlookCommand::Mail { command: MailCommand::List { folder, limit, unread, from, since } } => mail.list_with_options(&folder, ListMessagesOptions { limit, unread, from: from.as_deref(), since: since.as_deref() }),
        OutlookCommand::Mail { command: MailCommand::Search { query, folder, limit } } => mail.search_in_folder(&query, folder.as_deref(), limit),
        OutlookCommand::Mail { command: MailCommand::Get { id, body, include_headers } } => mail.get_with_options(&id, GetMessageOptions { body: body.preference(), include_headers }),
        OutlookCommand::Mail { command: MailCommand::Mark { id, read } } => mail.mark(&id, read),
        OutlookCommand::Mail { command: MailCommand::Move { id, folder } } => mail.move_message(&id, &folder),
        OutlookCommand::Mail { command: MailCommand::Copy { id, folder } } => mail.copy_message(&id, &folder),
        OutlookCommand::Mail { command: MailCommand::Delete { id } } => mail.soft_delete(&id),
        OutlookCommand::Mail { command: MailCommand::DraftCreate { to, subject, body, body_type, .. } } => mail.draft_create(message(&to, &subject, body.as_deref().expect("validated body"), &body_type)),
        OutlookCommand::Mail { command: MailCommand::DraftUpdate { id, subject, body, body_type, .. } } => {
            let mut update = serde_json::Map::new();
            if let Some(subject) = subject {
                update.insert("subject".to_owned(), json!(subject));
            }
            if let Some(content) = body {
                update.insert("body".to_owned(), json!({"contentType":body_type.graph(),"content":content}));
            }
            mail.draft_update(
                &id,
                Value::Object(update),
            )
        }
        OutlookCommand::Mail { command: MailCommand::Send { to, subject, body, body_type, .. } } => mail.send(&message(&to, &subject, body.as_deref().expect("validated body"), &body_type)),
        OutlookCommand::Mail { command: MailCommand::SendDraft { id, .. } } => mail.send_draft(&id).map(|_| json!({"accepted":true})),
        OutlookCommand::Mail { command: MailCommand::Reply { id, comment, body, body_type, .. } } => mail.reply(&id, comment.map_or_else(|| json!({"message":{"body":{"contentType":body_type.graph(),"content":body.expect("validated body")}}}), |comment| json!({"comment":comment})), false).map(|_| json!({"accepted":true})),
        OutlookCommand::Mail { command: MailCommand::ReplyAll { id, comment, body, body_type, .. } } => mail.reply(&id, comment.map_or_else(|| json!({"message":{"body":{"contentType":body_type.graph(),"content":body.expect("validated body")}}}), |comment| json!({"comment":comment})), true).map(|_| json!({"accepted":true})),
        OutlookCommand::Mail { command: MailCommand::Forward { id, to, comment, body, body_type, .. } } => mail.forward(&id, comment.map_or_else(|| json!({"message":{"body":{"contentType":body_type.graph(),"content":body.expect("validated body")},"toRecipients":to.iter().map(|address| json!({"emailAddress":{"address":address}})).collect::<Vec<_>>()}}), |comment| json!({"comment":comment,"toRecipients":to.iter().map(|address| json!({"emailAddress":{"address":address}})).collect::<Vec<_>>() }))).map(|_| json!({"accepted":true})),
        OutlookCommand::Mail { command: MailCommand::Sync { folder, reset, limit } } => {
            if reset {
                vault.delete_delta_link(profile_name, &graph_base, &folder).map_err(storage_error)?;
            }
            return sync_delta_cycle(&mail, vault, profile_name, &graph_base, &folder, limit);
        }
    };
    result.map_err(graph_error)
}
fn graph_error(error: GraphError) -> AppError {
    AppError::new(
        if matches!(error, GraphError::OutcomeUnknown { .. }) {
            "outcome_unknown"
        } else if matches!(error, GraphError::UnsupportedAttachment { .. }) {
            "unsupported_attachment"
        } else {
            "graph"
        },
        error.to_string(),
    )
}

fn local_error(error: moff_outlook_local::LocalError) -> AppError {
    AppError::new(error.code(), error.to_string())
}
fn local_outlook_data(command: OutlookCommand, dry_run: bool) -> Result<Value, AppError> {
    let limit = match &command {
        OutlookCommand::Mail {
            command: MailCommand::Visible { limit },
        } => *limit,
        _ => 25,
    };
    if matches!(
        &command,
        OutlookCommand::Mail {
            command: MailCommand::Visible { .. }
        }
    ) && limit > 100
    {
        return Err(AppError::new(
            "invalid_limit",
            "visible result limit must be <= 100",
        ));
    }
    if dry_run {
        let name = match command {
            OutlookCommand::Doctor => "outlook.doctor",
            OutlookCommand::Mail {
                command: MailCommand::Visible { .. },
            } => "outlook.mail.visible",
            _ => {
                return Err(AppError::new(
                    "ui_unsupported",
                    "local-macos supports only doctor and mail visible",
                ));
            }
        };
        return Ok(
            json!({"command":name,"provider":"local-macos","network":false,"auth_calls":false,"dry_run":true}),
        );
    }
    let ui = OsaScriptOutlookUi;
    match command {
        OutlookCommand::Doctor => local_doctor(&ui).map_err(local_error),
        OutlookCommand::Mail {
            command: MailCommand::Visible { limit },
        } => local_visible_mail(&ui, limit).map_err(local_error),
        _ => Err(AppError::new(
            "ui_unsupported",
            "local-macos supports only doctor and mail visible",
        )),
    }
}

fn command_data(
    command: Option<Command>,
    profile: String,
    dry_run: bool,
    yes: bool,
    vault: &dyn TokenVault,
    test_endpoints: bool,
) -> Result<Value, AppError> {
    match command {
        None => Ok(json!({"command":"root","status":"scaffold"})),
        Some(Command::Auth { command }) => auth_data(command, profile, dry_run, test_endpoints),
        Some(Command::Doctor { network }) => Ok(
            json!({"command":"doctor","status":"local_only","network_requested":network,"auth_calls":false}),
        ),
        Some(Command::Schema { source }) => {
            if let Some(value) = source.as_deref()
                && !known_source(value)
            {
                return Err(AppError::new(
                    ErrorCode::UnknownSource.as_str(),
                    format!("unknown source: {value}"),
                ));
            }
            Ok(schema(source.as_deref()))
        }
        Some(Command::Outlook { provider, command }) => {
            if provider == Provider::LocalMacos {
                local_outlook_data(command, dry_run)
            } else {
                outlook_data(command, dry_run, yes, &profile, vault, test_endpoints)
            }
        }
    }
}
fn run(cli: Cli, vault: &dyn TokenVault, test_endpoints: bool) -> Result<(), AppError> {
    if cli.json && cli.pretty {
        return Err(AppError::new(
            ErrorCode::IncompatibleOutputModes.as_str(),
            "--json and --pretty cannot be used together",
        ));
    }
    let profile = profile_name(cli.profile)?;
    let data = command_data(
        cli.command,
        profile,
        cli.dry_run,
        cli.yes,
        vault,
        test_endpoints,
    )?;
    let response = if cli.envelope {
        json!({"ok":true,"data":data,"warnings":[],"request_id":"moff-local"})
    } else {
        data
    };
    emit_json(&response, cli.pretty).expect("stdout is writable");
    Ok(())
}
pub(crate) fn main_with_vault(vault: &dyn TokenVault) -> ExitCode {
    main_with_vault_and_test_endpoints(vault, false)
}

#[cfg(feature = "test-support")]
#[allow(
    dead_code,
    reason = "called only by the feature-gated moff-test binary"
)]
pub(crate) fn main_with_test_endpoints(vault: &dyn TokenVault) -> ExitCode {
    main_with_vault_and_test_endpoints(vault, true)
}

fn main_with_vault_and_test_endpoints(vault: &dyn TokenVault, test_endpoints: bool) -> ExitCode {
    let cli = Cli::parse();
    let envelope = cli.envelope;
    let pretty = cli.pretty;
    match run(cli, vault, test_endpoints) {
        Ok(()) => ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("moff: [{}] {}", error.code, error.message);
            if envelope {
                let response = json!({"ok":false,"data":{"error":{"code":error.code,"message":error.message}},"warnings":[],"request_id":"moff-local"});
                emit_json(&response, pretty).expect("stdout is writable");
            }
            ExitCode::from(2)
        }
    }
}

#[allow(
    dead_code,
    reason = "the integration-only helper calls main_with_vault directly"
)]
fn main() -> ExitCode {
    main_with_vault(&KeyringVault)
}

#[cfg(test)]
mod tests {
    use std::collections::VecDeque;
    use std::sync::Mutex;

    use moff_core::InMemoryVault;

    use super::*;

    struct FakeGraph {
        responses: Mutex<VecDeque<Result<GraphResponse, GraphError>>>,
        requests: Mutex<Vec<GraphRequest>>,
    }
    impl FakeGraph {
        fn with(responses: Vec<Result<GraphResponse, GraphError>>) -> Self {
            Self {
                responses: Mutex::new(responses.into()),
                requests: Mutex::new(vec![]),
            }
        }
    }
    impl GraphHttp for FakeGraph {
        fn request(&self, request: GraphRequest) -> Result<GraphResponse, GraphError> {
            self.requests.lock().unwrap().push(request);
            self.responses.lock().unwrap().pop_front().unwrap()
        }
    }

    #[test]
    fn inline_body_over_one_mib_is_rejected_before_authentication() {
        let mut body = Some("x".repeat(1_048_577));
        let error = resolve_body_source(&mut body, None, false, true).unwrap_err();
        assert_eq!(error.code, "body_too_large");
    }

    #[test]
    fn production_endpoint_selection_ignores_environment_overrides() {
        assert!(matches!(
            entra_base(false),
            Ok(base) if base == "https://login.microsoftonline.com"
        ));
        assert!(matches!(
            graph_base(false),
            Ok(base) if base == "https://graph.microsoft.com/v1.0"
        ));
    }

    #[cfg(feature = "test-support")]
    #[test]
    fn test_endpoints_reject_spoofed_and_malformed_loopback_urls() {
        for (endpoint, graph) in [
            ("http://localhost:ignored@evil.example", false),
            ("https://localhost:8080", false),
            ("http://localhost", false),
            ("http://localhost:0", false),
            ("http://localhost:8080?query", false),
            ("http://localhost:ignored@evil.example/v1.0", true),
            ("http://127.0.0.1:8080/v1.0#fragment", true),
            ("http://127.0.0.1:8080/not-v1.0", true),
        ] {
            assert!(validate_test_endpoint(endpoint, graph).is_err());
        }
        assert!(validate_test_endpoint("http://localhost:8080", false).is_ok());
        assert!(validate_test_endpoint("http://127.0.0.1:8080/v1.0", true).is_ok());
    }

    #[cfg(unix)]
    #[test]
    fn body_file_symlink_is_rejected_without_following_it() {
        use std::os::unix::fs::symlink;

        let directory = std::env::temp_dir().join(format!("moff-body-link-{}", std::process::id()));
        fs::create_dir_all(&directory).unwrap();
        let target = directory.join("target.txt");
        let link = directory.join("body.txt");
        fs::write(&target, "body").unwrap();
        symlink(&target, &link).unwrap();
        let link = link.to_string_lossy().into_owned();
        let mut body = None;
        let error = resolve_body_source(&mut body, Some(&link), false, true).unwrap_err();
        assert_eq!(error.code, "body_file");
        fs::remove_dir_all(directory).unwrap();
    }

    #[test]
    fn sync_consumes_every_page_emits_all_records_and_persists_only_terminal_cursor() {
        let http = FakeGraph::with(vec![
            Ok(GraphResponse::json(
                200,
                r#"{"value":[{"id":"first"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/next"}"#,
            )),
            Ok(GraphResponse::json(
                200,
                r#"{"value":[{"id":"second"}],"@odata.deltaLink":"https://graph.microsoft.com/v1.0/delta"}"#,
            )),
        ]);
        let vault = InMemoryVault::default();
        vault
            .set_delta_link(
                "work",
                "https://graph.microsoft.com/v1.0",
                "inbox",
                "https://graph.microsoft.com/v1.0/old",
            )
            .unwrap();
        let mail = MailService::new(
            GraphClient::new(&http, "https://graph.microsoft.com/v1.0", "token").unwrap(),
        );

        let result = sync_delta_cycle(
            &mail,
            &vault,
            "work",
            "https://graph.microsoft.com/v1.0",
            "inbox",
            100,
        );
        assert!(result.is_ok());
        let result = result.ok().unwrap();

        assert_eq!(result["value"], json!([{"id":"first"}, {"id":"second"}]));
        assert_eq!(result["content_origin"], "external_email_untrusted");
        assert_eq!(
            vault
                .get_delta_link("work", "https://graph.microsoft.com/v1.0", "inbox")
                .unwrap()
                .as_deref(),
            Some("https://graph.microsoft.com/v1.0/delta")
        );
        let requests = http.requests.lock().unwrap();
        assert_eq!(requests.len(), 2);
        assert!(requests.iter().all(|request| {
            request.headers.get("Authorization").map(String::as_str) == Some("Bearer token")
        }));
        assert!(requests[0].url.ends_with("/old"));
        assert_eq!(requests[1].url, "https://graph.microsoft.com/v1.0/next");
        let emitted = result.to_string();
        assert!(!emitted.contains("token"));
    }

    #[test]
    fn sync_failure_after_a_nonterminal_page_does_not_replace_cursor() {
        let http = FakeGraph::with(vec![
            Ok(GraphResponse::json(
                200,
                r#"{"value":[{"id":"first"}],"@odata.nextLink":"https://graph.microsoft.com/v1.0/next"}"#,
            )),
            Err(GraphError::Transport),
            Err(GraphError::Transport),
            Err(GraphError::Transport),
        ]);
        let vault = InMemoryVault::default();
        vault
            .set_delta_link(
                "work",
                "https://graph.microsoft.com/v1.0",
                "inbox",
                "https://graph.microsoft.com/v1.0/old",
            )
            .unwrap();
        let mail = MailService::new(
            GraphClient::new(&http, "https://graph.microsoft.com/v1.0", "token").unwrap(),
        );

        assert!(
            sync_delta_cycle(
                &mail,
                &vault,
                "work",
                "https://graph.microsoft.com/v1.0",
                "inbox",
                100,
            )
            .is_err()
        );

        assert_eq!(
            vault
                .get_delta_link("work", "https://graph.microsoft.com/v1.0", "inbox")
                .unwrap()
                .as_deref(),
            Some("https://graph.microsoft.com/v1.0/old")
        );
    }

    #[test]
    fn sync_rejects_missing_or_untrusted_terminal_cursor_without_updating_state() {
        for terminal in [None, Some("https://evil.example/delta")] {
            let body = match terminal {
                Some(link) => format!(r#"{{"value":[],"@odata.deltaLink":"{link}"}}"#),
                None => r#"{"value":[]}"#.to_owned(),
            };
            let http = FakeGraph::with(vec![Ok(GraphResponse::json(200, &body))]);
            let vault = InMemoryVault::default();
            vault
                .set_delta_link(
                    "work",
                    "https://graph.microsoft.com/v1.0",
                    "inbox",
                    "https://graph.microsoft.com/v1.0/old",
                )
                .unwrap();
            let mail = MailService::new(
                GraphClient::new(&http, "https://graph.microsoft.com/v1.0", "token").unwrap(),
            );

            assert!(
                sync_delta_cycle(
                    &mail,
                    &vault,
                    "work",
                    "https://graph.microsoft.com/v1.0",
                    "inbox",
                    100,
                )
                .is_err()
            );
            assert_eq!(
                vault
                    .get_delta_link("work", "https://graph.microsoft.com/v1.0", "inbox")
                    .unwrap()
                    .as_deref(),
                Some("https://graph.microsoft.com/v1.0/old")
            );
        }
    }

    #[cfg(unix)]
    #[test]
    fn attachment_files_are_created_with_owner_only_permissions() {
        use std::os::unix::fs::PermissionsExt;

        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let directory = std::env::temp_dir().join(format!(
            "moff-attachment-test-{}-{nonce}",
            std::process::id()
        ));
        fs::create_dir(&directory).unwrap();
        let output = directory.join("attachment.txt");

        assert!(write_attachment(&output, b"sensitive", false).is_ok());

        assert_eq!(
            fs::metadata(&output).unwrap().permissions().mode() & 0o777,
            0o600
        );
        fs::remove_dir_all(directory).unwrap();
    }
}
