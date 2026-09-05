#![allow(
    clippy::enum_variant_names,
    clippy::needless_pass_by_value,
    reason = "CLI-owned values and Microsoft scope vocabulary are clearer at call sites"
)]

use std::collections::BTreeMap;
use std::io::{self, Write};
use std::process::ExitCode;

use clap::{ArgAction, Parser, Subcommand, ValueEnum};
use moff_core::{
    AuthService, DeviceFlowError, ErrorCode, GraphError, GraphHttp, GraphRequest, GraphResponse,
    HttpClient, HttpResponse, KeyringVault, Profile, ProfileStore, ThreadSleeper, TokenVault,
    known_source, schema,
};
use serde_json::{Value, json};
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
        #[command(subcommand)]
        command: OutlookCommand,
    },
}
#[derive(Debug, Subcommand)]
enum OutlookCommand {
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
    List {
        #[arg(long, default_value = "inbox")]
        folder: String,
        #[arg(long, default_value_t = 25)]
        limit: u16,
    },
    Search {
        query: String,
        #[arg(long, default_value_t = 25)]
        limit: u16,
    },
    Get {
        id: String,
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
        body: String,
        #[arg(long, value_enum, default_value_t = BodyType::Text)]
        body_type: BodyType,
    },
    DraftUpdate {
        id: String,
        #[arg(long)]
        subject: Option<String>,
        #[arg(long)]
        body: Option<String>,
        #[arg(long, value_enum, default_value_t = BodyType::Text)]
        body_type: BodyType,
    },
    Send {
        #[arg(long)]
        to: Vec<String>,
        #[arg(long)]
        subject: Option<String>,
        #[arg(long)]
        body: String,
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
        id: String,
        #[arg(long)]
        comment: Option<String>,
        #[arg(long)]
        body: Option<String>,
        #[arg(long)]
        confirm_send: bool,
    },
    ReplyAll {
        id: String,
        #[arg(long)]
        comment: Option<String>,
        #[arg(long)]
        confirm_send: bool,
    },
    Forward {
        id: String,
        #[arg(long)]
        to: Vec<String>,
        #[arg(long)]
        comment: Option<String>,
        #[arg(long)]
        confirm_send: bool,
    },
    Sync {
        #[arg(long, default_value = "inbox")]
        folder: String,
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
#[derive(Debug, Subcommand)]
enum AttachmentsCommand {
    List {
        message_id: String,
    },
    Download {
        message_id: String,
        attachment_id: String,
        #[arg(long)]
        output: String,
        #[arg(long)]
        overwrite: bool,
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
            Self::MailWrite => "Mail.Write",
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
        let body = response.text().map_err(|_| GraphError::Transport)?;
        Ok(GraphResponse {
            status,
            body,
            request_id,
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

fn auth_data(command: AuthCommand, profile_name: String, dry_run: bool) -> Result<Value, AppError> {
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
            let base = std::env::var("MOFF_ENTRA_BASE_URL")
                .unwrap_or_else(|_| "https://login.microsoftonline.com".to_owned());
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
    command: OutlookCommand,
    dry_run: bool,
    yes: bool,
    _profile_name: &str,
) -> Result<Value, AppError> {
    let (name, mutation) = match &command {
        OutlookCommand::Folders {
            command: FoldersCommand::List,
        } => ("outlook.folders.list", false),
        OutlookCommand::Attachments {
            command: AttachmentsCommand::List { .. },
        } => ("outlook.attachments.list", false),
        OutlookCommand::Attachments {
            command: AttachmentsCommand::Download { overwrite, .. },
        } => {
            if *overwrite && !yes {
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
            command: MailCommand::Sync { .. },
        } => ("outlook.mail.sync", false),
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
            if let OutlookCommand::Mail {
                command: MailCommand::Reply { comment, body, .. },
            } = &command
                && comment.is_some() == body.is_some()
            {
                return Err(AppError::new(
                    "invalid_reply",
                    "reply requires exactly one of --comment or --body",
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
    if mutation {
        return Err(AppError::new(
            "authentication",
            "configure an authenticated profile before Graph mutations",
        ));
    }
    Err(AppError::new(
        "authentication",
        "configure an authenticated profile before Graph reads",
    ))
}

fn command_data(
    command: Option<Command>,
    profile: String,
    dry_run: bool,
    yes: bool,
) -> Result<Value, AppError> {
    match command {
        None => Ok(json!({"command":"root","status":"scaffold"})),
        Some(Command::Auth { command }) => auth_data(command, profile, dry_run),
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
        Some(Command::Outlook { command }) => outlook_data(command, dry_run, yes, &profile),
    }
}
fn run(cli: Cli) -> Result<(), AppError> {
    if cli.json && cli.pretty {
        return Err(AppError::new(
            ErrorCode::IncompatibleOutputModes.as_str(),
            "--json and --pretty cannot be used together",
        ));
    }
    let profile = profile_name(cli.profile)?;
    let data = command_data(cli.command, profile, cli.dry_run, cli.yes)?;
    let response = if cli.envelope {
        json!({"ok":true,"data":data,"warnings":[],"request_id":"moff-local"})
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
            eprintln!("moff: [{}] {}", error.code, error.message);
            if envelope {
                let response = json!({"ok":false,"data":{"error":{"code":error.code,"message":error.message}},"warnings":[],"request_id":"moff-local"});
                emit_json(&response, pretty).expect("stdout is writable");
            }
            ExitCode::from(2)
        }
    }
}
