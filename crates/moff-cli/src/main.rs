use std::{
    env, fs,
    path::{Path, PathBuf},
    process::ExitCode,
};

use clap::{Args, Parser, Subcommand};
use moff_outlook_owa::{
    DraftInput, EventInput, GraphWebSession, OutlookWebError, OutlookWebSession,
    calendar_folders_request, calendar_view_request, validate_read_action,
};
use serde_json::{Value, json};

#[derive(Debug, Parser)]
#[command(
    name = "moff",
    version,
    about = "Personal Microsoft 365 browser-session CLI"
)]
struct Cli {
    #[command(subcommand)]
    command: Option<Command>,
}

#[derive(Debug, Subcommand)]
enum Command {
    Outlook {
        #[command(flatten)]
        session: SessionArgs,
        #[command(subcommand)]
        command: OutlookCommand,
    },
    Schema,
}

#[derive(Debug, Args)]
struct SessionArgs {
    #[arg(
        long,
        value_name = "PATH",
        help = "Path to a browser-exported Outlook Web or Microsoft Graph cURL request"
    )]
    session_curl: Option<PathBuf>,
}

#[derive(Debug, Subcommand)]
enum OutlookCommand {
    Mail {
        #[command(subcommand)]
        command: MailCommand,
    },
    Calendar {
        #[command(subcommand)]
        command: CalendarCommand,
    },
    People {
        #[command(subcommand)]
        command: PeopleCommand,
    },
    Org {
        #[command(subcommand)]
        command: OrgCommand,
    },
    Onedrive {
        #[command(subcommand)]
        command: OneDriveCommand,
    },
    Groups {
        #[command(subcommand)]
        command: GroupsCommand,
    },
    Loop {
        #[command(subcommand)]
        command: LoopCommand,
    },
    Service {
        #[command(subcommand)]
        command: ServiceCommand,
    },
}

#[derive(Debug, Subcommand)]
enum MailCommand {
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
        #[arg(long, default_value = "inbox")]
        folder: String,
        #[arg(long, default_value_t = 0)]
        offset: u32,
    },
    Get {
        #[arg(long)]
        id: String,
    },
    /// Search a folder's subjects and text bodies, using a case-insensitive literal query.
    Search {
        #[arg(long)]
        query: String,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
        #[arg(long, default_value = "inbox")]
        folder: String,
    },
    /// Create an unsent draft from JSON containing optional subject and body strings.
    Create {
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Update only the specified subject/body fields of an unsent draft.
    Update {
        #[arg(long)]
        id: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Permanently delete a message or draft (`HardDelete`).
    Delete {
        #[arg(long)]
        id: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
}

#[derive(Debug, Subcommand)]
enum CalendarCommand {
    Folders,
    View {
        #[arg(long)]
        folder_id: String,
        #[arg(long)]
        start: String,
        #[arg(long)]
        end: String,
    },
    List {
        #[arg(long)]
        folder_id: String,
        #[arg(long)]
        start: String,
        #[arg(long)]
        end: String,
    },
    /// Search subjects and bodies within an explicit calendar time range.
    Search {
        #[arg(long)]
        folder_id: String,
        #[arg(long)]
        start: String,
        #[arg(long)]
        end: String,
        #[arg(long)]
        query: String,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        id: String,
    },
    /// Create a single personal appointment without attendees or invitations.
    Create {
        #[arg(long)]
        folder_id: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Update a single personal appointment without attendees.
    Update {
        #[arg(long)]
        id: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
        #[arg(long, required = true)]
        confirm: bool,
    },
    /// Move a single personal appointment without attendees to Deleted Items.
    Delete {
        #[arg(long)]
        id: String,
        #[arg(long, required = true)]
        confirm: bool,
    },
}

#[derive(Debug, Subcommand)]
enum PeopleCommand {
    Search {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        user: String,
    },
}

#[derive(Debug, Subcommand)]
enum OrgCommand {
    Manager {
        #[arg(long)]
        user: Option<String>,
    },
    Reports {
        #[arg(long)]
        user: Option<String>,
    },
}

#[derive(Debug, Subcommand)]
enum OneDriveCommand {
    List {
        #[arg(long)]
        item_id: Option<String>,
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        id: String,
    },
}

#[derive(Debug, Subcommand)]
enum GroupsCommand {
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
}

#[derive(Debug, Subcommand)]
enum LoopCommand {
    List {
        #[arg(long, default_value_t = 25, value_parser = clap::value_parser!(u16).range(1..=100))]
        limit: u16,
    },
    Get {
        #[arg(long)]
        id: String,
    },
}

#[derive(Debug, Subcommand)]
enum ServiceCommand {
    Read {
        #[arg(long)]
        action: String,
        #[arg(long, value_name = "PATH")]
        body_file: PathBuf,
    },
}

fn print_json(value: &Value) {
    println!(
        "{}",
        serde_json::to_string(value).expect("JSON values serialize")
    );
}

fn error(message: impl std::fmt::Display) -> ExitCode {
    eprintln!("[browser_session] {message}");
    ExitCode::from(2)
}

fn session(args: SessionArgs) -> Result<OutlookWebSession, String> {
    let path = args.session_curl.or_else(|| {
        env::var_os("HOME")
            .map(|home| PathBuf::from(home).join(".config/moff/outlook-session.curl"))
    });
    OutlookWebSession::from_curl_file(
        &path.ok_or("--session-curl PATH is required and HOME is unavailable")?,
    )
    .map_err(|error| error.to_string())
}

fn graph_session(args: SessionArgs) -> Result<GraphWebSession, String> {
    let path = args.session_curl.or_else(|| {
        env::var_os("HOME").map(|home| PathBuf::from(home).join(".config/moff/graph-session.curl"))
    });
    GraphWebSession::from_curl_file(
        &path.ok_or("--session-curl PATH is required and HOME is unavailable")?,
    )
    .map_err(|error| error.to_string())
}

fn encode_segment(segment: &str) -> String {
    url::form_urlencoded::byte_serialize(segment.as_bytes()).collect()
}

fn user_path(user: Option<&str>, relationship: &str) -> String {
    match user {
        Some(user) => format!("/v1.0/users/{}/{}", encode_segment(user), relationship),
        None => format!("/v1.0/me/{relationship}"),
    }
}

fn output(command: &str, value: &Value) {
    print_json(
        &json!({"schema":moff_core::SCHEMA_VERSION,"command":command,"provider":"browser-session","content_origin":"microsoft_browser_session_untrusted","value":value}),
    );
}

#[allow(clippy::collapsible_if, clippy::needless_pass_by_value)]
fn graph_result(
    args: SessionArgs,
    command: &str,
    path: String,
    only_loop_files: bool,
) -> Result<(&'static str, Value), String> {
    let mut value = graph_session(args)
        .and_then(|session| session.get(&path).map_err(|error| error.to_string()))?;
    if only_loop_files {
        if let Some(items) = value.get_mut("value").and_then(Value::as_array_mut) {
            items.retain(|item| {
                item["name"].as_str().is_some_and(|name| {
                    Path::new(name)
                        .extension()
                        .is_some_and(|ext| ext.eq_ignore_ascii_case("loop"))
                })
            });
        }
    }
    let command = match command {
        "outlook.people.search" => "outlook.people.search",
        "outlook.people.get" => "outlook.people.get",
        "outlook.org.manager" => "outlook.org.manager",
        "outlook.org.reports" => "outlook.org.reports",
        "outlook.onedrive.list" => "outlook.onedrive.list",
        "outlook.onedrive.get" => "outlook.onedrive.get",
        "outlook.groups.list" => "outlook.groups.list",
        "outlook.loop.list" => "outlook.loop.list",
        "outlook.loop.get" => "outlook.loop.get",
        _ => return Err("unsupported browser-session command".to_owned()),
    };
    Ok((command, value))
}

fn confirmed_json(path: PathBuf, confirm: bool) -> Result<Value, String> {
    if !confirm {
        return Err("mutations require --confirm".to_owned());
    }
    fs::read_to_string(path)
        .map_err(|_| "--body-file could not be read".to_owned())
        .and_then(|body| {
            serde_json::from_str(&body).map_err(|_| "--body-file must contain JSON".to_owned())
        })
}

fn with_owa(
    args: SessionArgs,
    operation: impl FnOnce(&OutlookWebSession) -> Result<Value, OutlookWebError>,
) -> Result<Value, String> {
    operation(&session(args)?).map_err(|error| error.to_string())
}

fn mail_result(args: SessionArgs, command: MailCommand) -> Result<(&'static str, Value), String> {
    match command {
        MailCommand::List {
            folder,
            limit,
            offset,
        } => Ok((
            "outlook.mail.list",
            with_owa(args, |s| s.mail_list(&folder, limit, offset))?,
        )),
        MailCommand::Get { id } => Ok(("outlook.mail.get", with_owa(args, |s| s.mail_get(&id))?)),
        MailCommand::Search {
            folder,
            query,
            limit,
        } => Ok((
            "outlook.mail.search",
            with_owa(args, |s| s.mail_search(&folder, &query, limit))?,
        )),
        MailCommand::Create { body_file, confirm } => {
            let input: DraftInput = serde_json::from_value(confirmed_json(body_file, confirm)?)
                .map_err(|_| "--body-file must contain only optional subject and body strings")?;
            Ok((
                "outlook.mail.create",
                with_owa(args, |s| s.create_draft(&input))?,
            ))
        }
        MailCommand::Update {
            id,
            body_file,
            confirm,
        } => {
            let input: DraftInput = serde_json::from_value(confirmed_json(body_file, confirm)?)
                .map_err(|_| "--body-file must contain only optional subject and body strings")?;
            Ok((
                "outlook.mail.update",
                with_owa(args, |s| s.update_draft(&id, &input))?,
            ))
        }
        MailCommand::Delete { id, confirm } => {
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            Ok((
                "outlook.mail.delete",
                with_owa(args, |s| s.delete_message(&id))?,
            ))
        }
    }
}

fn calendar_result(
    args: SessionArgs,
    command: CalendarCommand,
) -> Result<(&'static str, Value), String> {
    match command {
        CalendarCommand::Folders => Ok((
            "outlook.calendar.folders",
            with_owa(args, |s| {
                s.call("GetCalendarFoldersV2", &calendar_folders_request())
            })?,
        )),
        CalendarCommand::View {
            folder_id,
            start,
            end,
        } => Ok((
            "outlook.calendar.view",
            with_owa(args, |s| {
                s.call(
                    "GetCalendarView",
                    &calendar_view_request(&folder_id, &start, &end),
                )
            })?,
        )),
        CalendarCommand::List {
            folder_id,
            start,
            end,
        } => Ok((
            "outlook.calendar.list",
            with_owa(args, |s| s.calendar_list(&folder_id, &start, &end))?,
        )),
        CalendarCommand::Search {
            folder_id,
            start,
            end,
            query,
            limit,
        } => Ok((
            "outlook.calendar.search",
            with_owa(args, |s| {
                s.calendar_search(&folder_id, &start, &end, &query, limit)
            })?,
        )),
        CalendarCommand::Get { id } => Ok((
            "outlook.calendar.get",
            with_owa(args, |s| s.calendar_get(&id))?,
        )),
        CalendarCommand::Create {
            folder_id,
            body_file,
            confirm,
        } => {
            let input: EventInput = serde_json::from_value(confirmed_json(body_file, confirm)?)
                .map_err(
                    |_| "--body-file must contain only subject, body, start, and end strings",
                )?;
            Ok((
                "outlook.calendar.create",
                with_owa(args, |s| s.create_event(&folder_id, &input))?,
            ))
        }
        CalendarCommand::Update {
            id,
            body_file,
            confirm,
        } => {
            let input: EventInput = serde_json::from_value(confirmed_json(body_file, confirm)?)
                .map_err(
                    |_| "--body-file must contain only subject, body, start, and end strings",
                )?;
            Ok((
                "outlook.calendar.update",
                with_owa(args, |s| s.update_event(&id, &input))?,
            ))
        }
        CalendarCommand::Delete { id, confirm } => {
            if !confirm {
                return Err("mutations require --confirm".to_owned());
            }
            Ok((
                "outlook.calendar.delete",
                with_owa(args, |s| s.delete_event(&id))?,
            ))
        }
    }
}

#[allow(clippy::too_many_lines)]
fn main() -> ExitCode {
    let cli = Cli::parse();
    match cli.command {
        Some(Command::Schema) => {
            print_json(
                &json!({"schema":moff_core::SCHEMA_VERSION,"provider":"browser-session","commands":{"mail":{"list":true,"search":true,"get":true,"create":true,"update":true,"delete":true,"send":false},"calendar":{"folders":true,"view":true,"list":true,"search":true,"get":true,"create":true,"update":true,"delete":true},"people":{"search":true,"get":true},"org":{"manager":true,"reports":true},"onedrive":{"list":true,"get":true},"groups":{"list":true},"loop":{"list":true,"get":true},"service":{"read":true}},"unimplemented_products":[],"mutations":true,"mutation_confirmation":"--confirm","mail_delete":"permanent","calendar_delete":"deleted-items","calendar_mutation_scope":"single appointments without attendees","search_semantics":"case-insensitive literal subject/body search within the chosen folder or calendar range"}),
            );
            ExitCode::SUCCESS
        }
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Service {
                    command: ServiceCommand::Read { action, body_file },
                },
        }) => {
            if let Err(validation_error) = validate_read_action(&action) {
                return error(validation_error);
            }
            let Some(body) = fs::read_to_string(body_file)
                .ok()
                .and_then(|text| serde_json::from_str(&text).ok())
            else {
                return error("--body-file must contain a JSON request body");
            };
            match session(args).and_then(|s| s.call(&action, &body).map_err(|e| e.to_string())) {
                Ok(value) => {
                    output("outlook.service.read", &value);
                    ExitCode::SUCCESS
                }
                Err(e) => error(e),
            }
        }
        Some(Command::Outlook {
            session: args,
            command: OutlookCommand::Mail { command },
        }) => match mail_result(args, command) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command: OutlookCommand::Calendar { command },
        }) => match calendar_result(args, command) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command: PeopleCommand::Search { limit },
                },
        }) => match graph_result(
            args,
            "outlook.people.search",
            format!("/v1.0/me/people?$top={limit}"),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::People {
                    command: PeopleCommand::Get { user },
                },
        }) => match graph_result(
            args,
            "outlook.people.get",
            format!(
                "/v1.0/users/{}?$select=id,displayName,mail,userPrincipalName,userType",
                encode_segment(&user)
            ),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Org {
                    command: OrgCommand::Manager { user },
                },
        }) => match graph_result(
            args,
            "outlook.org.manager",
            user_path(user.as_deref(), "manager"),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Org {
                    command: OrgCommand::Reports { user },
                },
        }) => match graph_result(
            args,
            "outlook.org.reports",
            user_path(user.as_deref(), "directReports"),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Onedrive {
                    command: OneDriveCommand::List { item_id, limit },
                },
        }) => {
            let base = item_id.map_or_else(
                || "root".to_owned(),
                |id| format!("items/{}", encode_segment(&id)),
            );
            match graph_result(
                args,
                "outlook.onedrive.list",
                format!("/v1.0/me/drive/{base}/children?$top={limit}"),
                false,
            ) {
                Ok((command, value)) => {
                    output(command, &value);
                    ExitCode::SUCCESS
                }
                Err(e) => error(e),
            }
        }
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Onedrive {
                    command: OneDriveCommand::Get { id },
                },
        }) => match graph_result(
            args,
            "outlook.onedrive.get",
            format!("/v1.0/me/drive/items/{}", encode_segment(&id)),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Groups {
                    command: GroupsCommand::List { limit },
                },
        }) => match graph_result(
            args,
            "outlook.groups.list",
            format!(
                "/v1.0/groups?$filter=groupTypes%2Fany%28c%3Ac%20eq%20%27Unified%27%29&$top={limit}"
            ),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Loop {
                    command: LoopCommand::List { limit },
                },
        }) => match graph_result(
            args,
            "outlook.loop.list",
            format!("/v1.0/me/drive/root/search(q='.loop')?$top={limit}"),
            true,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        Some(Command::Outlook {
            session: args,
            command:
                OutlookCommand::Loop {
                    command: LoopCommand::Get { id },
                },
        }) => match graph_result(
            args,
            "outlook.loop.get",
            format!("/v1.0/me/drive/items/{}", encode_segment(&id)),
            false,
        ) {
            Ok((command, value)) => {
                output(command, &value);
                ExitCode::SUCCESS
            }
            Err(e) => error(e),
        },
        None => {
            print_json(&json!({"command":"moff","provider":"browser-session","status":"ready"}));
            ExitCode::SUCCESS
        }
    }
}
