use moff_outlook_owa::{ArchivePlan, validate_archive_ids};
use serde_json::{Value, json};

use super::SessionArgs;

pub fn tools() -> Value {
    let ids = json!({"type":"array","items":{"type":"string","minLength":1},"minItems":1,"maxItems":20,"uniqueItems":true});
    json!({"tools":[
        {
            "name":"mail_archive_folders",
            "description":"Discover actual immediate mailbox folder IDs. Find the folder named Archive; never invent its ID. A truncated folder list is not a complete inventory.",
            "inputSchema":{"type":"object","properties":{},"additionalProperties":false},
            "annotations":{"readOnlyHint":true,"destructiveHint":false,"openWorldHint":true}
        },
        {
            "name":"mail_archive_preview",
            "description":"Preview 1-20 distinct non-draft Inbox messages for archiving. Supply discovered IDs and the actual Archive folder ID. Returns an exact versioned plan. Give a compact count, matching rule and optional examples; keep the full subjects available in the preview rather than reciting them all. Approved batches may move before whole-folder enumeration finishes. Do not claim the whole request complete until remaining pages are checked; folder offsets shift after moves. No mailbox changes.",
            "inputSchema":{"type":"object","properties":{"ids":ids,"archive_folder_id":{"type":"string","minLength":1}},"required":["ids","archive_folder_id"],"additionalProperties":false},
            "annotations":{"readOnlyHint":true,"destructiveHint":false,"openWorldHint":true}
        },
        {
            "name":"mail_archive",
            "description":"After Paul's approval, move exactly the messages in a mail_archive_preview plan from Inbox into the existing Archive folder. Supply the unchanged plan. This is not deletion and never sends mail. Tavern must approve this write. Maximum 20 per batch; stops on failure and reports partial progress. Never retry an uncertain move automatically.",
            "inputSchema":{"type":"object","properties":{"plan":{"type":"object","properties":{
                "archive_folder_id":{"type":"string","minLength":1},
                "inbox_folder_id":{"type":"string","minLength":1},
                "messages":{"type":"array","minItems":1,"maxItems":20,"items":{"type":"object","properties":{
                    "id":{"type":"string","minLength":1},"change_key":{"type":"string","minLength":1},"subject":{"type":"string"}
                },"required":["id","change_key","subject"],"additionalProperties":false}}
            },"required":["archive_folder_id","inbox_folder_id","messages"],"additionalProperties":false}},
            "required":["plan"],"additionalProperties":false},
            "annotations":{"readOnlyHint":false,"destructiveHint":true,"idempotentHint":false,"openWorldHint":true}
        }
    ]})
}

pub fn call(params: &Value, mut args: SessionArgs) -> Value {
    let result = execute(params, &mut args);
    match result {
        Ok(value) => super::mcp::tool_result(value.to_string(), value["ok"] == false),
        Err(error) => super::mcp::tool_result(error, true),
    }
}

fn execute(params: &Value, args: &mut SessionArgs) -> Result<Value, String> {
    let arguments = params["arguments"]
        .as_object()
        .ok_or("arguments must be an object")?;
    match params["name"].as_str() {
        Some("mail_archive_folders") => {
            if !arguments.is_empty() {
                return Err("mail_archive_folders accepts no arguments".into());
            }
            super::with_owa(
                *args,
                moff_outlook_owa::OutlookWebSession::mail_archive_folders,
            )
        }
        Some("mail_archive_preview") => {
            if arguments.len() != 2 {
                return Err("provide only ids and archive_folder_id".into());
            }
            let ids: Vec<String> = serde_json::from_value(params["arguments"]["ids"].clone())
                .map_err(|_| "ids must be an array of strings")?;
            let folder = params["arguments"]["archive_folder_id"]
                .as_str()
                .ok_or("archive_folder_id must be a string")?;
            validate_archive_ids(&ids).map_err(|e| e.to_string())?;
            if folder.trim().is_empty() {
                return Err("archive_folder_id must not be empty".into());
            }
            let session = super::session(*args)?;
            let plan = session
                .preview_mail_archive(&ids, folder)
                .map_err(|e| e.to_string())?;
            serde_json::to_value(plan).map_err(|e| e.to_string())
        }
        Some("mail_archive") => {
            if arguments.len() != 1 || !arguments.contains_key("plan") {
                return Err("provide only the unchanged preview plan".into());
            }
            let plan: ArchivePlan = serde_json::from_value(params["arguments"]["plan"].clone())
                .map_err(|_| "plan must be an unchanged mail_archive_preview result")?;
            let ids: Vec<_> = plan.messages.iter().map(|m| m.id.clone()).collect();
            validate_archive_ids(&ids).map_err(|e| e.to_string())?;
            if plan.archive_folder_id.trim().is_empty()
                || plan.inbox_folder_id.trim().is_empty()
                || plan.messages.iter().any(|m| m.change_key.trim().is_empty())
            {
                return Err("plan folder IDs and change keys must not be empty".into());
            }
            // The separate archive server is an explicit host opt-in. Its only write tool
            // is annotated as destructive so Tavern's existing approval gate applies.
            args.allow_mutation = true;
            super::with_owa(*args, |s| s.apply_mail_archive(&plan))
        }
        _ => Err("Unknown archive tool".into()),
    }
}
