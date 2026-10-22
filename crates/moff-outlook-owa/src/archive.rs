use serde::{Deserialize, Serialize};
use serde_json::{Value, json};

use crate::{OutlookWebError, OutlookWebSession, header};

#[derive(Debug, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct ArchiveMessage {
    pub id: String,
    pub change_key: String,
    pub subject: String,
}

#[derive(Debug, Deserialize, Serialize)]
#[serde(deny_unknown_fields)]
pub struct ArchivePlan {
    pub archive_folder_id: String,
    pub inbox_folder_id: String,
    pub messages: Vec<ArchiveMessage>,
}

fn nonempty(value: &str) -> Result<(), OutlookWebError> {
    if value.trim().is_empty() {
        return Err(OutlookWebError::InvalidInput(
            "folder and message IDs must not be empty",
        ));
    }
    Ok(())
}

fn folder_request(id: &str, distinguished: bool) -> Value {
    json!({"Header":header(),"Body":{
        "__type":"GetFolderRequest:#Exchange",
        "FolderShape":{"__type":"FolderResponseShape:#Exchange","BaseShape":"Default",
            "AdditionalProperties":[{"__type":"PropertyUri:#Exchange","FieldURI":"FolderClass"}]},
        "FolderIds":[{"__type":if distinguished {"DistinguishedFolderId:#Exchange"} else {"FolderId:#Exchange"},"Id":id}]
    }})
}

fn first<'a>(response: &'a Value, field: &str) -> Result<&'a Value, OutlookWebError> {
    let messages = response
        .pointer("/Body/ResponseMessages/Items")
        .and_then(Value::as_array)
        .ok_or(OutlookWebError::InvalidResponse)?;
    if messages.len() != 1 {
        return Err(OutlookWebError::InvalidResponse);
    }
    let items = messages[0]
        .get(field)
        .and_then(Value::as_array)
        .ok_or(OutlookWebError::InvalidResponse)?;
    if items.len() != 1 {
        return Err(OutlookWebError::InvalidResponse);
    }
    Ok(&items[0])
}

fn string<'a>(value: &'a Value, pointer: &str) -> Result<&'a str, OutlookWebError> {
    value
        .pointer(pointer)
        .and_then(Value::as_str)
        .filter(|s| !s.trim().is_empty())
        .ok_or(OutlookWebError::InvalidResponse)
}

pub fn validate_archive_ids(ids: &[String]) -> Result<(), OutlookWebError> {
    if ids.is_empty() || ids.len() > 20 {
        return Err(OutlookWebError::InvalidInput(
            "archive requires 1-20 distinct message IDs",
        ));
    }
    let mut seen = std::collections::HashSet::new();
    for id in ids {
        nonempty(id)?;
        if !seen.insert(id) {
            return Err(OutlookWebError::InvalidInput(
                "duplicate archive message ID",
            ));
        }
    }
    Ok(())
}

pub fn move_to_archive_request(
    id: &str,
    key: &str,
    folder: &str,
) -> Result<Value, OutlookWebError> {
    for value in [id, key, folder] {
        nonempty(value)?;
    }
    Ok(
        json!({"__type":"MoveItemJsonRequest:#Exchange","Header":header(),"Body":{
            "__type":"MoveItemRequest:#Exchange",
            "ToFolderId":{"__type":"TargetFolderId:#Exchange","BaseFolderId":{"__type":"FolderId:#Exchange","Id":folder}},
            "ItemIds":[{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key}],
            "ReturnNewItemIds":true
        }}),
    )
}

impl OutlookWebSession {
    /// Enumerate only immediate mailbox folders, not an online archive mailbox.
    pub fn mail_archive_folders(&self) -> Result<Value, OutlookWebError> {
        let request = json!({"Header":header(),"Body":{
            "__type":"FindFolderRequest:#Exchange","Traversal":"Shallow",
            "FolderShape":{"__type":"FolderResponseShape:#Exchange","BaseShape":"Default"},
            "ParentFolderIds":[{"__type":"DistinguishedFolderId:#Exchange","Id":"msgfolderroot"}],
            "Paging":{"__type":"IndexedPageView:#Exchange","BasePoint":"Beginning","Offset":0,"MaxEntriesReturned":100}
        }});
        self.execute("FindFolder", &request, false)
    }
}

#[cfg(test)]
#[path = "archive_tests.rs"]
mod tests;

impl OutlookWebSession {
    fn archive_folders(&self, archive: &str) -> Result<String, OutlookWebError> {
        nonempty(archive)?;
        let response = self.execute("GetFolder", &folder_request(archive, false), false)?;
        let folder = first(&response, "Folders")?;
        if string(folder, "/FolderId/Id")? != archive
            || !string(folder, "/DisplayName")?.eq_ignore_ascii_case("Archive")
            || string(folder, "/FolderClass")? != "IPF.Note"
        {
            return Err(OutlookWebError::InvalidInput(
                "destination must be an existing mail folder named Archive; localized names are not yet supported",
            ));
        }
        let inbox = self.execute("GetFolder", &folder_request("inbox", true), false)?;
        let inbox_id = string(first(&inbox, "Folders")?, "/FolderId/Id")?;
        if inbox_id == archive {
            return Err(OutlookWebError::InvalidInput(
                "archive destination cannot be Inbox",
            ));
        }
        Ok(inbox_id.to_owned())
    }

    /// Preview exact, versioned Inbox messages without changing their read state.
    pub fn preview_mail_archive(
        &self,
        ids: &[String],
        archive: &str,
    ) -> Result<ArchivePlan, OutlookWebError> {
        validate_archive_ids(ids)?;
        let inbox = self.archive_folders(archive)?;
        let mut messages = Vec::new();
        for id in ids {
            let response = self.mail_get(id)?;
            let item = first(&response, "Items")?;
            if !crate::operations::is_mail_message(item)
                || item["IsDraft"].as_bool() != Some(false)
                || string(item, "/ParentFolderId/Id")? != inbox
                || string(item, "/ItemId/Id")? != id
            {
                return Err(OutlookWebError::InvalidInput(
                    "archive supports only non-draft messages currently in Inbox",
                ));
            }
            messages.push(ArchiveMessage {
                id: id.clone(),
                change_key: string(item, "/ItemId/ChangeKey")?.to_owned(),
                subject: item["Subject"].as_str().unwrap_or("").to_owned(),
            });
        }
        Ok(ArchivePlan {
            archive_folder_id: archive.to_owned(),
            inbox_folder_id: inbox,
            messages,
        })
    }

    /// Revalidate the entire plan before the first write. Stop at the first uncertain outcome.
    pub fn apply_mail_archive(&self, plan: &ArchivePlan) -> Result<Value, OutlookWebError> {
        let ids: Vec<_> = plan.messages.iter().map(|m| m.id.clone()).collect();
        let current = self.preview_mail_archive(&ids, &plan.archive_folder_id)?;
        if current.inbox_folder_id != plan.inbox_folder_id
            || current
                .messages
                .iter()
                .zip(&plan.messages)
                .any(|(a, b)| a.change_key != b.change_key || a.subject != b.subject)
        {
            return Err(OutlookWebError::InvalidInput(
                "archive plan changed; preview and approve again",
            ));
        }
        let mut completed = Vec::new();
        for message in &plan.messages {
            let result = (|| {
                let request = move_to_archive_request(
                    &message.id,
                    &message.change_key,
                    &plan.archive_folder_id,
                )?;
                let response = self.execute("MoveItem", &request, true)?;
                let new_id = string(first(&response, "Items")?, "/ItemId/Id")?.to_owned();
                let verified = self.mail_get(&new_id)?;
                let item = first(&verified, "Items")?;
                if string(item, "/ParentFolderId/Id")? != plan.archive_folder_id
                    || string(item, "/ItemId/Id")? != new_id
                {
                    return Err(OutlookWebError::VerificationFailed);
                }
                Ok(json!({"source_id":message.id,"archived_id":new_id}))
            })();
            match result {
                Ok(item) => completed.push(item),
                Err(error) => {
                    return Ok(json!({
                        "ok":false,"completed":completed,"failed_or_uncertain_id":message.id,
                        "error":error.to_string(),
                        "instructions":"Stopped without retrying. Inspect Inbox and Archive before retrying this message; earlier completed messages were moved."
                    }));
                }
            }
        }
        Ok(json!({"ok":true,"completed":completed}))
    }
}
