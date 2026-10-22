use std::{collections::HashSet, fs::File, io::Read, path::PathBuf};

use base64::{Engine, engine::general_purpose::STANDARD};
use serde_json::{Value, json};

use crate::operations::single_item;
use crate::{OutlookWebError, OutlookWebSession, SendMailInput, header, send_mail_request};

pub const MAX_MAIL_ATTACHMENT_BYTES: usize = 500_000;
pub const MAX_MAIL_ATTACHMENTS: usize = 5;

pub struct MailAttachment {
    name: String,
    bytes: Vec<u8>,
}

impl MailAttachment {
    pub fn new(name: String, bytes: Vec<u8>) -> Result<Self, OutlookWebError> {
        if name.trim().is_empty()
            || name.len() > 255
            || name
                .chars()
                .any(|c| c.is_control() || matches!(c, '/' | '\\' | '\u{2028}' | '\u{2029}'))
            || matches!(name.as_str(), "." | "..")
            || bytes.len() > MAX_MAIL_ATTACHMENT_BYTES
        {
            return Err(OutlookWebError::InvalidInput(
                "invalid attachment filename or size (max 500000 bytes)",
            ));
        }
        Ok(Self { name, bytes })
    }
}

pub fn load_mail_attachments(paths: &[PathBuf]) -> Result<Vec<MailAttachment>, OutlookWebError> {
    if paths.len() > MAX_MAIL_ATTACHMENTS {
        return Err(OutlookWebError::InvalidInput(
            "send accepts at most 5 attachments",
        ));
    }
    let mut attachments = Vec::new();
    let mut total = 0;
    for path in paths {
        let metadata = std::fs::symlink_metadata(path).map_err(OutlookWebError::FileTransferIo)?;
        if !metadata.is_file() || metadata.len() > (MAX_MAIL_ATTACHMENT_BYTES - total) as u64 {
            return Err(OutlookWebError::InvalidInput(
                "attachments must be regular non-symlink files, at most 500000 bytes combined",
            ));
        }
        let name =
            path.file_name()
                .and_then(|v| v.to_str())
                .ok_or(OutlookWebError::InvalidInput(
                    "attachment filename must be UTF-8",
                ))?;
        let mut options = File::options();
        options.read(true);
        #[cfg(unix)]
        {
            use std::os::unix::fs::OpenOptionsExt;
            options.custom_flags(libc::O_NOFOLLOW | libc::O_NONBLOCK);
        }
        let file = options
            .open(path)
            .map_err(OutlookWebError::FileTransferIo)?;
        if !file
            .metadata()
            .map_err(OutlookWebError::FileTransferIo)?
            .is_file()
        {
            return Err(OutlookWebError::InvalidInput(
                "attachment is not a regular file",
            ));
        }
        let mut bytes = Vec::new();
        file.take((MAX_MAIL_ATTACHMENT_BYTES - total + 1) as u64)
            .read_to_end(&mut bytes)
            .map_err(OutlookWebError::FileTransferIo)?;
        total += bytes.len();
        if total > MAX_MAIL_ATTACHMENT_BYTES {
            return Err(OutlookWebError::InvalidInput(
                "attachment files grew beyond the 500000-byte combined limit while reading",
            ));
        }
        attachments.push(MailAttachment::new(name.to_owned(), bytes)?);
    }
    validate_attachments(&attachments)?;
    Ok(attachments)
}

fn validate_attachments(attachments: &[MailAttachment]) -> Result<(), OutlookWebError> {
    let mut names = HashSet::new();
    if attachments.is_empty()
        || attachments.len() > MAX_MAIL_ATTACHMENTS
        || attachments.iter().map(|a| a.bytes.len()).sum::<usize>() > MAX_MAIL_ATTACHMENT_BYTES
        || attachments
            .iter()
            .any(|a| !names.insert(a.name.to_lowercase()))
    {
        return Err(OutlookWebError::InvalidInput(
            "provide 1-5 uniquely named attachments, at most 500000 bytes combined",
        ));
    }
    Ok(())
}

pub fn attachment_draft_request(
    input: &SendMailInput,
    attachments: &[MailAttachment],
) -> Result<Value, OutlookWebError> {
    validate_attachments(attachments)?;
    let mut request = send_mail_request(input)?;
    request["Body"]["MessageDisposition"] = json!("SaveOnly");
    request["Body"]["Items"][0]["Attachments"] = json!(attachments.iter().map(|a| json!({
        "__type":"FileAttachment:#Exchange","Name":a.name,"ContentType":"application/octet-stream",
        "IsInline":false,"IsContactPhoto":false,"Content":STANDARD.encode(&a.bytes)
    })).collect::<Vec<_>>());
    if crate::safari::request_script_bytes("CreateItem", &request) > 1_000_000 {
        return Err(OutlookWebError::InvalidInput(
            "serialized attachment draft exceeds Cider's 1000000-byte script limit; reduce body or attachment size",
        ));
    }
    Ok(request)
}

pub fn attachment_read_request(id: &str) -> Value {
    json!({"__type":"GetAttachmentJsonRequest:#Exchange","Header":header(),"Body":{"__type":"GetAttachmentRequest:#Exchange",
        "AttachmentIds":[{"__type":"AttachmentId:#Exchange","Id":id}]}})
}

fn verification(field: &'static str) -> OutlookWebError {
    OutlookWebError::AttachmentVerificationFailed { field }
}

fn omitted_false(value: Option<&Value>) -> bool {
    matches!(value, None | Some(Value::Bool(false)))
}

pub fn send_attachment_draft_request(id: &str, key: &str) -> Result<Value, OutlookWebError> {
    if id.trim().is_empty() || key.trim().is_empty() {
        return Err(OutlookWebError::InvalidInput(
            "sending attachment draft requires its ID and change key",
        ));
    }
    Ok(
        json!({"__type":"SendItemJsonRequest:#Exchange","Header":header(),
        "Body":{"__type":"SendItemRequest:#Exchange","SaveItemToFolder":true,
            "ItemIds":[{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key}],
            "SavedItemFolderId":{"__type":"TargetFolderId:#Exchange","BaseFolderId":{
                "__type":"DistinguishedFolderId:#Exchange","Id":"sentitems"}}}}),
    )
}

impl OutlookWebSession {
    pub fn send_mail_with_attachments(
        &self,
        input: &SendMailInput,
        attachments: &[MailAttachment],
    ) -> Result<Value, OutlookWebError> {
        let request = attachment_draft_request(input, attachments)?;
        let created = self.execute("CreateItem", &request, true)?;
        let id = single_item(&created)
            .ok()
            .and_then(|item| item.pointer("/ItemId/Id"))
            .and_then(Value::as_str)
            .filter(|id| !id.trim().is_empty())
            .ok_or(OutlookWebError::OutcomeUnknown {
                action: "CreateItem attachment draft",
            })?;
        let partial = |stage, source| OutlookWebError::AttachmentSendFailed {
            draft_id: id.to_owned(),
            stage,
            source: Box::new(source),
        };
        let verified = self
            .verify_attachment_draft(id, input, attachments)
            .map_err(|error| partial("verification; send not attempted", error))?;
        let result = self
            .execute("SendItem", &verified, true)
            .and_then(|result| crate::operations::send_acknowledgement(&result, "SendItem"))
            .map_err(|error| partial("send; submission may be uncertain", error))?;
        Ok(
            json!({"status":"submission_acknowledged","acknowledged":true,
            "delivery_verified":false,"sent_copy_verified":false,"attachment_count":attachments.len(),
            "attachment_bytes_verified":true,"draft_id":id,"submission":result}),
        )
    }

    fn verify_attachment_draft(
        &self,
        id: &str,
        input: &SendMailInput,
        attachments: &[MailAttachment],
    ) -> Result<Value, OutlookWebError> {
        let response = self.mail_get(id)?;
        let draft = single_item(&response).map_err(|_| verification("draft response shape"))?;
        if draft["__type"] != "Message:#Exchange"
            || draft["IsDraft"] != true
            || draft["ItemId"]["Id"] != id
        {
            return Err(verification("draft identity/type/state"));
        }
        if draft["Subject"] != input.subject {
            return Err(verification("subject"));
        }
        if draft["Body"]["Value"] != input.body
            || draft["Body"]["BodyType"] != "Text"
            || !omitted_false(draft["Body"].get("IsTruncated"))
        {
            return Err(verification("plain-text body or truncation"));
        }
        for (key, expected) in [
            ("ToRecipients", &input.to),
            ("CcRecipients", &input.cc),
            ("BccRecipients", &input.bcc),
        ] {
            // OWA suppresses empty recipient collections, but null/invalid values are not empty.
            let actual = match draft.get(key) {
                None if expected.is_empty() => &[][..],
                Some(Value::Array(values)) => values.as_slice(),
                _ => return Err(verification(key)),
            };
            if actual.len() != expected.len()
                || actual
                    .iter()
                    .zip(expected)
                    .any(|(a, b)| a["EmailAddress"].as_str() != Some(b.as_str()))
            {
                return Err(verification(key));
            }
        }
        let actual = draft["Attachments"]
            .as_array()
            .ok_or(verification("draft attachment collection"))?;
        if actual.len() != attachments.len() {
            return Err(verification("attachment count"));
        }
        for expected in attachments {
            let matches: Vec<_> = actual
                .iter()
                .filter(|a| a["Name"] == expected.name)
                .collect();
            if matches.len() != 1 {
                return Err(verification("attachment name/uniqueness"));
            }
            let attachment_id = matches[0]["AttachmentId"]["Id"]
                .as_str()
                .filter(|id| !id.trim().is_empty())
                .ok_or(verification("attachment ID"))?;
            let response = self.call("GetAttachment", &attachment_read_request(attachment_id))?;
            let returned = response
                .pointer("/Body/ResponseMessages/Items")
                .and_then(Value::as_array)
                .filter(|a| a.len() == 1)
                .and_then(|a| a[0]["Attachments"].as_array())
                .filter(|a| a.len() == 1)
                .ok_or(verification("attachment response shape"))?;
            let a = &returned[0];
            if a["__type"] != "FileAttachment:#Exchange"
                || a["Name"] != expected.name
                || a["AttachmentId"]["Id"] != attachment_id
                || !omitted_false(a.get("IsInline"))
            {
                return Err(verification("attachment identity/type/inline state"));
            }
            if a["Content"] != STANDARD.encode(&expected.bytes) {
                return Err(verification("attachment bytes"));
            }
        }
        let key = draft["ItemId"]["ChangeKey"]
            .as_str()
            .ok_or(verification("draft change key"))?;
        send_attachment_draft_request(id, key)
    }
}
#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::{message, mock};

    fn input() -> SendMailInput {
        SendMailInput {
            to: vec!["to@example.com".into()],
            cc: vec![],
            bcc: vec![],
            subject: "Subject".into(),
            body: "<literal>\nText".into(),
        }
    }

    fn files() -> Vec<MailAttachment> {
        vec![
            MailAttachment::new("avatar.png".into(), (0..=255).collect()).unwrap(),
            MailAttachment::new("empty.txt".into(), vec![]).unwrap(),
        ]
    }

    fn draft() -> Value {
        let mut draft =
            attachment_draft_request(&input(), &files()).unwrap()["Body"]["Items"][0].clone();
        draft["ItemId"] = json!({"Id":"draft","ChangeKey":"verified-key"});
        draft["IsDraft"] = json!(true);
        for (i, a) in draft["Attachments"]
            .as_array_mut()
            .unwrap()
            .iter_mut()
            .enumerate()
        {
            a["AttachmentId"] = json!({"Id":format!("attachment-{i}")});
        }
        draft
    }

    fn replies() -> Vec<(&'static str, Value)> {
        let draft = draft();
        vec![
            ("CreateItem", message(json!({"ItemId":{"Id":"draft"}}))),
            ("GetItem", message(draft.clone())),
            (
                "GetAttachment",
                json!({"Body":{"ResponseMessages":{"Items":[{
                    "ResponseCode":"NoError","Attachments":[draft["Attachments"][0].clone()]
                }]}}}),
            ),
            (
                "GetAttachment",
                json!({"Body":{"ResponseMessages":{"Items":[{
                    "ResponseCode":"NoError","Attachments":[draft["Attachments"][1].clone()]
                }]}}}),
            ),
            (
                "SendItem",
                json!({"Body":{"ResponseMessages":{"Items":[{
                    "ResponseCode":"NoError","ResponseClass":"Success","Items":[]
                }]}}}),
            ),
        ]
    }

    #[test]
    fn live_owa_shape_omits_empty_recipients_and_false_inline_without_losing_byte_verification() {
        let mut responses = replies();
        let draft = &mut responses[1].1["Body"]["ResponseMessages"]["Items"][0]["Items"][0];
        draft.as_object_mut().unwrap().remove("CcRecipients");
        draft.as_object_mut().unwrap().remove("BccRecipients");
        draft["Body"]["IsTruncated"] = json!(false);
        for index in [2, 3] {
            let a =
                &mut responses[index].1["Body"]["ResponseMessages"]["Items"][0]["Attachments"][0];
            a.as_object_mut().unwrap().remove("IsInline");
            a.as_object_mut().unwrap().remove("IsContactPhoto");
            a["AttachmentOriginalUrl"] = json!("");
        }
        let (session, requests, worker) = mock(responses);
        assert_eq!(
            session
                .send_mail_with_attachments(&input(), &files())
                .unwrap()["attachment_bytes_verified"],
            true
        );
        worker.join().unwrap();
        let requests = requests.lock().unwrap();
        for request in &requests[2..4] {
            assert_eq!(request["__type"], "GetAttachmentJsonRequest:#Exchange");
            assert!(request["Body"].get("AttachmentShape").is_none());
            assert_eq!(request["Body"]["__type"], "GetAttachmentRequest:#Exchange");
        }
    }

    #[test]
    fn omitted_defaults_never_accept_null_malformed_nonempty_or_inline_values() {
        for key in ["ToRecipients", "CcRecipients", "BccRecipients"] {
            for value in [
                Value::Null,
                json!({}),
                json!(""),
                json!([{"EmailAddress":"unexpected@example.com"}]),
            ] {
                let mut responses = replies();
                responses[1].1["Body"]["ResponseMessages"]["Items"][0]["Items"][0][key] = value;
                responses.truncate(2);
                let (session, _, worker) = mock(responses);
                let error = session
                    .send_mail_with_attachments(&input(), &files())
                    .unwrap_err()
                    .to_string();
                assert!(error.contains(key), "{error}");
                worker.join().unwrap();
            }
        }
        for value in [Value::Null, json!(true), json!("false"), json!(0)] {
            let mut responses = replies();
            responses[2].1["Body"]["ResponseMessages"]["Items"][0]["Attachments"][0]["IsInline"] =
                value;
            responses.truncate(3);
            let (session, _, worker) = mock(responses);
            let error = session
                .send_mail_with_attachments(&input(), &files())
                .unwrap_err()
                .to_string();
            assert!(error.contains("inline state"), "{error}");
            worker.join().unwrap();
        }
        let mut responses = replies();
        let a = &mut responses[2].1["Body"]["ResponseMessages"]["Items"][0]["Attachments"][0];
        a.as_object_mut().unwrap().remove("IsInline");
        a["Content"] = json!("Y2hhbmdlZA==");
        responses.truncate(3);
        let (session, _, worker) = mock(responses);
        let error = session
            .send_mail_with_attachments(&input(), &files())
            .unwrap_err()
            .to_string();
        assert!(error.contains("attachment bytes"), "{error}");
        worker.join().unwrap();
        let mut expected = input();
        expected.cc.push("cc@example.com".into());
        let mut responses = replies();
        responses[1].1["Body"]["ResponseMessages"]["Items"][0]["Items"][0]
            .as_object_mut()
            .unwrap()
            .remove("CcRecipients");
        responses.truncate(2);
        let (session, _, worker) = mock(responses);
        assert!(
            session
                .send_mail_with_attachments(&expected, &files())
                .unwrap_err()
                .to_string()
                .contains("CcRecipients")
        );
        worker.join().unwrap();
    }

    #[test]
    fn attachment_send_saves_verifies_exact_binary_and_empty_file_then_sends_versioned_draft() {
        let (session, requests, worker) = mock(replies());
        let result = session
            .send_mail_with_attachments(&input(), &files())
            .unwrap();
        worker.join().unwrap();
        assert_eq!(result["attachment_count"], 2);
        assert_eq!(result["attachment_bytes_verified"], true);
        assert_eq!(result["delivery_verified"], false);
        assert_eq!(result["sent_copy_verified"], false);
        let requests = requests.lock().unwrap();
        assert_eq!(requests.len(), 5);
        assert_eq!(requests[0]["Body"]["MessageDisposition"], "SaveOnly");
        assert_eq!(
            requests[0]["Body"]["Items"][0]["Attachments"][0]["Content"],
            STANDARD.encode((0..=255).collect::<Vec<u8>>())
        );
        assert_eq!(
            requests[4]["Body"]["ItemIds"][0],
            json!({
                "__type":"ItemId:#Exchange","Id":"draft","ChangeKey":"verified-key"
            })
        );
        assert_eq!(requests[4]["Body"]["SaveItemToFolder"], true);
        assert_eq!(
            requests[4]["Body"]["SavedItemFolderId"]["BaseFolderId"]["Id"],
            "sentitems"
        );
    }

    #[test]
    fn attachment_workflow_stops_at_each_failed_stage_without_retry_or_cleanup() {
        for stage in 0..5 {
            for failure in [
                json!({}),
                json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"ErrorAccessDenied"}]}}}),
                json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"NoError"},{"ResponseCode":"NoError"}]}}}),
            ] {
                let mut responses = replies();
                responses[stage].1 = failure;
                responses.truncate(stage + 1);
                let (session, requests, worker) = mock(responses);
                let error = session
                    .send_mail_with_attachments(&input(), &files())
                    .unwrap_err()
                    .to_string();
                if stage > 0 {
                    assert!(error.contains("draft ID draft"), "{error}");
                    assert!(error.contains(if stage == 4 {
                        "submission may be uncertain"
                    } else {
                        "send not attempted"
                    }));
                }
                worker.join().unwrap();
                assert_eq!(requests.lock().unwrap().len(), stage + 1);
            }
        }
    }

    #[test]
    fn verification_rejects_changed_recipients_body_files_ids_or_bytes_before_send() {
        for (pointer, value, reads) in [
            ("/__type", json!("CalendarItem:#Exchange"), 2),
            ("/Subject", json!("Changed"), 2),
            ("/Body/Value", json!("Changed"), 2),
            (
                "/ToRecipients/0/EmailAddress",
                json!("other@example.com"),
                2,
            ),
            (
                "/CcRecipients",
                json!([{"EmailAddress":"unexpected@example.com"}]),
                2,
            ),
            ("/IsDraft", json!(false), 2),
            ("/ItemId/Id", json!("different"), 2),
            ("/ItemId/ChangeKey", json!(""), 4),
            ("/Attachments/0/Name", json!("different.png"), 2),
            ("/Attachments/0/Content", json!("Y2hhbmdlZA=="), 3),
            ("/Attachments/0/IsInline", json!(true), 3),
            ("/Attachments/0/AttachmentId/Id", json!(""), 2),
        ] {
            let mut changed = draft();
            *changed.pointer_mut(pointer).unwrap() = value;
            let mut responses = replies();
            responses[1].1 = message(changed.clone());
            if reads == 3 {
                responses[2].1["Body"]["ResponseMessages"]["Items"][0]["Attachments"][0] =
                    changed["Attachments"][0].clone();
            }
            responses.truncate(reads);
            let (session, requests, worker) = mock(responses);
            assert!(
                session
                    .send_mail_with_attachments(&input(), &files())
                    .is_err(),
                "{pointer}"
            );
            worker.join().unwrap();
            assert_eq!(requests.lock().unwrap().len(), reads);
        }
    }

    #[test]
    fn bounded_payload_includes_base64_json_and_browser_wrapper_overhead() {
        let large = MailAttachment::new("document.pdf".into(), vec![255; 500_000]).unwrap();
        let request = attachment_draft_request(&input(), &[large]).unwrap();
        assert_eq!(
            request["Body"]["Items"][0]["Attachments"][0]["Content"]
                .as_str()
                .unwrap()
                .len(),
            666_668
        );
        let size = crate::safari::request_script_bytes("CreateItem", &request);
        assert!(size > 666_668 && size < 1_000_000, "{size}");
        let mut empty = request.clone();
        empty["Body"]["Items"][0]["Body"]["Value"] = json!("");
        let available = 1_000_000 - crate::safari::request_script_bytes("CreateItem", &empty);
        let mut exact = input();
        exact.body = "\u{1}".repeat(available / 6) + &"x".repeat(available % 6);
        let boundary = [MailAttachment::new("document.pdf".into(), vec![255; 500_000]).unwrap()];
        let request = attachment_draft_request(&exact, &boundary).unwrap();
        assert_eq!(
            crate::safari::request_script_bytes("CreateItem", &request),
            1_000_000
        );
        exact.body.push('x');
        assert!(attachment_draft_request(&exact, &boundary).is_err());
        assert!(MailAttachment::new("over.pdf".into(), vec![0; 500_001]).is_err());
        let too_much = vec![
            MailAttachment::new("a".into(), vec![0; 250_001]).unwrap(),
            MailAttachment::new("b".into(), vec![0; 250_000]).unwrap(),
        ];
        assert!(attachment_draft_request(&input(), &too_much).is_err());
        let mut text = input();
        text.body = "\u{1}".repeat(200_000);
        assert!(attachment_draft_request(&text, &files()).is_err());
        text.body = "x".repeat(200_000);
        assert!(
            attachment_draft_request(
                &text,
                &[MailAttachment::new("file".into(), vec![255; 500_000]).unwrap()]
            )
            .is_ok()
        );
        for name in [
            "",
            ".",
            "..",
            "/etc/passwd",
            "a\\b",
            "a\nb",
            "a\u{2028}b",
            &"é".repeat(128),
        ] {
            assert!(MailAttachment::new(name.into(), vec![]).is_err());
        }
        let too_many = (0..6)
            .map(|n| MailAttachment::new(format!("a{n}"), vec![]).unwrap())
            .collect::<Vec<_>>();
        assert!(attachment_draft_request(&input(), &too_many).is_err());
        let duplicate = vec![
            MailAttachment::new("A.txt".into(), vec![]).unwrap(),
            MailAttachment::new("a.txt".into(), vec![]).unwrap(),
        ];
        assert!(attachment_draft_request(&input(), &duplicate).is_err());
    }

    #[test]
    fn local_files_are_bounded_regular_snapshots_not_links_directories_or_devices() {
        let path = std::env::current_dir()
            .unwrap()
            .join("target")
            .join(format!("attachment-files-{}", std::process::id()));
        std::fs::create_dir_all(&path).unwrap();
        let file = path.join("avatar.png");
        std::fs::write(&file, vec![255; 500_000]).unwrap();
        let loaded = load_mail_attachments(std::slice::from_ref(&file)).unwrap();
        assert_eq!(loaded[0].bytes.len(), 500_000);
        std::fs::write(&file, b"changed").unwrap();
        assert_eq!(loaded[0].bytes.len(), 500_000);
        assert!(load_mail_attachments(std::slice::from_ref(&path)).is_err());
        assert!(load_mail_attachments(&[path.join("missing")]).is_err());
        std::fs::write(&file, vec![0; 500_001]).unwrap();
        assert!(load_mail_attachments(std::slice::from_ref(&file)).is_err());
        #[cfg(unix)]
        {
            let link = path.join("link");
            std::os::unix::fs::symlink(&file, &link).unwrap();
            assert!(load_mail_attachments(std::slice::from_ref(&link)).is_err());
            std::fs::remove_file(link).unwrap();
            assert!(load_mail_attachments(&[PathBuf::from("/dev/null")]).is_err());
        }
        std::fs::remove_file(file).unwrap();
        std::fs::remove_dir(path).unwrap();
    }
}
