use serde_json::{Value, json};

use crate::{OutlookWebError, OutlookWebSession, SessionTransport, get_item_request, header};

fn first(response: &Value) -> Result<&Value, OutlookWebError> {
    let messages = response
        .pointer("/Body/ResponseMessages/Items")
        .and_then(Value::as_array)
        .ok_or(OutlookWebError::InvalidResponse)?;
    if messages.len() != 1 {
        return Err(OutlookWebError::InvalidResponse);
    }
    let items = messages[0]["Items"]
        .as_array()
        .ok_or(OutlookWebError::InvalidResponse)?;
    if items.len() != 1 {
        return Err(OutlookWebError::InvalidResponse);
    }
    Ok(&items[0])
}

pub fn mail_action_item_request(id: &str) -> Result<Value, OutlookWebError> {
    if id.trim().is_empty() {
        return Err(OutlookWebError::InvalidInput(
            "message ID must not be empty",
        ));
    }
    let mut request = get_item_request(id);
    request["Body"]["ItemShape"]["BaseShape"] = json!("IdOnly");
    request["Body"]["ItemShape"]["AdditionalProperties"] = json!([
        {"__type":"PropertyUri:#Exchange","FieldURI":"Subject"},
        {"__type":"PropertyUri:#Exchange","FieldURI":"Flag"},
        {"__type":"PropertyUri:#Exchange","FieldURI":"ReceivedOrRenewTime"},
        {"__type":"PropertyUri:#Exchange","FieldURI":"IsDraft"},
        {"__type":"PropertyUri:#Exchange","FieldURI":"IsRead"}
    ]);
    Ok(request)
}

pub fn native_pin_state(item: &Value) -> Result<bool, OutlookWebError> {
    let value = item["ReceivedOrRenewTime"]
        .as_str()
        .ok_or(OutlookWebError::InvalidResponse)?;
    let date = chrono::DateTime::parse_from_rfc3339(value)
        .map_err(|_| OutlookWebError::InvalidResponse)?;
    Ok(chrono::Datelike::year(&date) == 4500)
}

pub fn flag_request(id: &str, change_key: &str, flagged: bool) -> Result<Value, OutlookWebError> {
    if id.trim().is_empty() || change_key.trim().is_empty() {
        return Err(OutlookWebError::InvalidInput(
            "flag requires a versioned message ID",
        ));
    }
    Ok(
        json!({"__type":"UpdateItemJsonRequest:#Exchange","Header":header(),"Body":{
            "__type":"UpdateItemRequest:#Exchange","MessageDisposition":"SaveOnly",
            "ConflictResolution":"NeverOverwrite","SuppressReadReceipts":true,
            "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
            "ItemChanges":[{"__type":"ItemChange:#Exchange",
                "ItemId":{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":change_key},
                "Updates":[{"__type":"SetItemField:#Exchange",
                    "Path":{"__type":"PropertyUri:#Exchange","FieldURI":"Flag"},
                    "Item":{"__type":"Message:#Exchange","Flag":{"__type":"FlagType:#Exchange",
                        "FlagStatus":if flagged {"Flagged"} else {"NotFlagged"}}}}]}]
        }}),
    )
}

impl OutlookWebSession {
    /// Exact desired state, not a toggle; never retry an uncertain write.
    pub fn set_mail_flag(
        &self,
        id: &str,
        key: &str,
        flagged: bool,
    ) -> Result<Value, OutlookWebError> {
        self.set_mail_action(id, key, flagged, false)
    }

    pub fn set_mail_pin(
        &self,
        id: &str,
        key: &str,
        pinned: bool,
    ) -> Result<Value, OutlookWebError> {
        self.set_mail_action(id, key, pinned, true)
    }

    fn set_mail_action(
        &self,
        id: &str,
        key: &str,
        desired: bool,
        pin: bool,
    ) -> Result<Value, OutlookWebError> {
        let request = mail_action_item_request(id)?;
        let before = self.call("GetItem", &request)?;
        let item = first(&before)?;
        if !crate::operations::is_mail_message(item)
            || item["IsDraft"] != false
            || item["ItemId"]["Id"] != id
            || item["ItemId"]["ChangeKey"] != key
            || item["IsRead"].as_bool().is_none()
        {
            return Err(OutlookWebError::InvalidInput(
                "message changed or is not a non-draft mail message; reload before acting",
            ));
        }
        let state = |item: &Value| -> Result<bool, OutlookWebError> {
            if pin {
                native_pin_state(item)
            } else {
                match item["Flag"]["FlagStatus"].as_str() {
                    Some("Flagged") => Ok(true),
                    Some("NotFlagged" | "Complete") => Ok(false),
                    _ => Err(OutlookWebError::InvalidResponse),
                }
            }
        };
        if state(item)? == desired {
            return Ok(
                json!({"ok":true,"id":id,"change_key":key,"changed":false,"verified":true,"value":desired}),
            );
        }
        if pin {
            match &self.transport {
                SessionTransport::Safari(session) => session.set_pin(id, desired)?,
                #[cfg(test)]
                SessionTransport::Mock(session) => {
                    session.call("PinItem", &json!({"id":id,"pinned":desired}));
                }
            }
        } else {
            self.execute("UpdateItem", &flag_request(id, key, desired)?, true)?;
        }
        let after = self
            .call("GetItem", &request)
            .map_err(|_| OutlookWebError::VerificationFailed)?;
        let updated = first(&after).map_err(|_| OutlookWebError::VerificationFailed)?;
        if updated["ItemId"]["Id"] != id
            || state(updated).ok() != Some(desired)
            || updated["IsRead"] != item["IsRead"]
        {
            return Err(OutlookWebError::VerificationFailed);
        }
        let new_key = updated["ItemId"]["ChangeKey"]
            .as_str()
            .filter(|s| !s.is_empty())
            .ok_or(OutlookWebError::VerificationFailed)?;
        Ok(
            json!({"ok":true,"id":id,"change_key":new_key,"changed":true,"verified":true,"value":desired}),
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::{message, mock};

    fn item(flag: bool, pin: bool, key: &str) -> Value {
        json!({"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":key},
            "IsRead":false,"IsDraft":false,"Flag":{"FlagStatus":if flag {"Flagged"} else {"NotFlagged"}},
            "ReceivedOrRenewTime":if pin {"4500-09-01T00:00:00Z"} else {"2026-10-21T00:00:00Z"}})
    }

    #[test]
    fn desired_flag_and_native_pin_are_verified_without_changing_read_state() {
        for pin in [false, true] {
            let (session, requests, completion) = mock(vec![
                ("GetItem", message(item(false, false, "key"))),
                (
                    if pin { "PinItem" } else { "UpdateItem" },
                    message(json!({})),
                ),
                ("GetItem", message(item(!pin, pin, "new"))),
            ]);
            let result = if pin {
                session.set_mail_pin("id", "key", true)
            } else {
                session.set_mail_flag("id", "key", true)
            }
            .unwrap();
            assert_eq!(result["verified"], true);
            assert_eq!(result["change_key"], "new");
            assert_eq!(requests.lock().unwrap().len(), 3);
            completion.join().unwrap();
        }
    }

    #[test]
    fn stale_keys_do_not_write_and_failed_readback_does_not_retry() {
        let (session, _, completion) = mock(vec![("GetItem", message(item(false, false, "new")))]);
        assert!(session.set_mail_flag("id", "old", true).is_err());
        completion.join().unwrap();
        let (session, requests, completion) = mock(vec![
            ("GetItem", message(item(false, false, "key"))),
            ("UpdateItem", message(json!({}))),
            ("GetItem", message(item(false, false, "new"))),
        ]);
        assert!(matches!(
            session.set_mail_flag("id", "key", true),
            Err(OutlookWebError::VerificationFailed)
        ));
        assert_eq!(requests.lock().unwrap().len(), 3);
        completion.join().unwrap();
    }

    #[test]
    fn clearing_states_and_noops_are_explicit_and_verified() {
        for pin in [false, true] {
            let (session, _, completion) = mock(vec![
                ("GetItem", message(item(true, true, "key"))),
                (
                    if pin { "PinItem" } else { "UpdateItem" },
                    message(json!({})),
                ),
                ("GetItem", message(item(pin, !pin, "new"))),
            ]);
            let result = if pin {
                session.set_mail_pin("id", "key", false)
            } else {
                session.set_mail_flag("id", "key", false)
            }
            .unwrap();
            assert_eq!(result["value"], false);
            completion.join().unwrap();
            let (session, requests, completion) =
                mock(vec![("GetItem", message(item(false, false, "key")))]);
            assert_eq!(
                session.set_mail_pin("id", "key", false).unwrap()["changed"],
                false
            );
            assert_eq!(requests.lock().unwrap().len(), 1);
            completion.join().unwrap();
        }
    }

    #[test]
    fn drafts_unknown_read_state_and_read_state_changes_are_not_success() {
        for field in ["IsDraft", "IsRead"] {
            let mut before = item(false, false, "key");
            before[field] = if field == "IsDraft" {
                json!(true)
            } else {
                Value::Null
            };
            let (session, requests, completion) = mock(vec![("GetItem", message(before))]);
            assert!(session.set_mail_flag("id", "key", true).is_err());
            assert_eq!(requests.lock().unwrap().len(), 1);
            completion.join().unwrap();
        }
        let mut after = item(true, false, "new");
        after["IsRead"] = json!(true);
        let (session, _, completion) = mock(vec![
            ("GetItem", message(item(false, false, "key"))),
            ("UpdateItem", message(json!({}))),
            ("GetItem", message(after)),
        ]);
        assert!(matches!(
            session.set_mail_flag("id", "key", true),
            Err(OutlookWebError::VerificationFailed)
        ));
        completion.join().unwrap();
    }
}
