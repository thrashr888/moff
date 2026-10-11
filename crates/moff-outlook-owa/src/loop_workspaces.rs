use crate::{OutlookWebError, OutlookWebSession, SessionTransport};
use serde_json::{Value, json};

/// Native Loop containers, distinct from drive-backed `.loop` files.
pub enum LoopWorkspaceRequest<'a> {
    List {
        query: Option<&'a str>,
        limit: u16,
    },
    Get {
        id: &'a str,
    },
    Create,
    Delete {
        id: &'a str,
        mru_id: Option<&'a str>,
    },
}

fn pod_id(id: &str) -> Result<(), OutlookWebError> {
    if id.is_empty()
        || id.len() > 4096
        || !id
            .bytes()
            .all(|b| b.is_ascii_alphanumeric() || b"+/=".contains(&b))
    {
        return Err(OutlookWebError::InvalidInput(
            "workspace ID must be a base64 PodId, not an SPO_ or drive ID",
        ));
    }
    Ok(())
}

impl OutlookWebSession {
    pub fn loop_workspace(
        &self,
        request: &LoopWorkspaceRequest<'_>,
    ) -> Result<Value, OutlookWebError> {
        let (config, mutation) = match request {
            LoopWorkspaceRequest::List { query, limit } => {
                if !(1..=100).contains(limit)
                    || query.is_some_and(|q| q.trim().is_empty() || q.len() > 4096)
                {
                    return Err(OutlookWebError::InvalidInput(
                        "workspace query must be nonempty and limit must be 1..=100",
                    ));
                }
                (json!({"action":"list","query":query,"limit":limit}), false)
            }
            LoopWorkspaceRequest::Get { id } => {
                pod_id(id)?;
                (json!({"action":"get","id":id}), false)
            }
            LoopWorkspaceRequest::Create => {
                (json!({"action":"create","allow_mutation":true}), true)
            }
            LoopWorkspaceRequest::Delete { id, mru_id } => {
                pod_id(id)?;
                if mru_id.is_some_and(|id| {
                    id.len() != 36 || !id.bytes().all(|b| b.is_ascii_hexdigit() || b == b'-')
                }) {
                    return Err(OutlookWebError::InvalidInput("mru-id must be a GUID"));
                }
                (
                    json!({"action":"delete","id":id,"mru_id":mru_id,"allow_mutation":true}),
                    true,
                )
            }
        };
        let value = match &self.transport {
            SessionTransport::Safari(safari) => {
                safari.page_request(include_str!("safari_loop.js"), &config, mutation)?
            }
            #[cfg(test)]
            SessionTransport::Mock(_) => return Err(OutlookWebError::InvalidAction),
        };
        if value["verification_failed"] == true {
            return Err(OutlookWebError::ItemVerificationFailed {
                id: value["drive_id"]
                    .as_str()
                    .ok_or(OutlookWebError::InvalidResponse)?
                    .to_owned(),
            });
        }
        Ok(value)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn workspace_ids_are_not_drive_or_recent_ids() {
        assert!(pod_id("T0RTUHx0ZW5hbnQuc2hhcmVwb2ludC5jb218YiFk cml2ZXxyb290").is_err());
        for id in ["", "SPO_123", "b!drive", "../root", "x\n"] {
            assert!(pod_id(id).is_err());
        }
        assert!(pod_id("T0RTUHx0ZW5hbnQuc2hhcmVwb2ludC5jb218YiFkcml2ZXxyb290").is_ok());
    }
}
