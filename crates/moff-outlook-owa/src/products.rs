use crate::{OutlookWebError, OutlookWebSession};
use serde_json::{Value, json};
use url::Url;

pub(crate) fn product_url(path: &str) -> Result<Url, OutlookWebError> {
    let reject = || OutlookWebError::InvalidInput("unsupported product read path or query");
    let raw_path = path.split_once('?').map_or(path, |(path, _)| path);
    let lower = raw_path.to_ascii_lowercase();
    if !path.starts_with("/v1.0/")
        || lower.contains("%2f")
        || lower.contains("%5c")
        || lower.contains("%25")
        || raw_path.contains("//")
        || raw_path.ends_with('/')
        || raw_path.contains('\\')
        || path.contains('#')
        || raw_path.split('/').any(|segment| {
            matches!(
                segment.to_ascii_lowercase().as_str(),
                "." | ".." | "%2e" | ".%2e" | "%2e." | "%2e%2e"
            )
        })
    {
        return Err(reject());
    }

    let url = Url::parse(&format!("https://graph.microsoft.com{path}")).map_err(|_| reject())?;
    let segments: Vec<_> = url
        .path()
        .split('/')
        .filter(|segment| !segment.is_empty())
        .collect();
    let paged = matches!(
        segments.as_slice(),
        ["v1.0", "me", "people" | "contacts"]
            | ["v1.0", "groups", _, "owners" | "members"]
            | ["v1.0", "groups" | "users"]
            | [
                "v1.0",
                "me",
                "drive",
                "root",
                "children" | "search(q='.loop')"
            ]
            | ["v1.0", "me", "drive", "items", _, "children"]
    ) || matches!(segments.as_slice(), ["v1.0","me","drive","root", search] if search.starts_with("search(q='") && search.ends_with("')"));
    let user = matches!(segments.as_slice(), ["v1.0", "users", _]);
    let allowed = paged
        || user
        || matches!(
            segments.as_slice(),
            ["v1.0", "me"]
                | ["v1.0", "me", "contacts", _]
                | ["v1.0", "groups", _]
                | ["v1.0", "me", "manager" | "directReports"]
                | ["v1.0", "users", _, "manager" | "directReports"]
                | ["v1.0", "me", "drive", "items", _]
        );
    if !allowed {
        return Err(reject());
    }
    let mut keys = std::collections::HashSet::new();
    for (key, value) in url.query_pairs() {
        if !keys.insert(key.to_string()) {
            return Err(reject());
        }
        let valid = match key.as_ref() {
            "$top" if paged => {
                value.bytes().all(|byte| byte.is_ascii_digit())
                    && value.parse::<u16>().is_ok_and(|n| (1..=100).contains(&n))
            }
            "$select" if user => value == "id,displayName,mail,userPrincipalName,userType",
            "$filter" if segments.as_slice() == ["v1.0", "groups"] => {
                value == "groupTypes/any(c:c eq 'Unified')"
            }
            "$skiptoken" if paged => !value.is_empty() && value.len() <= 4096,
            "$skip" if paged => {
                !value.is_empty()
                    && value.bytes().all(|b| b.is_ascii_digit())
                    && value.parse::<u32>().is_ok()
            }
            _ => false,
        };
        if !valid {
            return Err(reject());
        }
    }
    Ok(url)
}

impl OutlookWebSession {
    /// Read allowlisted product metadata inside Safari; never export authentication.
    pub fn product_read(&self, path: &str) -> Result<serde_json::Value, OutlookWebError> {
        let url = product_url(path)?;
        match &self.transport {
            crate::SessionTransport::Safari(safari) => safari.product_read(url.as_str()),
            #[cfg(test)]
            crate::SessionTransport::Mock(_) => Err(OutlookWebError::InvalidAction),
        }
    }

    pub fn product_write(&self, request: &ProductWrite) -> Result<Value, OutlookWebError> {
        match &self.transport {
            crate::SessionTransport::Safari(safari) => safari.product_write(request),
            #[cfg(test)]
            crate::SessionTransport::Mock(_) => Err(OutlookWebError::InvalidAction),
        }
    }

    pub fn drive_download(&self, id: &str) -> Result<Vec<u8>, OutlookWebError> {
        let url = format!(
            "https://graph.microsoft.com/v1.0/me/drive/items/{}",
            product_segment(id)?
        );
        let value = match &self.transport {
            crate::SessionTransport::Safari(safari) => safari.product_download(&url)?,
            #[cfg(test)]
            crate::SessionTransport::Mock(_) => return Err(OutlookWebError::InvalidAction),
        };
        let bytes: Vec<u8> = serde_json::from_value(value["bytes"].clone())
            .map_err(|_| OutlookWebError::InvalidResponse)?;
        if bytes.len() > 128_000 {
            return Err(invalid(
                "browser file transfers are limited to 128000 bytes",
            ));
        }
        Ok(bytes)
    }
}

/// A validated, personal-content mutation. No caller-supplied URLs or headers.
#[derive(Debug)]
pub struct ProductWrite {
    pub(crate) url: String,
    pub(crate) method: &'static str,
    pub(crate) body: Option<Value>,
    pub(crate) etag: Option<String>,
    pub(crate) bytes: Option<Vec<u8>>,
}

fn invalid(message: &'static str) -> OutlookWebError {
    OutlookWebError::InvalidInput(message)
}

pub fn product_segment(value: &str) -> Result<String, OutlookWebError> {
    if value.trim().is_empty()
        || matches!(value, "." | "..")
        || value.contains(['/', '\\', '%'])
        || value.chars().any(char::is_control)
    {
        return Err(invalid(
            "product IDs/names must be nonempty single path segments",
        ));
    }
    Ok(url::form_urlencoded::byte_serialize(value.as_bytes())
        .collect::<String>()
        .replace('+', "%20"))
}

fn object_fields<'a>(
    body: &'a Value,
    allowed: &[&str],
) -> Result<&'a serde_json::Map<String, Value>, OutlookWebError> {
    let object = body
        .as_object()
        .filter(|o| !o.is_empty())
        .ok_or(invalid("mutation input must be a nonempty JSON object"))?;
    if object.keys().any(|key| !allowed.contains(&key.as_str())) {
        return Err(invalid("mutation input contains an unsupported field"));
    }
    Ok(object)
}

fn string_fields(body: &Value, allowed: &[&str]) -> Result<(), OutlookWebError> {
    let object = object_fields(body, allowed)?;
    if object.values().any(|value| !value.is_string()) {
        return Err(invalid("mutation fields must be strings"));
    }
    Ok(())
}

impl ProductWrite {
    pub fn drive_upload(
        parent: Option<&str>,
        id: Option<&str>,
        name: Option<&str>,
        bytes: Vec<u8>,
        etag: Option<&str>,
    ) -> Result<Self, OutlookWebError> {
        if bytes.len() > 128_000 {
            return Err(invalid(
                "browser file transfers are limited to 128000 bytes",
            ));
        }
        let path = if let Some(id) = id {
            if parent.is_some() || name.is_some() || etag.is_none() {
                return Err(invalid(
                    "file replacement requires an ID and current entity tag",
                ));
            }
            format!("/v1.0/me/drive/items/{}/content", product_segment(id)?)
        } else {
            let name = product_segment(name.ok_or(invalid("file creation requires a name"))?)?;
            let base = parent.map_or_else(
                || Ok("root".to_owned()),
                |id| Ok(format!("items/{}", product_segment(id)?)),
            )?;
            format!("/v1.0/me/drive/{base}:/{name}:/content?@microsoft.graph.conflictBehavior=fail")
        };
        let mut request = Self::new(path, "PUT", None, etag)?;
        request.bytes = Some(bytes);
        Ok(request)
    }

    fn new(
        path: impl AsRef<str>,
        method: &'static str,
        body: Option<Value>,
        etag: Option<&str>,
    ) -> Result<Self, OutlookWebError> {
        if body
            .as_ref()
            .is_some_and(|value| value.to_string().len() > 200_000)
        {
            return Err(invalid("product mutation body exceeds 200000 bytes"));
        }
        if etag.is_some_and(|value| value.is_empty() || value.contains(['\r', '\n'])) {
            return Err(invalid("invalid entity tag"));
        }
        Ok(Self {
            url: format!("https://graph.microsoft.com{}", path.as_ref()),
            method,
            body,
            etag: etag.map(str::to_owned),
            bytes: None,
        })
    }

    pub fn contact(
        id: Option<&str>,
        method: &str,
        body: Option<Value>,
        etag: Option<&str>,
    ) -> Result<Self, OutlookWebError> {
        let path = id.map_or_else(
            || Ok("/v1.0/me/contacts".to_owned()),
            |id| Ok(format!("/v1.0/me/contacts/{}", product_segment(id)?)),
        )?;
        let method = match (id, method) {
            (None, "POST") => "POST",
            (Some(_), "PATCH") => "PATCH",
            (Some(_), "DELETE") => "DELETE",
            _ => return Err(invalid("unsupported contact operation")),
        };
        if method == "DELETE" {
            if body.is_some() {
                return Err(invalid("delete must not have a body"));
            }
        } else {
            let body = body.as_ref().ok_or(invalid("contact body required"))?;
            let object = object_fields(
                body,
                &[
                    "givenName",
                    "surname",
                    "displayName",
                    "companyName",
                    "jobTitle",
                    "department",
                    "officeLocation",
                    "mobilePhone",
                    "businessPhones",
                    "homePhones",
                    "emailAddresses",
                    "personalNotes",
                ],
            )?;
            for (key, value) in object {
                match key.as_str() {
                    "businessPhones" | "homePhones" => {
                        if !value
                            .as_array()
                            .is_some_and(|a| a.iter().all(Value::is_string))
                        {
                            return Err(invalid("phone fields must be arrays of strings"));
                        }
                    }
                    "emailAddresses" => {
                        let addresses = value
                            .as_array()
                            .ok_or(invalid("emailAddresses must be an array"))?;
                        for address in addresses {
                            string_fields(address, &["name", "address"])?;
                            if !address["address"].as_str().is_some_and(|s| {
                                s.contains('@') && !s.chars().any(char::is_control)
                            }) {
                                return Err(invalid("email address required"));
                            }
                        }
                    }
                    _ if value.is_string() => (),
                    _ => return Err(invalid("contact text fields must be strings")),
                }
            }
        }
        Self::new(path, method, body, etag)
    }

    pub fn drive_folder(parent: Option<&str>, name: &str) -> Result<Self, OutlookWebError> {
        product_segment(name)?;
        let base = parent.map_or_else(
            || Ok("root".to_owned()),
            |id| Ok(format!("items/{}", product_segment(id)?)),
        )?;
        Self::new(
            format!("/v1.0/me/drive/{base}/children"),
            "POST",
            Some(json!({"name":name,"folder":{},"@microsoft.graph.conflictBehavior":"fail"})),
            None,
        )
    }

    pub fn drive_item(
        id: &str,
        method: &str,
        body: Option<Value>,
        etag: &str,
    ) -> Result<Self, OutlookWebError> {
        let path = format!("/v1.0/me/drive/items/{}", product_segment(id)?);
        let method = match method {
            "PATCH" => "PATCH",
            "DELETE" => "DELETE",
            _ => return Err(invalid("unsupported drive operation")),
        };
        if method == "DELETE" {
            if body.is_some() {
                return Err(invalid("delete must not have a body"));
            }
        } else {
            let body = body.as_ref().ok_or(invalid("drive update body required"))?;
            let object = object_fields(body, &["name", "description", "parentReference"])?;
            for (key, value) in object {
                if key == "parentReference" {
                    string_fields(value, &["id"])?;
                    product_segment(
                        value["id"]
                            .as_str()
                            .ok_or(invalid("parentReference requires a folder ID"))?,
                    )?;
                } else if !value.is_string() {
                    return Err(invalid("name and description must be strings"));
                }
            }
            if let Some(name) = body.get("name") {
                product_segment(name.as_str().ok_or(invalid("name must be a string"))?)?;
            }
        }
        Self::new(path, method, body, Some(etag))
    }

    pub fn group(
        id: Option<&str>,
        method: &str,
        mut body: Option<Value>,
    ) -> Result<Self, OutlookWebError> {
        let path = id.map_or_else(
            || Ok("/v1.0/groups".to_owned()),
            |id| Ok(format!("/v1.0/groups/{}", product_segment(id)?)),
        )?;
        let method = match (id, method) {
            (None, "POST") => "POST",
            (Some(_), "PATCH") => "PATCH",
            (Some(_), "DELETE") => "DELETE",
            _ => return Err(invalid("unsupported group operation")),
        };
        if method == "DELETE" {
            if body.is_some() {
                return Err(invalid("delete must not have a body"));
            }
        } else {
            let fields = if method == "POST" {
                &["displayName", "description", "mailNickname", "visibility"][..]
            } else {
                &["displayName", "description", "visibility"][..]
            };
            let value = body.as_mut().ok_or(invalid("group body required"))?;
            string_fields(value, fields)?;
            if value.get("displayName").is_some_and(|v| {
                v.as_str()
                    .is_none_or(|s| s.trim().is_empty() || s.len() > 256)
            }) {
                return Err(invalid(
                    "displayName must be nonempty and at most 256 bytes",
                ));
            }
            if value
                .get("visibility")
                .is_some_and(|v| !matches!(v.as_str(), Some("Private" | "Public")))
            {
                return Err(invalid("visibility must be Private or Public"));
            }
            if method == "POST" {
                if value.get("displayName").is_none() {
                    return Err(invalid("group creation requires displayName"));
                }
                let alias = value["mailNickname"].as_str().filter(|s| !s.is_empty() && s.len() <= 64 && s.bytes().all(|b| b.is_ascii_alphanumeric() || b == b'-' || b == b'_'))
                    .ok_or(invalid("mailNickname must contain 1-64 ASCII letters, digits, hyphens or underscores"))?;
                let _ = alias;
                value["groupTypes"] = json!(["Unified"]);
                value["mailEnabled"] = json!(true);
                value["securityEnabled"] = json!(false);
            }
        }
        Self::new(path, method, body, None)
    }

    pub fn group_member(group: &str, user: &str, add: bool) -> Result<Self, OutlookWebError> {
        let group = product_segment(group)?;
        let user = product_segment(user)?;
        if add {
            Self::new(
                format!("/v1.0/groups/{group}/members/$ref"),
                "POST",
                Some(json!({"@odata.id":format!("https://graph.microsoft.com/v1.0/users/{user}")})),
                None,
            )
        } else {
            Self::new(
                format!("/v1.0/groups/{group}/members/{user}/$ref"),
                "DELETE",
                None,
                None,
            )
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn only_bounded_product_metadata_routes_are_allowed() {
        for path in [
            "/v1.0/me/people?$top=5",
            "/v1.0/me/contacts?$top=5&$skip=5",
            "/v1.0/me/contacts/id",
            "/v1.0/me",
            "/v1.0/groups/id/owners?$top=100",
            "/v1.0/groups/id/members?$top=5&$skiptoken=next",
            "/v1.0/me/drive/root/search(q='project')?$top=5",
            "/v1.0/me/manager",
            "/v1.0/me/directReports",
            "/v1.0/users/person%40example.com?$select=id,displayName,mail,userPrincipalName,userType",
            "/v1.0/users/person/manager",
            "/v1.0/users/person/directReports",
            "/v1.0/me/drive/root/children?$top=5",
            "/v1.0/me/drive/items/id/children?$top=5",
            "/v1.0/me/drive/items/id",
            "/v1.0/me/drive/root/search(q='.loop')?$top=5",
            "/v1.0/groups?$filter=groupTypes%2Fany%28c%3Ac%20eq%20%27Unified%27%29&$top=5",
        ] {
            assert!(product_url(path).is_ok(), "{path}");
        }
        for path in [
            "https://evil.example/v1.0/me/people",
            "//evil.example/v1.0/me/people",
            "/v1.0/me/messages",
            "/v1.0/me/events",
            "/v1.0/me/drive/items/id/content",
            "/v1.0/users/id/messages",
            "/v1.0/me/drive/items/%2fmessages",
            "/v1.0/me/drive/items/id/../root",
            "/v1.0/me/drive/items/%2e%2e/root/children",
            "/v1.0/me//people",
            "/v1.0/me/people/",
            "/v1.0/me/people?$top=%2B1",
            "/v1.0/me/people?$top=101",
            "/v1.0/me/people?$top=1&$top=2",
            "/v1.0/users/id?$select=password",
            "/v1.0/me/people?$expand=messages",
            "/v1.0/groups?$filter=anything",
            "/v1.0/me/people#fragment",
        ] {
            assert!(product_url(path).is_err(), "{path}");
        }
    }

    #[test]
    fn product_writes_are_typed_and_scoped_to_personal_objects() {
        let contact = json!({"givenName":"Synthetic","emailAddresses":[{"name":"Test","address":"test@example.com"}]});
        assert!(ProductWrite::contact(None, "POST", Some(contact), None).is_ok());
        for body in [
            json!({}),
            json!({"owners@odata.bind":[]}),
            json!({"givenName":42}),
            json!({"emailAddresses":[{"name":"No address"}]}),
            json!({"businessPhones":[42]}),
        ] {
            assert!(ProductWrite::contact(None, "POST", Some(body), None).is_err());
        }
        assert!(
            ProductWrite::contact(Some("id"), "DELETE", Some(json!({"givenName":"x"})), None)
                .is_err()
        );
        let upload =
            ProductWrite::drive_upload(None, None, Some("test.bin"), vec![0, 255], None).unwrap();
        assert_eq!(upload.method, "PUT");
        assert!(
            upload
                .url
                .ends_with("test.bin:/content?@microsoft.graph.conflictBehavior=fail")
        );
        assert_eq!(upload.bytes, Some(vec![0, 255]));
        assert!(
            ProductWrite::drive_upload(None, None, Some("test.bin"), vec![0; 128_001], None)
                .is_err()
        );
        assert!(ProductWrite::drive_upload(None, Some("id"), None, vec![], None).is_err());
        assert!(ProductWrite::contact(None, "DELETE", None, None).is_err());
        assert!(
            ProductWrite::group(
                None,
                "POST",
                Some(json!({"displayName":"Test","mailNickname":"test"}))
            )
            .is_ok()
        );
        for body in [
            json!({"displayName":"Test"}),
            json!({"displayName":"Test","mailNickname":"bad alias"}),
            json!({"displayName":"Test","mailNickname":"test","groupTypes":["Security"]}),
        ] {
            assert!(ProductWrite::group(None, "POST", Some(body)).is_err());
        }
        assert!(
            ProductWrite::group(
                Some("id"),
                "PATCH",
                Some(json!({"owners@odata.bind":["id"]}))
            )
            .is_err()
        );
        assert!(ProductWrite::drive_folder(None, "Synthetic").is_ok());
        for name in ["..", "a/b", "a\\b", "a%2f", ""] {
            assert!(ProductWrite::drive_folder(None, name).is_err());
        }
        assert!(
            ProductWrite::drive_item("id", "PATCH", Some(json!({"name":"new"})), "etag").is_ok()
        );
        assert!(
            ProductWrite::drive_item(
                "id",
                "PATCH",
                Some(json!({"parentReference":{"id":"elsewhere"}})),
                "etag"
            )
            .is_ok()
        );
        assert!(
            ProductWrite::drive_item(
                "id",
                "PATCH",
                Some(json!({"parentReference":{"driveId":"other","id":"elsewhere"}})),
                "etag"
            )
            .is_err()
        );
        assert!(ProductWrite::drive_item("id", "DELETE", None, "").is_err());
        assert!(ProductWrite::drive_item("id", "DELETE", None, "x\r\nheader").is_err());
        assert!(
            ProductWrite::contact(
                None,
                "POST",
                Some(json!({"personalNotes":"x".repeat(200_001)})),
                None
            )
            .is_err()
        );
    }
}
