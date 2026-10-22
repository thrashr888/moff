use crate::{OutlookWebError, OutlookWebSession};
use serde_json::{Value, json};
use url::Url;

pub const DEFAULT_MAX_FILE_BYTES: u64 = 250_000_000;
pub const MAX_FILE_BYTES: u64 = 250_000_000;

pub fn validate_file_size(size: u64, max_bytes: u64) -> Result<(), OutlookWebError> {
    if max_bytes == 0 || max_bytes > MAX_FILE_BYTES || size > max_bytes {
        return Err(invalid(
            "file exceeds --max-bytes or maximum is outside 1..=250000000",
        ));
    }
    Ok(())
}

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
            | [
                "v1.0",
                "me",
                "drive",
                "items",
                _,
                "children" | "permissions"
            ]
            | ["v1.0", "drives", _, "items", _, "permissions"]
            | ["v1.0", "sites", _, "drives"]
    ) || matches!(segments.as_slice(), ["v1.0","me","drive","root", search]
        | ["v1.0","drives",_,"root",search] if search.starts_with("search(q='") && search.ends_with("')"));
    let user = matches!(segments.as_slice(), ["v1.0", "users", _]);
    let allowed = paged
        || user
        || matches!(
            segments.as_slice(),
            ["v1.0", "me"]
                | ["v1.0", "me", "contacts", _]
                | ["v1.0", "groups" | "sites", _]
                | ["v1.0", "me", "manager" | "directReports"]
                | ["v1.0", "users", _, "manager" | "directReports"]
                | ["v1.0", "me", "drive", "items", _]
                | ["v1.0", "drives", _, "items", _]
                | ["v1.0", "me", "drive", "items", _, "permissions", _]
                | ["v1.0", "drives", _, "items", _, "permissions", _]
                | ["v1.0", "sites", _, "sites" | "teams", _]
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
    pub fn drive_share(
        &self,
        drive: Option<&str>,
        id: &str,
        email: &str,
        notify: bool,
    ) -> Result<Value, OutlookWebError> {
        let request = ProductWrite::drive_share(drive, id, email, notify)?;
        let path = drive_item_path(drive, id)?;
        share_file_with(
            &path,
            id,
            email,
            notify,
            |path| self.product_read(path),
            || self.product_write(&request),
        )
    }

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

    pub fn drive_download(
        &self,
        id: &str,
        output: &mut impl std::io::Write,
        max_bytes: u64,
    ) -> Result<u64, OutlookWebError> {
        let url = format!(
            "https://graph.microsoft.com/v1.0/me/drive/items/{}",
            product_segment(id)?
        );
        match &self.transport {
            crate::SessionTransport::Safari(safari) => {
                safari.product_download(&url, output, max_bytes)
            }
            #[cfg(test)]
            crate::SessionTransport::Mock(_) => Err(OutlookWebError::InvalidAction),
        }
    }

    pub fn drive_upload(
        &self,
        request: &ProductWrite,
        input: &mut impl std::io::Read,
        max_bytes: u64,
    ) -> Result<Value, OutlookWebError> {
        match &self.transport {
            crate::SessionTransport::Safari(safari) => {
                safari.product_upload(request, input, max_bytes)
            }
            #[cfg(test)]
            crate::SessionTransport::Mock(_) => Err(OutlookWebError::InvalidAction),
        }
    }
}

/// A validated, personal-content mutation. No caller-supplied URLs or headers.
#[derive(Debug)]
pub struct ProductWrite {
    pub(crate) url: String,
    pub(crate) method: &'static str,
    pub(crate) body: Option<Value>,
    pub(crate) etag: Option<String>,
    pub(crate) upload_size: Option<u64>,
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

pub fn drive_item_path(drive: Option<&str>, id: &str) -> Result<String, OutlookWebError> {
    let base = drive.map_or_else(
        || Ok("/v1.0/me/drive".to_owned()),
        |drive| Ok(format!("/v1.0/drives/{}", product_segment(drive)?)),
    )?;
    Ok(format!("{base}/items/{}", product_segment(id)?))
}

pub fn sharepoint_library_location(input: &str) -> Result<(String, String), OutlookWebError> {
    let valid = |url: &Url| {
        url.scheme() == "https"
            && url.username().is_empty()
            && url.password().is_none()
            && url.port().is_none()
            && url.fragment().is_none()
            && url
                .host_str()
                .is_some_and(|host| host.ends_with(".sharepoint.com"))
    };
    let mut url =
        Url::parse(input).map_err(|_| invalid("expected an HTTPS SharePoint library URL"))?;
    if !valid(&url) {
        return Err(invalid("expected an HTTPS SharePoint library URL"));
    }
    if url.path() == "/shared" {
        let links: Vec<_> = url
            .query_pairs()
            .filter(|(key, _)| key == "listurl")
            .map(|(_, value)| value.into_owned())
            .collect();
        if links.len() != 1 {
            return Err(invalid(
                "SharePoint shared URLs require one listurl library URL",
            ));
        }
        let library =
            Url::parse(&links[0]).map_err(|_| invalid("invalid SharePoint library URL"))?;
        if !valid(&library)
            || library.host_str()
                != url
                    .host_str()
                    .map(|host| {
                        host.strip_suffix("-my.sharepoint.com").map_or_else(
                            || host.to_owned(),
                            |tenant| format!("{tenant}.sharepoint.com"),
                        )
                    })
                    .as_deref()
        {
            return Err(invalid("library must belong to the same SharePoint tenant"));
        }
        url = library;
    }
    if url.query().is_some() {
        return Err(invalid(
            "use the library root URL, not a file or Forms view URL",
        ));
    }
    let mut path = url.path().to_owned();
    for byte in b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~" {
        let encoded = format!("%{byte:02X}");
        path = path
            .replace(&encoded, &char::from(*byte).to_string())
            .replace(
                &encoded.to_ascii_lowercase(),
                &char::from(*byte).to_string(),
            );
    }
    url.set_path(&path);
    let parts: Vec<_> = url.path().trim_end_matches('/').split('/').collect();
    if !matches!(parts.as_slice(), ["", "sites" | "teams", site, library]
        if !site.is_empty() && !library.is_empty())
        || url.path().to_ascii_lowercase().contains("%2f")
        || url.path().contains(['\\', ':'])
    {
        return Err(invalid(
            "expected a site library root such as /sites/Project/Shared%20Documents",
        ));
    }
    let site_path = format!(
        "/v1.0/sites/{}:/{}/{}",
        url.host_str().ok_or(invalid("missing host"))?,
        parts[1],
        parts[2]
    );
    Ok((site_path, url.as_str().trim_end_matches('/').to_owned()))
}

fn share_file_with(
    path: &str,
    id: &str,
    email: &str,
    notify: bool,
    mut read: impl FnMut(&str) -> Result<Value, OutlookWebError>,
    write: impl FnOnce() -> Result<Value, OutlookWebError>,
) -> Result<Value, OutlookWebError> {
    let before = read(path)?;
    if before["id"] != id || !before["file"].is_object() || before.get("remoteItem").is_some() {
        return Err(invalid(
            "sharing requires an exact file ID in its owning drive, not a folder or shortcut",
        ));
    }
    let result = write()?;
    let unknown = || OutlookWebError::OutcomeUnknown {
        action: "OneDrive sharing; inspect permissions before retrying",
    };
    let permissions = result["value"]
        .as_array()
        .filter(|values| values.len() == 1)
        .ok_or_else(unknown)?;
    let permission = &permissions[0];
    let permission_id = permission["id"]
        .as_str()
        .filter(|id| !id.is_empty())
        .ok_or_else(unknown)?;
    if permission.get("error").is_some() || permission["roles"] != json!(["read"]) {
        return Err(unknown());
    }
    let permission_id = product_segment(permission_id).map_err(|_| unknown())?;
    let after = read(&format!("{path}/permissions/{permission_id}")).map_err(|_| unknown())?;
    if after["id"] != permission["id"]
        || after["roles"] != json!(["read"])
        || after.get("link").is_some()
        || after.get("error").is_some()
    {
        return Err(unknown());
    }
    Ok(
        json!({"item":{"id":id,"name":before["name"],"webUrl":before["webUrl"]},
        "recipient":email,"roles":["read"],"permission":after,"verified":true,
        "notification":{"requested":notify,"delivery_verified":false}}),
    )
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
    pub fn drive_share(
        drive: Option<&str>,
        id: &str,
        email: &str,
        notify: bool,
    ) -> Result<Self, OutlookWebError> {
        let valid = email.is_ascii()
            && email.len() <= 254
            && !email
                .bytes()
                .any(|b| b.is_ascii_whitespace() || b.is_ascii_control())
            && email.split_once('@').is_some_and(|(local, domain)| {
                !local.is_empty()
                    && !domain.is_empty()
                    && !domain.contains('@')
                    && local
                        .bytes()
                        .all(|b| b.is_ascii_alphanumeric() || b".!#$%&'*+-/=?^_`{|}~".contains(&b))
                    && domain.split('.').all(|label| {
                        !label.is_empty()
                            && !label.starts_with('-')
                            && !label.ends_with('-')
                            && label
                                .bytes()
                                .all(|b| b.is_ascii_alphanumeric() || b == b'-')
                    })
            });
        if !valid {
            return Err(invalid(
                "sharing requires one explicit recipient email address",
            ));
        }
        Self::new(
            format!("{}/invite", drive_item_path(drive, id)?),
            "POST",
            Some(json!({"recipients":[{"email":email}],"roles":["read"],
                "requireSignIn":true,"sendInvitation":notify,"retainInheritedPermissions":true})),
            None,
        )
    }

    pub fn drive_upload(
        parent: Option<&str>,
        id: Option<&str>,
        name: Option<&str>,
        size: u64,
        etag: Option<&str>,
    ) -> Result<Self, OutlookWebError> {
        validate_file_size(size, MAX_FILE_BYTES)?;
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
        request.upload_size = Some(size);
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
            upload_size: None,
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
    fn sharepoint_library_links_resolve_to_the_named_site_not_personal_onedrive() {
        let expected = (
            "/v1.0/sites/ibm.sharepoint.com:/sites/PROJ-Meridian".to_owned(),
            "https://ibm.sharepoint.com/sites/PROJ-Meridian/Shared%20Documents".to_owned(),
        );
        assert_eq!(sharepoint_library_location(&expected.1).unwrap(), expected);
        assert_eq!(sharepoint_library_location("https://ibm-my.sharepoint.com/shared?id=%2Fsites%2FPROJ%2DMeridian%2FShared%20Documents&listurl=https%3A%2F%2Fibm%2Esharepoint%2Ecom%2Fsites%2FPROJ%2DMeridian%2FShared%20Documents").unwrap().0,
            "/v1.0/sites/ibm.sharepoint.com:/sites/PROJ-Meridian");
        for input in [
            "http://ibm.sharepoint.com/sites/Project/Documents",
            "https://evil.example/sites/Project/Documents",
            "https://user@ibm.sharepoint.com/sites/Project/Documents",
            "https://ibm-my.sharepoint.com/shared?listurl=https://other.sharepoint.com/sites/P/Docs",
            "https://ibm.sharepoint.com/sites/P/Docs/file.pdf",
            "https://ibm.sharepoint.com/sites/P/Docs?anything=yes",
            "https://ibm-my.sharepoint.com/shared",
        ] {
            assert!(sharepoint_library_location(input).is_err(), "{input}");
        }
    }

    #[test]
    fn only_bounded_product_metadata_routes_are_allowed() {
        for path in [
            "/v1.0/me/people?$top=5",
            "/v1.0/sites/ibm.sharepoint.com:/sites/PROJ-Meridian",
            "/v1.0/sites/site-id/drives?$top=100",
            "/v1.0/drives/drive-id/root/search(q='comparison')?$top=20",
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
            "/v1.0/me/drive/items/id/permissions?$top=20",
            "/v1.0/drives/drive-id/items/file-id",
            "/v1.0/drives/drive-id/items/file-id/permissions",
            "/v1.0/drives/drive-id/items/file-id/permissions/permission-id",
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
            "/v1.0/me/drive/items/id/invite",
            "/v1.0/drives/drive-id/items/file-id/invite",
            "/v1.0/drives/drive-id/root/children",
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
    fn sharing_is_one_named_recipient_with_read_access_and_optional_notification() {
        for (drive, notify, path) in [
            (None, false, "/v1.0/me/drive/items/file/invite"),
            (Some("drive"), true, "/v1.0/drives/drive/items/file/invite"),
        ] {
            let request =
                ProductWrite::drive_share(drive, "file", "sarah@example.com", notify).unwrap();
            assert_eq!(request.url, format!("https://graph.microsoft.com{path}"));
            assert_eq!(request.method, "POST");
            assert_eq!(
                request.body.unwrap(),
                json!({
                    "recipients":[{"email":"sarah@example.com"}],"roles":["read"],
                    "requireSignIn":true,"sendInvitation":notify,"retainInheritedPermissions":true
                })
            );
        }
        for email in [
            "",
            "Sarah",
            "a@b@c",
            "a@example.com,b@example.com",
            "Sarah <sarah@example.com>",
            "a@example.com\n",
            "a@-invalid.com",
            "a@x..com",
        ] {
            assert!(
                ProductWrite::drive_share(None, "file", email, false).is_err(),
                "{email}"
            );
        }
        assert!(
            ProductWrite::drive_share(Some("../drive"), "file", "a@example.com", false).is_err()
        );
        assert!(ProductWrite::drive_share(None, "", "a@example.com", false).is_err());
    }

    #[test]
    fn sharing_verifies_the_returned_permission_without_claiming_delivery() {
        let mut reads = Vec::new();
        let result = share_file_with("/v1.0/me/drive/items/file", "file", "a@example.com", true,
                |path| {
                    reads.push(path.to_owned());
                    Ok(if path.ends_with("/permissions/p1") {
                        json!({"id":"p1","roles":["read"],"invitation":{"email":"a@example.com"}})
                    } else {
                        json!({"id":"file","name":"comparison.pdf","file":{},"webUrl":"https://example.com/file"})
                    })
                },
                || Ok(json!({"value":[{"id":"p1","roles":["read"]}]}))).unwrap();
        assert_eq!(
            reads,
            [
                "/v1.0/me/drive/items/file",
                "/v1.0/me/drive/items/file/permissions/p1"
            ]
        );
        assert_eq!(result["verified"], true);
        assert_eq!(result["recipient"], "a@example.com");
        assert_eq!(result["notification"]["requested"], true);
        assert_eq!(result["notification"]["delivery_verified"], false);
    }

    #[test]
    fn sharing_rejects_folders_shortcuts_and_mismatched_files_before_writing() {
        for item in [
            json!({"id":"file","folder":{}}),
            json!({"id":"other","file":{}}),
            json!({"id":"file","file":{},"remoteItem":{"id":"remote"}}),
        ] {
            assert!(
                share_file_with(
                    "/v1.0/me/drive/items/file",
                    "file",
                    "a@example.com",
                    false,
                    |_| Ok(item.clone()),
                    || panic!("must not write")
                )
                .is_err()
            );
        }
    }

    #[test]
    fn sharing_stops_on_partial_invalid_or_unverified_results_without_retry() {
        for result in [
            json!({}),
            json!({"value":[]}),
            json!({"value":[{"id":"p1","roles":["read"]},{"error":{"code":"denied"}}]}),
            json!({"value":[{"id":"p1","roles":["write"]}]}),
            json!({"value":[{"id":"p1","roles":["read"],"error":{}}]}),
            json!({"value":[{"id":"p1","roles":["read"]}]}),
        ] {
            let mut writes = 0;
            let error = share_file_with(
                "/v1.0/me/drive/items/file",
                "file",
                "a@example.com",
                false,
                |path| {
                    if path.ends_with("/permissions/p1") {
                        Err(OutlookWebError::ProductRequestRejected(403))
                    } else {
                        Ok(json!({"id":"file","file":{}}))
                    }
                },
                || {
                    writes += 1;
                    Ok(result.clone())
                },
            )
            .unwrap_err();
            assert!(matches!(error, OutlookWebError::OutcomeUnknown { .. }));
            assert_eq!(writes, 1);
        }
        for after in [
            json!({"id":"other","roles":["read"]}),
            json!({"id":"p1","roles":["write"]}),
            json!({"id":"p1","roles":["read"],"link":{"scope":"anonymous"}}),
        ] {
            let error = share_file_with(
                "/v1.0/me/drive/items/file",
                "file",
                "a@example.com",
                false,
                |path| {
                    Ok(if path.ends_with("/permissions/p1") {
                        after.clone()
                    } else {
                        json!({"id":"file","file":{}})
                    })
                },
                || Ok(json!({"value":[{"id":"p1","roles":["read"]}]})),
            )
            .unwrap_err();
            assert!(matches!(error, OutlookWebError::OutcomeUnknown { .. }));
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
        let upload = ProductWrite::drive_upload(None, None, Some("test.bin"), 2, None).unwrap();
        assert_eq!(upload.method, "PUT");
        assert!(
            upload
                .url
                .ends_with("test.bin:/content?@microsoft.graph.conflictBehavior=fail")
        );
        assert_eq!(upload.upload_size, Some(2));
        assert!(ProductWrite::drive_upload(None, None, Some("large.bin"), 1_500_000, None).is_ok());
        assert!(
            ProductWrite::drive_upload(None, None, Some("test.bin"), MAX_FILE_BYTES + 1, None)
                .is_err()
        );
        assert!(ProductWrite::drive_upload(None, Some("id"), None, 0, None).is_err());
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
    #[test]
    fn transfer_limits_are_required_ceilings() {
        assert!(validate_file_size(0, 1).is_ok());
        assert!(validate_file_size(1_500_000, DEFAULT_MAX_FILE_BYTES).is_ok());
        assert!(validate_file_size(DEFAULT_MAX_FILE_BYTES, DEFAULT_MAX_FILE_BYTES).is_ok());
        for (size, maximum) in [
            (0, 0),
            (2, 1),
            (1, MAX_FILE_BYTES + 1),
            (MAX_FILE_BYTES + 1, MAX_FILE_BYTES),
        ] {
            assert!(validate_file_size(size, maximum).is_err());
        }
    }
}
