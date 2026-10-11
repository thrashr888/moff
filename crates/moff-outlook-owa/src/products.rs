use crate::{OutlookWebError, OutlookWebSession};
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
        ["v1.0", "me", "people"]
            | ["v1.0", "groups"]
            | [
                "v1.0",
                "me",
                "drive",
                "root",
                "children" | "search(q='.loop')"
            ]
            | ["v1.0", "me", "drive", "items", _, "children"]
    );
    let user = matches!(segments.as_slice(), ["v1.0", "users", _]);
    let allowed = paged
        || user
        || matches!(
            segments.as_slice(),
            ["v1.0", "me", "manager" | "directReports"]
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
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn only_bounded_product_metadata_routes_are_allowed() {
        for path in [
            "/v1.0/me/people?$top=5",
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
            "/v1.0/me/drive/root/search(q='anything')",
        ] {
            assert!(product_url(path).is_err(), "{path}");
        }
    }
}
