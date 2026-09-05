use std::collections::VecDeque;
use std::sync::Mutex;

use moff_core::{GraphClient, GraphError, GraphHttp, GraphRequest, GraphResponse, MailService};

struct FakeHttp {
    responses: Mutex<VecDeque<Result<GraphResponse, GraphError>>>,
    requests: Mutex<Vec<GraphRequest>>,
}
impl FakeHttp {
    fn with(responses: Vec<Result<GraphResponse, GraphError>>) -> Self {
        Self {
            responses: Mutex::new(responses.into()),
            requests: Mutex::new(vec![]),
        }
    }
}
impl GraphHttp for FakeHttp {
    fn request(&self, request: GraphRequest) -> Result<GraphResponse, GraphError> {
        self.requests.lock().unwrap().push(request);
        self.responses.lock().unwrap().pop_front().unwrap()
    }
}
fn client(http: &FakeHttp) -> GraphClient<'_> {
    GraphClient::new(http, "https://graph.microsoft.com/v1.0", "secret-token").unwrap()
}

#[test]
fn list_and_search_map_to_safe_graph_requests_and_label_untrusted_email() {
    let http = FakeHttp::with(vec![
        Ok(GraphResponse::json(
            200,
            r#"{"value":[{"id":"a","subject":"hello"}]}"#,
        )),
        Ok(GraphResponse::json(200, r#"{"value":[]}"#)),
    ]);
    let mail = MailService::new(client(&http));
    let listed = mail.list("inbox", 25).unwrap();
    assert_eq!(listed["content_origin"], "external_email_untrusted");
    mail.search("from:alice@example.test & urgent", 25).unwrap();
    let requests = http.requests.lock().unwrap();
    assert!(requests[0].url.contains(
        "/me/mailFolders/inbox/messages?$select=id%2Csubject%2Cfrom%2CreceivedDateTime%2CisRead"
    ));
    assert!(
        requests[1]
            .url
            .contains("$search=%22from%3Aalice%40example.test%20%26%20urgent%22")
    );
    assert!(!requests[1].headers.contains_key("ConsistencyLevel"));
    assert!(
        requests
            .iter()
            .all(|request| request.headers["Authorization"] != "secret-token")
    );
}

#[test]
fn mail_operations_use_documented_routes_and_bodies() {
    let http = FakeHttp::with(vec![
        Ok(GraphResponse::json(200, r#"{"id":"m"}"#)),
        Ok(GraphResponse::json(202, "")),
        Ok(GraphResponse::json(204, "")),
    ]);
    let mail = MailService::new(client(&http));
    mail.move_message("id", "folder").unwrap();
    mail.send_draft("id").unwrap();
    mail.soft_delete("id").unwrap();
    let requests = http.requests.lock().unwrap();
    assert_eq!(requests[0].method, "POST");
    assert_eq!(
        requests[0].url,
        "https://graph.microsoft.com/v1.0/me/messages/id/move"
    );
    assert_eq!(
        requests[0].body.as_deref(),
        Some(r#"{"destinationId":"folder"}"#)
    );
    assert_eq!(requests[1].body.as_deref(), Some(""));
    assert_eq!(
        requests[1]
            .headers
            .get("Content-Length")
            .map(String::as_str),
        Some("0")
    );
    assert_eq!(requests[2].method, "DELETE");
}

#[test]
fn continuation_requires_same_configured_host_and_delta_link_is_exposed_only_at_end() {
    let http = FakeHttp::with(vec![Ok(GraphResponse::json(
        200,
        r#"{"value":[],"@odata.deltaLink":"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=x"}"#,
    ))]);
    let mail = MailService::new(client(&http));
    let result = mail.delta("inbox", None).unwrap();
    assert_eq!(
        result["delta_link"],
        "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages/delta?$deltatoken=x"
    );
    assert!(matches!(
        mail.delta("inbox", Some("https://evil.example/next")),
        Err(GraphError::UntrustedNextLink)
    ));
}
