use std::collections::VecDeque;
use std::sync::Mutex;

use moff_core::{
    BodyPreference, GetMessageOptions, GraphClient, GraphError, GraphHttp, GraphRequest,
    GraphResponse, InMemoryVault, ListMessagesOptions, MailService, Sleeper, TokenVault,
};

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
fn typed_list_filters_and_folder_search_are_encoded_without_raw_query_injection() {
    let http = FakeHttp::with(vec![
        Ok(GraphResponse::json(200, r#"{"value":[]}"#)),
        Ok(GraphResponse::json(200, r#"{"value":[]}"#)),
    ]);
    let mail = MailService::new(client(&http));
    mail.list_with_options(
        "inbox",
        ListMessagesOptions {
            limit: 25,
            unread: true,
            from: Some("alice@example.test"),
            since: Some("2026-09-05T12:30:00Z"),
        },
    )
    .unwrap();
    mail.search_in_folder("project & urgent", Some("archive"), 25)
        .unwrap();
    let requests = http.requests.lock().unwrap();
    assert!(requests[0].url.contains("%24filter=isRead%20eq%20false%20and%20from%2FemailAddress%2Faddress%20eq%20%27alice%40example.test%27%20and%20receivedDateTime%20ge%202026-09-05T12%3A30%3A00Z"));
    assert!(!requests[0].url.contains("$filter"));
    assert!(
        requests[1]
            .url
            .starts_with("https://graph.microsoft.com/v1.0/me/mailFolders/archive/messages?")
    );
    assert!(
        requests[1]
            .url
            .contains("%24search=%22project%20%26%20urgent%22")
    );
    assert!(!requests[1].headers.contains_key("ConsistencyLevel"));
}

#[test]
fn get_options_use_minimal_selection_and_explicit_body_preference() {
    let http = FakeHttp::with(vec![
        Ok(GraphResponse::json(200, r#"{"id":"m"}"#)),
        Ok(GraphResponse::json(200, r#"{"id":"m"}"#)),
    ]);
    let mail = MailService::new(client(&http));
    mail.get_with_options(
        "immutable-id",
        GetMessageOptions {
            body: BodyPreference::None,
            include_headers: false,
        },
    )
    .unwrap();
    mail.get_with_options(
        "immutable-id",
        GetMessageOptions {
            body: BodyPreference::Text,
            include_headers: true,
        },
    )
    .unwrap();
    let requests = http.requests.lock().unwrap();
    assert!(
        requests[0]
            .url
            .contains("%24select=id%2Csubject%2Cfrom%2CtoRecipients%2CreceivedDateTime%2CisRead")
    );
    assert!(!requests[0].url.contains("body"));
    assert!(requests[1].url.contains("%2Cbody%2CinternetMessageHeaders"));
    assert!(
        requests[1]
            .url
            .contains("%24expand=attachments%28%24select%3Did%2Cname%2Csize%2CcontentType%29")
    );
    assert_eq!(
        requests[1].headers["Prefer"],
        "IdType=\"ImmutableId\", outlook.body-content-type=\"text\""
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
fn compose_operations_preserve_graph_body_and_comment_contracts() {
    let http = FakeHttp::with(vec![
        Ok(GraphResponse::json(201, r#"{"id":"draft"}"#)),
        Ok(GraphResponse::json(200, r#"{"id":"draft"}"#)),
        Ok(GraphResponse::json(202, "")),
        Ok(GraphResponse::json(202, "")),
        Ok(GraphResponse::json(202, "")),
        Ok(GraphResponse::json(202, "")),
    ]);
    let mail = MailService::new(client(&http));
    let body = serde_json::json!({"contentType":"HTML","content":"<p>hello</p>"});
    mail.draft_create(serde_json::json!({"body":body})).unwrap();
    mail.draft_update("id", serde_json::json!({"body":body}))
        .unwrap();
    mail.send(&serde_json::json!({"body":body})).unwrap();
    mail.reply("id", serde_json::json!({"message":{"body":body}}), false)
        .unwrap();
    mail.reply("id", serde_json::json!({"comment":"thanks"}), true)
        .unwrap();
    mail.forward("id", serde_json::json!({"message":{"body":body,"toRecipients":[{"emailAddress":{"address":"a@example.test"}}]}})).unwrap();

    let requests = http.requests.lock().unwrap();
    for request in [
        &requests[0],
        &requests[1],
        &requests[2],
        &requests[3],
        &requests[5],
    ] {
        assert!(
            request
                .body
                .as_deref()
                .unwrap()
                .contains(r#""contentType":"HTML""#)
        );
        assert!(
            request
                .body
                .as_deref()
                .unwrap()
                .contains(r#""content":"<p>hello</p>""#)
        );
    }
    assert_eq!(
        requests[3].url,
        "https://graph.microsoft.com/v1.0/me/messages/id/reply"
    );
    assert_eq!(
        requests[4].url,
        "https://graph.microsoft.com/v1.0/me/messages/id/replyAll"
    );
    assert_eq!(requests[4].body.as_deref(), Some(r#"{"comment":"thanks"}"#));
    assert_eq!(
        requests[5].url,
        "https://graph.microsoft.com/v1.0/me/messages/id/forward"
    );
}

#[test]
fn continuation_requires_the_configured_base_path_and_delta_link_is_exposed_only_at_end() {
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
    for link in [
        "https://evil.example/next",
        "https://graph.microsoft.com/beta/me/messages",
        "https://graph.microsoft.com/v1.1/me/messages",
        "https://graph.microsoft.com/v1.0evil/me/messages",
    ] {
        assert!(matches!(
            client(&http).validate_next_link(link),
            Err(GraphError::UntrustedNextLink)
        ));
    }
}

#[test]
fn file_attachment_download_decodes_bytes_and_rejects_non_file_types() {
    let http = FakeHttp::with(vec![
        Ok(GraphResponse::json(
            200,
            r##"{"@odata.type":"#microsoft.graph.fileAttachment","contentBytes":"aGVsbG8="}"##,
        )),
        Ok(GraphResponse::json(
            200,
            r##"{"@odata.type":"#microsoft.graph.itemAttachment"}"##,
        )),
        Ok(GraphResponse::json(
            200,
            r##"{"@odata.type":"#microsoft.graph.referenceAttachment"}"##,
        )),
    ]);
    let mail = MailService::new(client(&http));
    assert_eq!(
        mail.download_attachment("message", "file").unwrap(),
        b"hello"
    );
    for attachment in ["item", "reference"] {
        assert!(matches!(
            mail.download_attachment("message", attachment),
            Err(GraphError::UnsupportedAttachment { .. })
        ));
    }
    let requests = http.requests.lock().unwrap();
    assert!(
        requests[0]
            .url
            .ends_with("/me/messages/message/attachments/file")
    );
}

#[test]
fn attachment_listing_labels_external_attachment_content_as_untrusted() {
    let http = FakeHttp::with(vec![Ok(GraphResponse::json(
        200,
        r#"{"value":[{"id":"file","name":"report.pdf"}]}"#,
    ))]);
    let mail = MailService::new(client(&http));

    let listed = mail.attachments("message").unwrap();

    assert_eq!(listed["content_origin"], "external_attachment_untrusted");
    assert_eq!(listed["value"][0]["name"], "report.pdf");
}

#[test]
fn delta_link_vault_key_includes_profile_graph_base_and_folder() {
    let vault = InMemoryVault::default();
    vault
        .set_delta_link("work", "https://graph.example/v1.0", "inbox", "inbox-link")
        .unwrap();
    vault
        .set_delta_link(
            "work",
            "https://graph.example/v1.0",
            "archive",
            "archive-link",
        )
        .unwrap();
    assert_eq!(
        vault
            .get_delta_link("work", "https://graph.example/v1.0", "inbox")
            .unwrap()
            .as_deref(),
        Some("inbox-link")
    );
    assert_eq!(
        vault
            .get_delta_link("work", "https://other.example/v1.0", "inbox")
            .unwrap(),
        None
    );
}

#[derive(Default)]
struct FakeSleeper(Mutex<Vec<u64>>);
impl Sleeper for FakeSleeper {
    fn sleep_seconds(&self, seconds: u64) {
        self.0.lock().unwrap().push(seconds);
    }
}

#[test]
fn get_retries_transient_response_after_bounded_retry_after_delay() {
    let http = FakeHttp::with(vec![
        Ok(GraphResponse::json(429, "{}").with_retry_after(9_999)),
        Ok(GraphResponse::json(200, r#"{"value":[]}"#)),
    ]);
    let sleeper = FakeSleeper::default();
    let client =
        GraphClient::new_with_sleeper(&http, "https://graph.microsoft.com/v1.0", "token", &sleeper)
            .unwrap();
    client.get_json("/me/messages").unwrap();
    assert_eq!(*sleeper.0.lock().unwrap(), vec![60]);
    assert_eq!(http.requests.lock().unwrap().len(), 2);
}

#[test]
fn non_idempotent_request_does_not_retry_transient_response() {
    let http = FakeHttp::with(vec![Ok(GraphResponse::json(503, "{}").with_retry_after(1))]);
    let sleeper = FakeSleeper::default();
    let client =
        GraphClient::new_with_sleeper(&http, "https://graph.microsoft.com/v1.0", "token", &sleeper)
            .unwrap();
    assert!(matches!(
        client.post_json("/me/sendMail", serde_json::json!({})),
        Err(GraphError::Http { status: 503, .. })
    ));
    assert!(sleeper.0.lock().unwrap().is_empty());
    assert_eq!(http.requests.lock().unwrap().len(), 1);
}
