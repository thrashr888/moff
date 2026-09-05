use std::collections::VecDeque;
use std::sync::Mutex;

use moff_core::{GraphClient, GraphError, GraphHttp, GraphRequest, GraphResponse};

struct FakeHttp(Mutex<VecDeque<Result<GraphResponse, GraphError>>>);
impl GraphHttp for FakeHttp {
    fn request(&self, _: GraphRequest) -> Result<GraphResponse, GraphError> {
        self.0.lock().unwrap().pop_front().unwrap()
    }
}

#[test]
fn graph_get_uses_immutable_ids_and_retries_only_safe_reads() {
    let http = FakeHttp(Mutex::new(VecDeque::from(vec![
        Err(GraphError::Transport),
        Ok(GraphResponse::json(200, r#"{"value":[]}"#)),
    ])));
    let client = GraphClient::new(&http, "https://graph.microsoft.com/v1.0", "token").unwrap();
    assert_eq!(
        client.get_json("/me/messages").unwrap()["value"],
        serde_json::json!([])
    );
}

#[test]
fn graph_rejects_untrusted_next_links_and_marks_mutation_ambiguity() {
    let http = FakeHttp(Mutex::new(VecDeque::from(vec![Err(GraphError::Transport)])));
    let client = GraphClient::new(&http, "https://graph.microsoft.com/v1.0", "token").unwrap();
    assert!(matches!(
        client.next_json("https://evil.example/next"),
        Err(GraphError::UntrustedNextLink)
    ));
    assert!(matches!(
        client.post_json("/me/sendMail", serde_json::json!({})),
        Err(GraphError::OutcomeUnknown { .. })
    ));
}

#[test]
fn mutation_http_failures_are_outcome_unknown_without_retries() {
    let responses = [
        GraphResponse {
            status: 500,
            body: "{}".to_owned(),
            request_id: Some("post-request".to_owned()),
            retry_after_seconds: None,
        },
        GraphResponse {
            status: 503,
            body: "{}".to_owned(),
            request_id: Some("patch-request".to_owned()),
            retry_after_seconds: None,
        },
        GraphResponse {
            status: 429,
            body: "{}".to_owned(),
            request_id: Some("delete-request".to_owned()),
            retry_after_seconds: None,
        },
        GraphResponse {
            status: 504,
            body: "{}".to_owned(),
            request_id: Some("empty-post-request".to_owned()),
            retry_after_seconds: None,
        },
    ];
    let http = FakeHttp(Mutex::new(VecDeque::from(
        responses.into_iter().map(Ok).collect::<Vec<_>>(),
    )));
    let client = GraphClient::new(&http, "https://graph.microsoft.com/v1.0", "token").unwrap();

    for (method, expected_request_id) in [
        ("POST", "post-request"),
        ("PATCH", "patch-request"),
        ("DELETE", "delete-request"),
    ] {
        let result = client.mutate_json(method, "/me/messages/id", Some(serde_json::json!({})));
        assert!(matches!(
            result,
            Err(GraphError::OutcomeUnknown { request_id: Some(request_id) })
                if request_id == expected_request_id
        ));
    }
    assert!(matches!(
        client.post_empty("/me/messages/id/send"),
        Err(GraphError::OutcomeUnknown { request_id: Some(request_id) })
            if request_id == "empty-post-request"
    ));
}

#[test]
fn graph_base_allows_only_official_graph_endpoint() {
    let http = FakeHttp(Mutex::new(VecDeque::new()));
    for base in [
        "https://graph.microsoft.com",
        "https://graph.microsoft.com/beta",
        "https://graph.example/v1.0",
        "https://graph.microsoft.com.evil.example/v1.0",
    ] {
        assert!(matches!(
            GraphClient::new(&http, base, "token"),
            Err(GraphError::InvalidBase)
        ));
    }
    assert!(GraphClient::new(&http, "https://graph.microsoft.com/v1.0", "token").is_ok());
}

#[test]
fn production_graph_client_rejects_loopback_http_graph_base() {
    let http = FakeHttp(Mutex::new(VecDeque::new()));
    assert!(matches!(
        GraphClient::new(&http, "http://127.0.0.1:8080/v1.0", "token"),
        Err(GraphError::InvalidBase)
    ));
}

#[cfg(feature = "test-support")]
#[test]
fn test_graph_client_allows_only_loopback_http_graph_base() {
    let http = FakeHttp(Mutex::new(VecDeque::new()));
    assert!(GraphClient::new_for_test(&http, "http://127.0.0.1:8080/v1.0", "token").is_ok());
    assert!(matches!(
        GraphClient::new_for_test(&http, "http://example.test/v1.0", "token"),
        Err(GraphError::InvalidBase)
    ));
}

#[cfg(feature = "test-support")]
#[test]
fn test_graph_client_rejects_spoofed_or_malformed_loopback_bases() {
    let http = FakeHttp(Mutex::new(VecDeque::new()));
    for base in [
        "http://localhost:ignored@evil.example/v1.0",
        "http://127.0.0.1:8080/v1.0?query",
        "http://127.0.0.1:8080/v1.0#fragment",
        "http://127.0.0.1/v1.0",
        "http://127.0.0.1:0/v1.0",
        "http://example.test:8080/v1.0",
        "http://127.0.0.1:8080/not-v1.0",
        "http://127.0.0.1:8080/v1.0/",
    ] {
        assert!(matches!(
            GraphClient::new_for_test(&http, base, "token"),
            Err(GraphError::InvalidBase)
        ));
    }
}
