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
