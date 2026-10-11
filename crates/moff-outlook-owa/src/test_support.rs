use crate::{OutlookWebSession, SessionTransport};
use serde_json::Value;
use std::{
    collections::VecDeque,
    sync::{Arc, Mutex},
};

type Replies = Arc<Mutex<VecDeque<(&'static str, Value)>>>;

pub struct MockSession {
    replies: Replies,
    requests: Arc<Mutex<Vec<Value>>>,
}

impl MockSession {
    pub fn call(&self, action: &'static str, body: &Value) -> Value {
        let (expected, response) = self
            .replies
            .lock()
            .unwrap()
            .pop_front()
            .expect("unexpected extra provider call");
        assert_eq!(action, expected);
        self.requests.lock().unwrap().push(body.clone());
        response
    }
}

pub struct MockCompletion(Replies);

impl MockCompletion {
    pub fn join(self) -> Result<(), &'static str> {
        if self.0.lock().unwrap().is_empty() {
            Ok(())
        } else {
            Err("expected provider calls were not completed")
        }
    }
}

#[allow(clippy::needless_pass_by_value)]
pub fn message(item: Value) -> Value {
    serde_json::json!({"Body":{"ResponseMessages":{"Items":[{"ResponseCode":"NoError","Items":[item]}]}}})
}

pub fn mock(
    replies: Vec<(&'static str, Value)>,
) -> (OutlookWebSession, Arc<Mutex<Vec<Value>>>, MockCompletion) {
    let replies = Arc::new(Mutex::new(replies.into()));
    let requests = Arc::new(Mutex::new(Vec::new()));
    (
        OutlookWebSession {
            transport: SessionTransport::Mock(MockSession {
                replies: replies.clone(),
                requests: requests.clone(),
            }),
        },
        requests,
        MockCompletion(replies),
    )
}
