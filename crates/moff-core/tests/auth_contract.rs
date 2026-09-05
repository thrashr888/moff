use std::collections::{BTreeMap, VecDeque};
use std::sync::{Arc, Mutex};

use moff_core::{
    AuthService, DeviceFlowError, HttpClient, HttpResponse, InMemoryVault, ProfileStore, Sleeper,
    TokenVault,
};

#[derive(Clone, Default)]
struct FakeHttp {
    requests: Arc<Mutex<Vec<(String, BTreeMap<String, String>)>>>,
    responses: Arc<Mutex<VecDeque<HttpResponse>>>,
}

impl FakeHttp {
    fn with(responses: Vec<HttpResponse>) -> Self {
        Self {
            responses: Arc::new(Mutex::new(responses.into())),
            ..Self::default()
        }
    }
}

impl HttpClient for FakeHttp {
    fn post_form(
        &self,
        url: &str,
        form: &BTreeMap<String, String>,
    ) -> Result<HttpResponse, DeviceFlowError> {
        self.requests
            .lock()
            .unwrap()
            .push((url.to_owned(), form.clone()));
        Ok(self.responses.lock().unwrap().pop_front().unwrap())
    }
}

#[derive(Default)]
struct FakeSleeper {
    waits: Mutex<Vec<u64>>,
}
impl Sleeper for FakeSleeper {
    fn sleep_seconds(&self, seconds: u64) {
        self.waits.lock().unwrap().push(seconds);
    }
}

fn response(status: u16, body: &str) -> HttpResponse {
    HttpResponse {
        status,
        body: body.to_owned(),
    }
}

#[test]
fn profile_store_rejects_empty_or_unsafe_names() {
    for name in ["", " ", "../work", "a/b", "a\\b"] {
        assert!(
            ProfileStore::validate_profile_name(name).is_err(),
            "{name:?} must fail"
        );
    }
    assert_eq!(
        ProfileStore::validate_profile_name("default").unwrap(),
        "default"
    );
}

#[test]
fn device_flow_posts_expected_form_polls_and_saves_only_refresh_token() {
    let http = FakeHttp::with(vec![
        response(
            200,
            r#"{"device_code":"device-secret","user_code":"ABCD","verification_uri":"https://microsoft.com/devicelogin","expires_in":900,"interval":1,"message":"Go there"}"#,
        ),
        response(400, r#"{"error":"authorization_pending"}"#),
        response(400, r#"{"error":"slow_down"}"#),
        response(
            200,
            r#"{"access_token":"access-secret","refresh_token":"refresh-secret","expires_in":3600,"scope":"Mail.Read"}"#,
        ),
    ]);
    let vault = InMemoryVault::default();
    let sleeper = FakeSleeper::default();
    let service = AuthService::new(&http, &vault, &sleeper, "https://login.example");
    let result = service
        .login(
            "work",
            "11111111-1111-1111-1111-111111111111",
            "common",
            &["Mail.Read"],
        )
        .unwrap();
    assert_eq!(result.user_code, "ABCD");
    assert_eq!(
        vault.get("work").unwrap().as_deref(),
        Some("refresh-secret")
    );
    assert_eq!(*sleeper.waits.lock().unwrap(), vec![1, 1, 6]);
    let requests = http.requests.lock().unwrap();
    assert_eq!(
        requests[0].0,
        "https://login.example/common/oauth2/v2.0/devicecode"
    );
    assert_eq!(
        requests[0].1["client_id"],
        "11111111-1111-1111-1111-111111111111"
    );
    assert_eq!(requests[0].1["scope"], "offline_access Mail.Read");
    assert!(!format!("{result:?}").contains("secret"));
}

#[test]
fn device_flow_surfaces_declined_and_expired_errors() {
    for error in ["authorization_declined", "expired_token"] {
        let http = FakeHttp::with(vec![
            response(
                200,
                r#"{"device_code":"d","user_code":"U","verification_uri":"https://x","expires_in":10,"interval":1,"message":"m"}"#,
            ),
            response(400, &format!(r#"{{"error":"{error}"}}"#)),
        ]);
        let vault = InMemoryVault::default();
        let sleeper = FakeSleeper::default();
        let service = AuthService::new(&http, &vault, &sleeper, "https://login.example");
        assert!(matches!(
            service.login(
                "default",
                "11111111-1111-1111-1111-111111111111",
                "common",
                &[]
            ),
            Err(DeviceFlowError::Terminal(_))
        ));
    }
}

#[test]
fn refresh_rotates_refresh_token_without_exposing_it() {
    let http = FakeHttp::with(vec![response(
        200,
        r#"{"access_token":"new-access","refresh_token":"new-refresh","expires_in":3600}"#,
    )]);
    let vault = InMemoryVault::default();
    vault.set("one", "old-refresh").unwrap();
    let sleeper = FakeSleeper::default();
    let service = AuthService::new(&http, &vault, &sleeper, "https://login.example");
    let token = service
        .access_token("one", "11111111-1111-1111-1111-111111111111", "common")
        .unwrap();
    assert_eq!(token.expires_in_seconds(), 3600);
    assert_eq!(vault.get("one").unwrap().as_deref(), Some("new-refresh"));
    assert!(!format!("{token:?}").contains("new-access"));
}

#[test]
fn vault_isolates_profiles() {
    let vault = InMemoryVault::default();
    vault.set("first", "one").unwrap();
    vault.set("second", "two").unwrap();
    vault.delete("first").unwrap();
    assert_eq!(vault.get("first").unwrap(), None);
    assert_eq!(vault.get("second").unwrap().as_deref(), Some("two"));
}
