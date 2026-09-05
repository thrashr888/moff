#[path = "../main.rs"]
mod application;

fn main() -> std::process::ExitCode {
    use moff_core::{InMemoryVault, TokenVault};

    let Some(path) = std::env::var_os("MOFF_TEST_REFRESH_TOKEN_FILE") else {
        eprintln!("moff-test: controlled refresh-token file is required");
        return std::process::ExitCode::from(2);
    };
    let Ok(refresh_token) = std::fs::read_to_string(path) else {
        eprintln!("moff-test: unable to read controlled refresh-token file");
        return std::process::ExitCode::from(2);
    };
    let vault = InMemoryVault::default();
    if vault.set("default", refresh_token.trim_end()).is_err() {
        eprintln!("moff-test: unable to prepare controlled credentials");
        return std::process::ExitCode::from(2);
    }
    application::main_with_test_endpoints(&vault)
}
