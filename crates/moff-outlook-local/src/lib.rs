//! Read-only, local-only macOS Outlook accessibility boundary.
//!
//! This crate deliberately knows nothing about Moff credentials, Microsoft Graph,
//! Outlook profiles, or network transports. Its sole production process boundary is
//! `/usr/bin/osascript -l JavaScript` talking to the Accessibility tree.

use std::fmt;
use std::process::Command;

use serde_json::{Value, json};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum UiRequest {
    Doctor,
    VisibleMail { limit: u16 },
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UiReply {
    pub stdout: String,
}
impl UiReply {
    #[must_use]
    pub fn json(stdout: impl Into<String>) -> Self {
        Self {
            stdout: stdout.into(),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum LocalError {
    OutlookNotRunning,
    AccessibilityUntrusted,
    UiUnsupported,
    UiChanged,
}
impl LocalError {
    #[must_use]
    pub const fn code(&self) -> &'static str {
        match self {
            Self::OutlookNotRunning => "outlook_not_running",
            Self::AccessibilityUntrusted => "accessibility_untrusted",
            Self::UiUnsupported => "ui_unsupported",
            Self::UiChanged => "ui_changed",
        }
    }
}
impl fmt::Display for LocalError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(self.code())
    }
}
impl std::error::Error for LocalError {}

/// Injectable subprocess boundary. Implementations must not mutate Outlook UI.
pub trait OutlookUi {
    ///
    /// # Errors
    /// Returns a typed error when the process, accessibility permission, or UI shape is unavailable.
    fn run(&self, request: UiRequest) -> Result<UiReply, LocalError>;
}

/// Production implementation: only System Events AX inspection via JXA.
pub struct OsaScriptOutlookUi;
impl OutlookUi for OsaScriptOutlookUi {
    fn run(&self, request: UiRequest) -> Result<UiReply, LocalError> {
        #[cfg(not(target_os = "macos"))]
        {
            let _ = request;
            return Err(LocalError::UiUnsupported);
        }
        #[cfg(target_os = "macos")]
        {
            let script = match request {
                UiRequest::Doctor => DOCTOR_JXA.to_owned(),
                UiRequest::VisibleMail { limit } => visible_jxa(limit),
            };
            let output = Command::new("/usr/bin/osascript")
                .args(["-l", "JavaScript", "-e", &script])
                .output()
                .map_err(|_| LocalError::UiUnsupported)?;
            if !output.status.success() {
                let stderr = String::from_utf8_lossy(&output.stderr);
                return Err(
                    if stderr.contains("not authorized") || stderr.contains("-1743") {
                        LocalError::AccessibilityUntrusted
                    } else {
                        LocalError::UiUnsupported
                    },
                );
            }
            Ok(UiReply::json(String::from_utf8_lossy(&output.stdout)))
        }
    }
}

const DOCTOR_JXA: &str = r"
const se = Application('System Events');
const p = se.processes.byName('Microsoft Outlook');
if (!p.exists()) JSON.stringify({state:'outlook_not_running'});
else { let version='unknown'; try { version=String(p.applicationFile.version()); } catch (_) {} JSON.stringify({state:'ok', running:true, version:version, ui_available:p.windows().length > 0}); }
";
fn visible_jxa(limit: u16) -> String {
    // This is inspection only: no activate, keystroke, click, set, or navigation.
    format!(
        r"
const se = Application('System Events');
const p = se.processes.byName('Microsoft Outlook');
if (!p.exists()) JSON.stringify({{state:'outlook_not_running'}});
else {{
  const w = p.windows[0];
  if (!w || !w.exists()) JSON.stringify({{state:'ui_changed'}});
  else {{
    const found=[]; const walk=(node, depth)=>{{
      if (depth>8 || found.length>={limit}) return;
      let role=''; try {{ role=node.role(); }} catch (_) {{ return; }}
      if (role === 'AXRow') {{
        let title=''; try {{ title=node.name(); }} catch (_) {{}}
        if (!title) try {{ title=node.value(); }} catch (_) {{}}
        if (title) found.push({{title:String(title)}});
      }}
      try {{ node.UIElements().forEach(child=>walk(child, depth+1)); }} catch (_) {{}}
    }};
    walk(w,0); JSON.stringify({{state:'ok', rows:found}});
  }}
  }}
  "
    )
}

fn decoded(reply: &UiReply) -> Result<Value, LocalError> {
    serde_json::from_str(&reply.stdout).map_err(|_| LocalError::UiUnsupported)
}
fn state(value: &Value) -> Result<(), LocalError> {
    match value["state"].as_str() {
        Some("ok") => Ok(()),
        Some("outlook_not_running") => Err(LocalError::OutlookNotRunning),
        Some("ui_changed") => Err(LocalError::UiChanged),
        Some("accessibility_untrusted") => Err(LocalError::AccessibilityUntrusted),
        _ => Err(LocalError::UiUnsupported),
    }
}

/// # Errors
/// Returns a typed UI availability or accessibility error.
pub fn doctor(ui: &dyn OutlookUi) -> Result<Value, LocalError> {
    let reply = ui.run(UiRequest::Doctor)?;
    let value = decoded(&reply)?;
    state(&value)?;
    Ok(
        json!({"command":"outlook.doctor","provider":"local-macos","platform":"macos","supported":true,"bundle_id":"com.microsoft.Outlook","version":value["version"],"running":true,"ui_available":value["ui_available"],"ax_preflight":"usable","network":false,"auth_calls":false}),
    )
}

/// # Errors
/// Returns a typed UI availability/accessibility error or rejects a limit above 100.
pub fn visible_mail(ui: &dyn OutlookUi, limit: u16) -> Result<Value, LocalError> {
    if limit > 100 {
        return Err(LocalError::UiUnsupported);
    }
    let reply = ui.run(UiRequest::VisibleMail { limit })?;
    let value = decoded(&reply)?;
    state(&value)?;
    let rows = value["rows"].as_array().ok_or(LocalError::UiUnsupported)?;
    if rows.is_empty() {
        return Err(LocalError::UiUnsupported);
    }
    let rows: Vec<Value> = rows.iter().take(usize::from(limit)).enumerate().map(|(index, row)| {
        json!({"summary":row["title"],"locator":format!("visible-row:{}", index + 1),"locator_scope":"ephemeral_local_visible_ui"})
    }).collect();
    Ok(
        json!({"command":"outlook.mail.visible","provider":"local-macos","source":"outlook-ui-local","content_origin":"external_email_untrusted","value":rows}),
    )
}

#[cfg(test)]
mod tests {
    use super::*;
    struct Fixture {
        reply: UiReply,
    }
    impl OutlookUi for Fixture {
        fn run(&self, _: UiRequest) -> Result<UiReply, LocalError> {
            Ok(self.reply.clone())
        }
    }
    #[test]
    fn visible_mail_is_bounded_and_has_safe_provenance() {
        let ui = Fixture {
            reply: UiReply::json(r#"{"state":"ok","rows":[{"title":"One"},{"title":"Two"}]}"#),
        };
        let output = visible_mail(&ui, 1).unwrap();
        assert_eq!(output["source"], "outlook-ui-local");
        assert_eq!(output["content_origin"], "external_email_untrusted");
        assert_eq!(output["value"].as_array().unwrap().len(), 1);
        assert!(
            output["value"][0]["locator"]
                .as_str()
                .unwrap()
                .starts_with("visible-row:")
        );
    }
    #[test]
    fn missing_rows_is_typed_ui_unsupported() {
        let ui = Fixture {
            reply: UiReply::json(r#"{"state":"ok","rows":[]}"#),
        };
        assert_eq!(visible_mail(&ui, 25).unwrap_err().code(), "ui_unsupported");
    }
    #[test]
    fn doctor_does_not_return_mail_data() {
        let ui = Fixture {
            reply: UiReply::json(r#"{"state":"ok","ui_available":true}"#),
        };
        let result = doctor(&ui).unwrap();
        assert!(result.get("rows").is_none());
        assert_eq!(result["auth_calls"], false);
    }
}
