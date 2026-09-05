use moff_outlook_local::{LocalError, parse_visible_rows};

const STRUCTURED: &str = r#"[
 {"text":"Alice Example","x":0.10,"y":0.86,"w":0.20,"h":0.02},
 {"text":"Quarterly status","x":0.10,"y":0.83,"w":0.24,"h":0.02},
 {"text":"Agenda attached","x":0.10,"y":0.80,"w":0.30,"h":0.02},
 {"text":"10:42 AM","x":0.64,"y":0.86,"w":0.08,"h":0.02},
 {"text":"Bob Example","x":0.10,"y":0.70,"w":0.20,"h":0.02},
 {"text":"Planning","x":0.10,"y":0.67,"w":0.20,"h":0.02},
 {"text":"Tomorrow works","x":0.10,"y":0.64,"w":0.25,"h":0.02},
 {"text":"Yesterday","x":0.64,"y":0.70,"w":0.10,"h":0.02},
 {"text":"OPEN MAIL BODY MUST NOT LEAK","x":0.92,"y":0.50,"w":0.05,"h":0.02}
]"#;

#[test]
fn parses_only_high_confidence_bounded_message_rows() {
    let rows = parse_visible_rows(STRUCTURED, 25).unwrap();
    assert_eq!(rows.len(), 2);
    assert_eq!(rows[0].sender, "Alice Example");
    assert_eq!(rows[0].subject, "Quarterly status");
    assert_eq!(rows[0].preview, "Agenda attached");
    assert_eq!(rows[0].time, "10:42 AM");
    assert!(rows[0].unread.is_none());
    assert_eq!(rows[0].locator, "capture-row:1");
    assert!(!format!("{rows:?}").contains("OPEN MAIL BODY"));
}

#[test]
fn rejects_unstructured_ocr_instead_of_returning_arbitrary_lines() {
    assert_eq!(
        parse_visible_rows(
            r#"[{"text":"random text","x":0.2,"y":0.3,"w":0.2,"h":0.02}]"#,
            25
        )
        .unwrap_err(),
        LocalError::UiChanged
    );
}

#[test]
fn rejects_limits_over_one_hundred() {
    assert_eq!(
        parse_visible_rows(STRUCTURED, 101).unwrap_err(),
        LocalError::InvalidLimit
    );
}
