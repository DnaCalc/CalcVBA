use calcvba_host::portability::{Host, check_source};

#[test]
fn versioned_report_attributes_sites_to_procedures() {
    let report = check_source(
        "Main",
        "Sub Safe()\nDebug.Print Len(\"x\")\nEnd Sub\nSub RunCommand()\nShell \"example\"\nEnd Sub\n",
        Host::Linux,
    )
    .unwrap();
    let json = serde_json::to_value(report).unwrap();
    let shell = json["findings"]
        .as_array()
        .unwrap()
        .iter()
        .find(|row| row["feature"] == "Shell")
        .unwrap();
    assert_eq!(shell["line"], 5, "{shell}");
    assert_eq!(shell["procedure"], "RunCommand", "{shell}");
    assert_eq!(shell["verdict"], "reject", "{shell}");
    assert_eq!(json["schema_version"], 1, "{json}");
    assert_eq!(json["analysis_complete"], true, "{json}");
}

#[test]
fn report_marks_analysis_incomplete_when_source_does_not_parse() {
    let report = check_source("Main", "Sub Broken()\nx = (1 +\nEnd Sub\n", Host::Linux).unwrap();
    let json = serde_json::to_value(report).unwrap();
    assert!(
        !json["parse_errors"].as_array().unwrap().is_empty(),
        "{json}"
    );
    assert_eq!(json["analysis_complete"], false, "{json}");
}
