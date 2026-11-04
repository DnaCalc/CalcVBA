use calcvba_office::{OfficeFile, Project};
use std::{collections::BTreeMap, path::Path};

pub fn run(command: &str, args: &[String]) -> Result<(), String> {
    if args
        .first()
        .is_some_and(|s| matches!(s.as_str(), "--help" | "-h"))
    {
        println!("{}", usage(command));
        return Ok(());
    }
    let input = args
        .first()
        .filter(|s| !s.starts_with('-'))
        .ok_or_else(|| usage(command))?;
    let mut flags = BTreeMap::new();
    let mut at = 1;
    while at < args.len() {
        let key = args[at].as_str();
        let allowed = match command {
            "export-vba" => key == "--out-dir",
            "write-vba" => matches!(key, "--project" | "--out"),
            _ => false,
        };
        if !allowed {
            return Err(format!("unknown argument {key}\n{}", usage(command)));
        }
        let value = args
            .get(at + 1)
            .filter(|s| !s.starts_with('-'))
            .ok_or_else(|| usage(command))?;
        if flags.insert(key, value.as_str()).is_some() {
            return Err(format!("duplicate argument {key}"));
        }
        at += 2;
    }
    let required = |key: &str| {
        flags
            .get(key)
            .copied()
            .ok_or_else(|| format!("missing {key}\n{}", usage(command)))
    };
    if command == "export-vba" {
        required("--out-dir")?;
    }
    if command == "write-vba" {
        required("--project")?;
        required("--out")?;
    }
    let mut office = OfficeFile::open(input).map_err(|e| e.to_string())?;
    match command {
        "inspect-vba" => {
            let p = &office.project;
            println!("{}",serde_json::to_string_pretty(&serde_json::json!({
                "name":p.properties.name,"codepage":p.properties.codepage,"description":p.properties.description,
                "constants":p.properties.constants,"sys_kind":p.properties.sys_kind,
                "lcid":p.properties.lcid,"lcid_invoke":p.properties.lcid_invoke,
                "version":{"major":p.properties.version_major,"minor":p.properties.version_minor},
                "modules":p.modules.iter().map(|m|serde_json::json!({
                    "name":m.name,"kind":format!("{:?}",m.kind),"stream_name":m.stream_name,
                    "source_bytes_utf8":m.source.len(),"read_only":m.read_only,"private":m.private
                })).collect::<Vec<_>>(),
                "references":p.references.iter().map(reference_json).collect::<Vec<_>>()
            })).map_err(|e|e.to_string())?);
        }
        "export-vba" => {
            let path = office
                .project
                .export_folder(Path::new(required("--out-dir")?))
                .map_err(|e| e.to_string())?;
            println!("{}", path.display());
        }
        "write-vba" => {
            let project = required("--project")?;
            let output = required("--out")?;
            office.project = Project::from_folder(project).map_err(|e| e.to_string())?;
            office.save_as(output).map_err(|e| e.to_string())?;
            println!("{output}");
        }
        _ => return Err(usage(command)),
    }
    Ok(())
}
fn reference_json(reference: &calcvba_office::Reference) -> serde_json::Value {
    use calcvba_office::ReferenceKind;
    match &reference.kind {
        ReferenceKind::Registered { libid } => {
            serde_json::json!({"name":reference.name,"kind":"registered","libid":libid})
        }
        ReferenceKind::Project {
            absolute,
            relative,
            major,
            minor,
        } => {
            serde_json::json!({"name":reference.name,"kind":"project","absolute":absolute,"relative":relative,"major":major,"minor":minor})
        }
        ReferenceKind::Control {
            original,
            twiddled,
            extended_name,
            extended,
            guid,
            cookie,
        } => {
            serde_json::json!({"name":reference.name,"kind":"control","original":original,"twiddled":twiddled,"extended_name":extended_name,"extended":extended,"guid":guid,"cookie":cookie})
        }
    }
}
fn usage(command: &str) -> String {
    match command {
        "inspect-vba"=>"usage: calcvba-cli inspect-vba <workbook> (JSON output)",
        "export-vba"=>"usage: calcvba-cli export-vba <workbook> --out-dir <new-folder>",
        _=>"usage: calcvba-cli write-vba <workbook> --project <project.basproj> --out <new-workbook>",
    }.into()
}
