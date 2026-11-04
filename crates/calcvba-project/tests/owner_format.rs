use calcvba_project::{Instancing, load_basproj_from_str, parse_basproj_xml};
#[test]
fn duplicate_vb_attribute_metadata_is_refused_even_when_empty() {
    for name in [
        "VBExposed",
        "VBCreatable",
        "VBPredeclaredId",
        "VBGlobalNamespace",
        "VB_Name",
        "Description",
    ] {
        for value in [
            format!("<{name}/>"),
            format!("<{name}></{name}>"),
            format!("<{name}>True</{name}>"),
        ] {
            let xml = format!(
                "<Project Sdk=\"CalcVBA.Sdk\"><PropertyGroup><OutputType>Library</OutputType></PropertyGroup><ItemGroup><ClassModule Include=\"C.cls\">{value}</ClassModule></ItemGroup></Project>"
            );
            assert!(parse_basproj_xml(&xml).is_err(), "{xml}");
        }
    }
}
#[test]
fn all_six_source_instancing_values_follow_the_vb_attributes_and_output_type() {
    let root = std::env::temp_dir().join(format!("calcvba-instancing-{}", std::process::id()));
    std::fs::create_dir(&root).unwrap();
    for (exposed, creatable, global, multi, expected) in [
        (false, false, false, true, Instancing::Private),
        (true, false, false, true, Instancing::PublicNotCreatable),
        (true, true, false, false, Instancing::SingleUse),
        (true, true, true, false, Instancing::GlobalSingleUse),
        (true, true, false, true, Instancing::MultiUse),
        (true, true, true, true, Instancing::GlobalMultiUse),
    ] {
        let source = format!(
            "VERSION 1.0 CLASS\nBEGIN\n MultiUse = {}\nEND\nAttribute VB_Name = \"SemanticName\"\nAttribute VB_Exposed = {exposed}\nAttribute VB_Creatable = {creatable}\nAttribute VB_GlobalNameSpace = {global}\nAttribute VB_Description = \"Source description\"\nPublic Sub Work()\nEnd Sub\n",
            if multi { -1 } else { 0 }
        );
        std::fs::write(root.join("DifferentFilename.cls"), source).unwrap();
        let xml = "<Project Sdk=\"CalcVBA.Sdk\"><PropertyGroup><OutputType>ComExe</OutputType></PropertyGroup><ItemGroup><ClassModule Include=\"DifferentFilename.cls\"><ProgId>Example.SemanticName</ProgId></ClassModule></ItemGroup></Project>";
        let result = load_basproj_from_str(xml, &root);
        let single = matches!(
            expected,
            Instancing::SingleUse | Instancing::GlobalSingleUse
        );
        if single && !cfg!(windows) {
            assert!(result.is_err());
            continue;
        }
        let loaded = result.unwrap();
        assert_eq!(loaded.manifest.modules[0].module_name, "SemanticName");
        assert_eq!(
            loaded.manifest.modules[0]
                .attributes
                .instancing
                .map(|i| format!("{i:?}")),
            Some(format!("{expected:?}"))
        );
        assert_eq!(
            loaded.class_module_metadata["SemanticName"]
                .description
                .as_deref(),
            Some("Source description")
        );
        if single {
            assert!(load_basproj_from_str(&xml.replace("ComExe", "ComServer"), &root).is_err());
        }
    }
    std::fs::remove_dir_all(root).unwrap();
}

#[test]
fn startup_helpers_do_not_shift_source_class_packaging_metadata() {
    let root =
        std::env::temp_dir().join(format!("calcvba-startup-metadata-{}", std::process::id()));
    std::fs::create_dir(&root).unwrap();
    std::fs::write(root.join("Entry.bas"), "Attribute VB_Name = \"Entry\"\nPublic Sub Main()\nEnd Sub\nPublic Sub Alternate()\nEnd Sub\n").unwrap();
    std::fs::write(root.join("DifferentFilename.cls"), "Attribute VB_Name = \"SemanticName\"\nAttribute VB_Exposed = True\nAttribute VB_Creatable = False\nAttribute VB_Description = \"Source description\"\nPublic Sub Work()\nEnd Sub\n").unwrap();
    let xml = "<Project Sdk=\"CalcVBA.Sdk\"><PropertyGroup><OutputType>Exe</OutputType></PropertyGroup><ItemGroup><Module Include=\"Entry.bas\"/><ClassModule Include=\"DifferentFilename.cls\"><Instancing>PublicNotCreatable</Instancing><ProgId>Example.SemanticName</ProgId></ClassModule></ItemGroup></Project>";
    let path = root.join("Project.basproj");
    std::fs::write(&path, xml).unwrap();
    let loaded = load_basproj_from_str(xml, &root).unwrap();
    assert_eq!(loaded.class_module_metadata.len(), 1);
    assert_eq!(
        loaded.class_module_metadata["SemanticName"]
            .prog_id
            .as_deref(),
        Some("Example.SemanticName")
    );
    for entry in [None, Some("Entry.Alternate")] {
        let closure = calcvba_project::load_project_closure_with_entry(&path, entry).unwrap();
        let modules = &closure.last().unwrap().modules;
        let class = modules
            .iter()
            .find(|m| m.module_name == "SemanticName")
            .unwrap();
        assert_eq!(
            class.attributes.prog_id.as_deref(),
            Some("Example.SemanticName")
        );
        assert_eq!(
            class.attributes.description.as_deref(),
            Some("Source description")
        );
        assert_eq!(
            class.attributes.instancing,
            Some(calcvba_symbol::manifest::Instancing::PublicNotCreatable)
        );
        assert!(
            modules
                .iter()
                .filter(|m| m.module_name != "SemanticName")
                .all(|m| m.attributes.prog_id.is_none())
        );
    }
    std::fs::remove_dir_all(root).unwrap();
}
