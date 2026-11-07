//! Project context shared by the native census examples: a `Main` entry with the
//! sibling modules and referenced projects that the conformance fixtures keep
//! next to it (`projects/<case>/main`, `projects/<case>/references/<lib>`).
use std::path::{Path, PathBuf};

pub const PROBE: &str = "CvbCensusProbe";

/// One module of a source project: name, class or procedural, source text.
pub struct ProjectModule {
    name: String,
    class: bool,
    text: String,
}
/// A `Main` entry together with the sibling modules and referenced projects the
/// conformance fixture keeps next to it (`projects/<case>/main`, `references/<lib>`).
pub struct Project {
    modules: Vec<ProjectModule>,
    references: Vec<(String, Vec<ProjectModule>)>,
}
fn read_modules(directory: &Path) -> Result<Vec<ProjectModule>, String> {
    let mut modules = Vec::new();
    let mut files: Vec<_> = std::fs::read_dir(directory)
        .map_err(|e| e.to_string())?
        .filter_map(|e| e.ok().map(|e| e.path()))
        .filter(|p| p.extension().is_some_and(|x| x.eq_ignore_ascii_case("bas")))
        .collect();
    files.sort();
    for file in files {
        let stem = file.file_stem().unwrap().to_string_lossy().to_string();
        let (name, kind) = stem.split_once('.').unwrap_or((&stem, "proc"));
        if !matches!(kind, "proc" | "class") {
            return Err(format!(
                "project module {stem} is a {kind} module; the census builds procedural and class modules only"
            ));
        }
        modules.push(ProjectModule {
            name: name.into(),
            class: kind == "class",
            text: std::fs::read_to_string(&file).map_err(|e| e.to_string())?,
        });
    }
    Ok(modules)
}
pub fn project_context(path: &Path) -> Result<Option<Project>, String> {
    let is_entry = path
        .file_name()
        .is_some_and(|n| n.to_string_lossy().eq_ignore_ascii_case("Main.proc.bas"));
    let Some(main) = path
        .parent()
        .filter(|p| is_entry && p.file_name().is_some_and(|n| n == "main"))
    else {
        return Ok(None);
    };
    let modules = read_modules(main)?;
    let references_dir = main.parent().map(|p| p.join("references"));
    let mut references = Vec::new();
    if let Some(dir) = references_dir.filter(|d| d.is_dir()) {
        let mut libraries: Vec<_> = std::fs::read_dir(dir)
            .map_err(|e| e.to_string())?
            .filter_map(|e| e.ok().map(|e| e.path()))
            .filter(|p| p.is_dir())
            .collect();
        libraries.sort();
        for library in libraries {
            references.push((
                library.file_name().unwrap().to_string_lossy().to_string(),
                read_modules(&library)?,
            ));
        }
    }
    if modules.len() < 2 && references.is_empty() {
        return Ok(None);
    }
    Ok(Some(Project {
        modules,
        references,
    }))
}
fn module_file(module: &ProjectModule) -> String {
    format!(
        "{}.{}",
        module.name,
        if module.class { "cls" } else { "bas" }
    )
}
fn module_item(module: &ProjectModule) -> String {
    let element = if module.class {
        "ClassModule"
    } else {
        "Module"
    };
    format!("<{element} Include=\"{}\"/>", module_file(module))
}
fn module_text(module: &ProjectModule, text: &str) -> String {
    let header = if module.class {
        format!(
            "VERSION 1.0 CLASS\nBEGIN\n  MultiUse = -1\nEND\nAttribute VB_Name = \"{}\"\n",
            module.name
        )
    } else {
        format!("Attribute VB_Name = \"{}\"\n", module.name)
    };
    format!("{header}{text}")
}
impl Project {
    /// Writes the project (entry module `Main` carries `entry_text`) below
    /// `directory` and returns the root project file; `exports` adds the probe.
    pub fn write(
        &self,
        directory: &Path,
        entry_text: &str,
        exports: bool,
    ) -> Result<PathBuf, String> {
        std::fs::create_dir_all(directory).map_err(|e| e.to_string())?;
        let mut items = String::new();
        for module in &self.modules {
            let text = if module.name.eq_ignore_ascii_case("Main") {
                entry_text
            } else {
                &module.text
            };
            std::fs::write(
                directory.join(module_file(module)),
                module_text(module, text),
            )
            .map_err(|e| e.to_string())?;
            items.push_str(&module_item(module));
        }
        for (library, modules) in &self.references {
            let sub = directory.join(format!("ref_{library}"));
            std::fs::create_dir_all(&sub).map_err(|e| e.to_string())?;
            let mut library_items = String::new();
            for module in modules {
                std::fs::write(
                    sub.join(module_file(module)),
                    module_text(module, &module.text),
                )
                .map_err(|e| e.to_string())?;
                library_items.push_str(&module_item(module));
            }
            std::fs::write(
                sub.join(format!("{library}.basproj")),
                format!(
                    "<Project><PropertyGroup><OutputType>Library</OutputType><ProjectName>{library}</ProjectName></PropertyGroup><ItemGroup>{library_items}</ItemGroup></Project>"
                ),
            )
            .map_err(|e| e.to_string())?;
            items.push_str(&format!(
                "<ProjectReference Include=\"ref_{library}/{library}.basproj\"/>"
            ));
        }
        if exports {
            items.push_str(&format!(
                "<NativeExport Include=\"CensusProbe\"><Module>Main</Module><Procedure>{PROBE}</Procedure></NativeExport>"
            ));
        }
        let project = directory.join("Census.basproj");
        std::fs::write(
            &project,
            format!(
                "<Project><PropertyGroup><OutputType>NativeLibrary</OutputType><ProjectName>Census</ProjectName></PropertyGroup><ItemGroup>{items}</ItemGroup></Project>"
            ),
        )
        .map_err(|e| e.to_string())?;
        Ok(project)
    }
}
