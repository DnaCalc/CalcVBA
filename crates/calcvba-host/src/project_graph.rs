//! Compile a source closure once, preserving each project's identities and order.
use super::*;
use calcvba_host_api as api;
use std::collections::{HashMap, HashSet};

pub(super) fn manifests(
    graph: &api::ProjectGraph,
) -> Result<Vec<sym::SymbolProjectManifest>, api::Error> {
    let invalid = |message: &str| api::Error::Compile(message.into());
    let mut nodes = HashMap::new();
    for node in &graph.projects {
        let name = node.project.name.to_ascii_lowercase();
        if name.trim().is_empty() || nodes.insert(name, node).is_some() {
            return Err(invalid("empty/duplicate project identity"));
        }
        if !node.references.is_empty() && !node.project.references.is_empty() {
            return Err(invalid("graph node has two ordered reference lists"));
        }
        if node.references.is_empty() {
            let mut names = HashSet::new();
            for reference in &node.project.references {
                if reference.name.trim().is_empty()
                    || !names.insert(reference.name.to_ascii_lowercase())
                {
                    return Err(invalid("empty/duplicate type-library reference"));
                }
            }
        }
    }
    let mut order = Vec::new();
    let mut visited = HashSet::new();
    let mut visiting = HashSet::new();
    fn visit<'a>(
        name: &str,
        nodes: &HashMap<String, &'a api::ProjectNode>,
        visited: &mut HashSet<String>,
        visiting: &mut HashSet<String>,
        order: &mut Vec<&'a api::ProjectNode>,
    ) -> Result<(), api::Error> {
        let key = name.to_ascii_lowercase();
        if visited.contains(&key) {
            return Ok(());
        }
        if !visiting.insert(key.clone()) {
            return Err(api::Error::Compile(
                "cyclic source-project references".into(),
            ));
        }
        let node = *nodes
            .get(&key)
            .ok_or_else(|| api::Error::Compile(format!("missing referenced project {name}")))?;
        let mut references = HashSet::new();
        for reference in &node.references {
            let name = match reference {
                api::ProjectReference::Source(name) => name,
                api::ProjectReference::TypeLibrary(r) => &r.name,
            };
            if name.trim().is_empty() || !references.insert(name.to_ascii_lowercase()) {
                return Err(api::Error::Compile(
                    "empty/duplicate ordered reference".into(),
                ));
            }
            if let api::ProjectReference::Source(name) = reference {
                visit(name, nodes, visited, visiting, order)?;
            }
        }
        visiting.remove(&key);
        visited.insert(key);
        order.push(node);
        Ok(())
    }
    visit(&graph.root, &nodes, &mut visited, &mut visiting, &mut order)?;
    let sources: Vec<_> = order
        .iter()
        .map(|node| sym::ReferencedProjectManifest {
            project_name: node.project.name.clone(),
            project_kind: sym::ProjectKind::Library,
            modules: node
                .project
                .modules
                .iter()
                .map(|module| sym::ModuleUnit {
                    module_name: module.name.clone(),
                    module_kind: sym::ModuleKind::Procedural,
                    attributes: sym::ModuleAttributes::named(&module.name),
                    source: module.source.clone(),
                })
                .collect(),
        })
        .collect();
    let typelib = |r: &api::TypeLibraryReference| sym::ProjectReference::TypeLibrary {
        name: r.name.clone(),
        guid: r.libid.clone(),
        version_major: r.major,
        version_minor: r.minor,
        lcid: r.lcid,
        import_lib: r.import_library.clone(),
    };
    Ok(order
        .iter()
        .map(|node| sym::SymbolProjectManifest {
            project_name: node.project.name.clone(),
            project_kind: sym::ProjectKind::Library,
            modules: sources
                .iter()
                .find(|s| s.project_name == node.project.name)
                .expect("visited source")
                .modules
                .clone(),
            references: if node.references.is_empty() {
                node.project.references.iter().map(&typelib).collect()
            } else {
                node.references
                    .iter()
                    .map(|r| match r {
                        api::ProjectReference::Source(name) => sym::ProjectReference::Project {
                            referenced_project_name: name.clone(),
                        },
                        api::ProjectReference::TypeLibrary(r) => typelib(r),
                    })
                    .collect()
            },
            reference_projects: sources
                .iter()
                .filter(|s| !s.project_name.eq_ignore_ascii_case(&node.project.name))
                .cloned()
                .collect(),
            conditional_constants: Default::default(),
            conditional_compilation_target: Default::default(),
        })
        .collect())
}
