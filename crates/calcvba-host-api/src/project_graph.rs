//! Ordered source-project closures; the original Project descriptor is unchanged.
use crate::{Project, TypeLibraryReference};

#[derive(Clone)]
pub enum ProjectReference {
    Source(String),
    TypeLibrary(TypeLibraryReference),
}
#[derive(Clone)]
pub struct ProjectNode {
    pub project: Project,
    /// Ordered references. When nonempty, Project.references must be empty.
    /// Typed object libraries are supplied separately to load_graph.
    pub references: Vec<ProjectReference>,
}
#[derive(Clone)]
pub struct ProjectGraph {
    pub root: String,
    pub projects: Vec<ProjectNode>,
}
