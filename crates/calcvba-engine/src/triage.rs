//! Capability triage of a project closure: bind and elaborate it, read the predicted
//! capabilities from OxIR, and fall back to a lexical scan of the sources when it does not bind.

use calcvba_lib::{Prediction, predicted_capabilities, predicted_capabilities_lexical};
use calcvba_symbol::CatalogTypeLibResolver;
use calcvba_symbol::manifest::SymbolProjectManifest;

/// The prediction for a whole closure, and why the lexical fallback was used if it was.
#[derive(Debug, Clone)]
pub struct ClosurePrediction {
    pub prediction: Prediction,
    /// Set when binding or elaboration failed and the lexical scan was used instead.
    pub fallback_reason: Option<String>,
}

/// Predict the capabilities of `closure` (leaf-first, entry last). Never executes anything.
pub fn predict_project_closure(closure: &[SymbolProjectManifest]) -> ClosurePrediction {
    // Triage must never load workbook-selected native type libraries.
    let _scope =
        calcvba_com::ComPolicyScope::enter(|_, _| Err("native COM disabled for triage".into()));
    let lexical = |reason: String| {
        let sources: Vec<&str> = closure
            .iter()
            .flat_map(|manifest| manifest.modules.iter().map(|m| m.source.as_str()))
            .collect();
        ClosurePrediction {
            prediction: predicted_capabilities_lexical(&sources),
            fallback_reason: Some(reason),
        }
    };
    let programs = match calcvba_bind::bind_projects(closure, &CatalogTypeLibResolver) {
        Ok(programs) => programs,
        Err(err) => return lexical(format!("bind: {err}")),
    };
    let mut merged: Option<Prediction> = None;
    for program in &programs {
        let ox = match calcvba_oxir::elaborate::elaborate(program) {
            Ok(ox) => ox,
            Err(err) => return lexical(format!("elaborate: {err}")),
        };
        let one = predicted_capabilities(&ox);
        match merged.as_mut() {
            Some(total) => total.merge(&one),
            None => merged = Some(one),
        }
    }
    match merged {
        Some(prediction) => ClosurePrediction {
            prediction,
            fallback_reason: None,
        },
        None => lexical("no programs".to_string()),
    }
}

/// Predict a single source module run as project `Main`.
pub fn predict_source(source: &str) -> ClosurePrediction {
    use calcvba_symbol::manifest as sym;
    let manifest = sym::SymbolProjectManifest {
        project_name: "Main".to_string(),
        project_kind: sym::ProjectKind::Source,
        modules: vec![sym::ModuleUnit {
            module_name: "Main".to_string(),
            module_kind: sym::ModuleKind::Procedural,
            attributes: sym::ModuleAttributes::named("Main"),
            source: source.to_string(),
        }],
        references: Vec::new(),
        reference_projects: Vec::new(),
        conditional_constants: std::collections::BTreeMap::new(),
        conditional_compilation_target: Default::default(),
    };
    predict_project_closure(&[manifest])
}
