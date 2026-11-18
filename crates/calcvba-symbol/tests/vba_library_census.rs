//! Public names captured from VBE7.DLL and MS-VBAL must reach the binder's resolver.
use calcvba_symbol::{
    provider::Provider, providers::vba_library::VbaLibraryProvider, signature::VarTypeRef,
};

const FUNCTIONS: &str =
    include_str!("../../../conformance/excel-measurements/vba-library/functions.tsv");
const TYPELIB: &str =
    include_str!("../../../conformance/excel-measurements/vba-library/vbe7-library.tsv");

#[test]
fn every_library_function_and_member_resolves() {
    let provider = VbaLibraryProvider;
    let mut missing = Vec::new();
    let mut count = 0;
    for row in FUNCTIONS.lines().skip(1) {
        let columns: Vec<_> = row.split('\t').collect();
        let (owner, name) = (columns[0], columns[1]);
        let class = matches!(
            owner,
            "Collection" | "ErrObject" | "RegExp" | "Match" | "MatchCollection" | "SubMatches"
        );
        for name in std::iter::once(name.to_string())
            .chain((columns[2] == "yes").then(|| format!("{name}$")))
        {
            count += 1;
            let binding = if class {
                provider.resolve_member(&VarTypeRef::Object(owner.into()), &name, None)
            } else {
                provider.resolve(&name)
            };
            if binding.is_none() {
                missing.push(format!("{owner}.{name}"));
            }
            if !class && owner != "_HiddenModule" {
                for parts in [
                    vec![owner, name.as_str()],
                    vec!["VBA", owner, name.as_str()],
                ] {
                    if provider.resolve_qualified(&parts).is_none() {
                        missing.push(parts.join("."));
                    }
                }
            }
        }
    }
    assert_eq!(count, 228);
    assert!(
        missing.is_empty(),
        "Unresolved library names:\n{}",
        missing.join("\n")
    );
}

#[test]
fn every_typelib_constant_resolves() {
    let provider = VbaLibraryProvider;
    let mut missing = Vec::new();
    let mut count = 0;
    for row in TYPELIB.lines().skip(1).filter(|row| !row.is_empty()) {
        let columns: Vec<_> = row.split('\t').collect();
        if columns[4] != "CONST" {
            continue;
        }
        count += 1;
        let name = columns[3];
        for parts in [
            vec![name],
            vec![columns[0], name],
            vec!["VBA", columns[0], name],
        ] {
            let binding = if parts.len() == 1 {
                provider.resolve(name)
            } else {
                provider.resolve_qualified(&parts)
            };
            if binding.is_none() {
                missing.push(parts.join("."));
            }
        }
    }
    assert_eq!(count, 272);
    assert!(
        missing.is_empty(),
        "Unresolved constants:\n{}",
        missing.join("\n")
    );
}
