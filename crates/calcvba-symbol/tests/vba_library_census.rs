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

#[test]
fn intrinsic_parameter_names_follow_the_captured_typelib() {
    use calcvba_bundle::NativeImplId;
    use calcvba_symbol::catalog::{intrinsic_entry, name_to_intrinsic};
    let mut checked = 0;
    for row in FUNCTIONS.lines().skip(1) {
        let columns: Vec<_> = row.split('\t').collect();
        if matches!(
            columns[0],
            "Collection" | "ErrObject" | "RegExp" | "Match" | "MatchCollection" | "SubMatches"
        ) {
            continue;
        }
        if columns[3] != "function" || columns[7].contains("ParamArray") {
            continue;
        }
        let id = match columns[1] {
            "IIf" => Some(NativeImplId::IIf),
            name => name_to_intrinsic(name),
        };
        let Some(id) = id else {
            continue;
        }; // Class members and structural pointer helpers.
        let entry = intrinsic_entry(id);
        if matches!(
            id,
            NativeImplId::Abs
                | NativeImplId::Fix
                | NativeImplId::Int
                | NativeImplId::Sgn
                | NativeImplId::InStr
                | NativeImplId::InStrB
                | NativeImplId::StrComp
        ) {
            assert!(
                entry.param_names.is_empty(),
                "{} refuses named arguments in Excel",
                columns[1]
            );
            continue;
        }
        let Some(maximum) = entry.sig.max_args else {
            continue;
        };
        // CDec's extra pvar is an ABI output; source arity excludes it.
        let expected: Vec<_> = columns[7]
            .split(", ")
            .filter_map(|p| {
                p.split_once(':')
                    .map(|(name, _)| name.trim_start_matches('?'))
            })
            .take(maximum as usize)
            .collect();
        assert_eq!(entry.param_names, expected, "{}.{}", columns[0], columns[1]);
        checked += 1;
    }
    assert!(checked > 120, "only {checked} callable signatures checked");
}
