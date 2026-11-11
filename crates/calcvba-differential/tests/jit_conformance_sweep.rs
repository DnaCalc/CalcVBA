//! Sweep: every single-module .bas fixture under conformance/ runs on both executors; any
//! difference (JIT declines, different value/error) is listed. Ignored by default (it is a
//! survey, not an assertion): `cargo test -p calcvba-differential --test jit_conformance_sweep -- --ignored --nocapture`.

use calcvba_differential::{Executor, run};
use std::path::{Path, PathBuf};

fn collect(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(rd) = std::fs::read_dir(dir) else { return };
    let mut entries: Vec<_> = rd.filter_map(|e| e.ok()).map(|e| e.path()).collect();
    entries.sort();
    for p in entries {
        if p.is_dir() {
            if p.file_name().is_some_and(|n| n == "integration" || n == "com") {
                continue;
            }
            collect(&p, out);
        } else if p.extension().is_some_and(|e| e == "bas") {
            out.push(p);
        }
    }
}

fn source(path: &Path) -> String {
    let bytes = std::fs::read(path).unwrap();
    String::from_utf8(bytes.clone()).unwrap_or_else(|_| bytes.iter().map(|b| char::from(*b)).collect())
}

#[test]
#[ignore = "survey"]
fn sweep_conformance_bas_files_on_both_executors() {
    // VBA_SWEEP_DIR points the survey at another tree (e.g. DNA Calc's corpus modules).
    let root = std::env::var_os("VBA_SWEEP_DIR")
        .map(PathBuf::from)
        .unwrap_or_else(|| Path::new(env!("CARGO_MANIFEST_DIR")).join("../../conformance"));
    let mut files = Vec::new();
    collect(&root, &mut files);
    let (mut same, mut declined, mut differ, mut both_unsupported) = (0, 0, 0, 0);
    for f in &files {
        let src = source(f);
        let name = f.strip_prefix(&root).unwrap().display().to_string();
        let handle = std::thread::Builder::new().stack_size(64 << 20).spawn(move || {
            let v = run(Executor::Vm3, &src);
            let j = run(Executor::Jit, &src);
            (v, j)
        }).unwrap();
        let Ok((v, j)) = handle.join() else { println!("PANIC\t{name}"); differ += 1; continue };
        if v.unsupported.is_some() {
            both_unsupported += 1;
            if j.unsupported.is_none() { println!("VM3-UNSUPPORTED-ONLY\t{name}\t{:?}", v.unsupported); }
            continue;
        }
        if let Some(why) = &j.unsupported {
            declined += 1;
            println!("DECLINE\t{name}\t{}", why.chars().take(1500).collect::<String>());
            continue;
        }
        if v.result == j.result && v.raised == j.raised && v.err == j.err {
            same += 1;
        } else {
            differ += 1;
            println!("DIFFER\t{name}\n  vm3: {:?} raised={} err={:?}\n  jit: {:?} raised={} err={:?}", v.result, v.raised, v.err, j.result, j.raised, j.err);
        }
    }
    println!("files {} same {same} declined {declined} differ {differ} vm3-unsupported {both_unsupported}", files.len());
}
