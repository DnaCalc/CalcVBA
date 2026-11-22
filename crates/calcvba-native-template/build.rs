fn main() {
    let os = std::env::var("CARGO_CFG_TARGET_OS").unwrap();
    let windows = os == "windows";
    let linux = os == "linux"
        && std::env::var("CARGO_CFG_TARGET_ARCH").as_deref() == Ok("x86_64")
        && std::env::var("CARGO_CFG_TARGET_ENV").as_deref() == Ok("gnu");
    if !windows && !linux {
        return;
    }
    if windows {
        // Developer link only: leave room for generated PE section headers.
        for flag in ["/FILEALIGN:4096", "/GUARD:NO", "/INCREMENTAL:NO"] {
            println!("cargo:rustc-cdylib-link-arg={flag}");
            println!("cargo:rustc-link-arg-bin=calcvba-native-exe-template={flag}");
            println!("cargo:rustc-link-arg-bin=calcvba-native-com-exe-template={flag}");
        }
    } else {
        // Preserve symbol indices and use the SysV hash format, which can be
        // rebuilt without ordering the symbols used by dynamic relocations.
        println!("cargo:rustc-cdylib-link-arg=-Wl,--hash-style=sysv");
        println!("cargo:rustc-cdylib-link-arg=-Wl,-Bsymbolic");
        println!("cargo:rustc-link-arg-bin=calcvba-native-exe-template=-Wl,--hash-style=sysv");
    }
    let root = std::path::PathBuf::from(std::env::var_os("CARGO_MANIFEST_DIR").unwrap());
    let contract = root.join(if windows {
        "../../assets/native/helpers-v1.txt"
    } else {
        "../../assets/native/helpers-linux-x64-v1.txt"
    });
    println!("cargo:rerun-if-changed={}", contract.display());
    let text = std::fs::read_to_string(contract).expect("native helper contract");
    let names: Vec<_> = text
        .lines()
        .skip(1)
        .map(|line| line.split('|').next().unwrap())
        .collect();
    let mut source = String::from("pub fn retain_native_helpers() { unsafe extern \"C\" {\n");
    for name in &names {
        if windows {
            println!("cargo:rustc-link-arg-bin=calcvba-native-exe-template=/EXPORT:{name}");
            println!("cargo:rustc-link-arg-bin=calcvba-native-com-exe-template=/EXPORT:{name}");
        }
        source.push_str(&format!("fn {name}();\n"));
    }
    if windows {
        println!(
            "cargo:rustc-link-arg-bin=calcvba-native-exe-template=/EXPORT:cvb_template_abi,DATA"
        );
        println!(
            "cargo:rustc-link-arg-bin=calcvba-native-com-exe-template=/EXPORT:cvb_template_abi,DATA"
        );
    }
    source.push_str("}\nstd::hint::black_box([\n");
    for name in &names {
        source.push_str(&format!("{name} as *const (),\n"));
    }
    source.push_str("]);\nstd::hint::black_box(std::ptr::addr_of!(CVB_TEMPLATE_ABI));\n}\n");
    let output = std::path::PathBuf::from(std::env::var_os("OUT_DIR").unwrap());
    std::fs::write(output.join("native_keepalive.rs"), source)
        .expect("native helper retention source");
}
