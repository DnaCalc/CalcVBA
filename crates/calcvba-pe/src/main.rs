use calcvba_pe::{
    Export, HELPER_ABI, Program, Template, Unwind, finish_template, shipped_helper_specs,
};
use std::{error::Error, fs, path::Path};

fn run() -> Result<(), Box<dyn Error>> {
    let args: Vec<_> = std::env::args_os().skip(1).collect();
    if args.len() != 3 {
        return Err("usage: calcvba-pe <finish|stub> <input-template.dll> <output.dll>".into());
    }
    let input = fs::read(&args[1])?;
    let result = match args[0].to_str() {
        Some("finish") => finish_template(&input, &shipped_helper_specs())?,
        Some("stub") => {
            let template = Template::parse(&input, HELPER_ABI)?;
            let program = Program {
                // Return42 is a leaf. WalkFrame calls its callback with proper
                // Win64 shadow space and a real UNWIND_INFO allocation record.
                text: vec![
                    0xb8, 42, 0, 0, 0, 0xc3, 0x90, 0x90, 0x48, 0x83, 0xec, 0x28, 0xff, 0xd1, 0x48,
                    0x83, 0xc4, 0x28, 0xc3,
                ],
                exports: vec![
                    Export {
                        name: "Return42".into(),
                        ordinal: Some(7),
                        offset: 0,
                    },
                    Export {
                        name: "WalkFrame".into(),
                        ordinal: None,
                        offset: 8,
                    },
                ],
                unwind: vec![Unwind {
                    begin: 8,
                    end: 19,
                    info: vec![1, 4, 1, 0, 4, 0x42, 0, 0],
                }],
                ..Program::default()
            };
            let name = Path::new(&args[2])
                .file_name()
                .and_then(|v| v.to_str())
                .ok_or("invalid output DLL name")?;
            let patched = template.patch(name, &program)?;
            for (name, ordinal, rva) in patched.map {
                println!("{name}\t#{ordinal}\t{rva:#010x}");
            }
            patched.bytes
        }
        _ => return Err("expected finish or stub".into()),
    };
    fs::write(&args[2], result)?;
    Ok(())
}
fn main() {
    if let Err(error) = run() {
        eprintln!("{error}");
        std::process::exit(1);
    }
}
