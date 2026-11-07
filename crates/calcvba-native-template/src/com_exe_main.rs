//! Out-of-process COM startup. All VBA entries are already compiled in this
//! image; this binary contains no source loader, interpreter or JIT.
#[cfg(windows)]
mod server;
#[cfg(windows)]
fn main() -> std::process::ExitCode {
    calcvba_native_template::retain_native_helpers();
    match server::run() {
        Ok(()) => std::process::ExitCode::SUCCESS,
        Err(error) => {
            eprintln!("CalcVBA COM server: {error}");
            std::process::ExitCode::FAILURE
        }
    }
}
#[cfg(not(windows))]
fn main() {
    eprintln!("CalcVBA COM executable templates require Windows");
    std::process::exit(1);
}
