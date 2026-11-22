//! Compare the entry of a reviewed OxIR project on the two execution backends.

use std::sync::{Arc, Mutex};

use calcvba_hal::{
    HostOutputChannel, HostPolicy, HostServices, PolicyHost, callbacks::HostCallbacks,
    install_thread_output_tap,
};
use calcvba_oxir::OxProgram;

use crate::{Canon, Executor, RunOutcome, run_ox_programs_on_host, vm3_oracle_engine};

#[derive(Debug)]
pub struct Observation {
    pub outcome: RunOutcome,
    pub output: Vec<(HostOutputChannel, String)>,
}

struct RecordedOutput;

impl HostCallbacks for RecordedOutput {
    fn on_msg_box(&self, _prompt: &str, style: i32) -> i32 {
        style.max(1)
    }

    fn on_input_box(&self, _prompt: &str, default: &str) -> String {
        default.to_owned()
    }

    fn on_status_bar(&self, _text: &str) {}

    fn on_console_print(&self, _text: &str) -> bool {
        true
    }

    fn on_debug_print(&self, _text: &str) {}
}

fn observe(executor: Executor, programs: &[OxProgram]) -> Observation {
    let mut engine = vm3_oracle_engine();
    engine.set_host_policy(HostPolicy {
        allow_process_spawn: false,
        allow_com_activation: false,
        ..HostPolicy::default()
    });
    engine.set_host_callbacks(Some(Arc::new(RecordedOutput)));
    let host = PolicyHost::safe(engine.host_services());
    let _policy = host.enter_com_policy();
    let output = Arc::new(Mutex::new(Vec::new()));
    let observer = output.clone();
    let tap = install_thread_output_tap(Arc::new(move |channel, text: &str| {
        observer.lock().unwrap().push((channel, text.to_owned()));
    }));
    let outcome = run_ox_programs_on_host(executor, programs, &host, true);
    drop(tap);
    let output = std::mem::take(&mut *output.lock().unwrap());
    Observation { outcome, output }
}

#[derive(Debug)]
pub struct Comparison {
    pub vm3: Observation,
    pub jit: Observation,
}

impl Comparison {
    pub fn differences(&self) -> Vec<&'static str> {
        let vm3 = &self.vm3.outcome;
        let jit = &self.jit.outcome;
        if [vm3, jit]
            .iter()
            .any(|run| run.unsupported.is_some() || (!run.raised && run.result.is_err()))
        {
            return Vec::new();
        }
        let mut differences = Vec::new();
        if vm3.raised != jit.raised {
            differences.push("error propagation");
        }
        if vm3.err != jit.err {
            differences.push("Err");
        }
        if let (Ok(vm3_values), Ok(jit_values)) = (&vm3.result, &jit.result)
            && vm3_values != jit_values
        {
            differences.push("values");
        }
        if self.vm3.output != self.jit.output {
            differences.push("host output");
        }
        differences
    }

    pub fn has_opaque_values(&self) -> bool {
        [&self.vm3, &self.jit].iter().any(|run| {
            run.outcome.result.as_ref().is_ok_and(|values| {
                values
                    .iter()
                    .any(|value| matches!(value, Canon::Opaque { .. }))
            })
        })
    }
}

/// Capture entry locals, project globals, Err and host output for a reviewed project.
pub fn compare_entry(programs: &[OxProgram]) -> Comparison {
    Comparison {
        vm3: observe(Executor::Vm3, programs),
        jit: observe(Executor::Jit, programs),
    }
}
