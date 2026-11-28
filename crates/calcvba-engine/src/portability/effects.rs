use std::collections::BTreeSet;

use calcvba_symbol::portability::{EntryStatus, PortabilityEffect, ProcedureEffects, Support};

#[derive(Default)]
pub(super) struct Block {
    pub successors: Vec<usize>,
    pub calls: Vec<usize>,
    pub may_calls: Vec<usize>,
    pub may: BTreeSet<PortabilityEffect>,
    pub must: BTreeSet<PortabilityEffect>,
    pub unknown: bool,
}

pub(super) struct Procedure {
    pub project: String,
    pub module: String,
    pub name: String,
    pub entry: usize,
    pub blocks: Vec<Block>,
    pub dynamic_control: bool,
}

#[derive(Default, PartialEq, Eq)]
struct Effects {
    may: BTreeSet<PortabilityEffect>,
    must: BTreeSet<PortabilityEffect>,
    unknown: bool,
}

fn reachable(procedure: &Procedure) -> Vec<usize> {
    if procedure.entry >= procedure.blocks.len() {
        return Vec::new();
    }
    if procedure.dynamic_control {
        return (0..procedure.blocks.len()).collect();
    }
    let mut seen = vec![false; procedure.blocks.len()];
    let mut pending = vec![procedure.entry];
    while let Some(block) = pending.pop() {
        if !seen[block] {
            seen[block] = true;
            pending.extend(&procedure.blocks[block].successors);
        }
    }
    seen.into_iter()
        .enumerate()
        .filter_map(|(index, seen)| seen.then_some(index))
        .collect()
}

fn summarize(procedure: &Procedure, reachable: &[usize], callees: &[Effects]) -> Effects {
    let mut result = Effects::default();
    if reachable.is_empty() {
        result.unknown = true;
        return result;
    }
    let mut local_must = vec![BTreeSet::new(); procedure.blocks.len()];
    for &index in reachable {
        let block = &procedure.blocks[index];
        result.may.extend(block.may.iter().cloned());
        result.unknown |= block.unknown;
        local_must[index].extend(block.must.iter().cloned());
        for (&callee, definite) in block
            .calls
            .iter()
            .map(|call| (call, true))
            .chain(block.may_calls.iter().map(|call| (call, false)))
        {
            let effects = &callees[callee];
            result.may.extend(effects.may.iter().cloned());
            result.unknown |= effects.unknown;
            if definite {
                local_must[index].extend(effects.must.iter().cloned());
            }
        }
    }
    if procedure.dynamic_control {
        return result;
    }

    // Walk backward from normal returns. Each branch contributes the intersection
    // of its successors. Starting at the empty set also handles cycles that can
    // continue indefinitely without reaching a particular operation.
    let mut must = local_must.clone();
    loop {
        let mut changed = false;
        for &index in reachable.iter().rev() {
            let mut next = local_must[index].clone();
            let successors = &procedure.blocks[index].successors;
            if let Some((&first, rest)) = successors.split_first() {
                next.extend(
                    must[first]
                        .iter()
                        .filter(|effect| rest.iter().all(|&target| must[target].contains(*effect)))
                        .cloned(),
                );
            }
            if must[index] != next {
                must[index] = next;
                changed = true;
            }
        }
        if !changed {
            result.must = std::mem::take(&mut must[procedure.entry]);
            return result;
        }
    }
}

pub(super) fn solve(procedures: Vec<Procedure>) -> Vec<ProcedureEffects> {
    let reachable: Vec<_> = procedures.iter().map(reachable).collect();
    let mut caller_lists = vec![Vec::new(); procedures.len()];
    for (caller, procedure) in procedures.iter().enumerate() {
        for &block in &reachable[caller] {
            for &callee in procedure.blocks[block]
                .calls
                .iter()
                .chain(&procedure.blocks[block].may_calls)
            {
                caller_lists[callee].push(caller);
            }
        }
    }
    let mut pending: Vec<_> = (0..procedures.len()).collect();
    let mut queued = vec![true; procedures.len()];
    let mut effects: Vec<_> = procedures.iter().map(|_| Effects::default()).collect();
    while let Some(index) = pending.pop() {
        queued[index] = false;
        let next = summarize(&procedures[index], &reachable[index], &effects);
        if next != effects[index] {
            effects[index] = next;
            for &caller in &caller_lists[index] {
                if !std::mem::replace(&mut queued[caller], true) {
                    pending.push(caller);
                }
            }
        }
    }
    procedures
        .into_iter()
        .zip(effects)
        .map(|(procedure, effects)| ProcedureEffects {
            project: procedure.project,
            module: procedure.module,
            procedure: procedure.name,
            status: if effects
                .must
                .iter()
                .any(|effect| effect.verdict == Support::Reject)
            {
                EntryStatus::CannotRun
            } else if effects.unknown
                || effects
                    .may
                    .iter()
                    .any(|effect| effect.verdict == Support::Reject)
            {
                EntryStatus::MayStop
            } else {
                EntryStatus::Runs
            },
            may: effects.may.into_iter().collect(),
            must: effects.must.into_iter().collect(),
            unknown: effects.unknown,
        })
        .collect()
}
