use std::collections::HashMap;

use calcvba_bundle::NativeImplId;
use calcvba_oxir::OxProgram;
use calcvba_oxir::analysis::{EscapeFacts, escape_facts};
use calcvba_oxir::ids::FuncId;
use calcvba_oxir::inst::OxInst;
use calcvba_oxir::program::OxFunc;
use calcvba_oxir::value::{OxCoerceTarget, OxConst, OxOperand, OxPlace};

use super::{builtin_call, effects};

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) enum Value {
    Text(String),
    EnvironmentPath { name: String, suffix: String },
    WorkbookPath(String),
    UserPath(String),
    Application,
    Workbook,
    Procedure(FuncId),
}

impl Value {
    pub fn literal(&self) -> Option<&str> {
        match self {
            Self::Text(text) => Some(text),
            _ => None,
        }
    }

    pub fn path(&self) -> Option<String> {
        match self {
            Self::Text(text) => Some(text.clone()),
            Self::EnvironmentPath { name, suffix } => Some(format!("Env({name}){suffix}")),
            Self::WorkbookPath(suffix) => Some(format!("WorkbookDir{suffix}")),
            Self::UserPath(suffix) => Some(format!("UserChosen{suffix}")),
            _ => None,
        }
    }

    fn concat(self, other: &Self) -> Option<Self> {
        const MAX_PATH_TEXT: usize = 4096;
        let other = other.literal()?;
        let mut result = self;
        let text = match &mut result {
            Self::Text(text)
            | Self::EnvironmentPath { suffix: text, .. }
            | Self::WorkbookPath(text)
            | Self::UserPath(text) => text,
            _ => return None,
        };
        if text.len().saturating_add(other.len()) > MAX_PATH_TEXT {
            return None;
        }
        text.push_str(other);
        Some(result)
    }
}

#[derive(Clone, Default)]
pub(super) struct Values(HashMap<OxPlace, Value>);

impl Values {
    pub fn get(&self, operand: &OxOperand) -> Option<Value> {
        match operand {
            OxOperand::Const(OxConst::Str(text)) => Some(Value::Text(text.clone())),
            OxOperand::Const(OxConst::NullStr) => Some(Value::Text(String::new())),
            OxOperand::Use(place) => self.0.get(place).cloned(),
            _ => None,
        }
    }

    pub fn at(&self, place: OxPlace) -> Option<Value> {
        self.0.get(&place).cloned()
    }

    fn set(&mut self, place: OxPlace, value: Option<Value>) {
        if let Some(value) = value {
            self.0.insert(place, value);
        } else {
            self.0.remove(&place);
        }
    }

    fn retain_agreement(&mut self, other: &Self) -> bool {
        let previous_len = self.0.len();
        self.0
            .retain(|place, value| other.0.get(place) == Some(value));
        self.0.len() != previous_len
    }
}

pub(super) struct Flow<'a> {
    program: &'a OxProgram,
    library: &'a calcvba_bundle::Bundle,
    function: &'a OxFunc,
    escape: EscapeFacts,
}

impl<'a> Flow<'a> {
    pub fn new(
        program: &'a OxProgram,
        library: &'a calcvba_bundle::Bundle,
        function: &'a OxFunc,
    ) -> Self {
        Self {
            program,
            library,
            function,
            escape: escape_facts(function),
        }
    }

    pub fn entries(&self, procedure: &effects::Procedure) -> Vec<Values> {
        if procedure.dynamic_control || procedure.entry >= procedure.blocks.len() {
            return vec![Values::default(); procedure.blocks.len()];
        }
        let mut entries = vec![None; procedure.blocks.len()];
        entries[procedure.entry] = Some(Values::default());
        let mut pending = vec![procedure.entry];
        while let Some(index) = pending.pop() {
            let mut exit = entries[index].clone().unwrap_or_default();
            for inst in &self.function.blocks[index].instrs {
                self.step(&mut exit, inst);
            }
            for &successor in &procedure.blocks[index].successors {
                let changed = if let Some(entry) = &mut entries[successor] {
                    entry.retain_agreement(&exit)
                } else {
                    entries[successor] = Some(exit.clone());
                    true
                };
                if changed {
                    pending.push(successor);
                }
            }
        }
        entries.into_iter().map(Option::unwrap_or_default).collect()
    }

    pub fn step(&self, values: &mut Values, inst: &OxInst) {
        if matches!(
            inst,
            OxInst::CallNative {
                callee: calcvba_oxir::value::OxNativeCallee::Declare { .. },
                ..
            }
        ) {
            values.0.clear();
            return;
        }
        let (dst, value) = match inst {
            OxInst::Assign { dst, value }
            | OxInst::Box {
                dst, src: value, ..
            }
            | OxInst::Unbox {
                dst, src: value, ..
            }
            | OxInst::Coerce {
                dst,
                src: value,
                target: OxCoerceTarget::Str | OxCoerceTarget::ImplicitVariant,
            } => (*dst, values.get(value)),
            OxInst::Concat { dst, lhs, rhs } => (
                *dst,
                values
                    .get(lhs)
                    .and_then(|lhs| lhs.concat(&values.get(rhs)?)),
            ),
            OxInst::LoadProcRef { dst, proc } => (*dst, Some(Value::Procedure(*proc))),
            OxInst::CallNative { dst: Some(dst), .. }
            | OxInst::CallExtern { dst: Some(dst), .. } => {
                let value = builtin_call(self.program, self.library, inst).and_then(|call| {
                    let args = call.arguments(values);
                    let arg = args.first()?.literal()?;
                    match call.id() {
                        NativeImplId::Environ if !arg.is_empty() => Some(Value::EnvironmentPath {
                            name: arg.to_uppercase(),
                            suffix: String::new(),
                        }),
                        NativeImplId::AppObject
                            if [
                                "Excel.Application",
                                "Excel._Application",
                                "Excel.Global",
                                "Excel._Global",
                            ]
                            .iter()
                            .any(|root| arg.eq_ignore_ascii_case(root)) =>
                        {
                            Some(Value::Application)
                        }
                        _ => None,
                    }
                });
                (*dst, value)
            }
            OxInst::ComCallEarly {
                dst: Some(dst),
                recv,
                method,
                ..
            } => {
                let member = self
                    .program
                    .com_method(*method)
                    .map(|method| method.name.as_str());
                (*dst, self.object_result(values, recv, member))
            }
            OxInst::ComCallLate {
                dst: Some(dst),
                recv,
                name,
                default_member: false,
                ..
            } => (*dst, self.object_result(values, recv, Some(name))),
            OxInst::AsNew { place: dst, .. }
            | OxInst::Coerce { dst, .. }
            | OxInst::Truthy { dst, .. }
            | OxInst::VariantChanged { dst, .. }
            | OxInst::Arith { dst, .. }
            | OxInst::Div { dst, .. }
            | OxInst::Pow { dst, .. }
            | OxInst::Neg { dst, .. }
            | OxInst::Compare { dst, .. }
            | OxInst::CompareObjectIs { dst, .. }
            | OxInst::Logical { dst, .. }
            | OxInst::Not { dst, .. }
            | OxInst::TypeOfIs { dst, .. }
            | OxInst::CallProc { dst: Some(dst), .. }
            | OxInst::CallProcRef { dst: Some(dst), .. }
            | OxInst::CallByName { dst: Some(dst), .. }
            | OxInst::ComCallLate { dst: Some(dst), .. }
            | OxInst::NewObject { dst, .. }
            | OxInst::NewExtern { dst, .. }
            | OxInst::Predeclared { dst, .. }
            | OxInst::PredeclaredExtern { dst, .. }
            | OxInst::NewRecord { dst, .. }
            | OxInst::FieldGet { dst, .. }
            | OxInst::RecordGet { dst, .. }
            | OxInst::RecordArrayGet { dst, .. }
            | OxInst::ArrayLiteral { dst, .. }
            | OxInst::ArrayAppend { dst, .. }
            | OxInst::ArrayRedim { dst, .. }
            | OxInst::ArrayGet { dst, .. }
            | OxInst::FieldArrayGet { dst, .. }
            | OxInst::Bound { dst, .. }
            | OxInst::AddrOf { dst, .. }
            | OxInst::MakeRef { dst, .. }
            | OxInst::WithEventsGet { dst, .. }
            | OxInst::WithEventsSet { dst, .. }
            | OxInst::WithEventsFirstOwner { dst, .. }
            | OxInst::WithEventsNextOwner { dst, .. }
            | OxInst::WithEventsClearOwner { dst, .. }
            | OxInst::ErrFieldGet { dst, .. }
            | OxInst::ErlGet { dst, .. } => (*dst, None),
            OxInst::ForEachNext {
                iter,
                item,
                has_value,
            } => {
                for dst in [iter, item, has_value] {
                    values.set(*dst, None);
                }
                return;
            }
            OxInst::StmtBoundary {
                clear_temps_from, ..
            } => {
                values.0.retain(
                    |place, _| !matches!(place, OxPlace::Temp(id) if id.0 >= *clear_temps_from),
                );
                return;
            }
            _ => return,
        };
        let track = match dst {
            OxPlace::Local(id) => {
                id.0 >= self.function.param_count && self.escape.locals.get(id.0) == Some(&false)
            }
            OxPlace::Temp(id) => self.escape.temps.get(id.0) == Some(&false),
            OxPlace::Global(_) => false,
        };
        if track {
            values.set(dst, value);
        }
    }

    fn object_result(
        &self,
        values: &Values,
        recv: &OxOperand,
        name: Option<&str>,
    ) -> Option<Value> {
        let name = name?;
        match values.get(recv)? {
            Value::Application if name.eq_ignore_ascii_case("ThisWorkbook") => {
                Some(Value::Workbook)
            }
            Value::Application
                if name.eq_ignore_ascii_case("GetOpenFilename")
                    || name.eq_ignore_ascii_case("GetSaveAsFilename") =>
            {
                Some(Value::UserPath(String::new()))
            }
            Value::Workbook if name.eq_ignore_ascii_case("Path") => {
                Some(Value::WorkbookPath(String::new()))
            }
            _ => None,
        }
    }
}
