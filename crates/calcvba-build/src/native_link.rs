//! Resolve a source-project closure to one native address space. Calls and
//! constructors become ordinary typed instructions before Cranelift sees them.
use calcvba_bundle::ExportTarget;
use calcvba_oxir::{inst::OxAsNew, *};

#[derive(Clone, Copy, Default)]
struct Base {
    funcs: usize,
    globals: usize,
    classes: usize,
    records: usize,
    interfaces: usize,
    imports: usize,
    declares: u32,
}

struct Link<'a> {
    programs: &'a [OxProgram],
    bases: Vec<Base>,
}
impl Link<'_> {
    fn target(
        &self,
        unit: usize,
        import: ImportId,
    ) -> Result<Option<(usize, ExportTarget)>, String> {
        let imported = self.programs[unit]
            .imports
            .get(import.0)
            .ok_or("invalid native import")?;
        if imported.unit.eq_ignore_ascii_case("VBA") {
            return Ok(None);
        }
        let target_unit = self
            .programs
            .iter()
            .position(|p| p.unit_name.eq_ignore_ascii_case(&imported.unit))
            .ok_or_else(|| format!("unlinked native source project {}", imported.unit))?;
        let target = self.programs[target_unit]
            .exports
            .iter()
            .find(|e| e.token.matches(&imported.token))
            .ok_or_else(|| {
                format!(
                    "unresolved native source export {}::{:?}",
                    imported.unit, imported.token
                )
            })?;
        Ok(Some((target_unit, target.target)))
    }
    fn class(&self, unit: usize, import: ImportId) -> Result<Option<ClassId>, String> {
        self.target(unit, import)?
            .map(|(u, target)| match target {
                ExportTarget::Class(c) => Ok(ClassId(self.bases[u].classes + c)),
                _ => Err("native constructor import names a procedure".into()),
            })
            .transpose()
    }
    fn source_class(&self, unit: usize, name: &str) -> Option<(usize, usize)> {
        if let Some((project, class)) = name.rsplit_once('.') {
            let u = self
                .programs
                .iter()
                .position(|p| p.unit_name.eq_ignore_ascii_case(project))?;
            self.programs[u]
                .classes
                .iter()
                .position(|c| c.name.eq_ignore_ascii_case(class))
                .map(|c| (u, c))
        } else {
            self.programs[unit]
                .classes
                .iter()
                .position(|c| c.name.eq_ignore_ascii_case(name))
                .map(|c| (unit, c))
        }
    }
    fn name(&self, unit: usize, name: &mut String) {
        if let Some((u, c)) = self.source_class(unit, name) {
            *name = format!(
                "{}.{}",
                self.programs[u].unit_name, self.programs[u].classes[c].name
            );
        }
    }
    fn ty(&self, unit: usize, ty: &mut OxTy) {
        let base = self.bases[unit];
        match ty {
            OxTy::Object(ObjClass::Class(c)) => c.0 += base.classes,
            OxTy::Object(ObjClass::Iface(i)) => {
                if let Some((u, c)) =
                    self.source_class(unit, self.programs[unit].com_interfaces[i.0].name())
                {
                    *ty = OxTy::Object(ObjClass::Class(ClassId(self.bases[u].classes + c)));
                } else {
                    i.0 += base.interfaces;
                }
            }
            OxTy::Object(ObjClass::ComIface(i)) => i.0 += base.interfaces,
            OxTy::Record(r) => r.0 += base.records,
            OxTy::Array(element, _) => self.ty(unit, element),
            _ => {}
        }
    }
    fn binding(&self, unit: usize, binding: &mut OxAsNew) -> Result<(), String> {
        match binding {
            OxAsNew::ProjectClass { class } => class.0 += self.bases[unit].classes,
            OxAsNew::ExternClass { import } => {
                if let Some(class) = self.class(unit, *import)? {
                    *binding = OxAsNew::ProjectClass { class };
                } else {
                    import.0 += self.bases[unit].imports;
                }
            }
            OxAsNew::ComClass { .. } => {}
        }
        Ok(())
    }
    fn inst(&self, unit: usize, inst: &mut OxInst) -> Result<(), String> {
        let base = self.bases[unit];
        let allocating = matches!(inst, OxInst::NewExtern { .. });
        // Resolve imports before relocating the remaining table indices.
        match inst {
            OxInst::CallExtern { dst, import, args } => {
                if let Some((u, target)) = self.target(unit, *import)? {
                    let ExportTarget::Proc(p) = target else {
                        return Err("native procedure import names a class".into());
                    };
                    *inst = OxInst::CallProc {
                        dst: *dst,
                        proc: FuncId(self.bases[u].funcs + p),
                        args: std::mem::take(args),
                    };
                    relocate_places(inst, base.globals);
                    return Ok(());
                }
            }
            OxInst::NewExtern { dst, import } | OxInst::PredeclaredExtern { dst, import } => {
                if let Some(class) = self.class(unit, *import)? {
                    *inst = if allocating {
                        OxInst::NewObject { dst: *dst, class }
                    } else {
                        OxInst::Predeclared { dst: *dst, class }
                    };
                    relocate_places(inst, base.globals);
                    return Ok(());
                }
            }
            OxInst::PredeclaredExternSet { import, value } => {
                if let Some(class) = self.class(unit, *import)? {
                    *inst = OxInst::PredeclaredSet {
                        class,
                        value: value.clone(),
                    };
                    relocate_places(inst, base.globals);
                    return Ok(());
                }
            }
            _ => {}
        }
        relocate_places(inst, base.globals);
        match inst {
            OxInst::AsNew { binding, .. } => self.binding(unit, binding)?,
            OxInst::CallProc { proc, .. } | OxInst::LoadProcRef { proc, .. } => {
                proc.0 += base.funcs
            }
            OxInst::CallExtern { import, .. }
            | OxInst::NewExtern { import, .. }
            | OxInst::PredeclaredExtern { import, .. }
            | OxInst::PredeclaredExternSet { import, .. } => import.0 += base.imports,
            OxInst::NewObject { class, .. }
            | OxInst::Predeclared { class, .. }
            | OxInst::PredeclaredSet { class, .. } => class.0 += base.classes,
            OxInst::Box { from, .. } => self.ty(unit, from),
            OxInst::Unbox { to, .. } => self.ty(unit, to),
            OxInst::ComCallEarly { method, .. } => method.iface.0 += base.interfaces,
            OxInst::CallNative {
                callee: OxNativeCallee::Declare { descriptor_id, .. },
                ..
            } => *descriptor_id += base.declares,
            OxInst::ValidateAssignment {
                target_type_name, ..
            }
            | OxInst::TypeOfIs {
                type_name: target_type_name,
                ..
            } => self.name(unit, target_type_name),
            _ => {}
        }
        Ok(())
    }
}

fn place(place: &mut OxPlace, base: usize) {
    if let OxPlace::Global(id) = place {
        id.0 += base;
    }
}
fn operand(op: &mut OxOperand, base: usize) {
    if let OxOperand::Use(p) = op {
        place(p, base);
    }
}
fn args(args: &mut [OxArg], base: usize) {
    for arg in args {
        match arg {
            OxArg::ByVal(op) => operand(op, base),
            OxArg::ByRef(p) => place(p, base),
            OxArg::Omitted => {}
        }
    }
}
fn call_args(args: &mut [OxCallArg], base: usize) {
    for arg in args {
        match arg {
            OxCallArg::Operand(op) | OxCallArg::Named { value: op, .. } => operand(op, base),
            OxCallArg::ByRef(p) => place(p, base),
            OxCallArg::Omitted | OxCallArg::Const(_) => {}
        }
    }
}

// Exhaustive on purpose: adding an instruction must account for every global
// operand, including aliases, writebacks and operands of error terminators.
fn relocate_places(inst: &mut OxInst, base: usize) {
    use OxInst::*;
    match inst {
        AsNew { place: p, .. } => place(p, base),
        Assign { dst, value: src }
        | Box { dst, src, .. }
        | Unbox { dst, src, .. }
        | Coerce { dst, src, .. }
        | Truthy { dst, src }
        | Neg { dst, src, .. }
        | Not { dst, src }
        | Ptr { dst, src, .. } => {
            place(dst, base);
            operand(src, base);
        }
        Arith { dst, lhs, rhs, .. }
        | Div { dst, lhs, rhs }
        | Pow { dst, lhs, rhs }
        | Concat { dst, lhs, rhs }
        | Compare { dst, lhs, rhs, .. }
        | CompareObjectIs { dst, lhs, rhs }
        | Logical { dst, lhs, rhs, .. }
        | VariantChanged {
            dst,
            current: lhs,
            original: rhs,
        } => {
            place(dst, base);
            operand(lhs, base);
            operand(rhs, base);
        }
        ValidateAssignment { src, .. } | ErrFieldSet { src, .. } => operand(src, base),
        LoadProcRef { dst, .. }
        | NewObject { dst, .. }
        | NewExtern { dst, .. }
        | Predeclared { dst, .. }
        | PredeclaredExtern { dst, .. }
        | NewRecord { dst, .. }
        | ErrFieldGet { dst, .. }
        | ErlGet { dst }
        | WithEventsNextOwner { dst } => place(dst, base),
        PredeclaredSet { value, .. } | PredeclaredExternSet { value, .. } => operand(value, base),
        TypeOfIs { dst, object, .. }
        | FieldGet { dst, object, .. }
        | WithEventsGet {
            dst, owner: object, ..
        }
        | WithEventsClearOwner { dst, owner: object }
        | WithEventsFirstOwner {
            dst,
            source: object,
            ..
        } => {
            place(dst, base);
            operand(object, base);
        }
        CallProc { dst, args: a, .. } | CallExtern { dst, args: a, .. } => {
            if let Some(dst) = dst {
                place(dst, base);
            }
            args(a, base);
        }
        CallProcRef {
            dst,
            target,
            args: a,
        } => {
            if let Some(dst) = dst {
                place(dst, base);
            }
            operand(target, base);
            args(a, base);
        }
        CallNative { dst, callee, args } => {
            if let Some(dst) = dst {
                place(dst, base);
            }
            call_args(args, base);
            if let OxNativeCallee::Declare { ptr_writebacks, .. } = callee {
                for writeback in ptr_writebacks {
                    place(&mut writeback.target, base);
                }
            }
        }
        CallByName {
            dst,
            object,
            name,
            calltype,
            args,
        } => {
            if let Some(dst) = dst {
                place(dst, base);
            }
            operand(object, base);
            operand(name, base);
            operand(calltype, base);
            call_args(args, base);
        }
        ComCallEarly {
            dst, recv, args, ..
        }
        | ComCallLate {
            dst, recv, args, ..
        } => {
            if let Some(dst) = dst {
                place(dst, base);
            }
            operand(recv, base);
            call_args(args, base);
        }
        FieldSet { object, value, .. } => {
            operand(object, base);
            operand(value, base);
        }
        RecordGet { dst, record, .. } => {
            place(dst, base);
            operand(record, base);
        }
        RecordSet { record, value, .. } | RecordLSet { record, value } => {
            place(record, base);
            operand(value, base);
        }
        RecordArrayGet {
            dst,
            record: array,
            indices,
            ..
        }
        | ArrayGet {
            dst,
            array,
            indices,
        }
        | FieldArrayGet {
            dst,
            object: array,
            indices,
            ..
        } => {
            place(dst, base);
            operand(array, base);
            for i in indices {
                operand(i, base);
            }
        }
        RecordArraySet {
            record: array,
            indices,
            value,
            ..
        }
        | ArraySet {
            array,
            indices,
            value,
        } => {
            place(array, base);
            for i in indices {
                operand(i, base);
            }
            operand(value, base);
        }
        FieldArraySet {
            object,
            indices,
            value,
            ..
        } => {
            operand(object, base);
            for i in indices {
                operand(i, base);
            }
            operand(value, base);
        }
        ArrayLiteral {
            dst,
            values,
            aliases,
            ..
        } => {
            place(dst, base);
            for v in values {
                operand(v, base);
            }
            for a in aliases.iter_mut().flatten() {
                place(a, base);
            }
        }
        ArrayAppend { dst, array, item } => {
            place(dst, base);
            operand(array, base);
            operand(item, base);
        }
        ArrayRedim {
            dst,
            upper_bounds,
            lower_bounds,
            ..
        } => {
            place(dst, base);
            for v in upper_bounds.iter_mut().chain(lower_bounds) {
                operand(v, base);
            }
        }
        ArrayErase { array, .. } => place(array, base),
        Bound {
            dst,
            array,
            dimension,
            ..
        } => {
            place(dst, base);
            operand(array, base);
            if let Some(d) = dimension {
                operand(d, base);
            }
        }
        ForEachInit { iter, source } => {
            place(iter, base);
            operand(source, base);
        }
        ForEachNext {
            iter,
            item,
            has_value,
        } => {
            place(iter, base);
            place(item, base);
            place(has_value, base);
        }
        WithEventsSet {
            dst, owner, value, ..
        } => {
            place(dst, base);
            operand(owner, base);
            operand(value, base);
        }
        RaiseEvent {
            source, args: a, ..
        } => {
            operand(source, base);
            args(a, base);
        }
        AddRef { object } | Release { object, .. } => operand(object, base),
        SetErrorHandler(_)
        | ClearErr
        | StmtBoundary { .. }
        | SetLineNumber { .. }
        | DrainTerminations => {}
    }
}

pub(super) fn link(programs: &[OxProgram]) -> Result<OxProgram, String> {
    if programs.len() == 1 {
        return Ok(programs[0].clone());
    }
    let root = programs.last().ok_or("empty native project closure")?;
    let mut bases = Vec::new();
    let mut at = Base::default();
    for p in programs {
        if !p.event_routes.is_empty() {
            return Err("native source-project events are not implemented".into());
        }
        bases.push(at);
        at.funcs += p.funcs.len();
        at.globals += p.globals.len();
        at.classes += p.classes.len();
        at.records += p.record_layouts.len();
        at.interfaces += p.com_interfaces.len();
        at.imports += p.imports.len();
        at.declares = at
            .declares
            .checked_add(
                p.external_calls
                    .iter()
                    .map(|d| d.descriptor_id)
                    .max()
                    .map_or(0, |id| id + 1),
            )
            .ok_or("native Declare table overflow")?;
    }
    let link = Link { programs, bases };
    let mut out = OxProgram {
        unit_name: root.unit_name.clone(),
        ..OxProgram::empty()
    };
    let mut initializers = Vec::new();
    for (unit, original) in programs.iter().enumerate() {
        let base = link.bases[unit];
        let mut p = original.clone();
        for f in &mut p.funcs {
            f.name = format!("{}::{}", original.unit_name, f.name);
            for local in &mut f.locals {
                link.ty(unit, &mut local.ty);
            }
            for ty in &mut f.temps {
                link.ty(unit, ty);
            }
            for block in &mut f.blocks {
                for inst in &mut block.instrs {
                    link.inst(unit, inst)?;
                }
                match &mut block.terminator {
                    OxTerminator::Branch { cond, .. } => operand(cond, base.globals),
                    OxTerminator::Raise {
                        number,
                        source,
                        description,
                        help_file,
                        help_context,
                        ..
                    } => {
                        operand(number, base.globals);
                        for op in [source, description, help_file, help_context]
                            .into_iter()
                            .flatten()
                        {
                            operand(op, base.globals);
                        }
                    }
                    _ => {}
                }
            }
        }
        for g in &mut p.globals {
            link.ty(unit, &mut g.ty);
        }
        for class in &mut p.classes {
            let qualified = format!("{}.{}", original.unit_name, class.name);
            for name in &mut class.implements {
                link.name(unit, name);
            }
            class.implements.push(qualified.clone());
            class.name = qualified;
            for f in class.initialize.iter_mut().chain(&mut class.terminate) {
                f.0 += base.funcs;
            }
            for m in &mut class.methods {
                m.proc.0 += base.funcs;
            }
            for f in &mut class.fields {
                link.ty(unit, &mut f.ty);
            }
            for f in &mut class.as_new_fields {
                link.binding(unit, &mut f.binding)?;
            }
        }
        for interface in &mut p.com_interfaces {
            if let ComInterface::Project(p) = interface {
                link.name(unit, &mut p.name);
            }
        }
        for d in &mut p.external_calls {
            d.descriptor_id += base.declares;
        }
        if let Some(id) = p.global_initializer {
            initializers.push(FuncId(id.0 + base.funcs));
        }
        if unit + 1 == programs.len() {
            out.entry = p.entry.map(|id| FuncId(id.0 + base.funcs));
            out.exports = p
                .exports
                .into_iter()
                .map(|mut e| {
                    match &mut e.target {
                        ExportTarget::Proc(p) => *p += base.funcs,
                        ExportTarget::Class(c) => *c += base.classes,
                    };
                    e
                })
                .collect();
        }
        out.funcs.extend(p.funcs);
        out.globals.extend(p.globals);
        out.classes.extend(p.classes);
        out.record_layouts.extend(p.record_layouts);
        out.com_interfaces.extend(p.com_interfaces);
        out.external_calls.extend(p.external_calls);
        out.imports.extend(p.imports);
    }
    if !initializers.is_empty() {
        let id = FuncId(out.funcs.len());
        let mut blocks = Vec::new();
        for (index, proc) in initializers.into_iter().enumerate() {
            blocks.push(OxBlock {
                id: BlockId(index),
                instrs: vec![OxInst::CallProc {
                    dst: None,
                    proc,
                    args: vec![],
                }],
                fault_target: Some(BlockId(usize::MAX)),
                terminator: OxTerminator::Jump(BlockId(index + 1)),
            });
        }
        let end = blocks.len();
        blocks.push(OxBlock::new(BlockId(end), OxTerminator::Return));
        let fail = blocks.len();
        blocks.push(OxBlock::new(
            BlockId(fail),
            OxTerminator::FaultDispatch {
                resume: BlockId(0),
                resume_next: BlockId(end),
            },
        ));
        for block in &mut blocks[..end] {
            block.fault_target = Some(BlockId(fail));
        }
        out.funcs.push(OxFunc {
            name: "__NativeProjectInitializers".into(),
            kind: calcvba_bundle::ProcedureKind::Sub,
            locals: vec![],
            temps: vec![],
            param_count: 0,
            return_local: None,
            blocks,
            entry: BlockId(0),
        });
        out.global_initializer = Some(id);
    }
    calcvba_oxir::verify_program(&out).map_err(|e| format!("invalid linked native OxIR: {e:?}"))?;
    Ok(out)
}
