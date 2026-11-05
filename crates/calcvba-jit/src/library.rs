//! Shared lowering of library arguments. Storage and helper placement are
//! provided by the frame backend or the native-cell backend.
use super::*;

pub(crate) fn lower_arguments<T>(
    builder: &mut FunctionBuilder<'_>,
    args: &[OxArg],
    mut operand: impl FnMut(&mut FunctionBuilder<'_>, &OxOperand) -> Result<T, JitError>,
) -> Result<Vec<T>, JitError> {
    args.iter()
        .map(|argument| match argument {
            OxArg::ByVal(value) => operand(builder, value),
            OxArg::ByRef(place) => operand(builder, &OxOperand::Use(*place)),
            // Library optionals use Empty, matching the existing JIT library ABI.
            OxArg::Omitted => operand(builder, &OxOperand::Const(OxConst::Empty)),
        })
        .collect()
}
