//! Portable in-memory Excel host stub for differential VM3/JIT parity runs.
//!
//! Provides Excel typelib metadata (`ExcelStubTypeLibs`) and a portable
//! `PortableDispatch` object model (`Excel.Global`, `Excel.Application`,
//! `Excel.Workbook`, `Excel.Worksheets`, `Excel.Worksheet`, `Excel.Range`,
//! `Excel.Areas`, `Excel.WorksheetFunction`) backed by an in-memory blank
//! worksheet so corpus and host-object examples execute on both engines without
//! Office or Windows COM registrations.

use std::collections::{BTreeMap, HashMap};
use std::sync::{Arc, Mutex};

use calcvba_com::{
    ComDefaultValue, DynamicCallArg, DynamicCallKind, PortableComProjection, TypeLibAppObject,
    TypeLibConstantMetadata, TypeLibMemberInvokeKind, TypeLibMemberMetadata, TypeLibMetadataBlob,
    TypeLibParamType, TypeLibResolveRequest, TypeLibResolvedIdentity, TypeLibWireType,
    platform::portable::{
        PortableDispatch, PortableObjectFactory, PortableObjectResult, PortableTypedCallResult,
    },
};
use calcvba_runtime::{
    ObjectRefExt, RuntimeCallResult, VarType, Variant,
    safe_array::{SafeArray, SafeArrayBound},
    variant_to_vba_string,
};
use calcvba_symbol::{CatalogTypeLibResolver, TypeLibResolver, manifest::ProjectReference};

const EXCEL_LIBID: &str = "{00020813-0000-0000-C000-000000000046}";
const EXCEL_GLOBAL_CLASS: &str = "Excel.Global";

const MAX_ROWS: i32 = 1_048_576;
const MAX_COLS: i32 = 16_384;

const XL_A1: i32 = 1;
const XL_R1C1: i32 = -4150;
const XL_DOWN: i32 = -4121;
const XL_TO_LEFT: i32 = -4159;
const XL_TO_RIGHT: i32 = -4161;
const XL_UP: i32 = -4162;
const XL_BY_ROWS: i32 = 1;
const XL_BY_COLUMNS: i32 = 2;
const XL_NEXT: i32 = 1;
const XL_PREVIOUS: i32 = 2;
const XL_WHOLE: i32 = 1;
const XL_PART: i32 = 2;
const XL_FORMULAS: i32 = -4123;
const XL_VALUES: i32 = -4163;
const XL_CELL_TYPE_CONSTANTS: i32 = 2;
const XL_CELL_TYPE_FORMULAS: i32 = -4123;
const XL_CELL_TYPE_BLANKS: i32 = 4;
const XL_NUMBERS: i32 = 1;
const XL_TEXT_VALUES: i32 = 2;
const XL_LOGICAL: i32 = 4;
const XL_ERRORS: i32 = 16;

/// Project reference that brings the stubbed Excel object library into scope.
pub fn excel_project_reference() -> ProjectReference {
    ProjectReference::TypeLibrary {
        name: "Excel".into(),
        guid: Some(EXCEL_LIBID.into()),
        version_major: Some(1),
        version_minor: Some(9),
        lcid: Some(0),
        import_lib: None,
    }
}

/// TypeLib resolver that serves the portable Excel host surface and delegates
/// all other references to [`CatalogTypeLibResolver`].
pub struct ExcelStubTypeLibs;

impl TypeLibResolver for ExcelStubTypeLibs {
    fn resolve(&self, request: &TypeLibResolveRequest) -> Option<TypeLibMetadataBlob> {
        let is_excel = request.reference_name.eq_ignore_ascii_case("Excel")
            || request
                .libid_hint
                .as_deref()
                .is_some_and(|g| g.eq_ignore_ascii_case(EXCEL_LIBID));
        if !is_excel {
            return CatalogTypeLibResolver.resolve(request);
        }
        Some(excel_blob(request.requested_coclass.as_deref()))
    }
}

/// Registers the portable Excel host application object on `projection`.
pub fn install_excel_stub(projection: &PortableComProjection) {
    projection.register_object(EXCEL_GLOBAL_CLASS, Arc::new(ExcelHostFactory));
}

fn identity(coclass: Option<&str>) -> TypeLibResolvedIdentity {
    TypeLibResolvedIdentity {
        reference_name: "Excel".into(),
        requested_coclass: coclass.map(str::to_owned),
        importlib: "EXCEL.EXE".into(),
        libid: Some(EXCEL_LIBID.into()),
        major_version: 1,
        minor_version: 9,
        lcid: Some(0),
        cache_key: format!("excel-stub:{}", coclass.unwrap_or("library")),
    }
}

fn member(
    name: &str,
    token: i32,
    invoke_kind: TypeLibMemberInvokeKind,
    params: &[(&str, TypeLibParamType, bool)],
    ret_iface: Option<&str>,
    is_default: bool,
) -> TypeLibMemberMetadata {
    let (return_type, return_wire_type) = match ret_iface {
        Some(iface) => (
            Some(TypeLibParamType::Object),
            Some(TypeLibWireType::InterfacePointer {
                name: iface.to_owned(),
            }),
        ),
        None => (Some(TypeLibParamType::Variant), None),
    };
    TypeLibMemberMetadata {
        name: name.into(),
        token,
        vtable_slot: None,
        requires_argument: params.iter().any(|(_, _, opt)| !opt),
        invoke_kind,
        parameter_names: params.iter().map(|(n, _, _)| (*n).to_owned()).collect(),
        parameter_optional: params.iter().map(|(_, _, opt)| *opt).collect(),
        parameter_optional_defaults: Vec::new(),
        is_default_member: is_default,
        parameter_types: params.iter().map(|(_, t, _)| *t).collect(),
        parameter_wire_types: Vec::new(),
        parameter_iids: Vec::new(),
        return_type,
        return_wire_type,
        callconv_is_stdcall: false,
        is_dual: false,
        interface_iid: None,
        source_typekind: None,
        vtable_slot_bound: None,
    }
}

fn excel_constants() -> Vec<TypeLibConstantMetadata> {
    let entries: &[(&str, &str, i64)] = &[
        ("XlReferenceStyle", "xlA1", i64::from(XL_A1)),
        ("XlReferenceStyle", "xlR1C1", i64::from(XL_R1C1)),
        ("XlDirection", "xlDown", i64::from(XL_DOWN)),
        ("XlDirection", "xlToLeft", i64::from(XL_TO_LEFT)),
        ("XlDirection", "xlToRight", i64::from(XL_TO_RIGHT)),
        ("XlDirection", "xlUp", i64::from(XL_UP)),
        ("XlSearchOrder", "xlByRows", i64::from(XL_BY_ROWS)),
        ("XlSearchOrder", "xlByColumns", i64::from(XL_BY_COLUMNS)),
        ("XlSearchDirection", "xlNext", i64::from(XL_NEXT)),
        ("XlSearchDirection", "xlPrevious", i64::from(XL_PREVIOUS)),
        ("XlLookAt", "xlWhole", i64::from(XL_WHOLE)),
        ("XlLookAt", "xlPart", i64::from(XL_PART)),
        ("XlFindLookIn", "xlFormulas", i64::from(XL_FORMULAS)),
        ("XlFindLookIn", "xlValues", i64::from(XL_VALUES)),
        (
            "XlCellType",
            "xlCellTypeConstants",
            i64::from(XL_CELL_TYPE_CONSTANTS),
        ),
        (
            "XlCellType",
            "xlCellTypeFormulas",
            i64::from(XL_CELL_TYPE_FORMULAS),
        ),
        (
            "XlCellType",
            "xlCellTypeBlanks",
            i64::from(XL_CELL_TYPE_BLANKS),
        ),
        ("XlSpecialCellsValue", "xlNumbers", i64::from(XL_NUMBERS)),
        (
            "XlSpecialCellsValue",
            "xlTextValues",
            i64::from(XL_TEXT_VALUES),
        ),
        ("XlSpecialCellsValue", "xlLogical", i64::from(XL_LOGICAL)),
        ("XlSpecialCellsValue", "xlErrors", i64::from(XL_ERRORS)),
    ];
    entries
        .iter()
        .map(|(owner, name, value)| TypeLibConstantMetadata {
            owner: (*owner).to_owned(),
            name: (*name).to_owned(),
            value: ComDefaultValue::Int(*value),
        })
        .collect()
}

fn union_intersect_params() -> Vec<(&'static str, TypeLibParamType, bool)> {
    let mut params = vec![
        ("Arg1", TypeLibParamType::Object, false),
        ("Arg2", TypeLibParamType::Object, false),
    ];
    const NAMES: &[&str] = &[
        "Arg3", "Arg4", "Arg5", "Arg6", "Arg7", "Arg8", "Arg9", "Arg10", "Arg11", "Arg12", "Arg13",
        "Arg14", "Arg15", "Arg16", "Arg17", "Arg18", "Arg19", "Arg20",
    ];
    for name in NAMES {
        params.push((name, TypeLibParamType::Variant, true));
    }
    params
}

fn global_and_app_members() -> Vec<TypeLibMemberMetadata> {
    use TypeLibMemberInvokeKind::{Method, PropertyGet};
    use TypeLibParamType::Variant as V;
    let ui_params = union_intersect_params();
    vec![
        member(
            "Application",
            148,
            PropertyGet,
            &[],
            Some("Excel.Application"),
            false,
        ),
        member(
            "ThisWorkbook",
            778,
            PropertyGet,
            &[],
            Some("Excel.Workbook"),
            false,
        ),
        member(
            "ActiveWorkbook",
            308,
            PropertyGet,
            &[],
            Some("Excel.Workbook"),
            false,
        ),
        member(
            "ActiveSheet",
            307,
            PropertyGet,
            &[],
            Some("Excel.Worksheet"),
            false,
        ),
        member(
            "Worksheets",
            494,
            PropertyGet,
            &[],
            Some("Excel.Worksheets"),
            false,
        ),
        member(
            "Sheets",
            485,
            PropertyGet,
            &[],
            Some("Excel.Worksheets"),
            false,
        ),
        member(
            "WorksheetFunction",
            1440,
            PropertyGet,
            &[],
            Some("Excel.WorksheetFunction"),
            false,
        ),
        member(
            "Range",
            197,
            PropertyGet,
            &[("Cell1", V, false), ("Cell2", V, true)],
            Some("Excel.Range"),
            false,
        ),
        member("Cells", 238, PropertyGet, &[], Some("Excel.Range"), false),
        member("Union", 779, Method, &ui_params, Some("Excel.Range"), false),
        member(
            "Intersect",
            766,
            Method,
            &ui_params,
            Some("Excel.Range"),
            false,
        ),
    ]
}

fn workbook_members() -> Vec<TypeLibMemberMetadata> {
    use TypeLibMemberInvokeKind::PropertyGet;
    vec![
        member(
            "Worksheets",
            494,
            PropertyGet,
            &[],
            Some("Excel.Worksheets"),
            false,
        ),
        member(
            "Sheets",
            485,
            PropertyGet,
            &[],
            Some("Excel.Worksheets"),
            false,
        ),
        member(
            "ActiveSheet",
            307,
            PropertyGet,
            &[],
            Some("Excel.Worksheet"),
            false,
        ),
        member(
            "Application",
            148,
            PropertyGet,
            &[],
            Some("Excel.Application"),
            false,
        ),
    ]
}

fn worksheets_members() -> Vec<TypeLibMemberMetadata> {
    use TypeLibMemberInvokeKind::PropertyGet;
    use TypeLibParamType::Variant as V;
    vec![
        member(
            "Item",
            0,
            PropertyGet,
            &[("Index", V, false)],
            Some("Excel.Worksheet"),
            true,
        ),
        member("Count", 118, PropertyGet, &[], None, false),
    ]
}

fn worksheet_members() -> Vec<TypeLibMemberMetadata> {
    use TypeLibMemberInvokeKind::PropertyGet;
    use TypeLibParamType::Variant as V;
    vec![
        member(
            "Range",
            197,
            PropertyGet,
            &[("Cell1", V, false), ("Cell2", V, true)],
            Some("Excel.Range"),
            false,
        ),
        member("Cells", 238, PropertyGet, &[], Some("Excel.Range"), false),
        member("Rows", 258, PropertyGet, &[], Some("Excel.Range"), false),
        member("Columns", 241, PropertyGet, &[], Some("Excel.Range"), false),
        member("Name", 110, PropertyGet, &[], None, false),
        member(
            "Application",
            148,
            PropertyGet,
            &[],
            Some("Excel.Application"),
            false,
        ),
    ]
}

fn areas_members() -> Vec<TypeLibMemberMetadata> {
    use TypeLibMemberInvokeKind::PropertyGet;
    use TypeLibParamType::Variant as V;
    vec![
        member(
            "Item",
            0,
            PropertyGet,
            &[("Index", V, false)],
            Some("Excel.Range"),
            true,
        ),
        member("Count", 118, PropertyGet, &[], None, false),
    ]
}

fn range_members() -> Vec<TypeLibMemberMetadata> {
    use TypeLibMemberInvokeKind::{Method, PropertyGet, PropertyPut};
    use TypeLibParamType::{Long, Variant as V};
    vec![
        member(
            "Item",
            0,
            PropertyGet,
            &[("RowIndex", V, true), ("ColumnIndex", V, true)],
            Some("Excel.Range"),
            true,
        ),
        member(
            "Item",
            0,
            PropertyPut,
            &[
                ("RowIndex", V, true),
                ("ColumnIndex", V, true),
                ("RHS", V, false),
            ],
            None,
            true,
        ),
        member(
            "Range",
            197,
            PropertyGet,
            &[("Cell1", V, false), ("Cell2", V, true)],
            Some("Excel.Range"),
            false,
        ),
        member("Cells", 238, PropertyGet, &[], Some("Excel.Range"), false),
        member("Rows", 258, PropertyGet, &[], Some("Excel.Range"), false),
        member("Columns", 241, PropertyGet, &[], Some("Excel.Range"), false),
        member("Areas", 568, PropertyGet, &[], Some("Excel.Areas"), false),
        member("Count", 118, PropertyGet, &[], None, false),
        member("Row", 257, PropertyGet, &[], None, false),
        member("Column", 240, PropertyGet, &[], None, false),
        member(
            "Offset",
            254,
            PropertyGet,
            &[("RowOffset", V, true), ("ColumnOffset", V, true)],
            Some("Excel.Range"),
            false,
        ),
        member(
            "Resize",
            256,
            PropertyGet,
            &[("RowSize", V, true), ("ColumnSize", V, true)],
            Some("Excel.Range"),
            false,
        ),
        member(
            "Address",
            236,
            PropertyGet,
            &[
                ("RowAbsolute", V, true),
                ("ColumnAbsolute", V, true),
                ("ReferenceStyle", V, true),
                ("External", V, true),
                ("RelativeTo", V, true),
            ],
            None,
            false,
        ),
        member(
            "CurrentRegion",
            243,
            PropertyGet,
            &[],
            Some("Excel.Range"),
            false,
        ),
        member(
            "End",
            500,
            PropertyGet,
            &[("Direction", Long, false)],
            Some("Excel.Range"),
            false,
        ),
        member(
            "Find",
            398,
            Method,
            &[
                ("What", V, false),
                ("After", V, true),
                ("LookIn", V, true),
                ("LookAt", V, true),
                ("SearchOrder", V, true),
                ("SearchDirection", V, true),
                ("MatchCase", V, true),
                ("MatchByte", V, true),
                ("SearchFormat", V, true),
            ],
            Some("Excel.Range"),
            false,
        ),
        member(
            "SpecialCells",
            410,
            Method,
            &[("Type", Long, false), ("Value", V, true)],
            Some("Excel.Range"),
            false,
        ),
        member(
            "Value",
            6,
            PropertyGet,
            &[("RangeValueDataType", V, true)],
            None,
            false,
        ),
        member(
            "Value",
            6,
            PropertyPut,
            &[("RangeValueDataType", V, true), ("RHS", V, false)],
            None,
            false,
        ),
        member("Value2", 1388, PropertyGet, &[], None, false),
        member(
            "Value2",
            1388,
            PropertyPut,
            &[("RHS", V, false)],
            None,
            false,
        ),
        member("Formula", 261, PropertyGet, &[], None, false),
        member(
            "Formula",
            261,
            PropertyPut,
            &[("RHS", V, false)],
            None,
            false,
        ),
        member("NumberFormat", 193, PropertyGet, &[], None, false),
        member(
            "NumberFormat",
            193,
            PropertyPut,
            &[("RHS", V, false)],
            None,
            false,
        ),
    ]
}

fn worksheet_function_members() -> Vec<TypeLibMemberMetadata> {
    use TypeLibMemberInvokeKind::Method;
    use TypeLibParamType::Variant as V;
    vec![member(
        "Round",
        101,
        Method,
        &[("Arg1", V, false), ("Arg2", V, false)],
        None,
        false,
    )]
}

fn excel_blob(coclass: Option<&str>) -> TypeLibMetadataBlob {
    let all_coclasses = vec![
        "Global".into(),
        "Application".into(),
        "Workbook".into(),
        "Worksheets".into(),
        "Worksheet".into(),
        "Range".into(),
        "Areas".into(),
        "WorksheetFunction".into(),
    ];
    let app_objects = vec![TypeLibAppObject {
        coclass: "Global".into(),
        clsid: EXCEL_GLOBAL_CLASS.into(),
        creatable: true,
    }];
    let (members, constants, coclass_names, activation_prog_id) =
        match coclass.map(str::to_ascii_lowercase).as_deref() {
            None => (Vec::new(), excel_constants(), all_coclasses, None),
            Some("global") => (
                global_and_app_members(),
                Vec::new(),
                vec!["Global".into()],
                Some(EXCEL_GLOBAL_CLASS.into()),
            ),
            Some("application") => (
                global_and_app_members(),
                Vec::new(),
                vec!["Application".into()],
                Some("Excel.Application".into()),
            ),
            Some("workbook") => (
                workbook_members(),
                Vec::new(),
                vec!["Workbook".into()],
                Some("Excel.Workbook".into()),
            ),
            Some("worksheets") => (
                worksheets_members(),
                Vec::new(),
                vec!["Worksheets".into()],
                Some("Excel.Worksheets".into()),
            ),
            Some("worksheet") => (
                worksheet_members(),
                Vec::new(),
                vec!["Worksheet".into()],
                Some("Excel.Worksheet".into()),
            ),
            Some("range") => (
                range_members(),
                Vec::new(),
                vec!["Range".into()],
                Some("Excel.Range".into()),
            ),
            Some("areas") => (
                areas_members(),
                Vec::new(),
                vec!["Areas".into()],
                Some("Excel.Areas".into()),
            ),
            Some("worksheetfunction") => (
                worksheet_function_members(),
                Vec::new(),
                vec!["WorksheetFunction".into()],
                Some("Excel.WorksheetFunction".into()),
            ),
            Some(_) => (Vec::new(), Vec::new(), Vec::new(), None),
        };
    let member_name_to_token = members.iter().map(|m| (m.name.clone(), m.token)).collect();
    TypeLibMetadataBlob {
        constants,
        identity: identity(coclass),
        activation_prog_id,
        member_name_to_token,
        members,
        events: Vec::new(),
        coclass_names,
        enum_type_names: Vec::new(),
        coclass_prog_ids: Vec::new(),
        app_objects,
        extensible: coclass.is_some(),
    }
}

// ── In-memory workbook / range model ────────────────────────────────────────

#[derive(Clone, Debug)]
enum CellContent {
    Empty,
    Number(f64),
    Text(String),
    Bool(bool),
    Error(i32),
    Formula(String),
}

#[derive(Clone, Debug)]
struct CellData {
    content: CellContent,
    number_format: String,
}

impl Default for CellData {
    fn default() -> Self {
        Self {
            content: CellContent::Empty,
            number_format: "General".into(),
        }
    }
}

impl CellData {
    fn is_non_empty(&self) -> bool {
        !matches!(self.content, CellContent::Empty)
    }

    fn is_date_format(&self) -> bool {
        let mut in_quotes = false;
        let mut escaped = false;
        for ch in self.number_format.chars() {
            if escaped {
                escaped = false;
                continue;
            }
            if ch == '\\' {
                escaped = true;
                continue;
            }
            if ch == '"' {
                in_quotes = !in_quotes;
                continue;
            }
            if !in_quotes && matches!(ch.to_ascii_lowercase(), 'y' | 'd') {
                return true;
            }
        }
        false
    }

    fn to_value2(&self) -> Variant {
        match &self.content {
            CellContent::Empty => Variant::empty(),
            CellContent::Number(n) => Variant::from_f64(*n),
            CellContent::Text(s) => Variant::from_string(s.clone()),
            CellContent::Bool(b) => Variant::from_bool(*b),
            CellContent::Error(code) => Variant::from_error_code(*code),
            CellContent::Formula(f) => Variant::from_string(f.clone()),
        }
    }

    fn to_value(&self) -> Variant {
        match &self.content {
            CellContent::Number(n) if self.is_date_format() => Variant::from_date_f64(*n),
            _ => self.to_value2(),
        }
    }

    fn to_formula_string(&self) -> String {
        match &self.content {
            CellContent::Empty => String::new(),
            CellContent::Number(n) => {
                if n.fract() == 0.0 && n.abs() < 1e15 {
                    format!("{}", *n as i64)
                } else {
                    format!("{n}")
                }
            }
            CellContent::Text(s) => s.clone(),
            CellContent::Bool(b) => if *b { "TRUE" } else { "FALSE" }.into(),
            CellContent::Error(code) => format!("#ERR{code}"),
            CellContent::Formula(f) => f.clone(),
        }
    }

    fn display_text(&self, look_in: i32) -> String {
        if look_in != XL_FORMULAS && matches!(self.content, CellContent::Formula(_)) {
            return String::new();
        }
        self.to_formula_string()
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct Rect {
    r1: i32,
    c1: i32,
    r2: i32,
    c2: i32,
}

impl Rect {
    fn single(r: i32, c: i32) -> Self {
        Self {
            r1: r,
            c1: c,
            r2: r,
            c2: c,
        }
    }

    fn new(r1: i32, c1: i32, r2: i32, c2: i32) -> Result<Self, String> {
        let (r1, r2) = if r1 <= r2 { (r1, r2) } else { (r2, r1) };
        let (c1, c2) = if c1 <= c2 { (c1, c2) } else { (c2, c1) };
        if r1 < 1 || c1 < 1 || r2 > MAX_ROWS || c2 > MAX_COLS {
            return Err("application-defined or object-defined error (range out of bounds)".into());
        }
        Ok(Self { r1, c1, r2, c2 })
    }

    fn rows(self) -> i32 {
        self.r2 - self.r1 + 1
    }

    fn cols(self) -> i32 {
        self.c2 - self.c1 + 1
    }

    fn is_single_cell(self) -> bool {
        self.r1 == self.r2 && self.c1 == self.c2
    }

    fn intersect(self, other: Self) -> Option<Self> {
        let r1 = self.r1.max(other.r1);
        let c1 = self.c1.max(other.c1);
        let r2 = self.r2.min(other.r2);
        let c2 = self.c2.min(other.c2);
        (r1 <= r2 && c1 <= c2).then_some(Self { r1, c1, r2, c2 })
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum RangeMode {
    Cells,
    Rows,
    Columns,
}

#[derive(Clone, Debug)]
struct RangeSpec {
    sheet_index: usize,
    areas: Vec<Rect>,
    mode: RangeMode,
}

impl RangeSpec {
    fn first(&self) -> Rect {
        self.areas[0]
    }
}

#[derive(Default)]
struct SheetState {
    name: String,
    cells: BTreeMap<(i32, i32), CellData>,
}

struct WorkbookState {
    sheets: Vec<SheetState>,
    ranges_by_handle: HashMap<i32, RangeSpec>,
}

impl WorkbookState {
    fn new() -> Self {
        Self {
            sheets: vec![SheetState {
                name: "Sheet1".into(),
                cells: BTreeMap::new(),
            }],
            ranges_by_handle: HashMap::new(),
        }
    }
}

type SharedWorkbook = Arc<Mutex<WorkbookState>>;

struct ExcelHostFactory;

impl PortableObjectFactory for ExcelHostFactory {
    fn create(&self) -> Box<dyn PortableDispatch> {
        Box::new(ExcelAppObject {
            wb: Arc::new(Mutex::new(WorkbookState::new())),
        })
    }
}

fn kind_from_invoke(kind: TypeLibMemberInvokeKind) -> DynamicCallKind {
    match kind {
        TypeLibMemberInvokeKind::PropertyGet => DynamicCallKind::PropertyGet,
        TypeLibMemberInvokeKind::Method => DynamicCallKind::Method,
        TypeLibMemberInvokeKind::PropertyPut => DynamicCallKind::PropertyLet,
        TypeLibMemberInvokeKind::PropertyPutRef => DynamicCallKind::PropertySet,
    }
}

fn arg_variant(args: &[DynamicCallArg], index: usize) -> Option<&Variant> {
    args.get(index)
        .and_then(|arg| arg.value.as_ref())
        .map(|value| value.variant())
        .filter(|v| !matches!(v.vtype(), VarType::Error | VarType::Empty))
}

fn arg_variant_raw(args: &[DynamicCallArg], index: usize) -> Option<&Variant> {
    args.get(index)
        .and_then(|arg| arg.value.as_ref())
        .map(|value| value.variant())
}

fn arg_named_or_pos<'a>(args: &'a [DynamicCallArg], pos: usize, name: &str) -> Option<&'a Variant> {
    for arg in args {
        if arg
            .name
            .as_deref()
            .is_some_and(|n| n.eq_ignore_ascii_case(name))
        {
            return arg
                .value
                .as_ref()
                .map(|v| v.variant())
                .filter(|v| !matches!(v.vtype(), VarType::Error | VarType::Empty));
        }
    }
    if args.get(pos).is_some_and(|a| a.name.is_none()) {
        arg_variant(args, pos)
    } else {
        None
    }
}

fn variant_to_i32(v: &Variant) -> Result<i32, String> {
    if let Some(n) = v.as_i32() {
        return Ok(n);
    }
    if let Some(n) = v.as_i16() {
        return Ok(i32::from(n));
    }
    if let Some(n) = v.as_i64() {
        return i32::try_from(n).map_err(|_| "overflow".into());
    }
    if let Some(n) = v.as_f64() {
        return Ok(n.round() as i32);
    }
    if let Some(n) = v.as_f32() {
        return Ok(n.round() as i32);
    }
    if let Some(b) = v.as_bool() {
        return Ok(if b { -1 } else { 0 });
    }
    let s = variant_to_vba_string(v)?.as_str();
    s.trim()
        .parse::<i32>()
        .map_err(|_| format!("type mismatch: `{s}`"))
}

fn variant_to_f64(v: &Variant) -> Result<f64, String> {
    if let Some(n) = v.as_f64() {
        return Ok(n);
    }
    if let Some(n) = v.as_date_f64() {
        return Ok(n);
    }
    if let Some(n) = v.as_f32() {
        return Ok(f64::from(n));
    }
    if let Some(n) = v.as_i32() {
        return Ok(f64::from(n));
    }
    if let Some(n) = v.as_i16() {
        return Ok(f64::from(n));
    }
    if let Some(n) = v.as_u8() {
        return Ok(f64::from(n));
    }
    if let Some(n) = v.as_i64() {
        return Ok(n as f64);
    }
    if let Some(n) = v.as_currency_scaled_i64() {
        return Ok(n as f64 / 10_000.0);
    }
    if let Some(b) = v.as_bool() {
        return Ok(if b { -1.0 } else { 0.0 });
    }
    let s = variant_to_vba_string(v)?.as_str();
    s.trim()
        .parse::<f64>()
        .map_err(|_| format!("type mismatch: `{s}`"))
}

fn variant_to_bool(v: &Variant) -> Result<bool, String> {
    if let Some(b) = v.as_bool() {
        return Ok(b);
    }
    Ok(variant_to_f64(v)? != 0.0)
}

fn variant_to_string(v: &Variant) -> Result<String, String> {
    Ok(variant_to_vba_string(v)?.as_str())
}

fn variant_to_cell_content(v: &Variant) -> CellContent {
    match v.vtype() {
        VarType::Empty => CellContent::Empty,
        VarType::Boolean => CellContent::Bool(v.as_bool().unwrap_or(false)),
        VarType::String => {
            let s = variant_to_vba_string(v)
                .map(|b| b.as_str())
                .unwrap_or_default();
            CellContent::Text(s)
        }
        VarType::Error => CellContent::Error(v.as_error_code().unwrap_or(0)),
        VarType::Date => CellContent::Number(v.as_date_f64().unwrap_or(0.0)),
        _ => match variant_to_f64(v) {
            Ok(n) => CellContent::Number(n),
            Err(_) => CellContent::Text(
                variant_to_vba_string(v)
                    .map(|b| b.as_str())
                    .unwrap_or_default(),
            ),
        },
    }
}

fn parse_col(letters: &str) -> Result<i32, String> {
    if letters.is_empty() {
        return Err("invalid column".into());
    }
    let mut col: i32 = 0;
    for b in letters.bytes() {
        let up = b.to_ascii_uppercase();
        if !up.is_ascii_uppercase() {
            return Err(format!("invalid column `{letters}`"));
        }
        col = col * 26 + i32::from(up - b'A' + 1);
    }
    if !(1..=MAX_COLS).contains(&col) {
        return Err(format!("column out of range: `{letters}`"));
    }
    Ok(col)
}

fn col_to_letters(mut col: i32) -> String {
    let mut chars = Vec::new();
    while col > 0 {
        let rem = (col - 1) % 26;
        chars.push((b'A' + rem as u8) as char);
        col = (col - 1) / 26;
    }
    chars.reverse();
    chars.into_iter().collect()
}

fn parse_a1_cell(spec: &str) -> Result<(i32, i32), String> {
    let s = spec.trim().replace('$', "");
    let split = s
        .find(|c: char| c.is_ascii_digit())
        .ok_or_else(|| format!("invalid A1 cell `{spec}`"))?;
    let col = parse_col(&s[..split])?;
    let row: i32 = s[split..]
        .parse()
        .map_err(|_| format!("invalid A1 row `{spec}`"))?;
    if !(1..=MAX_ROWS).contains(&row) {
        return Err(format!("row out of range: `{spec}`"));
    }
    Ok((row, col))
}

fn parse_a1_area(spec: &str) -> Result<Rect, String> {
    let s = spec.trim();
    if let Some((a, b)) = s.split_once(':') {
        let (r1, c1) = parse_a1_cell(a)?;
        let (r2, c2) = parse_a1_cell(b)?;
        Rect::new(r1, c1, r2, c2)
    } else {
        let (r, c) = parse_a1_cell(s)?;
        Ok(Rect::single(r, c))
    }
}

fn parse_a1_areas(spec: &str) -> Result<Vec<Rect>, String> {
    let mut areas = Vec::new();
    for part in spec.split(',') {
        areas.push(parse_a1_area(part)?);
    }
    if areas.is_empty() {
        return Err("empty range address".into());
    }
    Ok(areas)
}

fn format_a1_cell(r: i32, c: i32, row_abs: bool, col_abs: bool) -> String {
    format!(
        "{}{}{}{}",
        if col_abs { "$" } else { "" },
        col_to_letters(c),
        if row_abs { "$" } else { "" },
        r
    )
}

fn format_r1c1_cell(
    r: i32,
    c: i32,
    row_abs: bool,
    col_abs: bool,
    ref_r: i32,
    ref_c: i32,
) -> String {
    let row_part = if row_abs {
        format!("R{r}")
    } else if r == ref_r {
        "R".into()
    } else {
        format!("R[{}]", r - ref_r)
    };
    let col_part = if col_abs {
        format!("C{c}")
    } else if c == ref_c {
        "C".into()
    } else {
        format!("C[{}]", c - ref_c)
    };
    format!("{row_part}{col_part}")
}

fn lookup_range(wb: &SharedWorkbook, v: &Variant) -> Result<RangeSpec, String> {
    let obj = v
        .as_object_ref()
        .ok_or_else(|| "expected Range object".to_string())?;
    wb.lock()
        .unwrap()
        .ranges_by_handle
        .get(&obj.raw())
        .cloned()
        .ok_or_else(|| format!("unknown Range handle {}", obj.raw()))
}

fn make_range_result(wb: SharedWorkbook, spec: RangeSpec) -> PortableTypedCallResult {
    PortableTypedCallResult::Object(PortableObjectResult::new(
        "Excel.Range",
        Box::new(ExcelRangeObject { wb, spec }),
    ))
}

fn resolve_range_arg(
    wb: &SharedWorkbook,
    sheet_index: usize,
    origin: (i32, i32),
    cell1: &Variant,
    cell2: Option<&Variant>,
) -> Result<RangeSpec, String> {
    let (dr, dc) = (origin.0 - 1, origin.1 - 1);
    let resolve_one = |v: &Variant| -> Result<Vec<Rect>, String> {
        if v.vtype() == VarType::Object {
            let spec = lookup_range(wb, v)?;
            return Ok(spec.areas);
        }
        let s = variant_to_string(v)?;
        let raw = parse_a1_areas(&s)?;
        raw.into_iter()
            .map(|rect| Rect::new(rect.r1 + dr, rect.c1 + dc, rect.r2 + dr, rect.c2 + dc))
            .collect()
    };
    let first_areas = resolve_one(cell1)?;
    if let Some(c2) = cell2 {
        let second_areas = resolve_one(c2)?;
        let a = first_areas[0];
        let b = second_areas[0];
        let rect = Rect::new(
            a.r1.min(b.r1),
            a.c1.min(b.c1),
            a.r2.max(b.r2),
            a.c2.max(b.c2),
        )?;
        Ok(RangeSpec {
            sheet_index,
            areas: vec![rect],
            mode: RangeMode::Cells,
        })
    } else {
        Ok(RangeSpec {
            sheet_index,
            areas: first_areas,
            mode: RangeMode::Cells,
        })
    }
}

fn excel_round(x: f64, digits: i32) -> f64 {
    if !x.is_finite() {
        return x;
    }
    if digits == 0 {
        let rounded = (x.abs() + 0.5).floor();
        if rounded == 0.0 {
            0.0
        } else {
            rounded.copysign(x)
        }
    } else if digits > 0 {
        let factor = 10_f64.powi(digits);
        let scaled = x * factor;
        let rounded = (scaled.abs() + 0.5).floor().copysign(scaled) / factor;
        if rounded == 0.0 { 0.0 } else { rounded }
    } else {
        let factor = 10_f64.powi(-digits);
        let scaled = x / factor;
        let rounded = (scaled.abs() + 0.5).floor().copysign(scaled) * factor;
        if rounded == 0.0 { 0.0 } else { rounded }
    }
}

// ── PortableDispatch implementations ────────────────────────────────────────

struct ExcelAppObject {
    wb: SharedWorkbook,
}

impl ExcelAppObject {
    fn handle_call(
        &self,
        member: &str,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        match member.to_ascii_lowercase().as_str() {
            "application" => Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                "Excel.Application",
                Box::new(ExcelAppObject {
                    wb: self.wb.clone(),
                }),
            ))),
            "thisworkbook" | "activeworkbook" => {
                Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                    "Excel.Workbook",
                    Box::new(ExcelWorkbookObject {
                        wb: self.wb.clone(),
                    }),
                )))
            }
            "activesheet" => Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                "Excel.Worksheet",
                Box::new(ExcelWorksheetObject {
                    wb: self.wb.clone(),
                    sheet_index: 0,
                }),
            ))),
            "worksheets" | "sheets" => {
                if let Some(idx) = arg_variant(args, 0) {
                    return ExcelWorksheetsObject {
                        wb: self.wb.clone(),
                    }
                    .get_sheet(idx);
                }
                Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                    "Excel.Worksheets",
                    Box::new(ExcelWorksheetsObject {
                        wb: self.wb.clone(),
                    }),
                )))
            }
            "worksheetfunction" => Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                "Excel.WorksheetFunction",
                Box::new(ExcelWorksheetFunctionObject),
            ))),
            "range" => {
                let c1 = arg_variant(args, 0).ok_or("Range requires Cell1")?;
                let c2 = arg_variant(args, 1);
                let spec = resolve_range_arg(&self.wb, 0, (1, 1), c1, c2)?;
                Ok(make_range_result(self.wb.clone(), spec))
            }
            "cells" => ExcelWorksheetObject {
                wb: self.wb.clone(),
                sheet_index: 0,
            }
            .handle_call("cells", args),
            "union" => {
                let mut areas = Vec::new();
                let mut sheet_index = None;
                for arg in args {
                    let Some(v) = arg
                        .value
                        .as_ref()
                        .map(|v| v.variant())
                        .filter(|v| !matches!(v.vtype(), VarType::Error | VarType::Empty))
                    else {
                        continue;
                    };
                    let spec = lookup_range(&self.wb, v)?;
                    sheet_index.get_or_insert(spec.sheet_index);
                    areas.extend(spec.areas);
                }
                if areas.is_empty() {
                    return Err("Union requires at least two Range arguments".into());
                }
                Ok(make_range_result(
                    self.wb.clone(),
                    RangeSpec {
                        sheet_index: sheet_index.unwrap_or(0),
                        areas,
                        mode: RangeMode::Cells,
                    },
                ))
            }
            "intersect" => {
                let mut specs = Vec::new();
                for arg in args {
                    let Some(v) = arg
                        .value
                        .as_ref()
                        .map(|v| v.variant())
                        .filter(|v| !matches!(v.vtype(), VarType::Error | VarType::Empty))
                    else {
                        continue;
                    };
                    specs.push(lookup_range(&self.wb, v)?);
                }
                if specs.len() < 2 {
                    return Err("Intersect requires at least two Range arguments".into());
                }
                let sheet_index = specs[0].sheet_index;
                let mut acc = specs[0].areas.clone();
                for spec in &specs[1..] {
                    if spec.sheet_index != sheet_index {
                        return Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                            Variant::nothing(),
                        )));
                    }
                    let mut next = Vec::new();
                    for a in &acc {
                        for b in &spec.areas {
                            if let Some(inter) = a.intersect(*b) {
                                next.push(inter);
                            }
                        }
                    }
                    acc = next;
                }
                if acc.is_empty() {
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        Variant::nothing(),
                    )))
                } else {
                    Ok(make_range_result(
                        self.wb.clone(),
                        RangeSpec {
                            sheet_index,
                            areas: acc,
                            mode: RangeMode::Cells,
                        },
                    ))
                }
            }
            other => Err(format!("unsupported Excel.Application member `{other}`")),
        }
    }
}

impl PortableDispatch for ExcelAppObject {
    fn invoke_dynamic(
        &self,
        member: &str,
        _kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(member, args)
    }

    fn invoke_early_bound(
        &self,
        member: &TypeLibMemberMetadata,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(&member.name, args)
    }

    fn invoke(&self, member: &str, _args: &[Variant]) -> Result<Variant, String> {
        Err(format!("unexpected ExcelApp invoke `{member}`"))
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unexpected ExcelApp get `{member}`"))
    }
    fn put(&self, member: &str, _value: Variant) -> Result<(), String> {
        Err(format!("unexpected ExcelApp put `{member}`"))
    }
    fn member_names(&self) -> Vec<String> {
        global_and_app_members()
            .into_iter()
            .map(|m| m.name)
            .collect()
    }
}

struct ExcelWorkbookObject {
    wb: SharedWorkbook,
}

impl ExcelWorkbookObject {
    fn handle_call(
        &self,
        member: &str,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        match member.to_ascii_lowercase().as_str() {
            "worksheets" | "sheets" => {
                if let Some(idx) = arg_variant(args, 0) {
                    return ExcelWorksheetsObject {
                        wb: self.wb.clone(),
                    }
                    .get_sheet(idx);
                }
                Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                    "Excel.Worksheets",
                    Box::new(ExcelWorksheetsObject {
                        wb: self.wb.clone(),
                    }),
                )))
            }
            "activesheet" => Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                "Excel.Worksheet",
                Box::new(ExcelWorksheetObject {
                    wb: self.wb.clone(),
                    sheet_index: 0,
                }),
            ))),
            "application" => Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                "Excel.Application",
                Box::new(ExcelAppObject {
                    wb: self.wb.clone(),
                }),
            ))),
            other => Err(format!("unsupported Excel.Workbook member `{other}`")),
        }
    }
}

impl PortableDispatch for ExcelWorkbookObject {
    fn invoke_dynamic(
        &self,
        member: &str,
        _kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(member, args)
    }

    fn invoke_early_bound(
        &self,
        member: &TypeLibMemberMetadata,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(&member.name, args)
    }

    fn invoke(&self, member: &str, _args: &[Variant]) -> Result<Variant, String> {
        Err(format!("unexpected Workbook invoke `{member}`"))
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unexpected Workbook get `{member}`"))
    }
    fn put(&self, member: &str, _value: Variant) -> Result<(), String> {
        Err(format!("unexpected Workbook put `{member}`"))
    }
    fn member_names(&self) -> Vec<String> {
        workbook_members().into_iter().map(|m| m.name).collect()
    }
}

struct ExcelWorksheetsObject {
    wb: SharedWorkbook,
}

impl ExcelWorksheetsObject {
    fn get_sheet(&self, idx: &Variant) -> Result<PortableTypedCallResult, String> {
        let state = self.wb.lock().unwrap();
        let sheet_index = if idx.vtype() == VarType::String {
            let name = variant_to_string(idx)?;
            state
                .sheets
                .iter()
                .position(|s| s.name.eq_ignore_ascii_case(&name))
                .ok_or_else(|| format!("subscript out of range: worksheet `{name}`"))?
        } else {
            let i = variant_to_i32(idx)?;
            if i < 1 || (i as usize) > state.sheets.len() {
                return Err(format!("subscript out of range: worksheet {i}"));
            }
            (i - 1) as usize
        };
        drop(state);
        Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
            "Excel.Worksheet",
            Box::new(ExcelWorksheetObject {
                wb: self.wb.clone(),
                sheet_index,
            }),
        )))
    }

    fn handle_call(
        &self,
        member: &str,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        match member.to_ascii_lowercase().as_str() {
            "item" => {
                let idx = arg_variant(args, 0).ok_or("Worksheets.Item requires an index")?;
                self.get_sheet(idx)
            }
            "count" => {
                let count = self.wb.lock().unwrap().sheets.len() as i32;
                Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                    Variant::from_i32(count),
                )))
            }
            other => Err(format!("unsupported Excel.Worksheets member `{other}`")),
        }
    }
}

impl PortableDispatch for ExcelWorksheetsObject {
    fn invoke_dynamic(
        &self,
        member: &str,
        _kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(member, args)
    }

    fn invoke_early_bound(
        &self,
        member: &TypeLibMemberMetadata,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(&member.name, args)
    }

    fn invoke(&self, member: &str, _args: &[Variant]) -> Result<Variant, String> {
        Err(format!("unexpected Worksheets invoke `{member}`"))
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unexpected Worksheets get `{member}`"))
    }
    fn put(&self, member: &str, _value: Variant) -> Result<(), String> {
        Err(format!("unexpected Worksheets put `{member}`"))
    }
    fn member_names(&self) -> Vec<String> {
        worksheets_members().into_iter().map(|m| m.name).collect()
    }
}

struct ExcelWorksheetObject {
    wb: SharedWorkbook,
    sheet_index: usize,
}

impl ExcelWorksheetObject {
    fn full_sheet_range(&self, mode: RangeMode) -> RangeSpec {
        RangeSpec {
            sheet_index: self.sheet_index,
            areas: vec![Rect {
                r1: 1,
                c1: 1,
                r2: MAX_ROWS,
                c2: MAX_COLS,
            }],
            mode,
        }
    }

    fn handle_call(
        &self,
        member: &str,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        match member.to_ascii_lowercase().as_str() {
            "range" => {
                let c1 = arg_variant(args, 0).ok_or("Worksheet.Range requires Cell1")?;
                let c2 = arg_variant(args, 1);
                let spec = resolve_range_arg(&self.wb, self.sheet_index, (1, 1), c1, c2)?;
                Ok(make_range_result(self.wb.clone(), spec))
            }
            "cells" => {
                let full = self.full_sheet_range(RangeMode::Cells);
                if arg_variant(args, 0).is_some() || arg_variant(args, 1).is_some() {
                    let obj = ExcelRangeObject {
                        wb: self.wb.clone(),
                        spec: full,
                    };
                    return obj.handle_item_get(args);
                }
                Ok(make_range_result(self.wb.clone(), full))
            }
            "rows" => {
                let full = self.full_sheet_range(RangeMode::Rows);
                if arg_variant(args, 0).is_some() {
                    let obj = ExcelRangeObject {
                        wb: self.wb.clone(),
                        spec: full,
                    };
                    return obj.handle_item_get(args);
                }
                Ok(make_range_result(self.wb.clone(), full))
            }
            "columns" => {
                let full = self.full_sheet_range(RangeMode::Columns);
                if arg_variant(args, 0).is_some() {
                    let obj = ExcelRangeObject {
                        wb: self.wb.clone(),
                        spec: full,
                    };
                    return obj.handle_item_get(args);
                }
                Ok(make_range_result(self.wb.clone(), full))
            }
            "name" => {
                let name = self.wb.lock().unwrap().sheets[self.sheet_index]
                    .name
                    .clone();
                Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                    Variant::from_string(name),
                )))
            }
            "application" => Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                "Excel.Application",
                Box::new(ExcelAppObject {
                    wb: self.wb.clone(),
                }),
            ))),
            other => Err(format!("unsupported Excel.Worksheet member `{other}`")),
        }
    }
}

impl PortableDispatch for ExcelWorksheetObject {
    fn invoke_dynamic(
        &self,
        member: &str,
        _kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(member, args)
    }

    fn invoke_early_bound(
        &self,
        member: &TypeLibMemberMetadata,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(&member.name, args)
    }

    fn invoke(&self, member: &str, _args: &[Variant]) -> Result<Variant, String> {
        Err(format!("unexpected Worksheet invoke `{member}`"))
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unexpected Worksheet get `{member}`"))
    }
    fn put(&self, member: &str, _value: Variant) -> Result<(), String> {
        Err(format!("unexpected Worksheet put `{member}`"))
    }
    fn member_names(&self) -> Vec<String> {
        worksheet_members().into_iter().map(|m| m.name).collect()
    }
}

struct ExcelAreasObject {
    wb: SharedWorkbook,
    sheet_index: usize,
    areas: Vec<Rect>,
}

impl ExcelAreasObject {
    fn handle_call(
        &self,
        member: &str,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        match member.to_ascii_lowercase().as_str() {
            "count" => Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                Variant::from_i32(self.areas.len() as i32),
            ))),
            "item" => {
                let idx = arg_variant(args, 0)
                    .map(variant_to_i32)
                    .transpose()?
                    .ok_or("Areas.Item requires an index")?;
                if idx < 1 || (idx as usize) > self.areas.len() {
                    return Err(format!("subscript out of range: area {idx}"));
                }
                let rect = self.areas[(idx - 1) as usize];
                Ok(make_range_result(
                    self.wb.clone(),
                    RangeSpec {
                        sheet_index: self.sheet_index,
                        areas: vec![rect],
                        mode: RangeMode::Cells,
                    },
                ))
            }
            other => Err(format!("unsupported Excel.Areas member `{other}`")),
        }
    }
}

impl PortableDispatch for ExcelAreasObject {
    fn invoke_dynamic(
        &self,
        member: &str,
        _kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(member, args)
    }

    fn invoke_early_bound(
        &self,
        member: &TypeLibMemberMetadata,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(&member.name, args)
    }

    fn invoke(&self, member: &str, _args: &[Variant]) -> Result<Variant, String> {
        Err(format!("unexpected Areas invoke `{member}`"))
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unexpected Areas get `{member}`"))
    }
    fn put(&self, member: &str, _value: Variant) -> Result<(), String> {
        Err(format!("unexpected Areas put `{member}`"))
    }
    fn member_names(&self) -> Vec<String> {
        areas_members().into_iter().map(|m| m.name).collect()
    }
}

struct ExcelRangeObject {
    wb: SharedWorkbook,
    spec: RangeSpec,
}

impl ExcelRangeObject {
    fn resolve_item_rect(&self, args: &[DynamicCallArg]) -> Result<Rect, String> {
        let first = self.spec.first();
        let a0 = arg_variant(args, 0);
        let a1 = arg_variant(args, 1);
        match self.spec.mode {
            RangeMode::Rows => {
                let row_idx = a0.map(variant_to_i32).transpose()?.unwrap_or(1);
                let r = first.r1 + row_idx - 1;
                Rect::new(r, first.c1, r, first.c2)
            }
            RangeMode::Columns => {
                let col_idx = a0.map(variant_to_i32).transpose()?.unwrap_or(1);
                let c = first.c1 + col_idx - 1;
                Rect::new(first.r1, c, first.r2, c)
            }
            RangeMode::Cells => match (a0, a1) {
                (Some(r_v), Some(c_v)) => {
                    let r = first.r1 + variant_to_i32(r_v)? - 1;
                    let c = if c_v.vtype() == VarType::String {
                        first.c1 + parse_col(&variant_to_string(c_v)?)? - 1
                    } else {
                        first.c1 + variant_to_i32(c_v)? - 1
                    };
                    Rect::new(r, c, r, c)
                }
                (Some(idx_v), None) => {
                    let idx = variant_to_i32(idx_v)?;
                    if idx < 1 {
                        return Err("application-defined or object-defined error".into());
                    }
                    let w = first.cols();
                    let r = first.r1 + (idx - 1) / w;
                    let c = first.c1 + (idx - 1) % w;
                    Rect::new(r, c, r, c)
                }
                (None, None) => Ok(first),
                (None, Some(_)) => Err("invalid Range.Item arguments".into()),
            },
        }
    }

    fn handle_item_get(&self, args: &[DynamicCallArg]) -> Result<PortableTypedCallResult, String> {
        let rect = self.resolve_item_rect(args)?;
        Ok(make_range_result(
            self.wb.clone(),
            RangeSpec {
                sheet_index: self.spec.sheet_index,
                areas: vec![rect],
                mode: RangeMode::Cells,
            },
        ))
    }

    fn read_values(&self, use_value2: bool) -> Variant {
        let first = self.spec.first();
        let state = self.wb.lock().unwrap();
        let sheet = &state.sheets[self.spec.sheet_index];
        if first.is_single_cell() {
            return sheet
                .cells
                .get(&(first.r1, first.c1))
                .map_or_else(Variant::empty, |c| {
                    if use_value2 {
                        c.to_value2()
                    } else {
                        c.to_value()
                    }
                });
        }
        let rows = first.rows() as u32;
        let cols = first.cols() as u32;
        // SafeArray stores the first dimension (row) fastest (column-major).
        let mut values = Vec::with_capacity((rows * cols) as usize);
        for c in first.c1..=first.c2 {
            for r in first.r1..=first.r2 {
                let v = sheet
                    .cells
                    .get(&(r, c))
                    .map_or_else(Variant::empty, |cell| {
                        if use_value2 {
                            cell.to_value2()
                        } else {
                            cell.to_value()
                        }
                    });
                values.push(v);
            }
        }
        let sa = SafeArray::from_variants_nd(
            vec![
                SafeArrayBound {
                    lower: 1,
                    count: rows,
                },
                SafeArrayBound {
                    lower: 1,
                    count: cols,
                },
            ],
            values,
        );
        Variant::from_safearray(sa)
    }

    fn write_values(&self, rhs: &Variant) -> Result<(), String> {
        let mut state = self.wb.lock().unwrap();
        let sheet = &mut state.sheets[self.spec.sheet_index];
        if rhs.vtype() == VarType::ArrayVariant {
            let Some((bounds, _len)) = rhs.safearray_bounds_len() else {
                return Err("invalid array for Range.Value write".into());
            };
            let first = self.spec.first();
            if bounds.len() == 1 {
                let lb = bounds[0].lower;
                let cnt = bounds[0].count as i32;
                for r in first.r1..=first.r2 {
                    for c in first.c1..=first.c2 {
                        let offset = c - first.c1;
                        let content = if offset < cnt {
                            let idx = calcvba_runtime::safe_array::element_index(
                                &bounds,
                                &[i64::from(lb + offset)],
                            )
                            .unwrap_or(0);
                            rhs.safearray_element(idx)
                                .and_then(Result::ok)
                                .map_or(CellContent::Empty, |v| variant_to_cell_content(&v))
                        } else {
                            CellContent::Empty
                        };
                        sheet.cells.entry((r, c)).or_default().content = content;
                    }
                }
                return Ok(());
            } else if bounds.len() == 2 {
                let r_lb = bounds[0].lower;
                let r_cnt = bounds[0].count as i32;
                let c_lb = bounds[1].lower;
                let c_cnt = bounds[1].count as i32;
                for r in first.r1..=first.r2 {
                    for c in first.c1..=first.c2 {
                        let ro = r - first.r1;
                        let co = c - first.c1;
                        let content = if ro < r_cnt && co < c_cnt {
                            let idx = calcvba_runtime::safe_array::element_index(
                                &bounds,
                                &[i64::from(r_lb + ro), i64::from(c_lb + co)],
                            )
                            .unwrap_or(0);
                            rhs.safearray_element(idx)
                                .and_then(Result::ok)
                                .map_or(CellContent::Empty, |v| variant_to_cell_content(&v))
                        } else {
                            CellContent::Empty
                        };
                        sheet.cells.entry((r, c)).or_default().content = content;
                    }
                }
                return Ok(());
            }
        }
        let content = variant_to_cell_content(rhs);
        for rect in &self.spec.areas {
            for r in rect.r1..=rect.r2 {
                for c in rect.c1..=rect.c2 {
                    sheet.cells.entry((r, c)).or_default().content = content.clone();
                }
            }
        }
        Ok(())
    }

    fn read_formula(&self) -> Variant {
        let first = self.spec.first();
        let state = self.wb.lock().unwrap();
        let sheet = &state.sheets[self.spec.sheet_index];
        if first.is_single_cell() {
            let s = sheet
                .cells
                .get(&(first.r1, first.c1))
                .map_or_else(String::new, CellData::to_formula_string);
            return Variant::from_string(s);
        }
        let rows = first.rows() as u32;
        let cols = first.cols() as u32;
        let mut values = Vec::with_capacity((rows * cols) as usize);
        for c in first.c1..=first.c2 {
            for r in first.r1..=first.r2 {
                let s = sheet
                    .cells
                    .get(&(r, c))
                    .map_or_else(String::new, CellData::to_formula_string);
                values.push(Variant::from_string(s));
            }
        }
        let sa = SafeArray::from_variants_nd(
            vec![
                SafeArrayBound {
                    lower: 1,
                    count: rows,
                },
                SafeArrayBound {
                    lower: 1,
                    count: cols,
                },
            ],
            values,
        );
        Variant::from_safearray(sa)
    }

    fn write_formula(&self, rhs: &Variant) -> Result<(), String> {
        let s = variant_to_string(rhs)?;
        let content = if s.starts_with('=') {
            CellContent::Formula(s)
        } else if s.is_empty() {
            CellContent::Empty
        } else if let Ok(n) = s.trim().parse::<f64>() {
            CellContent::Number(n)
        } else {
            CellContent::Text(s)
        };
        let mut state = self.wb.lock().unwrap();
        let sheet = &mut state.sheets[self.spec.sheet_index];
        for rect in &self.spec.areas {
            for r in rect.r1..=rect.r2 {
                for c in rect.c1..=rect.c2 {
                    sheet.cells.entry((r, c)).or_default().content = content.clone();
                }
            }
        }
        Ok(())
    }

    fn handle_address(&self, args: &[DynamicCallArg]) -> Result<Variant, String> {
        let row_abs = arg_named_or_pos(args, 0, "RowAbsolute")
            .map(variant_to_bool)
            .transpose()?
            .unwrap_or(true);
        let col_abs = arg_named_or_pos(args, 1, "ColumnAbsolute")
            .map(variant_to_bool)
            .transpose()?
            .unwrap_or(true);
        let style = arg_named_or_pos(args, 2, "ReferenceStyle")
            .map(variant_to_i32)
            .transpose()?
            .unwrap_or(XL_A1);
        let _external = arg_named_or_pos(args, 3, "External")
            .map(variant_to_bool)
            .transpose()?
            .unwrap_or(false);
        let (ref_r, ref_c) = if let Some(rel) = arg_named_or_pos(args, 4, "RelativeTo") {
            let spec = lookup_range(&self.wb, rel)?;
            (spec.first().r1, spec.first().c1)
        } else {
            (1, 1)
        };
        let mut parts = Vec::with_capacity(self.spec.areas.len());
        for rect in &self.spec.areas {
            let fmt = |r, c| {
                if style == XL_R1C1 {
                    format_r1c1_cell(r, c, row_abs, col_abs, ref_r, ref_c)
                } else {
                    format_a1_cell(r, c, row_abs, col_abs)
                }
            };
            if rect.is_single_cell() {
                parts.push(fmt(rect.r1, rect.c1));
            } else {
                parts.push(format!(
                    "{}:{}",
                    fmt(rect.r1, rect.c1),
                    fmt(rect.r2, rect.c2)
                ));
            }
        }
        Ok(Variant::from_string(parts.join(",")))
    }

    fn handle_current_region(&self) -> Result<PortableTypedCallResult, String> {
        let first = self.spec.first();
        let state = self.wb.lock().unwrap();
        let sheet = &state.sheets[self.spec.sheet_index];
        let mut r1 = first.r1;
        let mut c1 = first.c1;
        let mut r2 = first.r2;
        let mut c2 = first.c2;
        loop {
            let mut expanded = false;
            let lo_r = (r1 - 1).max(1);
            let hi_r = (r2 + 1).min(MAX_ROWS);
            let lo_c = (c1 - 1).max(1);
            let hi_c = (c2 + 1).min(MAX_COLS);
            for (&(r, c), cell) in &sheet.cells {
                if !cell.is_non_empty() {
                    continue;
                }
                if r >= lo_r && r <= hi_r && c >= lo_c && c <= hi_c {
                    if r < r1 {
                        r1 = r;
                        expanded = true;
                    }
                    if r > r2 {
                        r2 = r;
                        expanded = true;
                    }
                    if c < c1 {
                        c1 = c;
                        expanded = true;
                    }
                    if c > c2 {
                        c2 = c;
                        expanded = true;
                    }
                }
            }
            if !expanded {
                break;
            }
        }
        drop(state);
        Ok(make_range_result(
            self.wb.clone(),
            RangeSpec {
                sheet_index: self.spec.sheet_index,
                areas: vec![Rect { r1, c1, r2, c2 }],
                mode: RangeMode::Cells,
            },
        ))
    }

    fn handle_end(&self, args: &[DynamicCallArg]) -> Result<PortableTypedCallResult, String> {
        let dir = arg_variant(args, 0)
            .map(variant_to_i32)
            .transpose()?
            .ok_or("Range.End requires Direction")?;
        let (dr, dc) = match dir {
            XL_DOWN => (1, 0),
            XL_UP => (-1, 0),
            XL_TO_RIGHT => (0, 1),
            XL_TO_LEFT => (0, -1),
            other => return Err(format!("invalid Range.End direction {other}")),
        };
        let first = self.spec.first();
        let state = self.wb.lock().unwrap();
        let sheet = &state.sheets[self.spec.sheet_index];
        let filled = |r: i32, c: i32| sheet.cells.get(&(r, c)).is_some_and(CellData::is_non_empty);
        let in_bounds = |r: i32, c: i32| (1..=MAX_ROWS).contains(&r) && (1..=MAX_COLS).contains(&c);
        let (mut r, mut c) = (first.r1, first.c1);
        let (nr, nc) = (r + dr, c + dc);
        if in_bounds(nr, nc) {
            if filled(r, c) && filled(nr, nc) {
                while in_bounds(r + dr, c + dc) && filled(r + dr, c + dc) {
                    r += dr;
                    c += dc;
                }
            } else {
                // Jump over blanks until the first filled cell along the ray or the sheet edge.
                let mut found: Option<(i32, i32)> = None;
                for (&(cr, cc), cell) in &sheet.cells {
                    if !cell.is_non_empty() {
                        continue;
                    }
                    let on_ray = match (dr, dc) {
                        (1, 0) => cc == c && cr > r,
                        (-1, 0) => cc == c && cr < r,
                        (0, 1) => cr == r && cc > c,
                        (0, -1) => cr == r && cc < c,
                        _ => false,
                    };
                    if !on_ray {
                        continue;
                    }
                    found = Some(match found {
                        None => (cr, cc),
                        Some((br, bc)) => {
                            let better = match (dr, dc) {
                                (1, 0) => cr < br,
                                (-1, 0) => cr > br,
                                (0, 1) => cc < bc,
                                (0, -1) => cc > bc,
                                _ => false,
                            };
                            if better { (cr, cc) } else { (br, bc) }
                        }
                    });
                }
                if let Some((fr, fc)) = found {
                    r = fr;
                    c = fc;
                } else {
                    match (dr, dc) {
                        (1, 0) => r = MAX_ROWS,
                        (-1, 0) => r = 1,
                        (0, 1) => c = MAX_COLS,
                        (0, -1) => c = 1,
                        _ => {}
                    }
                }
            }
        }
        drop(state);
        Ok(make_range_result(
            self.wb.clone(),
            RangeSpec {
                sheet_index: self.spec.sheet_index,
                areas: vec![Rect::single(r, c)],
                mode: RangeMode::Cells,
            },
        ))
    }

    fn handle_find(&self, args: &[DynamicCallArg]) -> Result<PortableTypedCallResult, String> {
        let what = arg_named_or_pos(args, 0, "What").ok_or("Range.Find requires What")?;
        let what_text = variant_to_string(what)?;
        let after = arg_named_or_pos(args, 1, "After")
            .map(|v| lookup_range(&self.wb, v))
            .transpose()?;
        let look_in = arg_named_or_pos(args, 2, "LookIn")
            .map(variant_to_i32)
            .transpose()?
            .unwrap_or(XL_VALUES);
        let look_at = arg_named_or_pos(args, 3, "LookAt")
            .map(variant_to_i32)
            .transpose()?
            .unwrap_or(XL_PART);
        let search_order = arg_named_or_pos(args, 4, "SearchOrder")
            .map(variant_to_i32)
            .transpose()?
            .unwrap_or(XL_BY_ROWS);
        let search_dir = arg_named_or_pos(args, 5, "SearchDirection")
            .map(variant_to_i32)
            .transpose()?
            .unwrap_or(XL_NEXT);
        let match_case = arg_named_or_pos(args, 6, "MatchCase")
            .map(variant_to_bool)
            .transpose()?
            .unwrap_or(false);

        let first = self.spec.first();
        let mut coords = Vec::new();
        if search_order == XL_BY_COLUMNS {
            for c in first.c1..=first.c2 {
                for r in first.r1..=first.r2 {
                    coords.push((r, c));
                }
            }
        } else {
            for r in first.r1..=first.r2 {
                for c in first.c1..=first.c2 {
                    coords.push((r, c));
                }
            }
        }
        if coords.is_empty() {
            return Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                Variant::nothing(),
            )));
        }
        let after_pos = after
            .and_then(|spec| {
                let a = spec.first();
                coords.iter().position(|&pt| pt == (a.r1, a.c1))
            })
            .unwrap_or(0);
        let n = coords.len();
        let state = self.wb.lock().unwrap();
        let sheet = &state.sheets[self.spec.sheet_index];
        let needle = if match_case {
            what_text.clone()
        } else {
            what_text.to_lowercase()
        };
        for step in 1..=n {
            let idx = if search_dir == XL_PREVIOUS {
                (after_pos + n - (step % n)) % n
            } else {
                (after_pos + step) % n
            };
            let (r, c) = coords[idx];
            let Some(cell) = sheet.cells.get(&(r, c)) else {
                continue;
            };
            if !cell.is_non_empty() {
                continue;
            }
            let hay = cell.display_text(look_in);
            let hay_cmp = if match_case { hay } else { hay.to_lowercase() };
            let matched = if look_at == XL_WHOLE {
                hay_cmp == needle
            } else {
                hay_cmp.contains(&needle)
            };
            if matched {
                drop(state);
                return Ok(make_range_result(
                    self.wb.clone(),
                    RangeSpec {
                        sheet_index: self.spec.sheet_index,
                        areas: vec![Rect::single(r, c)],
                        mode: RangeMode::Cells,
                    },
                ));
            }
        }
        Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
            Variant::nothing(),
        )))
    }

    fn handle_special_cells(
        &self,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        let cell_type = arg_named_or_pos(args, 0, "Type")
            .map(variant_to_i32)
            .transpose()?
            .ok_or("Range.SpecialCells requires Type")?;
        let value_mask = arg_named_or_pos(args, 1, "Value")
            .map(variant_to_i32)
            .transpose()?
            .unwrap_or(XL_NUMBERS | XL_TEXT_VALUES | XL_LOGICAL | XL_ERRORS);
        let state = self.wb.lock().unwrap();
        let sheet = &state.sheets[self.spec.sheet_index];
        let mut matched_areas = Vec::new();
        for rect in &self.spec.areas {
            for r in rect.r1..=rect.r2 {
                for c in rect.c1..=rect.c2 {
                    let cell = sheet.cells.get(&(r, c));
                    let matches = match cell_type {
                        XL_CELL_TYPE_BLANKS => cell.is_none_or(|d| !d.is_non_empty()),
                        XL_CELL_TYPE_FORMULAS => {
                            cell.is_some_and(|d| matches!(d.content, CellContent::Formula(_)))
                        }
                        XL_CELL_TYPE_CONSTANTS => cell.is_some_and(|d| match &d.content {
                            CellContent::Number(_) => (value_mask & XL_NUMBERS) != 0,
                            CellContent::Text(_) => (value_mask & XL_TEXT_VALUES) != 0,
                            CellContent::Bool(_) => (value_mask & XL_LOGICAL) != 0,
                            CellContent::Error(_) => (value_mask & XL_ERRORS) != 0,
                            CellContent::Empty | CellContent::Formula(_) => false,
                        }),
                        other => return Err(format!("unsupported SpecialCells type {other}")),
                    };
                    if matches {
                        matched_areas.push(Rect::single(r, c));
                    }
                }
            }
        }
        drop(state);
        if matched_areas.is_empty() {
            return Err("No cells were found.".into());
        }
        Ok(make_range_result(
            self.wb.clone(),
            RangeSpec {
                sheet_index: self.spec.sheet_index,
                areas: matched_areas,
                mode: RangeMode::Cells,
            },
        ))
    }

    fn handle_call(
        &self,
        member: &str,
        kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        let is_put = matches!(
            kind,
            Some(DynamicCallKind::PropertyLet | DynamicCallKind::PropertySet)
        );
        match member.to_ascii_lowercase().as_str() {
            "item" => {
                if is_put {
                    let rhs = args
                        .last()
                        .and_then(|a| a.value.as_ref())
                        .map(|v| v.variant().clone())
                        .unwrap_or_else(Variant::empty);
                    let idx_args = &args[..args.len().saturating_sub(1)];
                    let rect = self.resolve_item_rect(idx_args)?;
                    let cell_range = ExcelRangeObject {
                        wb: self.wb.clone(),
                        spec: RangeSpec {
                            sheet_index: self.spec.sheet_index,
                            areas: vec![rect],
                            mode: RangeMode::Cells,
                        },
                    };
                    cell_range.write_values(&rhs)?;
                    return Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        Variant::empty(),
                    )));
                }
                self.handle_item_get(args)
            }
            "range" => {
                let first = self.spec.first();
                let c1 = arg_variant(args, 0).ok_or("Range.Range requires Cell1")?;
                let c2 = arg_variant(args, 1);
                let spec = resolve_range_arg(
                    &self.wb,
                    self.spec.sheet_index,
                    (first.r1, first.c1),
                    c1,
                    c2,
                )?;
                Ok(make_range_result(self.wb.clone(), spec))
            }
            "cells" => {
                let base = RangeSpec {
                    sheet_index: self.spec.sheet_index,
                    areas: vec![self.spec.first()],
                    mode: RangeMode::Cells,
                };
                if arg_variant(args, 0).is_some() || arg_variant(args, 1).is_some() {
                    let obj = ExcelRangeObject {
                        wb: self.wb.clone(),
                        spec: base,
                    };
                    return obj.handle_item_get(args);
                }
                Ok(make_range_result(self.wb.clone(), base))
            }
            "rows" => {
                let base = RangeSpec {
                    sheet_index: self.spec.sheet_index,
                    areas: vec![self.spec.first()],
                    mode: RangeMode::Rows,
                };
                if arg_variant(args, 0).is_some() {
                    let obj = ExcelRangeObject {
                        wb: self.wb.clone(),
                        spec: base,
                    };
                    return obj.handle_item_get(args);
                }
                Ok(make_range_result(self.wb.clone(), base))
            }
            "columns" => {
                let base = RangeSpec {
                    sheet_index: self.spec.sheet_index,
                    areas: vec![self.spec.first()],
                    mode: RangeMode::Columns,
                };
                if arg_variant(args, 0).is_some() {
                    let obj = ExcelRangeObject {
                        wb: self.wb.clone(),
                        spec: base,
                    };
                    return obj.handle_item_get(args);
                }
                Ok(make_range_result(self.wb.clone(), base))
            }
            "areas" => Ok(PortableTypedCallResult::Object(PortableObjectResult::new(
                "Excel.Areas",
                Box::new(ExcelAreasObject {
                    wb: self.wb.clone(),
                    sheet_index: self.spec.sheet_index,
                    areas: self.spec.areas.clone(),
                }),
            ))),
            "count" => {
                let first = self.spec.first();
                let count = match self.spec.mode {
                    RangeMode::Rows => first.rows(),
                    RangeMode::Columns => first.cols(),
                    RangeMode::Cells => first.rows().saturating_mul(first.cols()),
                };
                Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                    Variant::from_i32(count),
                )))
            }
            "row" => Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                Variant::from_i32(self.spec.first().r1),
            ))),
            "column" => Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                Variant::from_i32(self.spec.first().c1),
            ))),
            "offset" => {
                let dr = arg_named_or_pos(args, 0, "RowOffset")
                    .map(variant_to_i32)
                    .transpose()?
                    .unwrap_or(0);
                let dc = arg_named_or_pos(args, 1, "ColumnOffset")
                    .map(variant_to_i32)
                    .transpose()?
                    .unwrap_or(0);
                let areas = self
                    .spec
                    .areas
                    .iter()
                    .map(|rect| Rect::new(rect.r1 + dr, rect.c1 + dc, rect.r2 + dr, rect.c2 + dc))
                    .collect::<Result<Vec<_>, _>>()?;
                Ok(make_range_result(
                    self.wb.clone(),
                    RangeSpec {
                        sheet_index: self.spec.sheet_index,
                        areas,
                        mode: RangeMode::Cells,
                    },
                ))
            }
            "resize" => {
                let first = self.spec.first();
                let rows = arg_named_or_pos(args, 0, "RowSize")
                    .map(variant_to_i32)
                    .transpose()?
                    .unwrap_or(first.rows());
                let cols = arg_named_or_pos(args, 1, "ColumnSize")
                    .map(variant_to_i32)
                    .transpose()?
                    .unwrap_or(first.cols());
                if rows < 1 || cols < 1 {
                    return Err("Resize dimensions must be positive".into());
                }
                let rect = Rect::new(first.r1, first.c1, first.r1 + rows - 1, first.c1 + cols - 1)?;
                Ok(make_range_result(
                    self.wb.clone(),
                    RangeSpec {
                        sheet_index: self.spec.sheet_index,
                        areas: vec![rect],
                        mode: RangeMode::Cells,
                    },
                ))
            }
            "address" => self
                .handle_address(args)
                .map(RuntimeCallResult::value)
                .map(PortableTypedCallResult::Value),
            "currentregion" => self.handle_current_region(),
            "end" => self.handle_end(args),
            "find" => self.handle_find(args),
            "specialcells" => self.handle_special_cells(args),
            "value" => {
                if is_put {
                    let rhs = arg_variant_raw(args, args.len().saturating_sub(1))
                        .cloned()
                        .unwrap_or_else(Variant::empty);
                    self.write_values(&rhs)?;
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        Variant::empty(),
                    )))
                } else {
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        self.read_values(false),
                    )))
                }
            }
            "value2" => {
                if is_put {
                    let rhs = arg_variant_raw(args, args.len().saturating_sub(1))
                        .cloned()
                        .unwrap_or_else(Variant::empty);
                    self.write_values(&rhs)?;
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        Variant::empty(),
                    )))
                } else {
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        self.read_values(true),
                    )))
                }
            }
            "formula" => {
                if is_put {
                    let rhs = arg_variant_raw(args, args.len().saturating_sub(1))
                        .cloned()
                        .unwrap_or_else(Variant::empty);
                    self.write_formula(&rhs)?;
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        Variant::empty(),
                    )))
                } else {
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        self.read_formula(),
                    )))
                }
            }
            "numberformat" => {
                if is_put {
                    let rhs = arg_variant_raw(args, args.len().saturating_sub(1))
                        .cloned()
                        .unwrap_or_else(Variant::empty);
                    let fmt = variant_to_string(&rhs)?;
                    let mut state = self.wb.lock().unwrap();
                    let sheet = &mut state.sheets[self.spec.sheet_index];
                    for rect in &self.spec.areas {
                        for r in rect.r1..=rect.r2 {
                            for c in rect.c1..=rect.c2 {
                                sheet.cells.entry((r, c)).or_default().number_format = fmt.clone();
                            }
                        }
                    }
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        Variant::empty(),
                    )))
                } else {
                    let first = self.spec.first();
                    let state = self.wb.lock().unwrap();
                    let fmt = state.sheets[self.spec.sheet_index]
                        .cells
                        .get(&(first.r1, first.c1))
                        .map_or_else(|| "General".into(), |c| c.number_format.clone());
                    Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                        Variant::from_string(fmt),
                    )))
                }
            }
            other => Err(format!("unsupported Excel.Range member `{other}`")),
        }
    }
}

impl PortableDispatch for ExcelRangeObject {
    fn on_registered(&self, handle: i32) {
        self.wb
            .lock()
            .unwrap()
            .ranges_by_handle
            .insert(handle, self.spec.clone());
    }

    fn invoke_dynamic(
        &self,
        member: &str,
        kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(member, kind, args)
    }

    fn invoke_early_bound(
        &self,
        member: &TypeLibMemberMetadata,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(
            &member.name,
            Some(kind_from_invoke(member.invoke_kind)),
            args,
        )
    }

    fn invoke(&self, member: &str, _args: &[Variant]) -> Result<Variant, String> {
        Err(format!("unexpected Range invoke `{member}`"))
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unexpected Range get `{member}`"))
    }
    fn put(&self, member: &str, _value: Variant) -> Result<(), String> {
        Err(format!("unexpected Range put `{member}`"))
    }
    fn member_names(&self) -> Vec<String> {
        range_members().into_iter().map(|m| m.name).collect()
    }
}

struct ExcelWorksheetFunctionObject;

impl ExcelWorksheetFunctionObject {
    fn handle_call(
        &self,
        member: &str,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        match member.to_ascii_lowercase().as_str() {
            "round" => {
                let x = arg_variant(args, 0)
                    .map(variant_to_f64)
                    .transpose()?
                    .ok_or("WorksheetFunction.Round requires Arg1")?;
                let d = arg_variant(args, 1)
                    .map(variant_to_f64)
                    .transpose()?
                    .ok_or("WorksheetFunction.Round requires Arg2")?
                    .trunc() as i32;
                Ok(PortableTypedCallResult::Value(RuntimeCallResult::value(
                    Variant::from_f64(excel_round(x, d)),
                )))
            }
            other => Err(format!(
                "unsupported Excel.WorksheetFunction member `{other}`"
            )),
        }
    }
}

impl PortableDispatch for ExcelWorksheetFunctionObject {
    fn invoke_dynamic(
        &self,
        member: &str,
        _kind: Option<DynamicCallKind>,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(member, args)
    }

    fn invoke_early_bound(
        &self,
        member: &TypeLibMemberMetadata,
        args: &[DynamicCallArg],
    ) -> Result<PortableTypedCallResult, String> {
        self.handle_call(&member.name, args)
    }

    fn invoke(&self, member: &str, _args: &[Variant]) -> Result<Variant, String> {
        Err(format!("unexpected WorksheetFunction invoke `{member}`"))
    }
    fn get(&self, member: &str) -> Result<Variant, String> {
        Err(format!("unexpected WorksheetFunction get `{member}`"))
    }
    fn put(&self, member: &str, _value: Variant) -> Result<(), String> {
        Err(format!("unexpected WorksheetFunction put `{member}`"))
    }
    fn member_names(&self) -> Vec<String> {
        worksheet_function_members()
            .into_iter()
            .map(|m| m.name)
            .collect()
    }
}
