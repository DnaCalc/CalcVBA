use calcvba_host_api as api;
pub const IID: api::InterfaceId = *b"Range-like-host1";
pub fn project() -> api::Project {
    let mut project = api::Project::module(
        "Api",
        r#"
Public Function ReadRange(ByVal root As HostObjects.Range, ByVal offset As Long) As Double
    ReadRange = root.Read(offset) + 1#
End Function
Public Function ReadValue(ByVal root As HostObjects.Range) As Double
    ReadValue = root.NumericValue * 2#
End Function
Public Function Twice(ByVal x As Double) As Double
    Twice = x * 2#
End Function
Public Function Mixed(ByVal left As Double, ByVal root As HostObjects.Range, ByVal offset As Long, ByVal right As Double) As Double
    Mixed = left + root.Read(offset) + right
End Function
"#,
    );
    project.libraries.push(api::ObjectLibrary {
        name: "HostObjects".into(),
        interfaces: vec![api::ObjectInterface {
            name: "Range".into(),
            iid: IID,
            members: vec![
                api::ObjectMember {
                    name: "Read".into(),
                    slot: 0,
                    kind: api::MemberKind::Method,
                    parameters: vec![api::ScalarType::Long],
                    result: api::ScalarType::Double,
                    is_default: false,
                },
                api::ObjectMember {
                    name: "NumericValue".into(),
                    slot: 1,
                    kind: api::MemberKind::PropertyGet,
                    parameters: vec![],
                    result: api::ScalarType::Double,
                    is_default: true,
                },
            ],
        }],
    });
    project
}
pub fn backends() -> Vec<api::Backend> {
    let mut backends = vec![api::Backend::Interpreter];
    if cfg!(feature = "jit") {
        backends.push(api::Backend::Jit);
    }
    backends
}
