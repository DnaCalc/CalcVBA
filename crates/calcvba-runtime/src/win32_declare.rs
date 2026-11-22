//! Win32 Declare shapes shared by the portability report and the Unix host.

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Win32Declare {
    Sleep,
    GetCurrentProcessId,
    GetTickCount,
    GetTickCount64,
    QueryPerformanceCounter,
    QueryPerformanceFrequency,
    StringLengthWide,
    MoveMemory,
    MoveMemoryLongCount,
}

impl Win32Declare {
    /// Match the ABI as well as the entry name before calling a built-in shim.
    pub fn resolve(
        library: &str,
        entry: &str,
        params: &[(&str, bool)],
        result: Option<&str>,
    ) -> Option<Self> {
        if !library.eq_ignore_ascii_case("kernel32")
            && !library.eq_ignore_ascii_case("kernel32.dll")
        {
            return None;
        }

        let pointer = |(ty, by_ref): (&str, bool)| {
            if by_ref {
                matches!(
                    ty,
                    "Any"
                        | "Byte"
                        | "Boolean"
                        | "Integer"
                        | "Long"
                        | "LongLong"
                        | "LongPtr"
                        | "Single"
                        | "Double"
                        | "Currency"
                        | "Date"
                )
            } else {
                ty == "LongPtr"
            }
        };

        match (entry, params, result) {
            ("Sleep", [("Long", false)], None) => Some(Self::Sleep),
            ("GetCurrentProcessId", [], Some("Long")) => Some(Self::GetCurrentProcessId),
            ("GetTickCount", [], Some("Long")) => Some(Self::GetTickCount),
            ("GetTickCount64", [], Some("LongLong")) => Some(Self::GetTickCount64),
            (
                "QueryPerformanceCounter" | "QueryPerformanceFrequency",
                [("LongLong" | "Currency", true) | ("LongPtr", false)],
                Some("Long"),
            ) => {
                if entry == "QueryPerformanceCounter" {
                    Some(Self::QueryPerformanceCounter)
                } else {
                    Some(Self::QueryPerformanceFrequency)
                }
            }
            ("lstrlenW", [("LongPtr", false)], Some("Long")) => Some(Self::StringLengthWide),
            ("RtlMoveMemory", [dest, src, ("LongPtr", false)], None)
                if pointer(*dest) && pointer(*src) =>
            {
                Some(Self::MoveMemory)
            }
            ("RtlMoveMemory", [dest, src, ("Long", false)], None)
                if pointer(*dest) && pointer(*src) =>
            {
                Some(Self::MoveMemoryLongCount)
            }
            _ => None,
        }
    }
}
