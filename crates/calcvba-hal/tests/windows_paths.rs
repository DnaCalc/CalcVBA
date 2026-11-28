#![cfg(target_os = "linux")]
use calcvba_hal::{
    HalProfileId, HostPolicy, adapters::builder::HostBuilder, model::WindowsMountAccess,
};
use calcvba_runtime::Variant;
use std::fs;

#[test]
fn mapped_drive_uses_windows_case_and_remembers_project_directory() {
    let root = std::env::temp_dir().join(format!("calcvba-hal-paths-{}", std::process::id()));
    let data = root.join("Data");
    fs::create_dir_all(&data).unwrap();
    fs::write(data.join("Notes.txt"), b"notes").unwrap();
    let host = HostBuilder::new()
        .profile(HalProfileId::Linux)
        .policy(HostPolicy::native().with_windows_drive_root('c', &root))
        .build();
    let path = Variant::from_string;
    let filesystem = host.fs();
    let size = filesystem
        .file_len_variant(path("c:\\DATA\\notes.TXT"))
        .unwrap();
    filesystem.ch_dir_variant(path("C:\\dAta")).unwrap();
    let directory = filesystem.cur_dir_variant(path("")).unwrap();
    let entry = host
        .process()
        .dir_variant(path("*.TXT"), Variant::from_i32(0))
        .unwrap();
    let missing = filesystem.ch_drive_variant(path("Q:")).unwrap_err();
    let bad_drive = filesystem.ch_drive_variant(path("1")).unwrap_err();
    let missing_path = filesystem
        .file_len_variant(path(r"Q:\Notes.txt"))
        .unwrap_err();
    filesystem
        .file_copy_variant(path("C:\\DATA\\notes.TXT"), path("copy.txt"))
        .unwrap();
    let copy = fs::read(data.join("copy.txt")).unwrap();
    let copy_attributes = filesystem.get_attr_variant(path("COPY.TXT")).unwrap();
    fs::create_dir(root.join("data")).unwrap();
    let collision = filesystem
        .file_len_variant(path("C:\\DATA\\notes.TXT"))
        .unwrap_err();
    fs::remove_dir_all(root).unwrap();
    assert_eq!(size.as_i32(), Some(5));
    assert_eq!(directory.as_bstr().unwrap().as_str(), "C:\\dAta");
    assert_eq!(entry.as_bstr().unwrap().as_str(), "Notes.txt");
    assert_eq!(
        (missing.host_error_code, bad_drive.host_error_code),
        (Some(68), Some(5))
    );
    assert_eq!(missing_path.host_error_code, Some(76));
    assert_eq!(copy, b"notes");
    assert_eq!(copy_attributes.as_i32(), Some(32));
    assert_eq!(collision.host_error_code, Some(75));
    let diagnostic = collision.to_string();
    assert!(
        diagnostic.contains("Data") && diagnostic.contains("data"),
        "{diagnostic}"
    );
}

#[test]
fn mapped_directory_suffixes_match_excel() {
    let root = std::env::temp_dir().join(format!("calcvba-hal-path-suffix-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    fs::create_dir(root.join("sub")).unwrap();
    let host = HostBuilder::new()
        .profile(HalProfileId::Linux)
        .policy(HostPolicy::native().with_windows_drive_root('C', &root))
        .build();
    let path = Variant::from_string;
    let filesystem = host.fs();
    let created = filesystem.mkdir_variant(path(r"C:\dot."));
    let dot_exists = root.join("dot").is_dir();
    let changed = filesystem.ch_dir_variant(path(r"C:\sub "));
    let timestamp = filesystem.file_date_time_variant(path(r"C:\sub\"));
    fs::remove_dir_all(root).unwrap();
    created.unwrap();
    assert_eq!(
        (
            dot_exists,
            changed.is_ok(),
            timestamp.err().and_then(|error| error.host_error_code)
        ),
        (true, true, Some(53))
    );
}

#[test]
fn mapped_hidden_system_and_archive_bits_survive_a_new_host() {
    let root = std::env::temp_dir().join(format!("calcvba-hal-attrs-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    fs::write(root.join("doc.txt"), b"hi").unwrap();
    let path = || Variant::from_string("C:\\doc.txt");
    let make_host = || {
        HostBuilder::new()
            .profile(HalProfileId::Linux)
            .policy(HostPolicy::native().with_windows_drive_root('C', &root))
            .build()
    };
    let host = make_host();
    let archive = host.fs().get_attr_variant(path()).unwrap();
    host.fs()
        .set_attr_variant(path(), Variant::from_i32(6))
        .unwrap();
    drop(host);
    let host = make_host();
    let stored = host.fs().get_attr_variant(path()).unwrap();
    let normal = host
        .process()
        .dir_variant(Variant::from_string("*"), Variant::from_i32(0))
        .unwrap();
    let filtered = host
        .process()
        .dir_variant(Variant::from_string("*"), Variant::from_i32(6))
        .unwrap();
    fs::remove_dir_all(root).unwrap();
    assert_eq!(archive.as_i32(), Some(32));
    assert_eq!(stored.as_i32(), Some(6));
    assert_eq!(normal.as_bstr().unwrap().as_str(), "");
    assert_eq!(filtered.as_bstr().unwrap().as_str(), "doc.txt");
}

#[test]
fn mapped_dir_and_kill_report_both_names_in_a_wildcard_collision() {
    let root = std::env::temp_dir().join(format!(
        "calcvba-hal-wildcard-collision-{}",
        std::process::id()
    ));
    fs::create_dir(&root).unwrap();
    fs::write(root.join("Notes.txt"), b"one").unwrap();
    fs::write(root.join("notes.TXT"), b"two").unwrap();
    let host = HostBuilder::new()
        .profile(HalProfileId::Linux)
        .policy(HostPolicy::native().with_windows_drive_root('C', &root))
        .build();
    let pattern = || Variant::from_string("C:\\*.txt");
    let listed = host.process().dir_variant(pattern(), Variant::from_i32(0));
    let deleted = host.fs().kill_variant(pattern());
    let still_present = [
        root.join("Notes.txt").exists(),
        root.join("notes.TXT").exists(),
    ];
    fs::remove_dir_all(&root).unwrap();

    for result in [listed, deleted] {
        let error = result.unwrap_err().to_string();
        assert!(
            error.contains("Notes.txt") && error.contains("notes.TXT"),
            "{error}"
        );
    }
    assert_eq!(still_present, [true, true]);
}

#[test]
fn longest_drive_mount_resolves_relative_paths_and_enforces_read_only() {
    let root = std::env::temp_dir().join(format!("calcvba-hal-drive-mount-{}", std::process::id()));
    let drive = root.join("drive");
    let temp = root.join("overlay");
    let restricted = root.join("restricted");
    for directory in [&drive, &temp, &restricted] {
        fs::create_dir_all(directory).unwrap();
    }
    fs::write(temp.join("a.txt"), b"temp overlay").unwrap();
    fs::write(restricted.join("b.txt"), b"nested").unwrap();
    let policy = HostPolicy::native()
        .with_windows_drive_root('C', &drive)
        .with_windows_drive_mount(r"C:\Temp", &temp, WindowsMountAccess::ReadWrite)
        .unwrap()
        .with_windows_drive_mount(r"C:\Temp\Shared", &restricted, WindowsMountAccess::Read)
        .unwrap();
    let host = HostBuilder::new()
        .profile(HalProfileId::Linux)
        .policy(policy)
        .build();
    let file = host.fs();
    let path = Variant::from_string;
    file.ch_dir_variant(path(r"C:\tEMP\SHARED")).unwrap();
    assert_eq!(
        file.cur_dir_variant(path(""))
            .unwrap()
            .as_bstr()
            .unwrap()
            .as_str(),
        r"C:\Temp\Shared"
    );
    assert_eq!(
        file.file_len_variant(path(r"c:B.TXT")).unwrap().as_i32(),
        Some(6)
    );
    assert_eq!(
        file.file_len_variant(path(r"..\A.txt")).unwrap().as_i32(),
        Some(12)
    );
    let blocked = file
        .file_copy_variant(path(r"b.txt"), path(r"no.txt"))
        .unwrap_err();
    assert_eq!(blocked.host_error_code, Some(75));
    assert!(!restricted.join("no.txt").exists());
    file.file_copy_variant(path(r"b.txt"), path(r"..\yes.txt"))
        .unwrap();
    assert_eq!(fs::read(temp.join("yes.txt")).unwrap(), b"nested");
    fs::remove_dir_all(&root).unwrap();
}

#[test]
fn unc_mount_access_is_enforced_case_insensitively() {
    let root = std::env::temp_dir().join(format!("calcvba-hal-unc-mount-{}", std::process::id()));
    let input = root.join("input");
    let output = root.join("output");
    fs::create_dir_all(&input).unwrap();
    fs::create_dir_all(&output).unwrap();
    fs::write(input.join("Data.txt"), b"test data").unwrap();
    let host = HostBuilder::new()
        .profile(HalProfileId::Linux)
        .policy(
            HostPolicy::native()
                .with_windows_unc_mount("Server", "Input", &input, WindowsMountAccess::Read)
                .with_windows_unc_root("Server", "Output", &output),
        )
        .build();
    let file = host.fs();
    let path = Variant::from_string;
    let data = r"\\SERVER\input\data.TXT";
    assert_eq!(file.file_len_variant(path(data)).unwrap().as_i32(), Some(9));
    for denied in [
        file.file_copy_variant(path(data), path(r"\\Server\INPUT\blocked.txt")),
        file.name_variant(path(data), path(r"\\Server\Output\moved.txt")),
        file.kill_variant(path(r"\\Server\INPUT\*.txt")),
        file.set_attr_variant(path(data), Variant::from_i32(2)),
    ] {
        assert_eq!(denied.unwrap_err().host_error_code, Some(75));
    }
    file.file_copy_variant(path(data), path(r"\\Server\output\copy.txt"))
        .unwrap();
    assert_eq!(fs::read(output.join("copy.txt")).unwrap(), b"test data");
    assert_eq!(fs::read_dir(&input).unwrap().count(), 1);
    assert!(input.join("Data.txt").exists());
    fs::remove_dir_all(&root).unwrap();
}

#[test]
fn faithful_filesystem_with_no_mounts_cannot_read_host_paths() {
    let root =
        std::env::temp_dir().join(format!("calcvba-hal-empty-mounts-{}", std::process::id()));
    fs::create_dir(&root).unwrap();
    let input = root.join("input.txt");
    fs::write(&input, b"outside the Windows namespace").unwrap();
    let mut policy = HostPolicy::native();
    policy.windows_faithful = true;
    let host = HostBuilder::new()
        .profile(HalProfileId::Linux)
        .policy(policy)
        .build();
    let error = host
        .fs()
        .file_len_variant(Variant::from_string(input.to_str().unwrap()))
        .unwrap_err();
    assert_eq!(error.host_error_code, Some(76));
    fs::remove_dir_all(root).unwrap();
}
