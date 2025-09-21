fn main() {
    // When `cargo` is invoked with `--manifest-path src-tauri/Cargo.toml` the
    // build script inherits the caller's working directory (the app root).
    // Tauri resolves `bundle.resources` globs relative to the current
    // directory, so ensure we're inside the crate root before delegating to
    // `tauri_build`.
    if let Ok(manifest_dir) = std::env::var("CARGO_MANIFEST_DIR") {
        let _ = std::env::set_current_dir(manifest_dir);
    }

    tauri_build::build()
}
