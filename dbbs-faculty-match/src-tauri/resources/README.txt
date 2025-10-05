This directory is packaged with the Tauri bundle. The Windows build stages the CPython
embeddable distribution here (python.exe, python311.dll, Lib/, DLLs/, site-packages/, and
app/). Non-Windows builds continue to create platform-specific virtual environments under
`python/<platform>-<arch>`. A non-hidden placeholder file is kept so the bundler's glob
patterns always match even before the runtime is generated.
