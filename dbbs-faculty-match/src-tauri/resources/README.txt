This directory is packaged with the Tauri bundle. The `prepare-python` script downloads
the Windows CPython embeddable runtime into `python/windows-<arch>/`, installs
dependencies into its `site-packages/` folder, and copies application sources into
`python/windows-<arch>/app/`. A non-hidden placeholder file is kept here so the bundler's
glob patterns still match before the embedded Python runtime is generated.
