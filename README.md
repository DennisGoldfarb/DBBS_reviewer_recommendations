# DDBS Reviewer Recommendations

Desktop tooling for configuring student-to-faculty reviewer recommendations. The
initial milestone delivers a Tauri-based interface that captures matching
settings, validates inputs, and confirms which options would be submitted to the
matching backend.

## Repository layout

- `tauri-gui/` – Tauri + React application that hosts the configuration UI.
- `tauri-gui/src-tauri/` – Rust commands that validate configuration requests.

## Getting started

```bash
cd tauri-gui
npm install
npm run tauri dev
```

Use the form to select the student data source, constrain the faculty pool, and
adjust recommendation limits. Submitting the form runs server-side validation in
Rust; the UI displays the confirmation payload without executing any matching
logic.
