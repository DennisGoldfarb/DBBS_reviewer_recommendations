mod commands;
mod config;
mod dataset;
mod directory_matching;
mod documents;
mod embeddings;
mod exclusions;
mod exports;
mod matching;
mod models;
mod paths;
mod scoring;
mod spreadsheet_matching;
mod spreadsheets;

use crate::embeddings::{shutdown_embedding_helper, warmup_embedding_helper};

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_shell::init())
        .on_window_event(|window, event| {
            #[cfg(debug_assertions)]
            if let tauri::WindowEvent::DragDrop(tauri::DragDropEvent::Drop { paths, position }) =
                event
            {
                eprintln!(
                    "Native upload drop: {} item(s), position=({}, {}), display_scale={}",
                    paths.len(),
                    position.x,
                    position.y,
                    window.scale_factor().unwrap_or(1.0)
                );
            }
            #[cfg(not(debug_assertions))]
            let _ = (window, event);
        })
        .setup(|app| {
            let handle = app.handle().clone();
            tauri::async_runtime::spawn(async move {
                if let Err(err) = warmup_embedding_helper(handle).await {
                    eprintln!("Failed to warm up embedding helper: {err}");
                }
            });
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            commands::validate_upload_path,
            commands::submit_matching_request,
            commands::update_faculty_embeddings,
            commands::ensure_embedding_helper_ready,
            commands::analyze_spreadsheet,
            commands::get_faculty_dataset_status,
            commands::preview_faculty_roster,
            commands::preview_faculty_dataset_replacement,
            commands::replace_faculty_dataset,
            commands::restore_default_faculty_dataset,
            commands::save_generated_spreadsheet
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");

    shutdown_embedding_helper();
}

#[cfg(test)]
mod test_support;
