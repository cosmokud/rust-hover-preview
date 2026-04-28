# Graph Report - G:\Clouds\Github\rust-hover-preview  (2026-04-28)

## Corpus Check
- Corpus is ~16,627 words - fits in a single context window. You may not need a graph.

## Summary
- 193 nodes · 407 edges · 7 communities detected
- Extraction: 86% EXTRACTED · 14% INFERRED · 0% AMBIGUOUS · INFERRED: 55 edges (avg confidence: 0.8)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Preview Rendering Pipeline|Preview Rendering Pipeline]]
- [[_COMMUNITY_Explorer State Resolution|Explorer State Resolution]]
- [[_COMMUNITY_Runtime Config Control|Runtime Config Control]]
- [[_COMMUNITY_Explorer Cache Flow|Explorer Cache Flow]]
- [[_COMMUNITY_Tray Settings Control|Tray Settings Control]]
- [[_COMMUNITY_Preview Loop Core|Preview Loop Core]]
- [[_COMMUNITY_Startup Registry Toggle|Startup Registry Toggle]]

## God Nodes (most connected - your core abstractions)
1. `run_preview_window()` - 23 edges
2. `run_explorer_hook()` - 18 edges
3. `run_explorer_hook` - 16 edges
4. `load_media()` - 13 edges
5. `is_media_file()` - 12 edges
6. `run_preview_window` - 12 edges
7. `MediaData` - 11 edges
8. `tray_window_proc()` - 11 edges
9. `get_item_under_cursor()` - 10 edges
10. `try_get_item_from_parent()` - 8 edges

## Surprising Connections (you probably didn't know these)
- `QTTabBar Hover Preview Inspiration` --conceptually_related_to--> `run_explorer_hook`  [AMBIGUOUS]
  README.md → src/explorer_hook.rs
- `Layered Topmost GDI Preview Window` --references--> `run_preview_window`  [EXTRACTED]
  README.md → src/preview_window.rs
- `run_explorer_hook()` --calls--> `hide_preview()`  [INFERRED]
  src\explorer_hook.rs → src\preview_window.rs
- `run_explorer_hook()` --calls--> `show_preview_keyboard()`  [INFERRED]
  src\explorer_hook.rs → src\preview_window.rs
- `run_explorer_hook()` --calls--> `show_preview()`  [INFERRED]
  src\explorer_hook.rs → src\preview_window.rs

## Hyperedges (group relationships)
- **Runtime Thread Orchestration** — main_main, preview_window_run_preview_window, explorer_hook_run_explorer_hook, tray_run_tray, main_running_state [INFERRED 0.90]
- **Explorer Hover Resolution Flow** — explorer_hook_get_item_under_cursor, explorer_hook_get_current_explorer_folder, explorer_hook_find_media_in_folder, explorer_hook_get_file_under_cursor [EXTRACTED 1.00]
- **Async Media Loading Pipeline** — preview_window_run_preview_window, preview_window_queue_load_request, preview_window_spawn_load_worker, preview_window_load_media, preview_window_streaming_frame_queue [INFERRED 0.87]

## Communities

### Community 0 - "Preview Rendering Pipeline"
Cohesion: 0.08
Nodes (45): composite_gif_frame(), compute_keyboard_layout(), compute_mouse_layout(), decode_gif_frame_to_image(), decode_image_with_header_check(), decode_webp_frame_to_image(), detect_video_crop(), ensure_noactivate_monitor() (+37 more)

### Community 1 - "Explorer State Resolution"
Cohesion: 0.13
Nodes (39): AccessibilityResult, build_folder_media_index(), clear_variant(), ExplorerFoldersCache, ExplorerState, find_media_in_folder(), FocusedItemInfo, FolderMediaIndex (+31 more)

### Community 2 - "Runtime Config Control"
Cohesion: 0.1
Nodes (28): Windows Embedded Resource Metadata, Confirm File Type Option, 300-Frame Limit and 30 FPS Ceiling, Queue-Based GIF/WebP Frame Transfer, AppConfig, AppConfig::config_path, AppConfig::load, AppConfig::save (+20 more)

### Community 3 - "Explorer Cache Flow"
Cohesion: 0.1
Nodes (26): Explorer Folder Cache Optimization, Post-Folder-Navigation Input Gate, ExplorerFoldersCache, find_media_in_folder, FolderMediaIndex Cache, get_all_explorer_folders, get_current_explorer_folder, get_explorer_state (+18 more)

### Community 4 - "Tray Settings Control"
Cohesion: 0.18
Nodes (16): AppConfig, disable_startup(), enable_startup(), is_startup_enabled(), add_tray_icon(), open_config_file(), remove_tray_icon(), run_tray() (+8 more)

### Community 5 - "Preview Loop Core"
Cohesion: 0.15
Nodes (11): main(), clear_load_request(), create_loading_media(), is_video_process_running(), MediaData, overlay_loading_spinner(), queue_load_request(), render_loading_frame() (+3 more)

### Community 6 - "Startup Registry Toggle"
Cohesion: 0.5
Nodes (5): HKCU Run Registry Startup Control, disable_startup, enable_startup, HKCU\Software\Microsoft\Windows\CurrentVersion\Run, toggle_startup

## Ambiguous Edges - Review These
- `QTTabBar Hover Preview Inspiration` → `run_explorer_hook`  [AMBIGUOUS]
  README.md · relation: conceptually_related_to

## Knowledge Gaps
- **28 isolated node(s):** `FolderMediaIndex`, `ExplorerFoldersCache`, `AccessibilityResult`, `ExplorerState`, `FocusedItemInfo` (+23 more)
  These have ≤1 connection - possible missing edges or undocumented components.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **What is the exact relationship between `QTTabBar Hover Preview Inspiration` and `run_explorer_hook`?**
  _Edge tagged AMBIGUOUS (relation: conceptually_related_to) - confidence is low._
- **Why does `run_explorer_hook()` connect `Explorer State Resolution` to `Preview Rendering Pipeline`, `Preview Loop Core`?**
  _High betweenness centrality (0.109) - this node is a cross-community bridge._
- **Why does `run_preview_window()` connect `Preview Loop Core` to `Preview Rendering Pipeline`, `Explorer State Resolution`?**
  _High betweenness centrality (0.079) - this node is a cross-community bridge._
- **Why does `run_explorer_hook` connect `Explorer Cache Flow` to `Runtime Config Control`?**
  _High betweenness centrality (0.050) - this node is a cross-community bridge._
- **Are the 3 inferred relationships involving `run_preview_window()` (e.g. with `main()` and `.default()`) actually correct?**
  _`run_preview_window()` has 3 INFERRED edges - model-reasoned connections that need verification._
- **Are the 8 inferred relationships involving `run_explorer_hook()` (e.g. with `.default()` and `.load()`) actually correct?**
  _`run_explorer_hook()` has 8 INFERRED edges - model-reasoned connections that need verification._
- **What connects `FolderMediaIndex`, `ExplorerFoldersCache`, `AccessibilityResult` to the rest of the system?**
  _28 weakly-connected nodes found - possible documentation gaps or missing edges._