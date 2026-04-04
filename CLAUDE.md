# CLAUDE.md - wuli 图片处理工具

## Project Overview
Python GUI app (PySide6) for image processing with ComfyUI integration, template compositing, and OSS upload. Distributed as PyInstaller exe with OTA auto-update via GitHub Releases.

## Key Files
- `gui_app.py` — Main GUI app, version defined as `APP_VERSION` at top
- `updater.py` — OTA update logic, downloads zip from GitHub Release
- `oss_uploader.py` — Alibaba Cloud OSS upload module
- `comfyui_client.py` — ComfyUI API client
- `image_processor.py` — Image processing utilities
- `build.bat` — PyInstaller build script
- `config.ini` — Runtime config (OSS credentials, paths, etc.)

## Build & Release Process

### Build
```
build.bat
```
Runtime files (config.ini, fonts/, styles/, templates/, workflow_i2i.json) are copied INTO `dist/图片处理工具/`, NOT to `dist/` root.

### Create Release Zip
```
cd dist && powershell Compress-Archive -Path '图片处理工具' -DestinationPath '..\wuli_vX.X.X.zip'
```

**CRITICAL: The zip MUST have exactly ONE top-level folder.** The OTA updater (`updater.py`) checks if the extracted zip has a single subdirectory — if it does, it enters that directory as the source. If there are multiple top-level entries, the updater copies the wrong structure and the update fails silently.

### Publish Release
```
gh release create vX.X.X wuli_vX.X.X.zip --title "vX.X.X" --notes "changelog here"
```

## Version Bumping
- Version is in `gui_app.py` line ~8: `APP_VERSION = "x.y.z"`
- Use patch version increments (e.g. 1.2.5 -> 1.2.6)
- Commit message format: `Release vX.X.X - short description`

## Common Pitfalls
- Output files are unified to JPG — watch for filename collisions when source folder has same-name files with different extensions (e.g. `1.jpg` and `1.png` both become `1.jpg`)
- OSS upload path is `{prefix}/{folder}/{filename}` — duplicate filenames in same folder will overwrite
- `config.ini` contains OSS credentials — never commit user-specific secrets to git
