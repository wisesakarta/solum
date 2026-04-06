# Technical Standard Note

<p align="left">
  <img src="assets/branding/technical-standard-note-icon.png" alt="Technical Standard Note icon" width="96">
</p>

Technical Standard Note is a lightweight Windows text editor built with C++17 and Win32 API.  
Technical Standard Note adalah editor teks Windows ringan yang dibangun dengan C++17 dan Win32 API.

Project goal: keep plain-text editing fast, stable, and clean without feature bloat.  
Tujuan proyek: menjaga pengeditan plain text tetap cepat, stabil, dan bersih tanpa feature bloat.

## Language Policy / Kebijakan Bahasa

- Primary documentation language: English.
- Bahasa dokumentasi utama: English.
- Companion language: Indonesian for collaboration clarity.
- Bahasa pendamping: Indonesia untuk kejelasan kolaborasi.
- Standard reference: `docs/language-standard.md`.
- Referensi standar: `docs/language-standard.md`.

## Open Source Attribution / Atribusi Open Source

This project is based on the upstream repository below.  
Proyek ini berbasis pada repositori upstream berikut.

- Upstream repository: https://github.com/ForLoopCodes/legacy-notepad

Attribution details and modification log:  
Detail atribusi dan log modifikasi:
- `docs/upstream-and-modifications.md`

## Core Features / Fitur Inti

- Multi-tab editing with startup behavior options. / Pengeditan multi-tab dengan opsi perilaku startup.
- Encoding support: UTF-8, UTF-8 BOM, UTF-16 LE/BE, ANSI. / Dukungan encoding: UTF-8, UTF-8 BOM, UTF-16 LE/BE, ANSI.
- Line ending support: CRLF, LF, CR. / Dukungan line ending: CRLF, LF, CR.
- Find, replace, go-to-line, zoom, word-wrap, font settings. / Find, replace, go-to-line, zoom, word-wrap, dan pengaturan font.
- Optional background image and opacity controls. / Opsi gambar latar belakang dan kontrol opacity.
- Built-in performance benchmark (`--benchmark-ci`). / Benchmark performa bawaan (`--benchmark-ci`).
- Large file mode (automatic behavior for big files). / Mode file besar (perilaku otomatis untuk file besar).

## Roadmap

- See detailed roadmap: `docs/roadmap.md`.
- Lihat roadmap detail: `docs/roadmap.md`.

## Testing and Research Docs

- Claim test protocol / Protokol uji klaim: `docs/testing/claim-test-protocol.md`.
- Tab visual parity checklist / Checklist parity visual tab: `docs/testing/tab-visual-parity-checklist.md`.
- Librarian research notes / Catatan riset: `docs/research/librarian-research-notepad-2026-02.md`.

## Requirements / Kebutuhan

- Windows 7 SP1+ (Windows 10/11 recommended). / Windows 7 SP1+ (Windows 10/11 disarankan).
- CMake 3.16+.
- C++17 toolchain:
  - MinGW-w64, or / atau
  - MSVC (Visual Studio 2022+).

## Build

### Quick Start (VS Code)

1. `Ctrl+Shift+P` -> `CMake: Select Configure Preset` -> pilih `MinGW Debug`.
2. `Ctrl+Shift+P` -> `CMake: Configure`.
3. `Ctrl+Shift+P` -> `CMake: Build`.

If compiler path becomes stale, run:
- `Ctrl+Shift+P` -> `CMake: Delete Cache and Reconfigure`.

### Quick Start (CLI, 3 commands)

```powershell
cmake --preset mingw-debug --fresh
cmake --build --preset mingw-debug
ctest --preset mingw-debug
```

### Alternative (Auto-detect generator/toolchain)

Use **Developer PowerShell for Visual Studio** (2026/2022) or shell with MinGW toolchain in `PATH`.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\dev-build.ps1 -Config Debug -RunTests
```

The script tries generators in this order:
- `Visual Studio 18 2026`
- `Visual Studio 17 2022`
- `NMake Makefiles`
- `Ninja Multi-Config`
- `Ninja`
- `MinGW Makefiles`

If you see `could not find any instance of Visual Studio`, use:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\dev-build.ps1 -ForceNMake -Config Debug -RunTests
```

### Manual MSVC (Visual Studio 2026)

```powershell
cmake -S . -B build/vs2026-debug --fresh -G "Visual Studio 18 2026" -A x64
cmake --build build/vs2026-debug --config Debug
.\build\vs2026-debug\Debug\technical-standard-note.exe
```

### Manual MSVC (Build Tools / Developer PowerShell fallback via NMake)

```powershell
cmake -S . -B build/nmake-debug --fresh -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Debug
cmake --build build/nmake-debug
.\build\nmake-debug\technical-standard-note.exe
```

## Test

### Unit/Logic tests via script (recommended)

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\dev-build.ps1 -Config Debug -RunTests
```

### Unit/Logic tests via CTest (Visual Studio 2026)

```powershell
cmake -S . -B build/vs2026-debug --fresh -G "Visual Studio 18 2026" -A x64
cmake --build build/vs2026-debug --config Debug
ctest --test-dir build/vs2026-debug -C Debug --output-on-failure
```

### Unit/Logic tests via CTest (NMake)

```powershell
cmake -S . -B build/nmake-debug --fresh -G "NMake Makefiles" -DCMAKE_BUILD_TYPE=Debug
cmake --build build/nmake-debug
ctest --test-dir build/nmake-debug --output-on-failure
```

### Manual smoke test checklist

1. Flicker test:
   - enable background image,
   - type continuously for 20-30 seconds,
   - verify editor and tab strip do not blink aggressively.
2. Auto-list test:
   - type `1. item` then press Enter -> next line should start `2. `,
   - type `- item` then press Enter -> next line should start `- `,
   - on an empty list line (`2. ` or `- `), press Enter -> list marker should be removed (exit list mode).
3. Standard editing regression:
   - undo/redo, copy/paste, Ctrl+Backspace, Ctrl+Delete, save/open.

## Distribution / Distribusi

Recommended distribution model for public users:  
Model distribusi yang disarankan untuk pengguna publik:

1. Do not commit compiled `.exe` directly to `main`.
2. Jangan commit file `.exe` hasil build langsung ke branch `main`.
3. Publish compiled binaries via GitHub Releases (x64/ARM64).
4. Publikasikan binary hasil build melalui GitHub Releases (x64/ARM64).
5. Keep source repo clean (code + docs), and let CI publish artifacts on tags.
6. Jaga repo source tetap bersih (code + docs), dan biarkan CI mempublikasikan artifact saat tag dirilis.

Current CI workflow: `.github/workflows/build.yml`.  
Workflow CI saat ini: `.github/workflows/build.yml`.

## Weekly Test Build / Build Uji Mingguan

After your repo is connected to GitHub:  
Setelah repo kamu terhubung ke GitHub:

1. Create and push a tag:
   - `git tag v1.3.9`
   - `git push origin v1.3.9`
2. Wait for `Build and Release` GitHub Actions workflow to finish.
3. Buka halaman `Releases` dan unduh salah satu file:
   - `technical-standard-note-setup-x64.exe` (recommended for most users; creates Start Menu entry)
   - `technical-standard-note-portable-x64.zip` (portable; no install required)
   - `technical-standard-note-setup-ARM64.exe` / `technical-standard-note-portable-ARM64.zip` (ARM devices)
4. For portable zip:
   - extract zip,
   - run `technical-standard-note.exe`.

For a one-week trial, keep a simple checklist:  
Untuk uji coba satu minggu, simpan checklist sederhana:
- startup speed / kecepatan startup
- memory usage / penggunaan memori
- save/open stability / stabilitas save/open
- tab/session behavior / perilaku tab/sesi
- crash/freeze reproduction steps / langkah reproduksi crash/freeze

## Benchmark

Internal benchmark / Benchmark internal:

```powershell
.\build\mingw-debug\technical-standard-note.exe --benchmark-ci
```

Report output / Lokasi output laporan:
- `%LOCALAPPDATA%\TechnicalStandardNote\benchmarks\benchmark-YYYYMMDD-HHMMSS.txt`

A/B benchmark (Technical Standard Note vs Microsoft Notepad):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\compare-notepad.ps1 -Iterations 3 -SampleSeconds 8 -SampleIntervalMs 500 -WarmupMs 3000 -ForceCloseBeforeEachRun
```

Generated outputs / Output yang dihasilkan:
- `research/perf-runs/run-YYYYMMDD-HHMMSS/summary.csv`
- `research/perf-runs/run-YYYYMMDD-HHMMSS/samples.csv`
- `research/perf-runs/run-YYYYMMDD-HHMMSS/comparison.md`

Cleanup old benchmark runs / Bersihkan run benchmark lama:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\cleanup-perf-runs.ps1 -KeepLatest 2
```

## Project Structure / Struktur Proyek

```text
src/
  core/
  lang/
  modules/
  main.cpp
  notepad.rc
tools/
docs/
research/ (generated benchmark outputs; gitignored)

```

## License / Lisensi

MIT License. See `LICENSE`.  
Lisensi MIT. Lihat `LICENSE`.


