# Saka Note

Saka Note is a lightweight Windows text editor built with C++17 and Win32 API.  
Saka Note adalah editor teks Windows ringan yang dibangun dengan C++17 dan Win32 API.

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

## Requirements / Kebutuhan

- Windows 7 SP1+ (Windows 10/11 recommended). / Windows 7 SP1+ (Windows 10/11 disarankan).
- CMake 3.16+.
- C++17 toolchain:
  - MinGW-w64, or / atau
  - MSVC (Visual Studio 2022+).

## Build

### MinGW (Debug)

```powershell
cmake -S . -B build/mingw-debug -G "MinGW Makefiles" -DCMAKE_BUILD_TYPE=Debug
cmake --build build/mingw-debug
.\build\mingw-debug\saka-note.exe
```

### MSVC (Debug)

```powershell
cmake -S . -B build/msvc-debug -G "Visual Studio 17 2022" -A x64
cmake --build build/msvc-debug --config Debug
.\build\msvc-debug\Debug\saka-note.exe
```

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
   - `git tag v1.3.0`
   - `git push origin v1.3.0`
2. Wait for `Build and Release` GitHub Actions workflow to finish.
3. Buka halaman `Releases` dan unduh salah satu file:
   - `saka-note-setup-x64.exe` (recommended for most users; creates Start Menu entry)
   - `saka-note-portable-x64.zip` (portable; no install required)
   - `saka-note-setup-ARM64.exe` / `saka-note-portable-ARM64.zip` (ARM devices)
4. For portable zip:
   - extract zip,
   - run `saka-note.exe`.

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
.\build\mingw-debug\saka-note.exe --benchmark-ci
```

Report output / Lokasi output laporan:
- `%LOCALAPPDATA%\SakaNote\benchmarks\benchmark-YYYYMMDD-HHMMSS.txt`

A/B benchmark (Saka Note vs Microsoft Notepad):

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
research/
docs/
```

## License / Lisensi

MIT License. See `LICENSE`.  
Lisensi MIT. Lihat `LICENSE`.
