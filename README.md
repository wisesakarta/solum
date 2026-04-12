# Technical Standard Note

<p align="left">
  <img src="src/logo_wordmark.png" alt="Technical Standard Note wordmark" width="420">
</p>
<p align="left">
  <img src="assets/branding/technical-standard-note-icon.png" alt="Technical Standard Note icon" width="96">
</p>

Technical Standard Note is a lightweight Win32 text editor focused on speed, stability, and clarity.

Technical Standard Note adalah editor teks Win32 yang ringan dengan fokus pada kecepatan, stabilitas, dan kejelasan.

## Scope

- Primary product in this repository: **Desktop (Win32 C++17)**.
- Mobile release artifacts are distributed via GitHub Releases, while source ownership can be managed separately.
- Public release channel policy:
  - Desktop release assets: `.exe`
  - Mobile release assets: `.apk`
  - Source of truth remains in repository source code, not release bundles.

## Core Features (Desktop)

- Multi-tab editing with startup resume behavior.
- Encoding support: UTF-8, UTF-8 BOM, UTF-16 LE/BE, ANSI.
- Line ending support: CRLF, LF, CR.
- Find/Replace/Go To, word wrap, zoom, font settings.
- Performance benchmark mode: `--benchmark-ci`.
- Large file mode with plain-text-first behavior.

## Build (Desktop)

Quick start via presets:

```powershell
cmake --preset mingw-debug --fresh
cmake --build --preset mingw-debug
ctest --preset mingw-debug
```

Alternative helper script:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\dev-build.ps1 -Config Debug -RunTests
```

## Runtime Benchmark

```powershell
.\build\mingw-debug\technical-standard-note.exe --benchmark-ci
```

Benchmark reports are written to:

- `%LOCALAPPDATA%\TechnicalStandardNote\benchmarks\benchmark-YYYYMMDD-HHMMSS.txt`

## Release Channels

Repository CI workflow: `.github/workflows/build.yml`

- Desktop tags: `desktop/v*`
- Mobile tags: `mobile/v*`

Release assets are intentionally minimal:

- Desktop: `technical-standard-note-desktop-... .exe`
- Mobile: `technical-standard-note-mobile-... .apk`

## Repository Hygiene Rules

- Do not commit runtime logs (`hs_err_pid*`, `replay_pid*`) and local temp audit folders.
- Do not commit compiled release bundles to `main`.
- Keep docs under `docs/`, source under `src/`, tests under `tests/`, scripts under `tools/`.
- All user-facing UI strings must go through `src/lang/*` localization tables.

## Documentation Index

- Roadmap: `docs/roadmap.md`
- Upstream attribution: `docs/upstream-and-modifications.md`
- Language policy: `docs/language-standard.md`
- Test protocol: `docs/testing/claim-test-protocol.md`
- Desktop release notes: `docs/releases/`
- Mobile/internal learning notes: `docs/mobile/`
- Debt/audit data: `docs/audits/`

## Upstream Attribution

This project is based on:

- https://github.com/ForLoopCodes/legacy-notepad

Modification and attribution details:

- `docs/upstream-and-modifications.md`

## License

MIT License. See `LICENSE`.
