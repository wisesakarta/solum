# Otso

<p align="center">
  <img src="assets/icons/otso_icon.png" alt="Otso" width="128">
</p>

**Otso** is a high-fidelity, Win32 text editor focused on speed, stability, and the "Renaissance of Software" craftsmanship.

**Otso** adalah editor teks Win32 beresolusi tinggi yang berfokus pada kecepatan, stabilitas, dan keahlian rekayasa perangkat lunak "Renaissance".

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
.\build\mingw-debug\Otso.exe --benchmark-ci
```

Benchmark reports are written to:

- `%LOCALAPPDATA%\Otso\benchmarks\benchmark-YYYYMMDD-HHMMSS.txt`

## Release Channels

Repository CI workflow: `.github/workflows/build.yml`

- Desktop tags: `desktop/v*`
- Mobile tags: `mobile/v*`

Release assets are intentionally minimal:

- Desktop: `Otso-desktop-... .exe`
- Mobile: `Otso-mobile-... .apk`

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
- Debt/audit data: `docs/audits/`

## Upstream Attribution

This project is based on:

- https://github.com/ForLoopCodes/legacy-notepad

Modification and attribution details:

- `docs/upstream-and-modifications.md`

## License

MIT License. See `LICENSE`.
