# Saka Note Roadmap

Roadmap ini disusun dari audit codebase saat ini (Win32 C++), fokus pada stabilitas, performa, dan UX modern tanpa mengorbankan lightweight footprint.

## Current Baseline (v1.3.x)

- Core editing: tabs, session restore, encoding/line-ending, find/replace, zoom, status bar.
- Release pipeline: setup installer + portable zip (x64/ARM64) sudah tersedia.
- Benchmarking: internal benchmark (`--benchmark-ci`) + A/B benchmark script (`tools/compare-notepad.ps1`).

## Why This Roadmap

Audit highlights:

- `src/main.cpp` masih monolitik (~2698 lines) dan jadi bottleneck maintainability.
- Beberapa modul besar (`src/modules/commands.cpp` ~1274 lines) menampung terlalu banyak concern.
- Belum ada automated unit/regression test untuk parser/encoding/session logic.
- Localization app baru EN/JA; Indonesian belum tersedia di runtime UI.

## P1 - Stability and Architecture (Now)

Goal / Tujuan: turunkan risiko bug regresi, rapikan arsitektur internal.

- Split `src/main.cpp` into focused units:
  - tab/session controller
  - window/layout controller
  - command routing
- Break down `src/modules/commands.cpp`:
  - update-check module
  - benchmark module
  - icon/menu actions module
- Add small test harness for pure logic:
  - encoding detection/conversion
  - version compare/update parser
  - session serialization/deserialization
- Add crash diagnostics minimal:
  - startup-safe logging path
  - optional unhandled-exception dump (opt-in)

## P2 - UX and Product Fit (Next)

Goal / Tujuan: modern UX polish tetap ringan.

- Tab UX refinement:
  - tighter hover/active transitions
  - keyboard tab management parity
  - better overflow behavior
- Add Indonesian UI language option in-app (`LangID::ID`).
- Settings usability:
  - reset-to-default
  - import/export settings profile
- Improve discoverability for non-technical users:
  - first-run tips
  - clearer menu wording for setup/portable behaviors

## P3 - Performance and Large Files (Next+)

Goal / Tujuan: scale lebih baik untuk file besar dan long sessions.

- Large-file mode improvements:
  - incremental loading strategy
  - memory cap heuristics per scenario
- Session persistence tuning:
  - autosave batching
  - contention-safe writes
- Expand benchmark scenarios:
  - repeated open/close loops
  - long-line rendering stress
  - multi-tab stress profiles

## P4 - Distribution and Trust (Ongoing)

Goal / Tujuan: rilis makin siap pakai dan kredibel.

- Signed binaries (code signing) for setup and portable executables.
- Release checklist in CI:
  - artifact naming policy
  - smoke install check
  - hash verification output
- Public performance dashboard from benchmark outputs.

## Proposed Backlog Items

- Autosave frequency setting (user configurable).
- Better accessibility pass (high contrast, keyboard-first checks).
- Safer update channel metadata (release channel separation stable/prerelease).
- Optional plugin/extension boundary (strictly minimal and opt-in).

## Done Definition (per phase)

- P1 done when:
  - `main.cpp` and `commands.cpp` responsibilities are split
  - core logic tests run in CI
  - no regression on build/release pipeline
- P2 done when:
  - in-app language includes Indonesian
  - tab UX polish shipped + documented shortcuts
- P3 done when:
  - large-file scenarios show measurable memory/latency improvement
- P4 done when:
  - release artifacts are signed and validation checklist is automated
