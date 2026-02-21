# Librarian Research - Real-world Notepad Issues and Saka Note Direction
# Librarian Research - Problem Nyata Notepad Modern dan Arah Saka Note

Research date / Tanggal riset: 20 Februari 2026
Focus / Fokus: public user complaints, official Microsoft signals, and implementation decisions for Saka Note while staying lightweight.

## Language Note / Catatan Bahasa
- This research file is bilingual by intent: primary analysis in Indonesian with key technical terms in English.
- File riset ini sengaja bilingual: analisa utama dalam Bahasa Indonesia dengan istilah teknis kunci dalam English.
- Reference standard: `docs/language-standard.md`.
- Referensi standar: `docs/language-standard.md`.

## 1) Metode
- Kumpulkan sumber resmi Microsoft (Windows Insider Blog, NVD).
- Kumpulkan feedback user publik (Reddit) untuk pain point UX/performa harian.
- Sinkronkan dengan guideline UI/UX internal di `.agent` (konsistensi, kontras, anti visual noise, performa).

## 2) Temuan Utama (Evidence-based)

### A. Session restore + tab behavior menimbulkan friksi
- Sumber resmi Microsoft menyebut Notepad otomatis restore tab dan unsaved content, bisa dimatikan di Settings.
- Evidence:
  - 31 Agustus 2023: Notepad mulai auto-save session + restore tab/unsaved content.
    - https://blogs.windows.com/windows-insider/2023/08/31/new-updates-for-snipping-tool-and-notepad-for-windows-insiders/
- Feedback user menunjukkan sebagian user tidak suka app dibuka dengan state lama terus-menerus.
  - https://www.reddit.com/r/windows/comments/18zcuku

Implikasi untuk Saka Note:
- Session policy harus jelas, sederhana, dan user-control.
- Jangan "memaksa" restore behavior tanpa opsi yang mudah ditemukan.

### B. Respons scroll/typing dianggap lebih berat dari notepad klasik
- Banyak complaint user pada scroll smoothing/lag dan input lag, terutama di file lebih besar.
- Evidence:
  - https://www.reddit.com/r/Windows11/comments/1b3w1bm
  - https://www.reddit.com/r/Windows11/comments/su4ia1
  - https://www.reddit.com/r/techsupport/comments/1eqolea

Implikasi untuk Saka Note:
- Latency harus jadi KPI utama (typing, scroll, open file).
- Hindari animasi/smoothing non-esensial.

### C. Affordance tab/window area tidak intuitif di edge-case
- Ada laporan area klik tab tidak bekerja di top-edge tertentu (friksi kecil tapi nyata).
- Evidence:
  - https://www.reddit.com/r/techsupport/comments/1h17dgy

Implikasi untuk Saka Note:
- Hit target tab harus konsisten dan predictable.
- Area interaksi tidak boleh ambigu antara title bar vs tab.

### D. Feature creep meningkatkan risiko (security + complexity)
- Microsoft memperluas Notepad ke formatting Markdown, tabel, AI write/rewrite/summarize.
- Evidence resmi:
  - 30 Mei 2025 (lightweight formatting/Markdown):
    - https://blogs.windows.com/windows-insider/2025/05/30/text-formatting-in-notepad-begin-rolling-out-to-windows-insiders/
  - 21 November 2025 (tables + streaming AI):
    - https://blogs.windows.com/windows-insider/2025/11/21/notepad-update-begins-rolling-out-to-windows-insiders/
  - 21 Januari 2026 (tambahan Markdown features + welcome screen):
    - https://blogs.windows.com/windows-insider/2026/01/21/notepad-and-paint-updates-begin-rolling-out-to-windows-insiders/
- Risiko keamanan terkait Markdown link handling:
  - CVE-2026-20841 (NVD, published 10 Februari 2026):
    - https://nvd.nist.gov/vuln/detail/CVE-2026-20841

Implikasi untuk Saka Note:
- Setiap fitur baru wajib lolos "cost-to-risk ratio".
- Core editor harus tetap plain-text first. Fitur non-core harus optional/isolated.

## 3) Sinkronisasi dengan Guideline UI/UX Internal (.agent)
- `.agent/skills/rams/designEngineer/SKILL.md` menekankan:
  - dark mode consistency
  - contrast
  - border/shadow consistency
- `.agent/skills/uiSkills/SKILL.md` menekankan minim visual noise dan kontrol performa.
- `.agent/skills/vercel/interfaceGuidelines/SKILL.md` menekankan:
  - clarity of interaction
  - crisp borders
  - depth secukupnya (bukan dekoratif berlebihan)

Kesimpulan UI: gaya visual harus bersih, kontras jelas, shadow minimal, state interaksi jelas, tanpa ornamen berat.

## 4) Rekomendasi Prioritas Implementasi (untuk Saka Note)

### P0 (harus duluan)
1. Interaction latency budget
- Target: typing latency nyaris instan pada file kecil-menengah.
- Instrumentasi: ukur waktu open file, waktu first paint, dan scroll jitter.

2. Session behavior matrix yang eksplisit
- Mode A: Always fresh start.
- Mode B: Restore last session.
- Mode C: Restore only saved files.
- UI setting harus langsung jelas (1-2 klik, bukan tersembunyi).

3. Tab hit-target correctness
- Audit seluruh area klik tab (termasuk top-edge pixel boundary).
- Pastikan close button/tab drag tidak konflik.

### P1 (setelah stabil)
4. Large-file mode (graceful degradation)
- Untuk file besar, nonaktifkan fitur mahal (mis. live syntax-like extras, heavy repaint path).
- Tampilkan indikator "Large file mode".

5. UX consistency pass untuk dialog
- Spacing, button placement, radius, border, shadow pakai token yang konsisten.
- DPI-aware di semua dialog custom.

### P2 (opsional, hanya jika cost rendah)
6. Minimal formatting sandbox
- Jika suatu saat butuh markdown-lite, harus optional, non-default, dan parser hardened.
- Security review wajib sebelum default-on.

## 5) Prinsip Produk (Hard Guardrails)
1. Plain-text first, everything else optional.
2. No feature shipped tanpa metrik performa sebelum/sesudah.
3. No visual effect yang menambah latency tapi tidak menambah kejelasan.
4. No security-risky behavior dari link/protocol handling.

## 6) Backlog Candidate yang langsung actionable
- Tambah benchmark internal sederhana:
  - Open 1MB, 5MB, 20MB file.
  - Scroll 10k lines.
  - Typing burst 60s.
- Tambah setting "Startup Behavior" yang eksplisit di menu settings.
- Tambah regression checklist tab hit-testing + DPI.
- Tambah visual consistency checklist untuk seluruh dialog custom.

## 7) Catatan Riset
- Feedback publik cenderung konsisten pada 3 hal: responsiveness, predictability, dan anti-bloat.
- Ini sangat sejalan dengan positioning Saka Note sebagai lightweight replacement.

## 8) Update Riset Lanjutan (20 Februari 2026)

### A. Friksi taskbar + tabs (discoverability)
- Ada feedback bahwa user ingin model klasik: setiap file muncul terpisah di taskbar (tanpa tab dalam satu window).
- Evidence:
  - https://www.reddit.com/r/Windows11/comments/182dbf6/

Implikasi:
- Pertimbangkan mode opsional "single-file windows" untuk user yang memang anti-tab workflow.

### B. Konfirmasi publik untuk "lag terasa berat"
- Feedback tambahan menguatkan narasi lag/scroll tidak se-responsif notepad lama.
- Evidence:
  - https://www.reddit.com/r/windows/comments/1949c5r/
  - https://www.reddit.com/r/windows/comments/1jm4h0j/

Implikasi:
- Benchmark harus jadi bagian release gate, bukan sekadar testing ad-hoc.

### C. Validasi resmi atas session-restore behavior
- Microsoft menegaskan notepad tabs + auto-save session default, dengan opsi mematikan restore.
- Evidence resmi:
  - https://blogs.windows.com/windows-insider/2023/08/31/new-updates-for-snipping-tool-and-notepad-for-windows-insiders/

Implikasi:
- Saka Note harus lebih eksplisit dari awal soal policy startup/session.

## 9) Rekomendasi Eksperimen Produk (A/B internal, low-cost)
1. Startup Policy Presets
- `Classic`: selalu dokumen baru kosong.
- `Resume`: lanjutkan sesi terakhir.
- `Hybrid`: lanjutkan hanya file saved, drop unsaved scratch tab.

2. Tab vs Multi-Window Mode
- Toggle: `Use Tabs` ON/OFF.
- OFF = satu dokumen per window (mendekati alur klasik).

3. Input/Scroll Responsiveness Budget
- Gate rilis jika:
  - open 5MB text > target waktu maksimum,
  - scroll jitter melebihi ambang,
  - typing burst memicu stutter.

## 10) Status Eksekusi Implementasi (20 Februari 2026)

- P0.1 (`Interaction latency budget`) - **DONE**
  - Benchmark internal open/typing/scroll sudah ada.
  - CI gate headless (`--benchmark-ci`) sudah aktif di workflow build x64.
- P0.1b (`Claim benchmark harness`) - **DONE**
  - Script A/B `tools/compare-notepad.ps1` untuk Legacy vs Microsoft Notepad.
  - Output otomatis: `summary.csv`, `samples.csv`, `comparison.md`.
- P0.1c (`Reliability & session perf hardening`) - **DONE (Pass-2/3)**
  - Tangani partial read/write file I/O (looped I/O).
  - Hindari retry autosave agresif saat persist gagal (retry backoff).
  - Session persist tidak lagi memaksa simpan full text untuk file yang sudah saved + unmodified.
  - Perbaikan lifecycle `LoadLibrary`/`FreeLibrary` untuk RichEdit module.
- P0.2 (`Session behavior matrix`) - **IN PROGRESS**
  - Struktur state/session sudah ada, masih perlu validasi UX menu + edge-case QA.
- P0.3 (`Tab hit-target correctness`) - **IN PROGRESS**
  - Fondasi tab custom sudah ada, masih perlu regression pass multi-DPI.

## 11) Update Deep Research - Word Wrap vs Horizontal Scroll (20 Februari 2026)

### A. Fakta teknis Win32/RichEdit (sumber primer Microsoft)
- Style `ES_AUTOHSCROLL` menjaga satu baris panjang dan men-scroll horizontal saat kursor melewati batas kanan.
- Untuk multiline edit control, jika `ES_AUTOHSCROLL` tidak dipakai, teks akan di-wrap otomatis ke baris baru.
- RichEdit mendukung mode no-wrap lewat style horizontal (`WS_HSCROLL`/`ES_AUTOHSCROLL`) dan mode wrap dengan menonaktifkan style tersebut.
- Evidence:
  - https://learn.microsoft.com/en-us/windows/win32/controls/edit-control-styles
  - https://learn.microsoft.com/en-us/windows/win32/controls/rich-edit-control-styles
  - https://learn.microsoft.com/en-us/windows/win32/controls/em-settargetdevice

Implikasi:
- Jika target UX adalah notepad modern tanpa horizontal scrollbar pada ngetik normal, `Word Wrap` harus aktif default.
- Horizontal scrollbar sebaiknya hanya muncul saat user memang mematikan wrap.

### B. Eksekusi ke Saka Note
- Default state `wordWrap` digeser ke ON.
- Migrasi setting lama ditambahkan supaya install lama yang masih OFF ikut naik ke default baru.
- Path pembuatan RichEdit awal sekarang mengikuti state wrap (tidak selalu membuat `WS_HSCROLL` dulu).

### C. Observasi lokal Notepad original (Windows 11)
- Paket Notepad terdeteksi sebagai `Microsoft.WindowsNotepad` versi `11.2510.14.0`.
- State tab/window disimpan di:
  - `%LOCALAPPDATA%\\Packages\\Microsoft.WindowsNotepad_8wekyb3d8bbwe\\LocalState\\TabState`
  - `%LOCALAPPDATA%\\Packages\\Microsoft.WindowsNotepad_8wekyb3d8bbwe\\LocalState\\WindowState`
- Temuan ini konsisten dengan model modern notepad yang session-oriented (tab + restore state).

Status:
- P0.4 (`Word-wrap parity with modern notepad expectation`) - **DONE**

## 12) Deep Research - Tab Menyatu dengan Content Area (21 Februari 2026)

### A. Problem statement
- Gejala: tab terlihat "melayang" (floating chip), tidak menyatu dengan area editor/page.
- Dampak UX: affordance tab aktif lemah, visual hierarchy terasa putus.

### B. Evidence teknis (sumber primer Microsoft)
- Win32 tab control punya dua area berbeda: `tab strip` dan `display area/page`.
- Microsoft merekomendasikan pakai `TabCtrl_AdjustRect` untuk mendapatkan rectangle display area yang benar.
- Tab control menggambar border area page saat `WM_PAINT`; jika layout child page dihitung dari tinggi strip mentah, hasil visual mudah terlihat seperti ada celah.
- Evidence:
  - Tab control overview: https://learn.microsoft.com/en-us/windows/win32/controls/tab-controls
  - TabCtrl_AdjustRect macro: https://learn.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-tabctrl_adjustrect
  - About tab controls: https://learn.microsoft.com/en-us/windows/win32/controls/about-tab-controls
  - WinUI TabView guidance (state aktif harus jelas terhubung ke content): https://learn.microsoft.com/en-us/windows/apps/design/controls/tab-view

### C. Analisa Saka Note (lokal codebase)
- Saka Note memakai `WC_TABCONTROLW` + custom paint (`NM_CUSTOMDRAW`) di `src/main.cpp`.
- Sebelum fix, layout editor dihitung dari `tabsH` tetap di `ResizeControls()` (`src/modules/ui.cpp`), bukan dari display-area tab control.
- Akibatnya, garis separator strip + posisi top editor berpotensi membentuk seam/celah visual.

### D. Solusi implementasi yang dijalankan
1. Layout editor dipindah ke basis `TabCtrl_AdjustRect`:
- `ResizeControls()` sekarang menghitung `displayRect.top` dari tab control lalu memakai nilai itu sebagai `editorTop`.
- Ini mengikuti contract Win32 tab control, bukan hardcoded strip height.

2. Separator strip diputus di bawah tab aktif:
- Garis bawah strip tidak lagi full-width; area tepat di bawah tab aktif dihilangkan.

3. Tab aktif diperpanjang hingga baseline strip:
- Background tab aktif di-extend ke bawah agar transisi tab -> page continuous.

### E. Status
- P0.5 (`Tab-to-editor visual continuity`) - **DONE**
- Catatan QA:
  - Verifikasi lagi pada DPI 100%, 125%, 150% dan tema light/dark.
  - Uji saat tab label sangat panjang + hover close button.

## 13) Deep Research - Divider dan Ruang Strip Tab (21 Februari 2026)

### A. Problem statement
- Setelah tab dibuat "menyatu", muncul masalah baru: strip tab terasa tidak punya ruang sendiri (hierarki visual hilang).
- Gejala: tab dan area editor terasa terlalu menempel, sehingga tidak ada pemisahan yang membantu scanning.

### B. Evidence teknis dari platform Microsoft
- Tab control Win32 punya konsep `display area` yang terpisah dari `window rectangle`; layout page seharusnya mengikuti `TCM_ADJUSTRECT/TabCtrl_AdjustRect`.
- Dokumentasi Win32 juga menjelaskan tab control menggambar border area display saat paint, sehingga ruang strip + boundary memang bagian natural dari kontrol.
- Guidance TabView (Windows apps) membedakan jelas `tab strip` vs content area sebagai dua bagian anatomi.
- Evidence:
  - https://learn.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-tabctrl_adjustrect
  - https://learn.microsoft.com/en-us/windows/win32/controls/tab-controls
  - https://learn.microsoft.com/en-us/windows/apps/develop/ui/controls/tab-view
  - https://blogs.windows.com/windows-insider/2023/01/19/tabs-in-notepad-begins-rolling-out-to-windows-insiders/

### C. Analisa Saka Note
- Sebelumnya strip dibuat terlalu blend dengan editor (separator minim), jadi "container feel" hilang.
- Untuk menjaga lightweight sekaligus clarity:
  - strip perlu tone background sendiri (subtle),
  - divider tipis di bawah strip tetap ada,
  - ada gutter tipis antara display area tab dan top editor.

### D. Solusi implementasi yang dijalankan
1. Strip background diberi tone halus berbeda dari editor.
2. Divider bawah strip dikembalikan full-width (tipis 1px).
3. Tab item diberi top/bottom inset agar punya ruang strip sendiri.
4. Layout editor diberi gutter `+1px` dari `displayRect.top` hasil `TabCtrl_AdjustRect`.
5. Tinggi strip/padding tab disesuaikan untuk hierarchy yang lebih jelas tanpa efek berat.

### E. Status
- P0.6 (`Tab strip hierarchy and divider clarity`) - **DONE**

## 14) Implementation Note - Attached Tab Emulation Pass (21 Februari 2026)

Tujuan:
- Meniru pola visual tab Notepad modern yang "menempel" ke content page tanpa efek floating card berlebihan.

Eksekusi di Saka Note:
1. Divider strip bawah dipecah kiri/kanan area tab aktif, bukan full-width.
2. Tab aktif diextend ke bawah (melewati baseline strip) agar transisi ke page terlihat kontinu.
3. Tab inactive dibuat lebih flat (menggunakan tone strip) supaya hierarki fokus ke tab aktif.
4. `editorTop` kembali di-anchor ke `displayRect.top - 1` untuk memperkuat efek attached.

Catatan batasan:
- Posisi tab di title bar (seperti UWP/WinUI Notepad) belum bisa identik karena Saka Note masih memakai menu bar Win32 klasik (`SetMenu`) di non-client flow.
- Untuk identik 1:1, perlu migrasi ke custom title bar + custom menu host.

Status:
- P0.7 (`Attached-tab emulation with classic Win32 menu`) - **DONE**

## 15) UX Refinement - Close Button on Hover Only (21 Februari 2026)

Tujuan:
- Menyamakan affordance tab dengan notepad modern: tombol close tidak terus-menerus tampil di semua tab.

Prinsip:
- Progressive disclosure: kontrol destruktif ditampilkan saat relevan (hover), sehingga noise visual turun.
- Lightweight-safe: perubahan hanya pada paint/hit-test logic, tanpa timer animasi atau layer komposisi tambahan.

Implementasi:
1. Glyph close hanya dirender pada tab yang sedang hover.
2. Hit-test close juga aktif hanya saat tab sedang hover (mencegah accidental close area "tak terlihat").
3. Text layout tab otomatis melebar ketika close tidak tampil.

Status:
- P0.8 (`Hover-only close affordance`) - **DONE**

## 16) Polish Batch - Premium Tab Rhythm (21 Februari 2026)

Tujuan:
- Naikkan kualitas visual tab (lebih intentional/premium) tanpa menambah cost runtime yang berat.

Implementasi ringan:
1. Typography hierarchy:
- Tab aktif memakai font semibold, tab non-aktif regular.
- Font dibangun sekali dan di-refresh hanya saat DPI/theme/style change.

2. Rhythm & spacing:
- Padding tab disesuaikan berbasis DPI (`TCM_SETPADDING`) agar konsisten antar scaling.

3. Active affordance:
- Ditambahkan accent line tipis pada tab aktif (2px) untuk focus cue yang cepat terbaca.
- Tetap tanpa animasi, blur, atau compositing berat.

Guardrail performa:
- Tidak ada timer animasi.
- Tidak ada bitmap offscreen baru.
- Hanya operasi GDI sederhana (`FillRect`/`RoundRect`) pada event paint tab.

Status:
- P0.9 (`Premium tab polish with lightweight render path`) - **DONE**

## 17) Visual Decision - No Dot, Continuous Divider (21 Februari 2026)

Catatan:
- Keputusan ini kemudian direvisi pada Section 18 setelah evaluasi visual lanjutan untuk mengembalikan integrasi tab aktif-page (segmented divider + baseline sync).

Perubahan keputusan:
- Indikator aktif berbentuk dot/line di tab dihapus total.
- Divider strip tab dikembalikan full-width (tidak diputus di area tab aktif).
- Offset top editor diselaraskan ke `displayRect.top` agar garis divider terlihat kontinu dan tidak terputus.

Alasan:
- Menurunkan ambiguitas makna indikator visual (user tidak perlu menebak arti dot/line).
- Menjaga boundary tab/content tetap jelas, rapi, dan konsisten.
- Tetap lightweight karena hanya perubahan paint/layout sederhana.

Status:
- P1.0 (`No-indicator tab style with seamless divider continuity`) - **DONE**

## 18) Deep Research - Active Tab/Page Integration (21 Februari 2026)

### A. Masalah saat ini
- Gejala: tab aktif terlihat sebagai komponen terpisah dari editor (tidak terasa satu surface).
- Ini terjadi walau warna sudah mirip, karena boundary geometry belum sinkron.

### B. Temuan teknis dari referensi primer
1. Win32 TabControl memisahkan `tab strip` dan `display area`.
- Rectangle page ideal harus dihitung via `TabCtrl_AdjustRect` (bukan hardcoded tinggi strip).
- Sumber: https://learn.microsoft.com/en-us/windows/win32/api/commctrl/nf-commctrl-tabctrl_adjustrect

2. Pada custom draw, urutan paint menentukan continuity visual.
- Jika separator strip digambar full-width lalu tab aktif tidak overdraw baseline dengan tepat, akan terlihat seam.
- Sumber: https://learn.microsoft.com/en-us/windows/win32/controls/about-custom-draw

3. Guidance tab modern Microsoft tetap menekankan hubungan kuat selected tab <-> content.
- Implikasi desain: selected tab harus terbaca sebagai pintu ke page, bukan chip terpisah.
- Sumber:
  - https://learn.microsoft.com/en-us/windows/apps/design/controls/tab-view
  - https://learn.microsoft.com/en-us/windows/apps/design/controls/navigationview

### C. Analisa akar masalah di Saka Note
- Root cause bukan hanya warna, tapi 3 baseline yang tidak selalu sinkron:
  1. posisi divider strip,
  2. posisi bottom edge tab aktif,
  3. posisi top edge editor.
- Saat salah satu baseline bergeser 1px, mata membaca "putus".

### D. Solusi arsitektural yang dipakai
1. Divider strip dipotong hanya pada area tab aktif (segmented divider).
2. Tab aktif diperpanjang sampai baseline page, lalu bottom stroke ditutup (`joinRect`) agar menyatu.
3. `editorTop` diikat ke `displayRect.top - 1` dari `TabCtrl_AdjustRect` agar edge editor sejajar dengan tab/page baseline.

### E. Status
- P1.1 (`Baseline-synced attached tab rendering`) - **DONE**
- QA lanjutan yang disarankan:
  - Uji DPI 100/125/150/175.
  - Uji tema light/dark + scaling font Windows.
  - Uji 1 tab vs banyak tab + label panjang.

## 19) Micro-Spacing Fix - Tab Title vs Editor Text (21 Februari 2026)

Masalah:
- Secara visual, judul tab aktif dan baris pertama teks editor terasa terlalu rapat/tanpa napas.

Pendekatan solusi:
- Jangan geser boundary tab-editor (karena bisa merusak efek attached).
- Tambah `internal text viewport padding` pada RichEdit, khususnya top inset.

Implementasi:
1. Tambah `ApplyEditorViewportPadding()` di modul editor.
2. Padding di-scale berdasarkan DPI (agar konsisten antar 100/125/150%).
3. Fungsi dipanggil saat:
   - editor dibuat ulang (word-wrap toggle/recreate),
   - resize editor (`WM_SIZE`),
   - layout pass global (`ResizeControls`).

Hasil:
- Boundary tab-editor tetap attached.
- Konten teks memiliki gap vertikal yang nyaman dari atas editor.

Status:
- P1.2 (`Editor viewport top-padding without breaking attached tabs`) - **DONE**
