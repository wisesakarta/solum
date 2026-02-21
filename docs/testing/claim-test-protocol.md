# Claim Test Protocol - Saka Note vs Microsoft Notepad
# Protokol Uji Klaim - Saka Note vs Microsoft Notepad

Purpose: compare performance fairly with identical test conditions.  
Tujuan: membandingkan performa secara adil dengan kondisi uji yang sama.

## 1. Required Baseline / Baseline Wajib

1. Use the same machine, same user account, and same power mode.  
   Gunakan mesin, akun pengguna, dan mode daya yang sama.
2. Close heavy background apps (many browser tabs, game launchers, IDE indexing).  
   Tutup aplikasi berat di latar belakang (browser tab banyak, game launcher, indexing IDE).
3. Ensure both editors are closed before each run.  
   Pastikan kedua editor tertutup sebelum setiap run.
4. Run at least 3 iterations to reduce noise.  
   Jalankan minimal 3 iterasi untuk mengurangi noise.

## 2. Standard Scenarios / Skenario Standar

1. Empty start (no file).
2. Open 1MB file.
3. Open 5MB file.
4. Open 20MB file.

`tools/compare-notepad.ps1` auto-generates shared corpus files in:  
`tools/compare-notepad.ps1` membuat corpus file otomatis di:
`research/perf-runs/corpus-shared`.

## 3. Automated Execution (Recommended) / Eksekusi Otomatis (Disarankan)

```powershell
cmake -S . -B build/mingw-debug -G "MinGW Makefiles" -DCMAKE_BUILD_TYPE=Debug
cmake --build build/mingw-debug
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\compare-notepad.ps1 -Iterations 3 -SampleSeconds 8 -SampleIntervalMs 500 -WarmupMs 3000 -ForceCloseBeforeEachRun
```

Note / Catatan:
- `-ForceCloseBeforeEachRun` will terminate active Notepad processes.  
  `-ForceCloseBeforeEachRun` akan menghentikan proses Notepad yang sedang aktif.
- Save important unsaved documents first.  
  Simpan dulu dokumen penting yang belum tersimpan.

Outputs / Keluaran:
- `research/perf-runs/run-*/summary.csv`
- `research/perf-runs/run-*/samples.csv`
- `research/perf-runs/run-*/comparison.md`

Optional cleanup / Cleanup opsional:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\cleanup-perf-runs.ps1 -KeepLatest 2
```

## 4. Measured Metrics / Metrik yang Diukur

1. CPU (%)
2. Working Set (MB)
3. Private Memory (MB)
4. Paged Memory (MB)
5. Handles
6. Threads
7. Startup time (window handle available)

## 5. Manual Task Manager Audit / Audit Manual Task Manager

1. Open `taskmgr`.  
   Buka `taskmgr`.
2. Go to `Details` tab.  
   Masuk tab `Details`.
3. Right-click column header -> `Select columns`.  
   Klik kanan header kolom -> `Select columns`.
4. Enable at minimum: `CPU`, `Memory (active private working set)`, `Commit size`, `Handles`, `Threads`.  
   Aktifkan minimal: `CPU`, `Memory (active private working set)`, `Commit size`, `Handles`, `Threads`.
5. Run identical scenarios for both applications.  
   Jalankan skenario yang sama untuk kedua aplikasi.
6. Record peak values per metric for each scenario.  
   Catat nilai puncak per metrik di setiap skenario.

## 6. Winner Rule / Aturan Penentuan Pemenang

1. Lower is better for all metrics above.  
   Nilai lebih rendah lebih baik untuk semua metrik di atas.
2. Use average or median of 3 iterations.  
   Gunakan rata-rata atau median dari 3 iterasi.
3. Do not finalize claims from a single anomalous run.  
   Jangan finalisasi klaim hanya dari satu run anomali.
4. Keep all evidence (`summary.csv`, `samples.csv`, Task Manager screenshots).  
   Simpan semua bukti (`summary.csv`, `samples.csv`, screenshot Task Manager).

## 7. Technical Note / Catatan Teknis

Modern Microsoft Notepad can show single-instance behavior on some Windows builds.  
Microsoft Notepad modern dapat menunjukkan perilaku single-instance pada beberapa build Windows.

If startup time appears as `N/A`, use resource metrics (CPU/memory/handles/threads) for comparison.  
Jika startup time terbaca `N/A`, gunakan metrik resource (CPU/memory/handles/threads) untuk komparasi.
