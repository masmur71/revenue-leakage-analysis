# Menyelesaikan sisa Phase 4 tanpa mengulang Verify penuh

Hasil native yang sudah disimpan telah diperiksa: **282 nilai utama dan 131 nilai tambahan semuanya cocok**. Jangan menjalankan ulang Build/Finish atau rangkaian Verify penuh. Bukti dan batas cakupan ada di `docs/excel/partial-verification-review.md`.

## A. Jalankan hanya empat sampel cohort

1. Simpan **RevenueLeakage.xlsx** dengan **Ctrl+S**. Biarkan terbuka.
2. **Alt+F11 → Ctrl+R**. Di **VBAProject (DashboardRunner.xlsm) → Modules**, klik kanan **DashboardRunner → Remove DashboardRunner… → No**.
3. **File → Import File…**, pilih **excel\DashboardRunner.bas** terbaru.
4. **Debug → Compile VBAProject**. Jika Compile abu-abu tanpa error, lanjut. Simpan runner, lalu **Alt+Q**.
5. **Alt+F8 → Phase4_CheckCohortSamples → Run**.
6. Jika diminta folder, pilih folder lokal project **excel**. Konfirmasi workbook OneDrive hanya bila benar menunjuk workbook Phase 4 Anda.
7. Status menampilkan **cohort sample 1/4** sampai **4/4**. Hanya empat kasus dijalankan; bukan 343 kasus.
8. Setelah pesan selesai, kirim lokasi folder run atau cukup beri tahu agent agar membaca log terbaru. Tidak perlu menjalankan macro kedua.

Kasus yang diperiksa:

| Kasus | Pemeriksaan |
|---|---|
| Cohort Januari 2025, aktivitas Januari 2025 | Jumlah outlet, outlet aktif dan retention cocok SQL |
| Cohort Desember 2025, aktivitas Desember 2025 | Jumlah outlet, outlet aktif dan retention cocok SQL |
| Cohort Desember 2025, aktivitas Januari 2025 | Sel sebelum onboarding kosong, bukan nol |
| Cohort Januari 2025, aktivitas Januari + Februari 2025 | Ketiga measure kosong untuk total lintas bulan |

Macro memakai sheet teknis baru **Cohort Verification**, PivotTable terpisah berisi tiga measures, dan tidak menghubungkannya ke slicer dashboard. Tidak ada filter DimDate karena cohort memakai tanggal sendiri. Dashboard dan hasil parsial lama dipertahankan. Backup dibuat lebih dahulu; file capture ditutup setelah setiap kasus agar hasil tersimpan bila terjadi interupsi.

Hasilnya **cohort-samples.csv**, berisi **10 nilai**. File `.partial.csv` berarti rangkaian kecil belum selesai. Untuk pembandingan oleh agent atau PowerShell:

```powershell
$nodePath = 'C:\Users\masmu\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
$sampleRun = 'C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis\excel\dashboard-runs\REPLACE-WITH-COHORT-SAMPLE-RUN'
& $nodePath excel/check_dashboard_capture.mjs --cohort-samples (Join-Path $sampleRun 'cohort-samples.csv')
```

Memilih **Phase4_Verify** pada modul terbaru hanya menampilkan pesan pengaman. Rangkaian exhaustive lama tidak lagi muncul sebagai macro publik untuk dijalankan tanpa sengaja.

## B. Setelah sampel cohort cocok: refresh terpisah

Jangan jalankan langkah B bersamaan dengan langkah A. Setelah hasil A diperiksa:

1. Simpan target, lalu jalankan **Phase4_TestRefresh**.
2. Tahap ini menguji sumber asli, folder alternatif, missing-folder failure, lalu mengembalikan sumber asli. Definisi pFolder asli disimpan dan backup dibuat sebelum perubahan.
3. Setelah selesai, periksa **Data → Queries & Connections**: tidak ada error/refresh tertinggal. Jangan menyimpulkan berhasil hanya dari jumlah baris.
4. Agent membandingkan **refresh-ORIGINAL.csv**, **refresh-ALTERNATE.csv**, dan **refresh-RESTORED.csv**, masing-masing 46 nilai, menggunakan opsi `--refresh`.
5. Tidak perlu menjalankan `Phase4_Verify` setelah refresh. Tiga capture referensi tadi merekonsiliasi hasil refresh, sementara bukti kombinasi filter yang sudah lolos tetap disimpan.

Default periode tahunan penuh dipilih melalui satu anggota year, bukan mengulang daftar 365 tanggal. Refresh model tetap dapat memerlukan waktu sesuai Excel; pengujian ini terpisah dari cohort dan tidak mengulang ratusan kasus.

Jika muncul **RECOVERY FAILED**, jangan simpan parameter sumber yang salah. Pulihkan query pFolder melalui Advanced Editor dari file **original-pFolder.m** dalam folder run, lalu refresh sumber asli. Lihat bagian recovery pada panduan VBA dashboard untuk rincian.

## C. Pemeriksaan tampilan dan simpan/buka ulang

- Periksa Executive Summary, Leakage Diagnostics, Operations, Data & Calculations, dan Methodology. Tidak ada `####`, error, label terpotong, grafik kosong atau objek bertumpuk.
- Pada cohort, aktivitas masa depan tidak muncul sebagai nol; grand total lintas activity month dinonaktifkan.
- Slicer tahun 2025 → 2024 → 2025 dan cabang Medan mengubah kartu/tabel/grafik bersama. Clear Filter memulihkan nasional. Jangan memilih SKU/outlet ketika membandingkan target.
- Simpan, tutup hanya RevenueLeakage.xlsx, buka lagi, Refresh All, lalu ulangi Medan/2025 (**net sales 862.259.449,78**) dan Jakarta/kuartal I 2025 (**274.557.868,19**). Lepaskan filter dan pastikan nasional 2025 kembali **15.762.528.589,51**.
- Catat atau kirim screenshot hasil. Agent mencatat bukti, memeriksa sisa acceptance, lalu baru melakukan commit, arsip status dan handoff jika semua syarat terpenuhi.

Fase tetap terbuka sampai pemeriksaan native ini berhasil. Tidak ada klaim bahwa semua 276 kasus cohort exhaustive yang dilewati sudah diuji.
