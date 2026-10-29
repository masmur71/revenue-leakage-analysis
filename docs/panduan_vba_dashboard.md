# Otomatisasi dashboard Phase 4 melalui VBA

**Pembaruan 2026-10-30:** Verify penuh sudah dihentikan dan hasil parsialnya telah direkonsiliasi. Untuk pekerjaan saat ini, ikuti [panduan penyelesaian singkat](panduan_penyelesaian_phase4.md), bukan mengulang bagian Verify di bawah. Macro publik Phase4_Verify kini hanya menampilkan pesan pengaman; gunakan Phase4_CheckCohortSamples (4 kasus/10 nilai), kemudian refresh terpisah. Bagian lama dipertahankan sebagai catatan rancangan awal.

Paket ini melanjutkan workbook **RevenueLeakage.xlsx** yang sudah Anda kerjakan. Modul dijalankan oleh Anda di Excel Windows melalui **DashboardRunner.xlsm**, terpisah dari hasil utama. Jangan memasukkan modul ke RevenueLeakage.xlsx.

Kode sudah disiapkan dan diuji di luar Excel. **Compile VBA, pembangunan dashboard, dan seluruh pengujian native modul baru belum dinyatakan berhasil.** Pesan selesai dari satu tahap bukan penutupan Phase 4.

## 1. Siapkan runner satu kali

1. Buka RevenueLeakage.xlsx dan tekan **Ctrl+S**. Biarkan terbuka.
2. Buat workbook kosong dengan **Ctrl+N**.
3. Tekan **F12 → Browse**. Pilih folder lokal:

   ```text
   C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis\excel
   ```

4. Nama: **DashboardRunner.xlsm**. Pada **Save as type**, pilih **Excel Macro-Enabled Workbook (*.xlsm)**, lalu **Save**. MeasureInstaller.xlsm yang lama tetap terpisah.
5. Tekan **Alt+F11**, kemudian **Ctrl+R** untuk menampilkan Project Explorer.
6. Klik **VBAProject (DashboardRunner.xlsm)**. Pilih menu atas **File → Import File…** dan impor **excel\DashboardRunner.bas**. Ini mengimpor file kode, bukan workbook.
7. Di bawah **Modules**, klik dua kali **DashboardRunner** untuk membuka kode. Pada menu atas editor, pilih **Debug → Compile VBAProject**. Jika pilihan Compile abu-abu dan tidak ada error, lanjut. Jika muncul error, catat pesan dan baris yang disorot; hentikan tahap ini.
8. Tekan **Ctrl+S**, lalu **Alt+Q** untuk kembali ke Excel.

Jika ada pembatasan macro dari Excel/organisasi, jangan mengubah pengaturan keamanan. Catat pesan yang ditampilkan sebagai blocker. Jika modul diperbarui nanti, lepas hanya modul DashboardRunner lama dari runner, lalu impor file .bas terbaru; jangan mengimpor dua salinan.

## 2. Jalankan Audit terlebih dahulu

1. Pastikan RevenueLeakage.xlsx sudah disimpan.
2. Tekan **Alt+F8**. Pada **Macros in**, pilih **All Open Workbooks** atau **DashboardRunner.xlsm**.
3. Pilih **Phase4_Audit → Run**.
4. Jika muncul pemilih folder, pilih folder **excel** lokal pada langkah 1. Folder tersebut harus berisi RevenueLeakage.xlsx, measures.dax, relationships.csv, dan dashboard-specs.
5. Jika Excel mengembalikan alamat OneDrive berbentuk URL, konfirmasi **Yes** hanya jika dialog menunjuk workbook Phase 4 Anda. Pilihan awal adalah **No**.
6. Tunggu pesan **Audit passed**. Catat lokasi folder bukti yang muncul.

Audit mencatat tabel dan jumlah baris, hubungan aktif, formula measures, PivotTables/field/filter, grafik, dan hubungan slicer. Nama dan susunan field digunakan untuk mengenali objek lama meskipun caption-nya diubah menjadi “Monthly Net Sales” atau “Target”. Audit tidak membangun dashboard dan tidak mengubah definisi model.

**Jika berhenti:** kirim pesan error dan file **events.csv** dari folder yang disebutkan. Konflik area, objek ambigu, formula berbeda, atau tabel tidak cocok harus diselesaikan sebelum Build. Jangan menghapus PivotTable/grafik lama agar macro bisa melewati pemeriksaan.

## 3. Bangun dashboard

1. Setelah Audit berhasil, tekan **Ctrl+S** pada RevenueLeakage.xlsx jika diperlukan.
2. **Alt+F8 → Phase4_BuildDashboard → Run**.
3. Modul mengulang preflight, lalu membuat backup `.xlsx` sebelum perubahan pertama.
4. Tunggu pesan **Dashboard stage finished**, kemudian periksa hasil dan tekan **Ctrl+S** pada RevenueLeakage.xlsx.

Yang dibangun atau digunakan kembali:

| Sheet | Isi |
|---|---|
| Executive Summary | Delapan kartu KPI, tren bulanan aktual/target, pencapaian region/cabang, slicer dimensi |
| Leakage Diagnostics | Diagnostik cabang, grafik price shortfall, 10 SKU/outlet teratas berdasarkan shortfall, alasan retur tanpa anggota kosong |
| Operations | Piutang, kunjungan salesman, RFM, dan cohort tanpa grand total activity month |
| Data & Calculations | pControls lama dipertahankan; kontrol lengkap dan sumber kartu/grafik ditambahkan mulai kolom AA |
| Methodology | Definisi periode/snapshot, BLANK, pembulatan, dan pencegahan hitung ganda |

PivotTable baru menggunakan Data Model dan measures eksplisit. Grafik tren dan slicer cabang lama digunakan kembali jika identitasnya jelas. Default setelah Build adalah tahun 2025 dan semua cabang; filter manual pada objek yang dikelola direset. Pemilihan slicer berikutnya akan memengaruhi objek yang terhubung.

Jalankan Build sekali lagi setelah menyimpan untuk memeriksa bahwa objek lama digunakan kembali dan tidak ada duplikasi. Jika proses sebelumnya berhenti di tengah, penambahan parsial dapat tetap ada; log menyebut tahap yang gagal. Jangan menyimpulkan bahwa rerun memperbaiki setiap konflik secara otomatis.

## 4. Rekam hasil native Excel

**Melanjutkan Build yang sudah membuat semua PivotTable tetapi berhenti pada slicer:** simpan target, impor versi modul terbaru, lalu jalankan **Phase4_FinishDashboard**. Macro ini memvalidasi dan menggunakan PivotTables yang sudah lengkap, membuat backup, kemudian melanjutkan sambungan slicer/kartu/metodologi tanpa menyusun ulang field tabel. Jika ada PivotTable yang belum lengkap, proses berhenti dan meminta Build terlebih dahulu. Pembuatan slicer dan pengaturan posisi dicatat terpisah; error 5 tidak dianggap sudah teratasi sebelum percobaan native berhasil.

1. Simpan RevenueLeakage.xlsx.
2. **Alt+F8 → Phase4_Verify → Run**.
3. Modul membuat **Native Verification** dan PivotTable **rlNativeProbe**, terpisah dari sambungan slicer dashboard. Jangan menghubungkan slicer ke PivotTable ini.
4. Tunggu seluruh kasus selesai. Status bar menampilkan kasus yang sedang berjalan. Proses mencakup seluruh sel cohort yang teramati sehingga dapat memakan waktu.
5. Pesan **Native captures written** berarti pengambilan nilai selesai, belum berarti angka cocok dengan SQL.

File pada folder run:

- **native-results.csv**: 46 KPI default dan kombinasi filter komersial/pendukung, sesuai template proyek.
- **extended-results.csv**: guard BLANK, residual, snapshot lintas tahun, segmen RFM, dan cohort.
- **events.csv**: kasus yang benar-benar dieksekusi.

VBA tidak membaca angka harapan SQL untuk mengisi hasil. File berakhiran **.partial.csv** menandakan pengambilan belum lengkap; jangan mengubah namanya menjadi hasil final.

Untuk pemeriksaan, buka PowerShell pada folder project. Ganti nilai `$captureRun` dengan folder yang benar dari dialog, kemudian jalankan:

```powershell
$nodePath = 'C:\Users\masmu\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'
$captureRun = 'C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis\excel\dashboard-runs\REPLACE-WITH-VERIFICATION-RUN'
& $nodePath excel/verify_exports.mjs --native (Join-Path $captureRun 'native-results.csv')
& $nodePath excel/check_dashboard_capture.mjs (Join-Path $captureRun 'extended-results.csv')
```

Kedua pemeriksaan harus berhasil. Anda juga dapat mengirim lokasi folder run agar agent menjalankan pembandingan. Pembulatan: uang/AOV 2 desimal, persentase 2 desimal tampilan, hari/produktivitas 2 desimal, counts tepat; BLANK harus tetap BLANK. Nilai RFM direkonsiliasi terhadap snapshot CSV yang sudah diperiksa validator; sel cohort terhadap SQLCohort.

## 5. Uji folder dan refresh secara terpisah

1. Simpan workbook dan pastikan tidak ada refresh lain yang berjalan.
2. **Alt+F8 → Phase4_TestRefresh → Run**.
3. Modul membuat backup dan menyimpan formula pFolder asli ke **original-pFolder.m**.
4. Modul merefresh sumber asli, menyalin CSV ke folder alternatif pada folder run, menguji sumber alternatif, lalu menguji path yang sengaja tidak ada.
5. Setelah pengujian, modul mengembalikan **seluruh formula pFolder asli**, lalu merefresh sumber asli kembali. Definisi query lain tidak diubah.
6. Buka **Data → Queries & Connections**. Pastikan tidak ada refresh berjalan atau error tersisa. Tekan **Data → Refresh All**, tunggu selesai, lalu simpan.
7. Jalankan **Phase4_Verify** lagi dan bandingkan kedua hasil capture terbaru.

Folder run refresh juga berisi **refresh-ORIGINAL.csv**, **refresh-ALTERNATE.csv**, dan **refresh-RESTORED.csv**, masing-masing berisi 46 nilai aktual Excel. Pada PowerShell, arahkan `$refreshRun` ke folder run refresh yang benar dan periksa ketiganya:

```powershell
$refreshRun = 'C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis\excel\dashboard-runs\REPLACE-WITH-REFRESH-RUN'
foreach ($refreshStage in @('ORIGINAL', 'ALTERNATE', 'RESTORED')) {
    & $nodePath excel/check_dashboard_capture.mjs --refresh (Join-Path $refreshRun "refresh-$refreshStage.csv")
    if ($LASTEXITCODE -ne 0) { throw "Refresh comparison failed: $refreshStage" }
}
```

Jika API tidak melaporkan error missing-folder secara sinkron, status pengujian menjadi **INCONCLUSIVE**, bukan PASS. Pengujian perlu ditinjau melalui Queries & Connections. Jumlah baris yang tetap sama saja tidak membuktikan sumber berhasil direfresh.

Jika tertulis **RECOVERY FAILED**, atau Excel/macro dihentikan ketika pengujian berlangsung, jangan simpan workbook dalam keadaan sumber salah. Buka query **pFolder → Advanced Editor**, pulihkan teks dari **original-pFolder.m** di folder run, refresh sumber asli, dan pastikan sehat sebelum menyimpan. Backup tetap tersedia. Jika tidak yakin, kirim log/error terlebih dahulu.

## 6. Pemeriksaan akhir yang tetap diperlukan

- Pilih Medan/2025: kartu, tabel, dan grafik berubah bersama; net sales **862.259.449,78**. Clear Filter mengembalikan nasional **15.762.528.589,51**.
- Ubah tahun ke 2024, lalu kembalikan 2025. Uji Jakarta pada kuartal I melalui pilihan tiga bulan pada slicer year_month; net sales **274.557.868,19**.
- Uji satu SKU: target/achievement, AR, visits, dan RFM yang tidak didukung menjadi kosong. Setelah filter dilepas, nilai pulih.
- Periksa seluruh lima sheet dashboard: tidak ada `####`, error formula, judul terpotong, grafik kosong, atau objek bertumpuk. Native Verification adalah sheet teknis tambahan.
- Tekan Ctrl+S, tutup RevenueLeakage.xlsx, buka lagi, Refresh All, lalu ulangi Medan/2025 dan Jakarta/kuartal I. Pastikan format tetap `.xlsx`.
- Kirim log dan screenshot untuk pencatatan bukti. Phase 4 hanya ditutup setelah seluruh acceptance checklist berhasil; commit/arsip status dilakukan setelah itu.

## Referensi teknis

Modul menggunakan API resmi [Model.DataModelConnection](https://learn.microsoft.com/en-us/office/vba/api/excel.model.datamodelconnection), [PivotCaches.Create](https://learn.microsoft.com/en-us/office/vba/api/excel.pivotcaches.create), [SlicerCaches.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.slicercaches.add), dan [Model.Refresh](https://learn.microsoft.com/en-us/office/vba/api/excel.model.refresh). Kesamaan API tidak menggantikan pengujian pada versi Excel yang Anda gunakan.
