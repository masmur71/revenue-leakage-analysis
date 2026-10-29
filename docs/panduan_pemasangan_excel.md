# Panduan pemasangan model dan dashboard Excel

Panduan ini untuk **Microsoft Excel desktop di Windows**. Nama menu di bawah memakai bahasa Inggris seperti pada Excel yang sudah diperiksa. Jika Excel Anda berbahasa Indonesia, gunakan menu dengan fungsi yang sama.

**Status:** berkas sumber sudah tersedia, tetapi workbook native belum dibuat dan belum diverifikasi. Ikuti urutannya; jangan langsung membuat dashboard sebelum tabel, relasi, dan measures berhasil dimuat.

## 1. Kenali berkas yang dipakai

Folder proyek:

```text
C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis
```

| Lokasi dari folder proyek | Kegunaan |
|---|---|
| `data\excel` | Sumber CSV yang dibaca Excel; jangan membuka semua CSV sebagai worksheet |
| `excel\power_query` | Kode Power Query, masing-masing dalam berkas `.m` |
| `excel\relationships.csv` | Daftar 28 relasi yang harus dibuat |
| `excel\measures.dax` | Rumus measures yang dimasukkan satu per satu |
| `docs\excel\csv-verification.json` | Jumlah baris yang diharapkan untuk setiap tabel |
| `docs\excel\native-execution.md` | Tempat mencatat hasil pengujian Excel yang sebenarnya |

Buka berkas `.m` dan `.dax` menggunakan **Notepad atau Visual Studio Code** untuk menyalin isinya. Jangan menempelkan nama/path berkas ke Advanced Editor; yang ditempel adalah **seluruh isi kodenya**. Nama tabel, query, kolom, dan measure tetap memakai bahasa Inggris agar sesuai dengan kode.

CSV untuk snapshot **31 Desember 2025** sudah diekspor. Pemasangan awal tidak memerlukan koneksi langsung Excel ke PostgreSQL atau Docker. Jangan mengedit CSV untuk mengubah hasil KPI.

## 2. Buat dan simpan workbook baru

1. Buka Excel → **Blank workbook**. Gunakan workbook baru, terpisah dari `Book1` yang sebelumnya terhubung.
2. Tekan **F12** atau pilih **File → Save As → Browse**.
3. Pilih folder `excel` di dalam proyek. Isi nama **RevenueLeakage.xlsx** dan jenis **Excel Workbook (*.xlsx)**, lalu klik **Save**.
4. Buat lima worksheet dengan tombol **+**, lalu ganti namanya dengan klik kanan tab → **Rename**:
   - `Executive Summary`
   - `Leakage Diagnostics`
   - `Operations`
   - `Data & Calculations`
   - `Methodology`
5. Simpan lagi dengan **Ctrl+S**.

**Periksa:** judul workbook adalah `RevenueLeakage.xlsx`, dan kelima tab di atas tersedia. Workbook ini tidak memerlukan macro.

## 3. Pastikan Power Pivot tersedia

1. Periksa apakah tab **Power Pivot** sudah terlihat di ribbon. Jika ada, lanjut ke bagian 4.
2. Jika belum ada: **File → Options → Add-ins**.
3. Pada pilihan **Manage**, pilih **COM Add-ins**, lalu klik **Go…**.
4. Centang **Microsoft Power Pivot for Excel** atau nama Microsoft Power Pivot yang tersedia, lalu **OK**.
5. Periksa kembali tab **Power Pivot**.

Jika Power Pivot tidak tercantum atau tidak dapat diaktifkan, catat pesan yang terlihat. Jangan mengganti Data Model dengan PivotTable biasa. Cara aktivasi mengacu pada [petunjuk Microsoft](https://support.microsoft.com/en-us/excel/start-the-power-pivot-add-in-for-excel).

## 4. Buat parameter folder: pFolder

Parameter ini membuat lokasi CSV bisa diganti tanpa mengubah seluruh query.

1. Di Excel pilih **Data → Get Data → From Other Sources → Blank Query**. Jendela **Power Query Editor** akan terbuka.
2. Pilih **Home → Advanced Editor**.
3. Buka `excel\power_query\pFolder.m` di editor teks. Salin seluruh isinya.
4. Pada Advanced Editor tekan **Ctrl+A**, tempel kode, kemudian klik **Done**.
5. Di panel kanan **Query Settings**, ubah **Name** menjadi tepat `pFolder`. Jika panel tidak terlihat, aktifkan melalui **View → Query Settings**.
6. Pastikan nilai parameter menunjuk ke:

   ```text
   C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis\data\excel
   ```

   Jika proyek dipindah, ganti nilai parameter dengan folder `data\excel` di lokasi baru. Pilih folder, bukan satu berkas CSV.
7. Pilih panah di samping **Home → Close & Load → Close & Load To…**. Gunakan **Only Create Connection** dan **jangan** centang **Add this data to the Data Model**.

Jika opsi pemuatan tidak ditampilkan untuk parameter, cukup tutup editor tanpa memuatnya ke worksheet. Parameter bukan tabel data.

**Periksa:** pada **Data → Queries & Connections** ada query `pFolder`, tanpa worksheet baru berisi path folder.

## 5. Buat fungsi pembaca: fnLoadCsv

1. Buat **Blank Query** lagi seperti bagian 4.
2. Buka **Advanced Editor**. Ganti seluruh kode bawaan dengan seluruh isi `excel\power_query\fnLoadCsv.m`.
3. Klik **Done**, lalu ubah **Name** menjadi tepat `fnLoadCsv`.
4. Jika tampil kotak isian `fileName`, `schema`, dan `primaryKey`, itu normal: query ini adalah **fungsi**. Tidak perlu mengisi atau menekan **Invoke**.
5. Tutup editor. Jadikan **connection only** jika diminta; jangan muat fungsi ke worksheet atau Data Model.

**Periksa:** query bernama `fnLoadCsv` tersedia, biasanya dengan ikon `fx`. Tidak ada pesan error.

## 6. Muat 14 tabel ke Data Model

Lakukan contoh `DimBranch` berikut terlebih dahulu.

1. Buat **Blank Query** → **Advanced Editor**.
2. Tempel seluruh isi `excel\power_query\DimBranch.m`, lalu **Done**.
3. Ubah **Name** menjadi tepat `DimBranch`, tanpa akhiran `.m`.
4. Pastikan preview berisi kolom seperti `branch_id`, `branch_name`, dan `region_name`.
5. Pilih panah **Close & Load → Close & Load To…**.
6. Dalam jendela **Import Data**:
   - Pilih **Only Create Connection**.
   - **Centang Add this data to the Data Model**.
   - Klik **OK** dan tunggu sampai pemuatan selesai.
7. Buka **Power Pivot → Manage**. Pastikan tabel `DimBranch` ada dan berisi 15 baris, kemudian kembali ke jendela Excel utama.

**Penting:** connection only saja belum cukup. Tabel juga harus ditambahkan ke **Data Model**. Panel Queries & Connections dapat menampilkan “Connection only” meskipun tabel sudah masuk model; periksa keberadaannya melalui Power Pivot. Pilihan pemuatan ini mengikuti [petunjuk Power Query Microsoft](https://support.microsoft.com/en-us/excel/create-load-or-edit-a-query-in-excel-power-query).

Ulangi langkah yang sama untuk semua baris berikut, memakai berkas `.m` dengan nama yang persis sama:

| Urutan | Nama query/tabel | Jumlah baris untuk snapshot awal |
|---|---|---:|
| 1 | DimBranch | 15 |
| 2 | DimChannel | 3 |
| 3 | DimDate | 731 |
| 4 | DimOutlet | 3.000 |
| 5 | DimProduct | 300 |
| 6 | DimSalesman | 50 |
| 7 | FactCommercial | 335.988 |
| 8 | FactVisits | 105.918 |
| 9 | FactInvoiceSnapshot | 129.653 |
| 10 | FactTargets | 360 |
| 11 | FactPayments | 121.924 |
| 12 | FactRFM | 3.000 |
| 13 | FactCohort | 10.457 |
| 14 | ReportContext | 1 |

Jumlah di atas berasal dari ekspor yang sudah diperiksa, bukan bukti bahwa Excel sudah memuatnya. Tabel besar dapat memerlukan waktu lebih lama. Jangan memuatnya ke worksheet. Jika pilihan **Close & Load To…** tidak aktif, tutup editor lalu klik kanan nama query pada **Queries & Connections → Load To…**.

**Titik pemeriksaan pertama:** sebelum melanjutkan, pastikan semua 14 tabel terlihat di Power Pivot, jumlah baris sesuai, dan tidak ada query error. Simpan workbook. Jika ingin pemeriksaan bersama, berhenti di sini dan beri tahu bahwa tahap 6 selesai.

## 7. Muat SQLReference ke worksheet

1. Buat **Blank Query**, tempel isi `SQLReference.m`, dan beri nama `SQLReference`.
2. Pada **Close & Load To…**, pilih **Table → Existing worksheet**.
3. Tentukan lokasi `'Data & Calculations'!$A$5` dengan tombol pemilih lokasi atau pilih sel A5 pada tab tersebut.
4. **Jangan centang Add this data to the Data Model**. Klik **OK**.

**Periksa:** tabel kontrol berisi 46 baris metrik dengan kolom `metric_id`, `numerator`, `denominator`, `value`, dan `unit`.

**Existing worksheet** berarti lokasi tabel dimuat, bukan nama worksheet yang wajib dipilih. Untuk SQLReference, pilih **Data & Calculations!A5**, bukan Executive Summary. Jika sudah terlanjur dimuat di Executive Summary, query tidak perlu dibuat ulang; pindahkan seluruh tabel beserta header ke Data & Calculations sebelum menyusun dashboard.

Total ada **17 query**: 14 tabel model, 1 tabel kontrol, 1 parameter, dan 1 fungsi. CSV lain seperti SQLFilterControls belum perlu dimasukkan ke model.

## 8. Buat 28 relasi antartabel

Bagian ini dilakukan pada **jendela Excel utama atau Power Pivot**, bukan Power Query Editor. Jika masih di editor query, pilih Close & Load terlebih dahulu. Di jendela Excel utama, maksimalkan jendela; tombol Relationships biasanya berada dalam grup **Data Tools** pada tab Data. Grup ini bisa terlipat pada ribbon yang sempit.

**Jika tombol Relationships tidak ditemukan, gunakan jalur alternatif berikut:**

1. Pada Excel utama pilih **Power Pivot → Manage**.
2. Pada jendela Power Pivot pilih **Home → Diagram View**.
3. Pastikan kotak tabel `DimDate` dan `FactCommercial` terlihat.
4. Seret kolom **DimDate[date]** ke **FactCommercial[event_date]**.
5. Periksa muncul garis relasi dengan sisi **1** pada DimDate dan sisi banyak (**\*** ) pada FactCommercial. Jika relasi sudah ada, jangan membuatnya lagi.
6. Ulangi untuk pasangan kolom pada `excel\relationships.csv`: seret `dimension[dimension_key]` ke `fact[fact_key]`. Ini alternatif untuk membuat 28 relasi yang sama, bukan 28 relasi tambahan.

Cara Diagram View mengikuti [petunjuk Microsoft](https://support.microsoft.com/en-au/excel/create-relationships-in-diagram-view-in-power-pivot). Jika tabel tidak terlihat di Power Pivot, kembali ke tahap 6 dan periksa **Add this data to the Data Model**. Jika tab Power Pivot sendiri tidak tersedia, lakukan tahap 3 terlebih dahulu.

**Jika tombol Relationships tersedia, gunakan dialog berikut sebagai alternatif Diagram View:**

1. Pilih **Data → Relationships → New**.
2. Untuk relasi contoh pertama, isi:

   | Isian pada dialog | Pilihan |
   |---|---|
   | Table | FactCommercial |
   | Column (Foreign) | event_date |
   | Related Table | DimDate |
   | Related Column (Primary) | date |

3. Klik **OK**. Buat relasi berikutnya dengan **New**.
4. Buka `excel\relationships.csv` di editor teks sebagai daftar kerja. **Setiap baris adalah satu relasi**, dengan pemetaan:
   - `fact` → **Table**.
   - `fact_key` → **Column (Foreign)**.
   - `dimension` → **Related Table**.
   - `dimension_key` → **Related Column (Primary)**.
5. Ulangi sampai semua **28 relasi** selesai. Jangan membuat ulang relasi yang sudah ada.

Kolom pada tabel dimensi harus unik. Jangan membalik sisi tabel. **Jangan membuat relasi tambahan** antara tabel fact, antara dimensi, atau dari DimDate ke FactInvoiceSnapshot/FactRFM/FactCohort.

**Periksa:** semua relasi dari CSV tercantum; tidak ada peringatan key duplikat atau tipe data berbeda. Jika Excel sudah membuat relasi otomatis, cocokkan dengan CSV sebelum menghapus atau menambah apa pun.

## 9. Tetapkan tabel kalender dan periksa tipe data

1. Pilih **Power Pivot → Manage**, kemudian tab tabel **DimDate**.
2. Pilih **Design → Mark as Date Table → Mark as Date Table**, lalu kolom **date**. Letak menu dapat sedikit berbeda menurut versi Excel.
3. Periksa tipe kolom di Power Pivot:
   - ID/nama → **Text**.
   - Tanggal → **Date**, bukan teks.
   - Jumlah/unit → **Whole Number**.
   - Nilai uang → **Currency / Fixed Decimal Number** sesuai pilihan yang tersedia.
4. Jangan mengganti isi tanggal dengan teks atau mengubah uang menjadi string berawalan `Rp`.

**Periksa:** tanggal tersusun kronologis dan nilai uang bisa dijumlahkan. Jika tipenya salah, periksa query sumber sebelum mengubah rumus.

## 10. Masukkan measures DAX

**Jalur yang sekarang dipilih: pemasangan otomatis sisa measures lewat VBA.** Ikuti [panduan VBA sekali jalan](panduan_vba_measures.md), lalu kembali ke tahap 11. Bagian manual di bawah tetap tersedia sebagai rujukan. Measures baru disimpan di **ReportContext**, bukan FactCommercial, untuk menghindari konflik nama measure dan kolom. Helper yang sudah dibuat sebelumnya tetap boleh berada di FactCommercial.

Measure adalah rumus yang dihitung ulang mengikuti filter PivotTable. Measures berbeda dari kolom sumber dan tidak dibuat di sel worksheet biasa.

### 10A. Buat satu measure pertama dahulu

Untuk langkah ini kembali ke **jendela Excel utama yang menampilkan worksheet**, bukan Power Query Editor atau jendela Power Pivot yang menampilkan diagram relasi.

1. Klik tab **Power Pivot** pada ribbon Excel utama.
2. Klik **Measures → New Measure**. Pada ribbon sempit, tombol ini bisa berada dalam grup **Calculations**.
3. Akan muncul dialog dengan kotak **Table name**, **Measure name**, **Description**, dan **Formula**. Isi sebagai berikut:

   | Kotak | Isi yang harus dipilih/diketik |
   |---|---|
   | Table name | Pilih `ReportContext` dari dropdown untuk measure baru |
   | Measure name | Ketik `_sales_lines` |
   | Description | Boleh dikosongkan |
   | Formula | Tempel rumus di bawah, dengan tepat satu tanda `=` di awal |

   ```dax
   =CALCULATE(COUNTROWS(FactCommercial), FactCommercial[event_type] = "Sale") + 0
   ```

4. Jika kotak Formula sudah berisi `=`, hapus isi kotaknya dahulu sebelum menempelkan rumus lengkap. Hasilnya tidak boleh diawali `==`.
5. Pada **Formatting Options**, pilih **Number**, lalu **Decimal places: 0**. Biarkan pemisah ribuan sesuai pengaturan Excel.
6. Klik **Check Formula** jika tersedia, kemudian **OK**.
7. Pilih **Power Pivot → Measures → Manage Measures**. Pastikan `_sales_lines` tercantum. Untuk helper yang sudah dibuat sebelumnya, Table Name FactCommercial tetap benar; untuk yang baru gunakan ReportContext. Jangan membuat nama yang sama dua kali.

**Berhasil jika:** nama `_sales_lines` tercantum dalam Manage Measures tanpa error. Tidak muncul angka baru di worksheet adalah normal; measure baru akan ditampilkan melalui PivotTable pada tahap 11. Jangan membuat kolom baru bernama `_sales_lines` pada tabel.

Cara ini mengikuti [petunjuk pembuatan measure dari Microsoft](https://support.microsoft.com/en-us/excel/create-a-measure-in-power-pivot).

### 10B. Pahami cara memisahkan nama dan formula

Satu baris pada berkas measures.dax berbentuk:

```dax
_returned_cogs := SUM(FactCommercial[returned_cogs]) + 0
```

Bagian **sebelum `:=`** dimasukkan ke kotak **Measure name**. Bagian **sesudah `:=`** dimasukkan ke kotak **Formula**, dengan tambahan `=` di awal. Jadi pengisiannya:

| Measure name | Formula | Format |
|---|---|---|
| `_returned_cogs` | `=SUM(FactCommercial[returned_cogs]) + 0` | Number, 2 desimal |
| `_cogs_sold` | `=SUM(FactCommercial[cogs_sold]) + 0` | Number, 2 desimal |
| `_snapshot_supported` | `=NOT ISCROSSFILTERED(DimProduct[sku_id])` | General |

Buat setiap baris sebagai **New Measure yang terpisah**; Table name untuk measures baru adalah **ReportContext**. Table name adalah tempat menyimpan measure; rumus tetap dapat membaca tabel lain. Jangan pindahkan helper lama yang sudah benar hanya untuk menyeragamkan lokasinya.

Awalan `_` adalah bagian nama dan harus disertakan. Measures berawalan `_` merupakan rumus pembantu yang dipakai KPI lain, jadi tetap harus dibuat. Hasil helper dapat berupa angka atau TRUE/FALSE; itu normal. Jangan menyembunyikannya sebelum pemeriksaan selesai.

### 10C. Lanjutkan seluruh definisi secara berurutan

Jika `_sales_lines` sudah dibuat, mulai dari baris `_returned_cogs` dalam berkas. Jangan membuat `_sales_lines` untuk kedua kalinya. Gunakan langkah berikut untuk melanjutkan:

1. Buka `excel\measures.dax` di editor teks.
2. Masukkan definisi **dari atas ke bawah**, termasuk helper yang namanya diawali `_`; lewati measure yang sudah dibuat pada 10A–10B. Lewati baris komentar `//` dan baris kosong.
3. Di Excel pilih **Power Pivot → Measures → New Measure**.
4. Pilih **Table name: ReportContext** sebagai tempat penyimpanan measures baru.
5. Contoh untuk baris `_returned_cogs := SUM(FactCommercial[returned_cogs]) + 0`:
   - **Measure name:** `_returned_cogs`
   - **Formula:** `=SUM(FactCommercial[returned_cogs]) + 0`
6. Klik **Check Formula** jika tersedia. Jika tidak ada error, klik **OK**.
7. Ulangi untuk setiap definisi. Contoh `net_sales := [billed_sales] - [return_amount]` berarti nama `net_sales` dan formula `=[billed_sales] - [return_amount]`.

Jangan menempelkan seluruh berkas sekaligus ke satu measure. Jangan menyertakan `nama :=` dalam kotak Formula. Jika measure merujuk nama yang belum ada, buat dependensinya terlebih dahulu. Jika Excel menolak pemisah argumen, periksa apakah locale Excel memerlukan `;` sebagai pengganti `,`; jangan mengganti tanda desimal pada data sumber.

Atur format melalui dialog measure:

| Jenis | Format tampilan |
|---|---|
| Uang IDR dan average_order_value | Angka 2 desimal; satuan IDR ditulis pada judul |
| Rate/rasio | Percentage, 2 desimal |
| Hari dan calls_per_salesman_day | Angka 2 desimal |
| Count/unit | Angka tanpa desimal |

**Periksa:** seluruh 46 KPI, helper, dan measures diagnostik berhasil dibuat tanpa error. Nilai kosong/BLANK tidak boleh dipaksa menjadi nol. Tulis keterangan **Blank = N/A / filter tidak didukung** pada dashboard.

## 11. Buat PivotTable pertama dan uji model

1. Pilih sel **J30** pada worksheet **Data & Calculations**.
2. Pilih **Insert → PivotTable → From Data Model**. Jika pilihan itu tidak terlihat, buka dialog PivotTable dan pilih **Use this workbook's Data Model**.
3. Pilih **Existing Worksheet** dan lokasi `'Data & Calculations'!$J$30`, lalu **OK**.
4. Pada **PivotTable Analyze → PivotTable Name**, beri nama `pControls`.
5. Dari daftar field pilih measures `gross_sales`, `billed_sales`, `return_amount`, `net_sales`, `gross_profit`, dan `price_bridge_residual` ke area **Values**.
6. Jika semua nilai melebar ke kanan, seret field **Σ Values / Values** dari **Columns** ke **Rows**. Hasilnya daftar metrik dua kolom.
7. Tambahkan `DimDate[year]` ke area **Filters**, lalu pilih **2025**.

Untuk ekspor awal, angka yang diharapkan:

| Measure | Nilai SQL pembanding, IDR |
|---|---:|
| net_sales | 15.762.528.589,51 |
| gross_profit | 3.851.071.463,51 |
| price_bridge_residual | 0,00 |

Angka ini hanya pembanding. Jangan mengetiknya ke PivotTable sebagai pengganti hasil perhitungan. Tanpa filter 2025, model mencakup 2024–2025 sehingga totalnya berbeda.

**Periksa:** daftar field memuat banyak tabel model dan ikon measure; nilai berasal dari model, bukan range worksheet. Jika angka tidak cocok, periksa filter tahun, relasi, dan rumus sebelum membuat grafik.

## 12. Susun dashboard dan slicer

**Opsi otomatisasi yang sekarang disetujui:** ikuti [panduan VBA dashboard bertahap](panduan_vba_dashboard.md). Modul mengaudit dan mempertahankan objek lama, melengkapi dashboard, serta merekam hasil native untuk rekonsiliasi. Langkah manual di bawah tetap menjadi referensi. Jangan menjalankan pembangunan manual dan VBA bersamaan.

Buat PivotTable tambahan dari Data Model, dengan pola langkah bagian 11. Nama field ditulis dalam format `tabel[kolom]`.

| Worksheet | Baris PivotTable | Values |
|---|---|---|
| Executive Summary | DimDate[year_month] | net_sales, target_net_sales |
| Executive Summary | DimBranch[region_name], DimBranch[branch_name] | net_sales, target_gap, achievement_rate |
| Leakage Diagnostics | DimBranch[branch_name] | reference_sales, price_upside, price_shortfall, discount_amount, billed_sales, return_amount, net_sales, gross_profit |
| Leakage Diagnostics | DimProduct[sku_name] | price_shortfall, return_amount, return_rate, billed_sales |
| Leakage Diagnostics | DimOutlet[outlet_name] | price_shortfall, return_amount, return_rate, billed_sales |
| Leakage Diagnostics | FactCommercial[return_reason] | return_amount, return_margin_reversal; keluarkan item kosong |
| Operations | DimBranch[branch_name] | outstanding_ar, overdue_ar, overdue_rate, dso_90d |
| Operations | DimSalesman[salesman_name] | completed_visits, visit_completion_rate, effective_call_rate, order_conversion_rate, calls_per_salesman_day |
| Operations | FactRFM[segment] | rfm_outlets, rfm_monetary |
| Operations | FactCohort[cohort_month] | retention_rate; masukkan FactCohort[activity_month] ke Columns |

Tambahkan seluruh 46 KPI dan ketiga residual ke `pControls` untuk pemeriksaan lengkap. Untuk cohort, matikan grand total yang menjumlahkan beberapa activity month melalui pengaturan **Design → Grand Totals**; total lintas bulan memang tidak boleh dipakai sebagai retention rate.

Untuk grafik: klik PivotTable tren → **PivotTable Analyze → PivotChart**, pilih grafik kolom atau garis. Tampilkan net_sales dan target_net_sales terhadap tahun-bulan. Buat kartu ringkasan dengan mengetik `=` pada sel output lalu mengeklik nilai PivotTable, sehingga Excel dapat membuat referensi GETPIVOTDATA; jangan menyalin angka sebagai nilai tetap.

Untuk slicer:

1. Klik salah satu PivotTable model → **PivotTable Analyze → Insert Slicer**.
2. Pilih field dimensi: `DimDate[year]`, `DimDate[year_month]`, `DimBranch[region_name]`, `DimBranch[branch_name]`, `DimProduct[category]`, `DimProduct[sku_name]`, `DimOutlet[outlet_id]`, `DimSalesman[salesman_name]`, dan `DimChannel[channel]`.
3. Klik kanan slicer → **Report Connections / PivotTable Connections**. Centang PivotTable model yang harus mengikuti pilihan tersebut, termasuk `pControls`.
4. Salin slicer yang sudah terhubung jika ingin menaruhnya di worksheet lain. Uji apakah pilihannya tersinkron.
5. Gunakan area **N:U** untuk slicer, **A:L** untuk isi dashboard; tempatkan cohort lebih bawah, mulai sekitar **A110**, agar tabel lebar tidak bertabrakan.
6. Hapus filter tahun lokal dari pControls yang dibuat saat bagian 11, kemudian gunakan slicer tahun bersama. Pilih **2025** sebagai tampilan awal.

**Periksa:** memilih cabang mengubah hasil dan grafik, lalu tombol **Clear Filter** mengembalikan hasil semula. Jika satu PivotTable tidak berubah, periksa Report Connections dan sumber Data Model-nya.

## 13. Isi Methodology dan pahami hasil kosong

Cantumkan periode default 1 Januari–31 Desember 2025 dan snapshot 31 Desember 2025. Tulis aturan berikut agar pembaca tidak salah menafsirkan dashboard:

- Net sales = billed sales dikurangi retur pada tanggal event retur. Jangan mengurangi retur dua kali.
- Price shortfall, seluruh diskon, dan retur adalah komponen diagnostik. Diskon promo invalid sudah termasuk diskon; jangan dijumlahkan lagi sebagai tambahan kerugian.
- AR dan RFM memakai snapshot tetap, terpisah dari periode operasi. Perubahan slicer tahun saja tidak mengubah snapshot AR/RFM.
- Target hanya tersedia pada grain cabang-bulan penuh. Target/achievement kosong untuk filter SKU, outlet, salesman, channel, serta tanggal parsial atau tidak berurutan.
- Filter produk membuat AR dan kunjungan kosong; filter produk/salesman membuat RFM dan cohort kosong karena grain tersebut tidak didukung.
- Rasio memakai komponen total, bukan rata-rata persentase anggota. Rasio tanpa denominator tetap kosong.
- Cohort bulan nol adalah aktivasi; tidak ada nilai nol untuk bulan masa depan yang belum diamati.

Tambahkan rujukan ke kpi_dictionary.md, core_analysis.md, optimization_report.md, serta manifest ekspor. Jangan menyatakan semua indikator sebagai satu nilai kerugian yang pasti dapat dipulihkan.

## 14. Refresh dan verifikasi native Excel

1. Pilih **Data → Refresh All**, tunggu semua proses selesai, lalu periksa **Queries & Connections**. Pastikan tidak ada error.
2. Cocokkan seluruh 46 KPI dengan SQLReference pada periode default. Bandingkan uang 2 desimal, rasio 2 desimal persen, hari/produktivitas 2 desimal, dan count secara tepat.
3. Ubah slicer tahun menjadi 2024, lalu kembali ke 2025. Catat hasil sebelum/sesudah dan pastikan grafik ikut berubah.
4. Pilih **Medan/B003** pada 2025. Price shortfall pembanding adalah **IDR 3.023.132,22**. Kembalikan filter ke semua cabang.
5. Uji kombinasi pada `data\excel\SQLFilterControls.csv` dan `SQLSupportControls.csv`. Kolom filter kosong berarti semua anggota. Untuk tanggal parsial, gunakan filter tanggal pada `DimDate[date]`, bukan hanya slicer bulan.
6. Uji pFolder: salin folder CSV ke lokasi lokal lain, ubah nilai pFolder melalui **Data → Get Data → Launch Power Query Editor**, pilih pFolder, ganti nilainya, lalu refresh. Angka harus tetap sama. Kembalikan lokasi asal dan refresh lagi.
7. Uji folder yang tidak ada: refresh harus gagal dengan pesan sumber tidak ditemukan. Kembalikan path yang benar dan pastikan refresh berhasil lagi.
8. Simpan workbook, tutup, buka kembali, refresh, lalu ulangi filter Medan/2025 dan Jakarta/kuartal I. Pastikan relasi, measures, dan sambungan slicer tetap berfungsi.
9. Periksa setiap worksheet pada zoom normal: tidak ada `####`, angka terpotong, grafik kosong, atau objek bertumpuk.
10. Catat hasil nyata dan lokasi screenshot pada `docs\excel\native-execution.md`. Daftar pengujian teknis lengkap ada di `docs\excel\native-verification.md`.

Untuk pencatatan angka, salin `docs\excel\native-results.template.csv` menjadi `native-results.csv` pada folder yang sama dan isi **actual_value dari hasil Excel**, bukan dari SQL. Gunakan angka tanpa pemisah ribuan/simbol mata uang; contoh `15762528589.51`. Persentase disimpan sebagai pecahan: 6,98% → `0.0698`. Pertahankan kosong untuk BLANK yang sebenarnya. Pemeriksaan CSV/capture dapat dijalankan dengan perintah pada dokumen teknis excel_model.md; pemeriksaan itu sendiri tidak membuktikan refresh Excel.

**Phase 4 belum selesai sampai seluruh pengujian native berhasil.** Jika gagal, catat pesan error dan tahapnya; jangan menandai fase selesai, membuat commit penyelesaian, atau mengarsipkan status fase.

## Jika menemui kendala

| Pesan/gejala | Tindakan |
|---|---|
| `The name 'pFolder' wasn't recognized` | Pastikan query parameter sudah dibuat dan namanya tepat, tanpa `.m` |
| `fnLoadCsv` tidak dikenali | Buat fungsi sebelum query tabel; periksa huruf besar/kecil namanya |
| File/folder tidak ditemukan | Periksa nilai pFolder dan keberadaan CSV; pastikan berkas OneDrive tersedia secara lokal |
| Schema mismatch | Jangan menghapus kolom atau mengabaikan error; pastikan M dan CSV berasal dari paket ekspor yang sama |
| Duplicate/invalid primary key | Catat tabel dan pesan; jangan menghapus duplikat secara manual untuk meloloskan model |
| Tabel muncul di worksheet, bukan model | Klik kanan query → Load To… → Only Create Connection + Add this data to the Data Model |
| Measure tidak mengenali tabel/kolom | Cocokkan nama query dengan nama berkas dan periksa dependensi measure |
| Slicer tidak memengaruhi pivot tertentu | Periksa Report Connections, filter lokal yang tersisa, dan sumber PivotTable |
| Target/AR tertentu kosong setelah filter | Periksa aturan grain pada bagian 13 sebelum menganggapnya error |
| Angka awal jauh berbeda | Pastikan tahun 2025 dipilih; model awal memuat riwayat 2024–2025 |

Saat melaporkan kendala, sebutkan **nomor tahap, nama query/tabel/measure, serta teks error lengkap** agar perbaikannya bisa langsung ditentukan.
