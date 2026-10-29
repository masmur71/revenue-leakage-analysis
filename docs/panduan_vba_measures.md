# Memasang sisa measures dengan satu kali menjalankan VBA

Modul sudah disiapkan di **excel\MeasureInstaller.bas**. Anda menjalankannya melalui workbook pembantu. Hasil utama tetap **RevenueLeakage.xlsx** tanpa macro. Modul belum dijalankan pada workbook Anda; pemasangan native dan hasil KPI masih perlu diperiksa.

## Jika sebelumnya berhenti pada “Save MeasureInstaller.xlsm in the project's local excel folder first”

Excel dapat mengembalikan alamat URL OneDrive meskipun berkas dibuka dari folder sinkronisasi lokal. Modul sudah diperbaiki agar lokasi sumber dapat dipilih langsung. **Kode yang sudah diimpor ke workbook tidak ikut berubah ketika berkas .bas diperbarui**; ganti modul lama terlebih dahulu:

1. Buka MeasureInstaller.xlsm → **Alt+F11** → **Ctrl+R**.
2. Pada **VBAProject (MeasureInstaller.xlsm) → Modules**, klik kanan **MeasureInstaller → Remove MeasureInstaller…**. Jika ditanya apakah ingin mengekspor modul lama, pilih **No**. Ini hanya melepas kode installer lama dari runner; tabel dan measures pada RevenueLeakage.xlsx tidak dihapus.
3. Pilih **File → Import File…**, impor kembali **excel\MeasureInstaller.bas** yang sudah diperbarui. Jangan mengimpor dua salinan modul dengan macro bernama sama.
4. Simpan runner, kembali ke Excel, simpan RevenueLeakage.xlsx, lalu **Alt+F8 → InstallMeasures → Run**.
5. Jika muncul pemilih folder, pilih folder lokal berikut lalu klik **OK/Select Folder**:

   ```text
   C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis\excel
   ```

6. Folder yang dipilih wajib memuat **measures.dax** dan **RevenueLeakage.xlsx**. Jika dialog dibatalkan, installer berhenti tanpa mengubah model.
7. Jika alamat workbook utama juga berupa URL, akan muncul **Confirm cloud target** yang menampilkan alamat workbook dan folder sumber. Pilih **Yes** hanya bila itu workbook RevenueLeakage.xlsx berisi model Phase 4 Anda; pilih **No** jika ragu. Identitas cloud tidak disamakan dengan file lokal hanya berdasarkan namanya. Setelah dikonfirmasi, pengecekan nama, format .xlsx, status tersimpan/read-only, tabel, dependensi, konflik formula, dan backup tetap dijalankan.

Selanjutnya ikuti bagian 5 untuk membaca hasil dan melakukan percobaan ulang.

## 1. Simpan pekerjaan yang sudah selesai

**Jika percobaan lama berhenti dengan `CREATE MEASURE ... == ...`:** masalah ada pada versi installer lama yang mengirim tanda `=` tambahan ke API native. Berkas MeasureInstaller.bas sudah diperbaiki; formula measures.dax dan formula yang dimasukkan melalui dialog Excel tetap memakai aturan semula. Ganti modul lama dengan berkas terbaru menggunakan langkah Remove/Import di atas, lalu jalankan kembali. Jangan mengubah 66 rumus sumber atau menghapus helper yang sudah berhasil dibuat. Jika pemasangan sebelumnya berhenti di tengah, measures yang sudah sesuai akan dilewati saat percobaan ulang.

**Jika berhenti pada ar_1_30 dengan potongan `>= 1. FactInvoiceSnapshot`:** log native memperlihatkan koma pemisah argumen setelah angka dibaca sebagai tanda desimal. Installer terbaru menyesuaikan pemisah argumen/angka dengan locale Excel, tanpa mengubah pengaturan sistem atau rumus sumber. Literal teks dan nama kolom tidak diubah. Simpan workbook untuk mempertahankan measures yang sudah berhasil, ganti modul dengan versi terbaru, lalu jalankan ulang. Periksa catatan LOCALE dan ADD_REQUEST pada log jika masih gagal.

1. Buka workbook **RevenueLeakage.xlsx** yang sudah berisi model tahap 1–9.
2. Tekan **Ctrl+S**. Jangan menutup workbook tersebut.
3. Pastikan workbook tersimpan pada folder ini:

   ```text
   C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis\excel
   ```

Jangan membuat ulang tabel atau measures yang sudah ada. `_snapshot_supported` adalah ejaan yang benar. Jika sebelumnya Anda membuat `_shapshot_supported`, modul mencatatnya sebagai peringatan, mempertahankannya, dan tetap menambahkan nama yang benar bila belum tersedia.

## 2. Buat workbook pembantu

1. Di Excel pilih **File → New → Blank workbook**.
2. Pilih **File → Save As → Browse**.
3. Pilih folder **excel** yang sama dengan RevenueLeakage.xlsx.
4. Isi nama **MeasureInstaller.xlsm**.
5. Pada **Save as type**, pilih **Excel Macro-Enabled Workbook (*.xlsm)**, lalu **Save**.

**Periksa:** sekarang ada dua workbook terbuka: RevenueLeakage.xlsx dan MeasureInstaller.xlsm. Jangan menyimpan workbook utama sebagai .xlsm. Workbook pembantu harus berada di samping measures.dax, bukan di Documents atau folder Downloads.

## 3. Impor modul VBA

1. Dengan MeasureInstaller.xlsm terbuka, tekan **Alt+F11** untuk membuka **Visual Basic Editor**.
2. Jika panel **Project Explorer** di kiri tidak terlihat, tekan **Ctrl+R**.
3. Cari dan klik **VBAProject (MeasureInstaller.xlsm)**. Pastikan bukan proyek RevenueLeakage.xlsx.
4. Pilih **File → Import File…**. Pada Visual Basic Editor, jangan memakai menu Open milik Excel utama.
5. Pilih berkas:

   ```text
   C:\Users\masmu\OneDrive\Documents\PersProject\leakageanalysis\excel\MeasureInstaller.bas
   ```

6. Klik **Open**. Pada proyek MeasureInstaller.xlsm akan muncul **Modules → MeasureInstaller**.
7. Klik dua kali modul itu. Pilih **Debug → Compile VBAProject**. Jika muncul pesan compile error, berhenti dan kirim pesan serta baris yang disorot; jangan lanjut menjalankan macro.
8. Tekan **Ctrl+S**, lalu **Alt+F11** untuk kembali ke Excel.

Tidak perlu menambah reference melalui Tools → References atau mengaktifkan akses ke VBA project object model. Modul tidak memasang dirinya sendiri ke workbook utama.

## 4. Jalankan pemasangan

1. Pastikan RevenueLeakage.xlsx masih terbuka dan sudah disimpan.
2. Tekan **Alt+F8**.
3. Pada **Macros in**, pilih **MeasureInstaller.xlsm**, atau pilih **All Open Workbooks** jika diperlukan.
4. Pilih **InstallMeasures** milik MeasureInstaller.xlsm, lalu **Run**.
5. Tunggu dialog hasil. Jangan mengedit workbook atau menjalankan refresh bersamaan.

Modul mengikat target lokal berdasarkan **nama dan path lengkap**, bukan workbook yang sedang aktif. Untuk workbook yang dilaporkan Excel sebagai URL OneDrive/SharePoint, modul meminta konfirmasi identitas workbook yang terbuka sebelum melanjutkan. Jadi memilih workbook pembantu pada layar tidak mengubah target pemasangan. Folder sumber lokal dipilih melalui dialog jika lokasi runner dilaporkan sebagai URL atau tidak memuat berkas sumber yang diperlukan.

Sebelum menambah measure, modul akan:

- Membaca 66 definisi pada measures.dax, memeriksa duplikat, tabel/kolom, dan urutan dependensi.
- Membandingkan seluruh measures yang sudah ada. Formula yang sesuai dilewati; formula berbeda menghentikan proses **sebelum ada penambahan**.
- Membuat salinan workbook pada **excel\measure-backups**. Jika backup gagal, penambahan tidak dimulai.
- Menambah hanya measures yang belum tersedia, lalu memeriksa ulang nama dan formula dari model Excel.

**Perubahan tempat penyimpanan measure:** measures baru ditempatkan pada tabel **ReportContext**, agar nama seperti gross_sales tidak bertabrakan dengan kolom FactCommercial. Nama/formula KPI tidak berubah. Helper lama pada FactCommercial tetap dipertahankan; tidak perlu dipindahkan atau dibuat ulang. Penggunaan ReportContext sebagai tempat penyimpanan tidak membuat measure mengikuti filter snapshot yang baru; semantik rumus tetap ditentukan oleh ekspresi DAX dan relasi model.

Format measures baru: rasio 2 desimal persen, count tanpa desimal, uang/hari/produktivitas 2 desimal, helper logika TRUE/FALSE. Format measures lama yang dilewati tetap dipertahankan; periksa format tersebut pada tahap PivotTable.

## 5. Baca hasil dan simpan

Jika berhasil, dialog menampilkan jumlah **installed**, **skipped**, serta pesan bahwa semua **66 expected names/formulas** tersedia.

1. Beralih ke RevenueLeakage.xlsx dan tekan **Ctrl+S**. Modul tidak menyimpan perubahan target secara otomatis.
2. Pilih **Power Pivot → Measures → Manage Measures**. Pastikan measures baru tersedia; kebanyakan berada pada ReportContext.
3. Jalankan **InstallMeasures** lagi melalui Alt+F8. Hasil kedua yang diharapkan: **installed 0**, **skipped 66**. Measure lain seperti salah eja yang dipertahankan dapat membuat total seluruh model lebih dari 66; yang diperiksa adalah 66 nama yang benar dari sumber.
4. Simpan workbook jika diperlukan, kemudian lanjut ke **tahap 11** pada panduan pemasangan untuk uji PivotTable dan rekonsiliasi SQL.

Log setiap percobaan berada pada **excel\measure-install-logs**, dengan timestamp pada nama berkas. CSV memuat kejadian, nama measure, pesan, dan formula aktual bila tersedia. Salinan `.source.dax` di folder yang sama merekam definisi yang dipakai saat percobaan itu.

## Jika pemasangan berhenti

| Pesan/tahap | Tindakan |
|---|---|
| Save target first / target belum disimpan | Simpan RevenueLeakage.xlsx dengan Ctrl+S, lalu jalankan lagi |
| Target tidak terbuka atau path lokal berbeda | Buka RevenueLeakage.xlsx yang dimaksud; path lokal target harus sesuai folder sumber yang dipilih |
| Runner tetap terbaca sebagai URL OneDrive | Impor ulang modul terbaru dan gunakan dialog pemilih folder; bila target juga berupa URL, periksa dialog Confirm cloud target |
| Required model table missing / kolom hilang | Periksa tahap 6; tabel harus benar-benar masuk Data Model dengan nama tepat |
| Formula conflict | Baca CONFLICT_ACTUAL dan CONFLICT_EXPECTED pada log; kirim log untuk diperiksa, jangan mengganti rumus secara tebakan |
| Backup gagal | Periksa lokasi/izin penulisan folder; modul belum menambah measures |
| Native Add/format gagal di tengah | Penambahan sebelumnya mungkin masih ada; kirim tahap, measure, pesan, dan log. Setelah diperbaiki, jalankan ulang untuk melanjutkan tanpa duplikasi |
| Macro diblokir Excel/kebijakan organisasi | Catat pesan yang terlihat dan laporkan; jangan mengaktifkan seluruh macro atau mengubah kebijakan keamanan global |

Modul tidak menghapus atau menimpa measures lama. Jika perlu memulihkan backup, simpan/tutup workbook terlebih dahulu, pertahankan file hasil percobaan secara terpisah, lalu buka salinan backup yang log-nya mencatat BACKUP_OK. Jangan menimpa workbook terbuka.

**Berhasil memasang measures belum berarti Phase 4 selesai.** Refresh, respons filter, angka SQL, dan save/close/reopen masih harus diuji di Excel. Jangan menandai checklist native hanya berdasarkan pengujian parser.

## Dasar teknis dan batas verifikasi

Pemasangan menggunakan [ModelMeasures.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.modelmeasures.add); backup menggunakan [SaveCopyAs](https://learn.microsoft.com/en-us/office/vba/api/excel.workbook.savecopyas). Pemisahan tempat measure/kolom mengikuti [aturan penamaan DAX](https://learn.microsoft.com/en-us/dax/best-practices/dax-column-measure-references).

Pengujian otomatis mencakup inti parser/normalisasi/dependensi dan klasifikasi path lokal/cloud melalui Windows Script Host dengan mock koleksi/skema, tanpa membuka Excel. Jumlah dan hasil terakhir tercatat pada docs/excel/measure-installer-tests.txt. Itu bukan kompilasi VBA atau eksekusi ModelMeasures.Add native. Dialog folder/konfirmasi cloud serta dua percobaan pemasangan di atas tetap perlu diuji pada Excel Anda.
