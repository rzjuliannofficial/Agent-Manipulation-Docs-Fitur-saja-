# Office Document Automation (PDF, DOCX, XLSX)

Repositori ini berisi sekumpulan perkakas dan skrip utilitas untuk melakukan ekstraksi teks dari PDF/Word serta pembuatan dokumen Word dan Excel secara otomatis sesuai dengan template standar industri dan akademik.

---

## Struktur Direktori

Proyek ini disusun secara modular untuk memisahkan logika pembacaan, pembuatan, berkas referensi, dan hasil output:

| Nama Folder | Fungsi & Deskripsi |
| :--- | :--- |
| **`pembaca/`** | Berisi skrip ekstraksi informasi dan pembaca berkas PDF atau Word. |
| **`pembuat/`** | Berisi logika pembentukan/pembuatan dokumen Excel baru dan dokumen Word berbasis template. |
| **`referensi/`** | Berkas panduan format, PDF artikel asli, dan data master rumus/beasiswa. |
| **`output/`** | Tempat penyimpanan berkas hasil generasi akhir (.docx dan .xlsx) siap pakai. |
| **`scratch/`** | Ruang kerja sementara untuk analisis data rumus, pengujian formula, dan skrip bantuan debugging. |

---

## Persyaratan Sistem

Pastikan lingkungan lokal Anda telah terinstal modul-modul berikut:

### 1. Python (Versi >= 3.8)
Instal dependensi Python untuk menangani pembacaan dokumen dan manipulasi Excel:
```bash
pip install python-docx openpyxl PyPDF2
```

### 2. Node.js (Versi >= 18)
Digunakan untuk beberapa modul builder berbasis JavaScript:
```bash
npm install
```

---

## Panduan Penggunaan

### 1. Ekstraksi Dokumen (Pembaca)

Untuk mengekstrak isi teks dari file PDF:
```bash
python pembaca/read_pdf.py referensi/3147-Article.pdf -o output/extracted_pdf.txt
```

Untuk mengekstrak paragraf dan tabel dari file Word:
```bash
python pembaca/read_docx.py referensi/artikel_kelompok_6.docx -o output/extracted_text.txt
```

### 2. Pembuatan Dokumen (Pembuat)

Untuk memformat lembar kerja Excel baru menggunakan data master:
```bash
python pembuat/create_excel.py
```

Untuk menghasilkan draf dokumen ilmiah dengan format standar JIP (Jurnal Informatika Polinema):
```bash
python pembuat/create_jip_docx.py -o output/Jurnal_JIP_Final.docx
```

---

## Pengembangan Lebih Lanjut
Semua penambahan skrip pengujian baru, eksplorasi formula matematika Excel, atau pembuangan logika sementara wajib dilakukan di dalam direktori `scratch/` agar root proyek tetap bersih dan teratur.