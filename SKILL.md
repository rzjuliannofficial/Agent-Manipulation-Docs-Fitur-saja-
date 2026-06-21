---
name: office_document_handling
description: "Use this skill whenever the user wants to read PDF files, create or edit Excel (.xlsx) spreadsheets, or create, read, edit, or manipulate Word documents (.docx files). Triggers include requests to read/extract text from PDFs, create/update spreadsheets, generate reports, memos, lists, tables of contents, or formatted Word documents. Do NOT use for general coding tasks unrelated to document generation or PDF reading."
license: Proprietary. LICENSE.txt has complete terms
---

# Document Manipulation (PDF, DOCX, XLSX)

## Overview & Project Folder Structure

This workspace is strictly organized into clean, structured directories to keep the project root clean and maintainable. All files must be placed in their appropriate folders according to this structure:

1. **`referensi/`**: Contains reference files, PDFs, original guideline documents, and the master Excel file (`referensi/SWARA_MAIRCA_DETAIL_RUMUS.xlsx`).
2. **`pembaca/`**: Contains utility reading scripts (e.g., `pembaca/read_pdf.py`, `pembaca/read_docx.py`).
3. **`pembuat/`**: Contains generation and creation scripts (e.g., `pembuat/create_excel.py`, `pembuat/create_docx.py`, `pembuat/create_jip_docx.py`, `pembuat/build._jip.js`).
4. **`scratch/`**: Contains temporary scripts, log files, dump scripts, and analysis files (e.g., `scratch/fix_excel_references.py`, `scratch/calculate_all.py`, formula text outputs). All temporary work, testing, and debugging files **MUST** be placed in this folder.
5. **`output/`**: Contains final production artifacts, generated spreadsheets, and Word documents (`output/Jurnal_JIP_SWARA_MAIRCA_Final_1Kolom.docx`, `output/Jurnal_JIP_SWARA_MAIRCA_Final_2Kolom.docx`, `output/SWARA_MAIRCA_DETAIL_RUMUS.xlsx`, etc.). Whenever the master Excel file in `referensi/` is updated, a copy of the result **MUST** be placed here as the final output.

---


## 1. Reading PDF Files

To extract text content from a PDF file:

```bash
# Print PDF text content directly to console
py pembaca/read_pdf.py path/to/document.pdf

# Save PDF text content to a text file
py pembaca/read_pdf.py path/to/document.pdf -o output_text.txt
```

---

## 2. Reading Word Documents (.docx)

To extract text and tables from a Word file:

```bash
# Print DOCX text content directly to console
py pembaca/read_docx.py path/to/document.docx

# Save DOCX text content to a text file
py pembaca/read_docx.py path/to/document.docx -o output_text.txt
```

## 3. Excel Spreadsheets (.xlsx)

To create or edit Excel workbooks, use the `openpyxl` Python library.

### Basic Excel Generation Example
Run the utility template:
```bash
py pembuat/create_excel.py -o output/sales_report.xlsx
```

### Scripting with `openpyxl`
```python
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook()
ws = wb.active
ws.title = "Data Sheet"

# Always enable gridlines explicitly
ws.views.sheetView[0].showGridLines = True

# Styling
title_font = Font(name="Segoe UI", size=14, bold=True)
ws["A1"] = "Styled Sheet"
ws["A1"].font = title_font

wb.save("output.xlsx")
```

---

## 4. Word Documents (.docx)

To generate new Word documents, use the `python-docx` library or JavaScript-based `docx-js`.

### Generating with Python (`python-docx`)
Run the utility template:
```bash
py pembuat/create_docx.py -o output/progress_report.docx
```

### Scripting with `python-docx`
```python
from docx import Document
from docx.shared import Inches, Pt

doc = Document()

# Set page margins
for section in doc.sections:
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)

# Add heading and paragraph
doc.add_heading("Project Report", level=1)
doc.add_paragraph("This is normal paragraph text formatted in Arial.")

doc.save("output.docx")
```

### Creating with JavaScript (`docx-js`)
If JavaScript is preferred, install globally: `npm install -g docx`

```javascript
const { Document, Packer, Paragraph, TextRun } = require('docx');
const fs = require('fs');

const doc = new Document({
  sections: [{
    children: [
      new Paragraph({
        children: [new TextRun("Hello World")]
      })
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => fs.writeFileSync("doc.docx", buffer));
```

---

## XML-Based DOCX Editing (Advanced)

For unpacking, editing raw XML, and repacking existing `.docx` files:

### Step 1: Unpack
```bash
python scripts/office/unpack.py document.docx unpacked/
```
Extracts XML, pretty-prints, and converts smart quotes to XML entities (`&#x201C;` etc.) so they survive editing.

### Step 2: Edit XML
Edit files under `unpacked/word/`.
- Use "Claude" as the author for tracked changes.
- Use XML entities for smart quotes:
  - `&#x201C;` and `&#x201D;` for double quotes (“ ”)
  - `&#x2018;` and `&#x2019;` for single quotes/apostrophes (‘ ’)

### Step 3: Pack
```bash
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```
Validates, condenses XML, and packages it back into a DOCX.

---

## 5. Jurnal Informatika Polinema (JIP) Templates

To generate a fully compliant Word document template based on the Jurnal Informatika Polinema (JIP) format:

```bash
# Generate the JIP template document
py pembuat/create_jip_docx.py -o output/template_jip_filled.docx
```

Guidelines and details are saved in:
* [JIP_TEMPLATE_GUIDELINES.md](file:///c:/laragon/www/Mandiri/Create%20Dokumen/referensi/JIP_TEMPLATE_GUIDELINES.md)

---

## 6. Claude-Style Output & Document Formatting Guidelines

Untuk menjaga kualitas visual, struktur informasi, dan efisiensi komunikasi, agen **MUST** mengikuti standar format berikut saat merespons atau membuat dokumen:

### Gaya Penulisan (Tone & Style)
* **No Fluff**: Hindari salam pembuka, pengantar panjang, atau penutup basa-basi. Langsung berikan hasil/analisis/kode.
* **Scannability**: Gunakan huruf tebal (**bolding**) hanya untuk istilah yang sangat penting, variabel, atau hasil akhir secara taktis. Hindari menebalkan setiap kata kerja/kata sifat secara acak.
* **Tulis Alami (Non-AI)**: Gunakan gaya bahasa mengalir dan organik seperti tulisan manusia/developer profesional. Hindari struktur paragraf/poin kaku yang terlalu terpola atau penataan visual/pewarnaan yang berlebihan yang memberikan kesan "dibuat otomatis oleh robot".
* **Bahasa**: Gunakan bahasa Indonesia formal dan istilah teknis bahasa Inggris secara natural (seperti *array*, *loop*, *endpoint*).

### Format Tampilan Dokumen
* **GitHub Alerts**: Gunakan alert secara strategis untuk informasi tambahan:
  - `> [!NOTE]` untuk info latar belakang/opsional.
  - `> [!TIP]` untuk saran optimasi dan cara efisien.
  - `> [!IMPORTANT]` untuk instruksi kritis yang harus dipatuhi.
* **Diagram Mermaid**: Gunakan diagram alur jika menjelaskan proses/arsitektur logika yang rumit.

### Penyajian Kode & Perubahan
* **Penamaan File**: Tuliskan nama file target di atas blok kode secara jelas.
* **Format Perubahan**: Gunakan blok `diff` jika mengubah kode yang sudah ada untuk memperjelas baris yang dihapus (`-`) dan ditambahkan (`+`).

---

## 7. Kebijakan Placeholder (Template Policy)
* **Placeholder Tabel Data**: Untuk dokumen draf, template, atau jurnal yang digenerate selanjutnya, **SELALU** gunakan format *placeholder dinamis* untuk tabel data/perhitungan besar (seperti alternatif A01-A20 atau matriks perhitungan) alih-alih memasukkan ribuan baris data nyata secara mentah.
* **Contoh Format Placeholder**: Tuliskan representasi ringkas, misalnya baris data cukup diwakili oleh:
  * Alternatif pertama (A01), alternatif kedua (A02), baris placeholder `... [Alternatif A03 - A19] ...`, dan alternatif terakhir (A20).
* **Tujuan**: Untuk menjaga efisiensi ruang dokumen, kebersihan visual draf, dan memastikan dokumen berfungsi optimal sebagai template siap pakai yang mudah dibaca.


