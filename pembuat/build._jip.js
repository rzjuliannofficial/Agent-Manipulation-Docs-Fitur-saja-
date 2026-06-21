const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, LevelFormat, SectionType } = require('docx');
const fs = require('fs');
const readline = require('readline');

// ── JIP Page spec (A4) ──────────────────────────────────────────────────────
// A4: 11906 × 16838 DXA
// Margins: top/left = 30mm = 1701 DXA, bottom/right = 20mm = 1134 DXA
// Content width = 11906 - 1701 - 1134 = 9071 DXA
// Two-col gap = 709 DXA (~12.5mm)
// One column = (9071 - 709) / 2 = 4181 DXA
const PAGE_W = 11906, PAGE_H = 16838;
const MARG_T = 1701, MARG_B = 1134, MARG_L = 1701, MARG_R = 1134;
const CONTENT_W = PAGE_W - MARG_L - MARG_R; // 9071
const COL_GAP  = 709;
const COL_W    = Math.floor((CONTENT_W - COL_GAP) / 2); // 4181

const PAGE_PROPS = {
  size: { width: PAGE_W, height: PAGE_H },
  margin: { top: MARG_T, right: MARG_R, bottom: MARG_B, left: MARG_L,
            header: 851, footer: 851 }
};

const F = "Times New Roman";
function sz(pt) { return pt * 2; } // DXA half-points

// ── Borders ─────────────────────────────────────────────────────────────────
const bThin  = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const bBlk   = { style: BorderStyle.SINGLE, size: 8, color: "000000" };
const bNone  = { style: BorderStyle.NONE, size: 0, color: "auto" };
const bNoneAll = { top:bNone, bottom:bNone, left:bNone, right:bNone };
const bHLine = { style: BorderStyle.SINGLE, size: 8, color: "000000" };

// ── Run helpers ──────────────────────────────────────────────────────────────
function r(text, o={}) {
  return new TextRun({ text, font:F,
    size: o.sz||sz(10), bold:o.b||false, italics:o.i||false,
    color: o.c||"000000",
    underline: o.u?{type:"single"}:undefined });
}

// ── Paragraph helpers ────────────────────────────────────────────────────────
function p(children, o={}) {
  const runs = typeof children==="string" ? [r(children,o)]
    : Array.isArray(children) ? children : [children];
  return new Paragraph({
    children: runs,
    alignment: o.align || AlignmentType.JUSTIFIED,
    spacing: { before: o.before||0, after: o.after||80, line: o.line||240 },
    indent: o.indent ? { firstLine: 720 } : undefined,
    border: o.border || undefined,
  });
}
function blank(after=80) { return new Paragraph({ children:[r("")], spacing:{after} }); }

function secHead(text, num) {
  return new Paragraph({
    children: [ r(`${num}.  ${text}`, { b:true, sz:sz(10) }) ],
    spacing: { before: 160, after: 80 },
    alignment: AlignmentType.LEFT,
  });
}
function subHead(text) {
  return new Paragraph({
    children: [ r(text, { b:true, sz:sz(10) }) ],
    spacing: { before: 120, after: 60 },
  });
}

function tblLabel(text) {
  return new Paragraph({
    children: [r(text, {b:true, sz:sz(8)})],
    alignment: AlignmentType.CENTER,
    spacing: { before:80, after:40 },
  });
}
function placeholder(label) {
  return new Paragraph({
    children: [r(`[Tempatkan Gambar ${label} di Sini]`, {b:true, i:true, sz:sz(10), c:"FF0000"})],
    alignment: AlignmentType.CENTER,
    spacing: { before:120, after:120 },
  });
}
function figLabel(text) {
  return new Paragraph({
    children: [r(text, {sz:sz(8)})],
    alignment: AlignmentType.CENTER,
    spacing: { before:40, after:80 },
  });
}

// ── Section 1: Single column (Header + Title + Abstract) ────────────────────
const section1 = {
  properties: { page: PAGE_PROPS, type: SectionType.CONTINUOUS },
  children: [
    new Paragraph({
      children:[r("Volume X, Edisi X, November XXXX",{sz:sz(9),b:true})],
      spacing:{after:0},
    }),
    new Paragraph({
      children:[
        r("JIP (Jurnal Informatika Polinema)",{sz:sz(9),b:true,i:true}),
        r("      ISSN: 2614-6371  E-ISSN: 2407-070X",{sz:sz(9)}),
      ],
      spacing:{after:120},
    }),
    new Paragraph({
      children:[r("SISTEM PENDUKUNG KEPUTUSAN SELEKSI BEASISWA LAZIZMU\nMENGGUNAKAN SWARA DAN MAIRCA",{b:true,sz:sz(14)})],
      alignment: AlignmentType.CENTER,
      spacing:{before:80, after:60},
    }),
    new Paragraph({
      children:[
        r("[Nama Penulis 1]",{b:true,sz:sz(10)}),
        r("1",{b:true,sz:sz(8)}),
        r(", [Nama Penulis 2]",{b:true,sz:sz(10)}),
        r("2",{b:true,sz:sz(8)}),
      ],
      alignment:AlignmentType.CENTER, spacing:{after:40},
    }),
    new Paragraph({
      children:[r("1,2 Program Studi Teknik Informatika, Jurusan Teknologi Informasi, Politeknik Negeri Malang",{sz:sz(10)})],
      alignment:AlignmentType.CENTER, spacing:{after:40},
    }),
    new Paragraph({
      children:[r("1[email1@polinema.ac.id], 2[email2@polinema.ac.id]",{sz:sz(10)})],
      alignment:AlignmentType.CENTER, spacing:{after:120},
    }),
    new Paragraph({
      children:[r("")],
      border:{ bottom: bHLine },
      spacing:{ before:0, after:80 },
    }),
    new Paragraph({
      children:[r("Abstrak",{b:true,sz:sz(10)})],
      alignment:AlignmentType.CENTER, spacing:{after:60},
    }),
    p([r("Proses seleksi beasiswa LAZIZMU (Lembaga Amil Zakat, Infaq dan Shadaqah Muhammadiyah) di Universitas Muhammadiyah Purworejo selama ini masih dilakukan secara manual sehingga berpotensi menimbulkan subjektivitas dan ketidakkonsistenan dalam pengambilan keputusan. Penelitian ini bertujuan membangun Sistem Pendukung Keputusan (SPK) yang mampu membantu seleksi beasiswa secara objektif dan terstruktur menggunakan kombinasi metode pembobotan ",{sz:sz(10)}),
       r("Step-wise Weight Assessment Ratio Analysis",{sz:sz(10),i:true}),
       r(" (SWARA) dan metode perangkingan ",{sz:sz(10)}),
       r("Multi-Attributive Ideal-Real Comparative Analysis",{sz:sz(10),i:true}),
       r(" (MAIRCA). Data penelitian terdiri dari 20 alternatif calon penerima beasiswa yang dievaluasi berdasarkan 10 kriteria, meliputi hasil wawancara, Indeks Prestasi Kumulatif (IPK), kepemilikan Kartu Indonesia Pintar (KIP), kepemilikan Kartu Keluarga Sejahtera (KKS), penghasilan orang tua, status kepemilikan rumah, daya listrik, luas tanah, dan jenis sumber air. Metode SWARA digunakan untuk menentukan bobot kriteria berdasarkan rasio kepentingan relatif, dengan hasil K3 (Status KIP) memperoleh bobot tertinggi sebesar 0,1456. Selanjutnya, metode MAIRCA digunakan untuk menghitung nilai gap antara preferensi teoritis dan preferensi nyata setiap alternatif. Hasil perangkingan menunjukkan alternatif A03 memperoleh nilai Qi terkecil sebesar 0,00799 sehingga menempati peringkat pertama. Untuk mengukur konsistensi dan validitas hasil, dilakukan uji korelasi Spearman antara perangkingan SWARA-MAIRCA dengan metode pembanding dari penelitian terdahulu yaitu SAW (rs = 0,9805), WP (rs = 0,9729), dan TOPSIS (rs = 0,9188) yang menunjukkan korelasi positif sangat kuat. Analisis sensitivitas juga dilakukan dengan memvariasikan bobot kriteria K3 sebesar -20%, -10%, +10%, dan +20%, yang menunjukkan stabilitas tinggi di mana alternatif A03 tetap konsisten berada di peringkat pertama.",{sz:sz(10)})],
      {align:AlignmentType.JUSTIFIED, after:60}),
    new Paragraph({
      children:[
        r("Kata kunci: ",{b:true,sz:sz(10)}),
        r("sistem pendukung keputusan, beasiswa, LAZIZMU, SWARA, MAIRCA, korelasi Spearman, analisis sensitivitas",{sz:sz(10)}),
      ],
      alignment:AlignmentType.JUSTIFIED, spacing:{after:80},
    }),
    new Paragraph({
      children:[r("")],
      border:{ bottom: bHLine },
      spacing:{ before:0, after:120 },
    }),
  ]
};

// ── Section 2 Children (Combined Body Content) ─────────────────────────────
const bodyChildren = [
  secHead("Pendahuluan", "1"),
  p([r("Lembaga Amil Zakat, Infaq dan Shadaqah Muhammadiyah (LAZIZMU) merupakan lembaga pengelola dana zakat yang secara aktif menyalurkan beasiswa bagi mahasiswa yang memenuhi kriteria tertentu. Di Universitas Muhammadiyah Purworejo (UMP), proses seleksi beasiswa LAZIZMU melibatkan banyak kriteria dan dilakukan secara manual, sehingga berpotensi menghasilkan keputusan yang tidak konsisten dan kurang transparan (Pasa et al., 2022).",{sz:sz(10)})] ,{indent:true,after:80}),
  p([r("Berbagai penelitian terdahulu telah menerapkan metode MCDM untuk seleksi beasiswa. Pasa et al. (2023) membandingkan metode SAW, WP, dan TOPSIS pada kasus yang sama dengan 20 alternatif, dengan hasil menunjukkan TOPSIS paling sensitif terhadap perubahan bobot. Ridho et al. (2021) mengombinasikan AHP dan TOPSIS untuk rekomendasi beasiswa SMK dan menghasilkan sistem yang cukup akurat. Namun, metode-metode tersebut belum mengeksplorasi SWARA sebagai pembobot adaptif berbasis penilaian pakar maupun MAIRCA sebagai metode perangkingan berbasis gap preferensi.",{sz:sz(10)})] ,{indent:true,after:80}),
  p([r("Metode SWARA ("),r("Step-wise Weight Assessment Ratio Analysis",{sz:sz(10),i:true}),r(") ditawarkan sebagai alternatif pembobotan yang lebih efisien karena tidak memerlukan perbandingan berpasangan seperti AHP (Ardiansyah et al., 2025). Metode MAIRCA ("),r("Multi-Attributive Ideal-Real Comparative Analysis",{sz:sz(10),i:true}),r(") bekerja dengan menghitung gap antara preferensi teoritis dan preferensi nyata, sehingga menghasilkan perangkingan yang lebih komprehensif dibandingkan metode berbasis nilai maksimum-minimum semata (Stević et al., 2018).",{sz:sz(10)})] ,{indent:true,after:80}),
  p([r("Penelitian ini menerapkan kombinasi SWARA-MAIRCA pada studi kasus seleksi beasiswa LAZIZMU menggunakan 20 alternatif dan 10 kriteria yang diadaptasi dari Pasa et al. (2023). Selain itu, dilakukan uji korelasi Spearman untuk mengukur konsistensi hasil perangkingan dibandingkan metode pembanding SAW, WP, dan TOPSIS dari penelitian Pasa et al. (2023). Analisis sensitivitas juga diintegrasikan untuk membuktikan stabilitas dan keandalan keputusan SPK yang dibangun terhadap variasi bobot kriteria utama.",{sz:sz(10)})] ,{indent:true,after:80}),

  secHead("Metodologi Penelitian", "2"),
  subHead("2.1  Data Penelitian"),
  p([r("Data penelitian diadaptasi dari Pasa et al. (2023) yang memuat data seleksi beasiswa LAZIZMU UMP Tahun Akademik 2021/2022 berdasarkan Surat Keputusan Rektor UMP Nomor 616/KEP/II.3.AU/F/2021. Data terdiri dari 20 alternatif mahasiswa (A01–A20) dan 10 kriteria penilaian.",{sz:sz(10)})] ,{indent:true,after:80}),

  subHead("2.2  Kriteria Penelitian"),
  p([r("Sepuluh kriteria yang digunakan terdiri dari 4 atribut ",{sz:sz(10)}),r("benefit",{sz:sz(10),i:true}),r(" (semakin besar semakin baik) dan 6 atribut ",{sz:sz(10)}),r("cost",{sz:sz(10),i:true}),r(" (semakin kecil semakin baik), seperti ditunjukkan pada Tabel 1.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 1. Kriteria Penelitian"),
  placeholder("Tabel 1. Kriteria Penelitian"),
  blank(),

  subHead("2.3  Metode SWARA"),
  p([r("SWARA adalah metode pembobotan berbasis penilaian pakar yang menentukan bobot melalui rasio kepentingan relatif antarkriteria secara bertahap (Gustientiedina, 2025). Langkah-langkah perhitungan SWARA adalah sebagai berikut:",{sz:sz(10)})] ,{indent:true,after:60}),
  p([r("a.  Mengurutkan kriteria dari yang paling penting hingga yang paling tidak penting berdasarkan penilaian pakar.",{sz:sz(10)})] ,{after:40}),
  p([r("b.  Menentukan nilai komparatif Sj (S",{sz:sz(10)}),r("j",{sz:sz(8)}),r(" = 0 untuk kriteria pertama).",{sz:sz(10)})] ,{after:40}),
  p([r("c.  Menghitung koefisien Kj menggunakan persamaan:",{sz:sz(10)})] ,{after:20}),
  new Paragraph({
    children:[r("Kj = Sj + 1  ……………………………………  (1)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED, spacing:{before:0,after:40},indent:{left:360}
  }),
  p([r("d.  Menghitung bobot sementara Qj (Q",{sz:sz(10)}),r("1",{sz:sz(8)}),r(" = 1):",{sz:sz(10)})] ,{after:20}),
  new Paragraph({
    children:[r("Qj = Qj−1 / Kj  ………………………………  (2)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED, spacing:{before:0,after:40},indent:{left:360}
  }),
  p([r("e.  Menghitung bobot akhir Wj:",{sz:sz(10)})] ,{after:20}),
  new Paragraph({
    children:[r("Wj = Qj / ΣQj  …………………………………  (3)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED, spacing:{before:0,after:80},indent:{left:360}
  }),

  subHead("2.4  Metode MAIRCA"),
  p([r("MAIRCA menghitung selisih (gap) antara preferensi teoritis dan preferensi nyata setiap alternatif (Stević et al., 2018). Alternatif dengan total gap terkecil (Qi minimum) merupakan alternatif terbaik. Langkah-langkah perhitungan MAIRCA:",{sz:sz(10)})] ,{indent:true,after:60}),
  p([r("a.  Normalisasi matriks keputusan:",{sz:sz(10)})] ,{after:20}),
  new Paragraph({
    children:[r("Benefit: x̄ij = (xij − min xj) / (max xj − min xj)  …(4)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED,spacing:{before:0,after:20},indent:{left:360}
  }),
  new Paragraph({
    children:[r("Cost: x̄ij = (max xj − xij) / (max xj − min xj)  …(5)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED,spacing:{before:0,after:40},indent:{left:360}
  }),
  p([r("b.  Menghitung preferensi teoritis TPij (m = jumlah alternatif = 20):",{sz:sz(10)})] ,{after:20}),
  new Paragraph({
    children:[r("TPij = Wj / m  ……………………………………………  (6)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED,spacing:{before:0,after:40},indent:{left:360}
  }),
  p([r("c.  Menghitung preferensi nyata PPij:",{sz:sz(10)})] ,{after:20}),
  new Paragraph({
    children:[r("PPij = TPij × x̄ij  ……………………………………………  (7)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED,spacing:{before:0,after:40},indent:{left:360}
  }),
  p([r("d.  Menghitung matriks gap Gij:",{sz:sz(10)})] ,{after:20}),
  new Paragraph({
    children:[r("Gij = TPij − PPij  ……………………………………………  (8)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED,spacing:{before:0,after:40},indent:{left:360}
  }),
  p([r("e.  Menghitung total gap per alternatif Qi:",{sz:sz(10)})] ,{after:20}),
  new Paragraph({
    children:[r("Qi = Σⁿj=1 Gij  ……………………………………………………  (9)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED,spacing:{before:0,after:80},indent:{left:360}
  }),

  subHead("2.5  Uji Korelasi Spearman"),
  p([r("Uji korelasi Spearman dilakukan untuk mengukur tingkat kesesuaian (konsistensi) antara hasil perangkingan SWARA-MAIRCA (R1) dengan metode-metode pembanding (R2) yaitu SAW, WP, dan TOPSIS. Koefisien korelasi Spearman (rs) bernilai antara −1 hingga +1, di mana nilai mendekati +1 menunjukkan hubungan positif yang sangat kuat (Fauzi et al., 2020). Persamaan koefisien korelasi Spearman adalah:",{sz:sz(10)})] ,{indent:true,after:60}),
  new Paragraph({
    children:[r("rs = 1 − (6 × Σdi²) / (n × (n² − 1))  ………  (10)",{sz:sz(10),i:true})],
    alignment:AlignmentType.JUSTIFIED,spacing:{before:0,after:40},indent:{left:360}
  }),
  p([r("di mana n = jumlah alternatif (n = 20), di = selisih peringkat ke-i antara metode SWARA-MAIRCA dan metode pembanding.",{sz:sz(10)})] ,{after:80}),

  // ══ 3. HASIL DAN PEMBAHASAN ═════════════════════════════════════════════
  secHead("Hasil dan Pembahasan", "3"),
  subHead("3.1  Data Kriteria dan Alternatif"),
  p([r("Penelitian ini menggunakan 20 alternatif calon penerima beasiswa yang dievaluasi berdasarkan 10 kriteria. Matriks data keputusan disajikan pada Tabel 2.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 2. Matriks Data Keputusan"),
  placeholder("Tabel 2. Matriks Data Keputusan"),
  blank(),

  subHead("3.2  Hasil Pembobotan SWARA"),
  p([r("Pembobotan SWARA dilakukan berdasarkan tingkat kepentingan relatif antarkriteria dalam konteks seleksi beasiswa. Status KIP (K3) ditetapkan sebagai prioritas pertama karena merupakan indikator utama kelayakan beasiswa sosial. Sebagai contoh, K4 (Status KKS) berada pada prioritas 2 dengan Sj = 0,05, sehingga diperoleh Kj = 1,05, Qj = 1/1,05 = 0,9524, dan Wj = 0,9524/6,8702 = 0,1386. Hasil lengkap pembobotan SWARA disajikan pada Tabel 3.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 3. Hasil Pembobotan SWARA"),
  placeholder("Tabel 3. Hasil Pembobotan SWARA"),
  blank(),

  p([r("Hasil pembobotan menunjukkan K3 (Status KIP) memperoleh bobot tertinggi 0,1456 (14,56%), diikuti K4 (Status KKS) 0,1386 (13,86%), dan K2 (IPK) 0,1260 (12,60%). Sebaliknya, K8 (Daya Listrik) mendapatkan bobot terendah 0,0648 (6,48%). Hal ini mencerminkan bahwa kepemilikan dokumen sosial-ekonomi dan prestasi akademik merupakan faktor paling determinan dalam seleksi beasiswa LAZIZMU.",{sz:sz(10)})] ,{indent:true,after:80}),

  subHead("3.3  Normalisasi Matriks Keputusan"),
  p([r("Normalisasi dilakukan menggunakan Persamaan (4) untuk atribut ",{sz:sz(10)}),r("benefit",{sz:sz(10),i:true}),r(" dan Persamaan (5) untuk atribut ",{sz:sz(10)}),r("cost",{sz:sz(10),i:true}),r(". Sebagai contoh, A01 pada K3 (Benefit): x̄ = (5−1)/(5−1) = 1,0000; A01 pada K7 (Cost): x̄ = (4−4)/(4−1) = 0,0000. Seluruh hasil normalisasi disajikan pada Tabel 4.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 4. Matriks Normalisasi"),
  placeholder("Tabel 4. Matriks Normalisasi"),
  blank(),

  subHead("3.4  Preferensi Teoritis (TP) dan Preferensi Nyata (PP)"),
  p([r("Preferensi teoritis TPij dihitung menggunakan Persamaan (6) dengan membagi bobot Wj oleh jumlah alternatif m = 20. Nilai TPij bersifat konstan untuk semua alternatif pada setiap kriteria. Sebagai contoh, TP(K3) = 0,1456/20 = 0,00728. Tabel 5 menyajikan nilai bobot dan preferensi teoritis seluruh kriteria.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 5. Bobot Wj dan Preferensi Teoritis (TPij)"),
  placeholder("Tabel 5. Bobot Wj dan Preferensi Teoritis (TPij)"),
  blank(),

  p([r("Preferensi nyata PPij dihitung menggunakan Persamaan (7). Sebagai contoh, PP(A03, K3) = 0,00728 × 1,0000 = 0,00728 dan PP(A01, K7) = 0,00412 × 0,0000 = 0,00000, menunjukkan A01 tidak memenuhi kondisi ideal pada kriteria K7.",{sz:sz(10)})] ,{indent:true,after:80}),

  subHead("3.5  Matriks Gap dan Perangkingan"),
  p([r("Gap dihitung menggunakan Persamaan (8). Nilai Qi merupakan penjumlahan seluruh gap per alternatif (Persamaan 9). Semakin kecil Qi, semakin baik peringkat alternatif. Tabel 6 menyajikan hasil perangkingan SWARA-MAIRCA.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 6. Hasil Perangkingan SWARA-MAIRCA"),
  placeholder("Tabel 6. Hasil Perangkingan SWARA-MAIRCA"),
  blank(),

  p([r("Berdasarkan Tabel 6, alternatif A03 menempati peringkat pertama (Qi = 0,00799), diikuti A14 (Qi = 0,01159) dan A01 (Qi = 0,01342). Alternatif A03 terpilih sebagai calon penerima beasiswa terbaik karena memiliki gap terkecil terhadap kondisi ideal di seluruh kriteria.",{sz:sz(10)})] ,{indent:true,after:80}),

  subHead("3.6  Analisis Perbandingan Metode"),
  p([r("Hasil perangkingan menggunakan metode usulan SWARA-MAIRCA kemudian dibandingkan dengan hasil dari metode pembanding SAW, WP, dan TOPSIS yang diadaptasi dari penelitian Pasa et al. (2023). Hasil perbandingan peringkat disajikan pada Tabel 7.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 7. Perbandingan Peringkat SWARA-MAIRCA vs SAW, WP, dan TOPSIS"),
  placeholder("Tabel 7. Perbandingan Peringkat SWARA-MAIRCA vs SAW, WP, dan TOPSIS"),
  figLabel("*Nilai SAW V, WP V, dan TOPSIS V diambil dari penelitian Pasa et al. (2023)"),
  blank(),

  p([r("Berdasarkan Tabel 7, alternatif A03 secara konsisten menempati peringkat pertama (Rank 1) pada seluruh metode pembanding utama (SAW, WP, TOPSIS, dan SWARA-MAIRCA). Hal ini menunjukkan keandalan metode usulan dalam menyaring alternatif terbaik secara objektif. Selanjutnya, dilakukan uji korelasi Spearman untuk mengetahui tingkat kesesuaian hasil perangkingan antarmetode. Hasil perhitungan koefisien korelasi Spearman disajikan pada Tabel 8.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 8. Koefisien Korelasi Spearman Antarmetode"),
  placeholder("Tabel 8. Koefisien Korelasi Spearman Antarmetode"),
  blank(),

  p([r("Berdasarkan Tabel 8, diperoleh nilai koefisien korelasi Spearman antara SWARA-MAIRCA dengan SAW sebesar rs = 0,9805; dengan WP sebesar rs = 0,9729; dan dengan TOPSIS sebesar rs = 0,9188. Seluruh koefisien korelasi bernilai > 0,80, yang mengindikasikan korelasi positif yang sangat kuat. Hal ini membuktikan bahwa metode SWARA-MAIRCA yang diusulkan memiliki tingkat validitas dan konsistensi yang sangat tinggi.",{sz:sz(10)})] ,{indent:true,after:80}),

  subHead("3.7  Analisis Sensitivitas"),
  p([r("Analisis sensitivitas dilakukan untuk menguji tingkat kestabilan hasil perangkingan terhadap perubahan bobot kriteria. Pengujian dilakukan dengan memvariasikan bobot kriteria Status KIP (K3) yang memiliki bobot terbesar sebesar -20%, -10%, +10%, dan +20%. Hasil pengujian sensitivitas ditunjukkan pada Tabel 9.",{sz:sz(10)})] ,{indent:true,after:60}),

  tblLabel("Tabel 9. Hasil Analisis Sensitivitas"),
  placeholder("Tabel 9. Hasil Analisis Sensitivitas"),
  blank(),

  p([r("Berdasarkan Tabel 9, perubahan bobot kriteria K3 dari -20% hingga +20% menunjukkan stabilitas yang sangat tinggi. Alternatif A03 tetap kokoh menempati peringkat pertama pada seluruh skenario perubahan bobot. Persentase perubahan peringkat alternatif lain hanya berkisar 10,0% (hanya 2 dari 20 alternatif yang bertukar posisi, yaitu A01 dan A05). Hal ini menunjukkan bahwa keputusan yang dihasilkan oleh metode SWARA-MAIRCA tidak sensitif terhadap fluktuasi bobot kriteria, sehingga sangat andal untuk diimplementasikan.",{sz:sz(10)})] ,{indent:true,after:80}),

  // ══ 4. KESIMPULAN DAN SARAN ═════════════════════════════════════════════
  secHead("Kesimpulan dan Saran", "4"),
  subHead("4.1  Kesimpulan"),
  p([r("Berdasarkan hasil penelitian dan pembahasan, diperoleh kesimpulan sebagai berikut. Pertama, metode SWARA berhasil menentukan bobot kriteria secara objektif dengan K3 (Status KIP) memperoleh bobot tertinggi 14,56% dan K8 (Daya Listrik) terendah 6,48%. Kedua, metode MAIRCA menghasilkan perangkingan dengan A03 sebagai alternatif terbaik (Qi = 0,00799), diikuti A14 (0,01159) dan A01 (0,01342). Ketiga, uji korelasi Spearman membuktikan konsistensi hasil perangkingan dengan koefisien korelasi sangat kuat terhadap SAW (rs = 0,9805), WP (rs = 0,9729), dan TOPSIS (rs = 0,9188). Keempat, analisis sensitivitas mengonfirmasi stabilitas model di mana A03 secara konsisten menempati peringkat pertama pada seluruh variasi bobot kriteria K3.",{sz:sz(10)})] ,{indent:true,after:80}),

  subHead("4.2  Saran"),
  p([r("Penelitian selanjutnya disarankan untuk: (1) menguji sensitivitas terhadap kriteria lainnya; (2) membandingkan metode SWARA dengan metode pembobotan objektif seperti Entropy atau CRITIC; dan (3) mengembangkan SPK seleksi beasiswa berbasis web untuk mempermudah operasional lembaga beasiswa LAZIZMU.",{sz:sz(10)})] ,{indent:true,after:100}),

  // ══ DAFTAR PUSTAKA ═══════════════════════════════════════════════════════
  new Paragraph({
    children:[r("Daftar Pustaka",{b:true,sz:sz(10)})],
    spacing:{before:120,after:80},
  }),

  // 15 References — APA style
  ...[
    "Ardiansyah, M., Fahmi, H., & Imtihan, K. (2025). Sistem Penunjang Keputusan Penerima Bantuan Rumah Tidak Layak Huni (RTLH) Menggunakan Metode SWARA. Journal of Information System Research (JOSH), 7(1), 12–20. https://doi.org/10.47065/josh.v7i1.8235",
    "Diana, D., & Seprina, I. (2019). Sistem Pendukung Keputusan untuk Menentukan Penerima Bantuan Sosial Menerapkan Weighted Product Method. Jurnal Edukasi Dan Penelitian Informatika (JEPIN), 5(3), 370. https://doi.org/10.26418/jp.v5i3.34971",
    "Fauzi, A. A., Zahro', H. Z., & Prasetya, R. P. (2020). Analisis Perbandingan Metode TOPSIS dan SAW dalam Penentuan Prioritas Perbaikan Jalan di Kabupaten Rembang. JATI (Jurnal Mahasiswa Teknik Informatika), 4(2), 29–36. https://doi.org/10.36040/JATI.V4I2.2676",
    "Gustientiedina, G. (2025). Penerapan Metode SWARA dan MOORA dalam Menentukan Insentif Karyawan. Jurnal Mahasiswa Aplikasi Teknologi Komputer dan Informasi (JMApTeKsi). https://ejournal.pelitaindonesia.ac.id/ojs32/index.php/jmapteksi/article/view/4165",
    "Hayat, M. (2023). Sistem Pendukung Keputusan Pemilihan Aplikasi Investasi Menggunakan Pendekatan WASPAS dan Rank Sum. Jurnal FASILKOM, 13(3). https://doi.org/10.37859/jf.v13i3.6080",
    "Ishak, I., Mesran, M., & Ndruru, E. (2023). Kombinasi Multi-Criteria Decision Making dan LOPCOW. Graha Mitra Edukasi. https://repository.nusamandiri.ac.id/repo/files/248484",
    "Jayawardani, W. R. K., & Maryam, M. (2022). Sistem Pendukung Keputusan Seleksi Penerima Program Keluarga Harapan dengan Implementasi Metode SAW dan Pembobotan ROC. Emitor: Jurnal Teknik Elektro, 22(2), 99–109. https://doi.org/10.23917/EMITOR.V22I2.18411",
    "Mandarani, P., Ramadhan, H., Yulianti, E., & Syahrani, A. (2022). Sistem Pendukung Keputusan Penulis Terbaik Menggunakan Metode Rank Order Centroid (ROC) dan EDAS. Journal of Information System Research (JOSH), 3(4), 686–694. https://doi.org/10.47065/josh.v3i4.1845",
    "Murtina, H., Hidayatun, N., & Susafa'ati, S. (2022). Implementasi Multi Attribute Decision Making Menggunakan Metode Weighted Product dalam Pemilihan Supervisor. Jurnal Teknoinfo, 16(2), 435–442. https://ejurnal.teknokrat.ac.id/index.php/teknoinfo/article/view/2006",
    "Pasa, I. Y., Prasetya, N. W. A., & Maharrani, R. H. (2022). Penerapan Metode SAW pada Penentuan Penerima Beasiswa Lazizmu. INTEK: Jurnal Informatika dan Teknologi Informasi, 5(1), 81–89. https://doi.org/10.37729/INTEK.V5I1.1971",
    "Pasa, I. Y., Prasetya, N. W. A., & Maharrani, R. H. (2023). Analisis Perbandingan Metode SAW, WP, dan TOPSIS untuk Optimasi SPK Proses Seleksi Beasiswa Lazizmu. Jurnal INTEK, 6(1), 65–76.",
    "Ridho, M. R., Hairani, H., Latif, K. A., & Hammad, R. (2021). Kombinasi Metode AHP dan TOPSIS untuk Rekomendasi Penerima Beasiswa SMK Berbasis Sistem Pendukung Keputusan. Jurnal Tekno Kompak, 15(1), 26–39. https://doi.org/10.33365/JTK.V15I1.905",
    "Sari, W. E., B, M., & Rani, S. (2021). Perbandingan Metode SAW dan TOPSIS pada Sistem Pendukung Keputusan Seleksi Penerima Beasiswa. Jurnal Sisfokom, 10(1), 52–58. https://doi.org/10.32736/sisfokom.v10i1.1027",
    "Setiawan, D. (2023). Implementasi Metode COPRAS (Complex Proportional Assessment) dalam Pemilihan Security. Jurnal Sistem Informasi Triguna Dharma (JURSI TGD). https://ojs.trigunadharma.ac.id/index.php/jsi/article/view/5332",
    "Stević, Ž., Pamučar, D., Subotić, M., Antuchevičiene, J., & Zavadskas, E. K. (2018). The Sustainable Supplier Selection Based on the Original MAIRCA and Interval Rough Numbers. Symmetry, 10(8), 305. https://doi.org/10.3390/sym10080305",
  ].map((ref,i)=>new Paragraph({
    children:[r(`${ref}`,{sz:sz(9)})],
    alignment:AlignmentType.JUSTIFIED,
    spacing:{before:0,after:60},
    indent:{left:180,hanging:180},
  })),
];

// ─── MAIN BUILDER FUNCTION ──────────────────────────────────────────────────
function buildDoc(isTwoColumn) {
  const section2 = {
    properties: {
      page: PAGE_PROPS,
      type: SectionType.CONTINUOUS,
      column: isTwoColumn ? { count: 2, space: COL_GAP, equalWidth: true } : undefined
    },
    children: bodyChildren
  };

  const doc = new Document({
    styles:{
      default:{document:{run:{font:F,size:sz(10)}}}
    },
    numbering:{config:[{
      reference:"bullet",
      levels:[{level:0,format:LevelFormat.BULLET,text:"•",
        style:{paragraph:{indent:{left:360,hanging:180}}}}]
    }]},
    sections:[section1, section2]
  });

  const outputFilename = isTwoColumn
    ? "output/Jurnal_JIP_SWARA_MAIRCA_Final_2Kolom.docx"
    : "output/Jurnal_JIP_SWARA_MAIRCA_Final_1Kolom.docx";

  Packer.toBuffer(doc).then(buf => {
    try {
      fs.writeFileSync(outputFilename, buf);
      console.log(`\nSukses! File berhasil dibuat: ${outputFilename}`);
      process.exit(0);
    } catch (writeErr) {
      if (writeErr.code === 'EBUSY' || writeErr.code === 'EPERM') {
        const fallbackFilename = outputFilename.replace(".docx", "_baru.docx");
        console.warn(`\n[Peringatan] File utama ${outputFilename} sedang dikunci (kemungkinan sedang dibuka di Word).`);
        console.log(`Mencoba menyimpan ke file alternatif: ${fallbackFilename}`);
        fs.writeFileSync(fallbackFilename, buf);
        console.log(`Sukses! File berhasil disimpan di: ${fallbackFilename}`);
        process.exit(0);
      } else {
        throw writeErr;
      }
    }
  }).catch(err => {
    console.error(`\nGagal menulis file: ${err.message}`);
    process.exit(1);
  });
}

// ─── INTERACTIVE PROMPT ──────────────────────────────────────────────────────
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

console.log("=============================================");
console.log("PILIH LAYOUT DOKUMEN JURNAL");
console.log("=============================================");
console.log("1. 1 Kolom (Dokumen Biasa)");
console.log("2. 2 Kolom (Jurnal JIP)");
console.log("=============================================");

rl.question("Pilihan Anda (1/2) [Default: 2]: ", (answer) => {
  const isTwoColumn = answer.trim() !== "1";
  rl.close();
  buildDoc(isTwoColumn);
});