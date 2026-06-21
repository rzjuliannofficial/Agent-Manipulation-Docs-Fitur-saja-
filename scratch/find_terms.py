import docx

doc = docx.Document("output/Jurnal_JIP_SWARA_MAIRCA_Final_2Kolom.docx")

for idx, p in enumerate(doc.paragraphs):
    if "spearman" in p.text.lower() or "korelasi" in p.text.lower() or "sensitivitas" in p.text.lower():
        print(f"P{idx}: {p.text}\n")
