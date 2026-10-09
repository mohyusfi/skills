"""
generate_report.py
Skrip CLI utama pembuat Laporan Praktikum (.docx) sesuai standar akademik:
- Template & format berbasis referensi f5512520089_prak_gis_5.docx
- Kertas A4, Margin 4-4-3-3 cm, Times New Roman 12 pt, 1.5 Line Spacing
- Hierarki heading native Word (kompatibel penuh MS Word Navigation Pane)
- Cetak miring otomatis istilah non-KBBI / bahasa asing
- Foto modul asdos dibersihkan dan diganti placeholder [Sisipkan Gambar: ...]
- Bab V kesimpulan kumulatif otomatis dari laporan-laporan sebelumnya
- Bagian Daftar Pustaka APA 7th Edition (hanging indent)
"""

import os
import sys
import json
import glob
import re
import argparse

# Pastikan modul di folder scripts dapat diimpor
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

import docx_styler

def scan_historical_conclusions(directory_path: str, current_module_num: int):
    """
    Memindai dokumen praktikum sebelumnya di direktori untuk mengumpulkan
    kesimpulan 5.1.1 s.d. 5.1.(current_module_num - 1).
    """
    if not directory_path or not os.path.isdir(directory_path):
        return []

    conclusions = {}
    files = glob.glob(os.path.join(directory_path, "*prak_*.docx")) + glob.glob(os.path.join(directory_path, "*modul*.docx"))
    files = sorted(list(set(files)))

    for f in files:
        try:
            doc = docx_styler.docx.Document(f)
            in_bab5 = False
            current_num = None
            current_title = None
            current_text = []

            for p in doc.paragraphs:
                txt = p.text.strip()
                if not txt:
                    continue
                if "BAB V" in txt.upper() or "PENUTUP" in txt.upper():
                    in_bab5 = True
                    continue
                if "DAFTAR PUSTAKA" in txt.upper():
                    in_bab5 = False
                    if current_num and current_text:
                        conclusions[current_num] = {
                            'title': current_title,
                            'text': " ".join(current_text)
                        }
                    break

                if in_bab5:
                    m = re.match(r'^(?:[\t\s]*)5\.1\.(\d+)(?:[\t\s]+)(.*)$', txt)
                    if m:
                        if current_num and current_text:
                            conclusions[current_num] = {
                                'title': current_title,
                                'text': " ".join(current_text)
                            }
                            current_text = []

                        sub_idx = int(m.group(1))
                        current_num = f"5.1.{sub_idx}"
                        current_title = m.group(2).strip()
                    elif current_num:
                        current_text.append(txt)

            if current_num and current_text:
                conclusions[current_num] = {
                    'title': current_title,
                    'text': " ".join(current_text)
                }
        except Exception as e:
            print(f"[Warning] Gagal membaca kesimpulan dari {f}: {e}")

    result = []
    sorted_keys = sorted(conclusions.keys(), key=lambda k: [int(x) for x in k.split('.')])
    for k in sorted_keys:
        try:
            idx = int(k.split('.')[-1])
            if idx < current_module_num:
                result.append({
                    'number': k,
                    'title': conclusions[k]['title'],
                    'text': conclusions[k]['text']
                })
        except ValueError:
            continue

    return result

def build_report_docx(data: dict, output_path: str, template_path: str = None, scan_dir: str = None):
    """Membangun dokumen Word dari data input JSON/dictionary."""
    # 1. Inisialisasi dokumen
    if not template_path:
        default_tpl = os.path.join(os.path.dirname(script_dir), "templates", "base_template.docx")
        if os.path.exists(default_tpl):
            template_path = default_tpl

    doc = docx_styler.create_document(template_path)

    module_num = data.get("module_number", 1)
    try:
        module_num_int = int(module_num)
    except ValueError:
        module_num_int = 1

    module_title = data.get("module_title", "Laporan Praktikum")
    header_title = data.get("header_title", module_title)
    nama = data.get("student_name", "MOH. YUSFI LAKHAFIDUN")
    nim = data.get("student_nim", "F5512520089")
    kelas = data.get("student_class", "TI C")

    # 2. Halaman Sampul (Identitas)
    docx_styler.add_identity_header(doc, header_title, nama, nim, kelas)
    docx_styler.add_page_break(doc)

    has_parent_headings = data.get("has_parent_headings", False)

    # 3. BAB II Landasan Teori (Halaman 2 Baru)
    bab2 = data.get("bab2", {})
    bab2_has_parent = bab2.get("has_parent_headings", has_parent_headings)
    bab2_chapter_title = bab2.get("chapter_title", "BAB II \nLANDASAN TEORI")
    bab2_parent_title = bab2.get("parent_title", "Landasan Teori")
    bab2_title = bab2.get("subbab_title", module_title)

    if bab2_has_parent:
        docx_styler.add_heading_1(doc, bab2_chapter_title)
        docx_styler.add_heading_2(doc, f"2.1\t{bab2_parent_title}")
    
    docx_styler.add_heading_3(doc, f"\t2.1.{module_num_int}\t{bab2_title}")
    
    # Isi Bab 2
    content_items = bab2.get("content", [])
    est_lines = sum(
        (len(item.get("text", "")) // 55 + 1) if isinstance(item, dict) else (len(item) // 55 + 1)
        for item in content_items
    )
    print(f"[Info Bab II] Total item: {len(content_items)}, estimasi baris: ~{est_lines} baris (kapasitas halaman A4: ~34 baris/halaman).")

    bab2_num_id = docx_styler.create_numbered_list_id(doc)
    current_content_indent_cm = 3.29
    for item in bab2.get("content", []):
        if isinstance(item, str):
            docx_styler.add_body_paragraph(doc, item, first_line_indent_cm=1.27, left_indent_cm=1.27)
        elif isinstance(item, dict):
            item_type = item.get("type", "paragraph")
            text = item.get("text", "")
            level = item.get("level", 1)
            if item.get("restart"):
                bab2_num_id = docx_styler.create_numbered_list_id(doc)
            if item_type == "list":
                prefix = item.get("prefix", "1.")
                list_type = item.get("list_type", "number")
                base_indent_cm = item.get("base_indent_cm", 2.54)
                text_indent_cm = item.get("text_indent_cm", 0.75)
                current_content_indent_cm = base_indent_cm + (level - 1) * text_indent_cm + text_indent_cm
                docx_styler.add_list_item(doc, text, prefix=prefix, level=level, num_id=bab2_num_id, list_type=list_type, base_indent_cm=base_indent_cm, text_indent_cm=text_indent_cm)
            elif item_type == "list_body":
                base_ind = item.get("base_indent_cm", 2.54)
                txt_ind = item.get("text_indent_cm", 0.75)
                first_ind = item.get("first_line_indent_cm", 0.75)
                docx_styler.add_list_body_paragraph(doc, text, level=level, base_indent_cm=base_ind, text_indent_cm=txt_ind, first_line_indent_cm=first_ind)
            elif item_type == "subheading":
                docx_styler.add_body_paragraph(doc, f"**{text}**", first_line_indent_cm=0, left_indent_cm=2.54)
            elif item_type == "page_break":
                docx_styler.add_page_break(doc)
            else:
                first_ind = item.get("first_line_indent_cm", 1.27)
                left_ind = item.get("left_indent_cm", 1.27)
                docx_styler.add_body_paragraph(doc, text, first_line_indent_cm=first_ind, left_indent_cm=left_ind)

    # 4. BAB III Langkah Kerja (Halaman Baru)
    docx_styler.add_page_break(doc)
    bab3 = data.get("bab3", {})
    bab3_has_parent = bab3.get("has_parent_headings", has_parent_headings)
    bab3_chapter_title = bab3.get("chapter_title", "BAB III \nLANGKAH KERJA")
    bab3_parent_title = bab3.get("parent_title", "Langkah Kerja")
    bab3_title = bab3.get("subbab_title", module_title)

    if bab3_has_parent:
        docx_styler.add_heading_1(doc, bab3_chapter_title)
        docx_styler.add_heading_2(doc, f"3.1\t{bab3_parent_title}")
    
    docx_styler.add_heading_3(doc, f"\t3.1.{module_num_int}\t{bab3_title}")

    bab3_num_id = docx_styler.create_numbered_list_id(doc)
    current_content_indent_cm = 3.29
    for item in bab3.get("content", []):
        if isinstance(item, str):
            docx_styler.add_body_paragraph(doc, item, first_line_indent_cm=0, left_indent_cm=current_content_indent_cm)
        elif isinstance(item, dict):
            item_type = item.get("type", "paragraph")
            text = item.get("text", "")
            level = item.get("level", 1)
            if item.get("restart"):
                bab3_num_id = docx_styler.create_numbered_list_id(doc)
            if item_type == "heading_step":
                p = docx_styler.add_body_paragraph(doc, f"**{text}**", first_line_indent_cm=0, left_indent_cm=2.54)
            elif item_type == "list":
                prefix = item.get("prefix", "1.")
                list_type = item.get("list_type", "number")
                base_indent_cm = item.get("base_indent_cm", 2.54)
                text_indent_cm = item.get("text_indent_cm", 0.75)
                current_content_indent_cm = base_indent_cm + (level - 1) * text_indent_cm + text_indent_cm
                docx_styler.add_list_item(doc, text, prefix=prefix, level=level, num_id=bab3_num_id, list_type=list_type, base_indent_cm=base_indent_cm, text_indent_cm=text_indent_cm)
            elif item_type == "list_body":
                base_ind = item.get("base_indent_cm", 2.54)
                txt_ind = item.get("text_indent_cm", 0.75)
                first_ind = item.get("first_line_indent_cm", 0)
                docx_styler.add_list_body_paragraph(doc, text, level=level, base_indent_cm=base_ind, text_indent_cm=txt_ind, first_line_indent_cm=first_ind)
            elif item_type == "placeholder":
                indent = item.get("left_indent_cm", current_content_indent_cm)
                docx_styler.add_image_placeholder(doc, text, left_indent_cm=indent)
            elif item_type == "page_break":
                docx_styler.add_page_break(doc)
            else:
                indent = item.get("left_indent_cm", current_content_indent_cm)
                docx_styler.add_body_paragraph(doc, text, first_line_indent_cm=0, left_indent_cm=indent)

    # 5. BAB IV Hasil dan Pembahasan (Halaman Baru)
    docx_styler.add_page_break(doc)
    bab4 = data.get("bab4", {})
    bab4_has_parent = bab4.get("has_parent_headings", has_parent_headings)
    bab4_chapter_title = bab4.get("chapter_title", "BAB IV \nHASIL DAN PEMBAHASAN")
    bab4_parent_title = bab4.get("parent_title", "Hasil dan Pembahasan")
    bab4_title = bab4.get("subbab_title", module_title)

    if bab4_has_parent:
        docx_styler.add_heading_1(doc, bab4_chapter_title)
        docx_styler.add_heading_2(doc, f"4.1\t{bab4_parent_title}")
    
    docx_styler.add_heading_3(doc, f"\t4.1.{module_num_int}\t{bab4_title}")

    bab4_num_id = docx_styler.create_numbered_list_id(doc)
    current_content_indent_cm = 3.29
    for item in bab4.get("content", []):
        if isinstance(item, str):
            docx_styler.add_body_paragraph(doc, item, first_line_indent_cm=0, left_indent_cm=current_content_indent_cm)
        elif isinstance(item, dict):
            item_type = item.get("type", "paragraph")
            text = item.get("text", "")
            level = item.get("level", 1)
            if item.get("restart"):
                bab4_num_id = docx_styler.create_numbered_list_id(doc)
            if item_type == "heading_output":
                prefix = item.get("prefix", "")
                full_head = f"{prefix} {text}".strip() if prefix else text
                p = docx_styler.add_body_paragraph(doc, f"**{full_head}**", first_line_indent_cm=0, left_indent_cm=2.54)
            elif item_type == "list":
                prefix = item.get("prefix", "1.")
                list_type = item.get("list_type", "number")
                base_indent_cm = item.get("base_indent_cm", 2.54)
                text_indent_cm = item.get("text_indent_cm", 0.75)
                current_content_indent_cm = base_indent_cm + (level - 1) * text_indent_cm + text_indent_cm
                docx_styler.add_list_item(doc, text, prefix=prefix, level=level, num_id=bab4_num_id, list_type=list_type, base_indent_cm=base_indent_cm, text_indent_cm=text_indent_cm)
            elif item_type == "list_body":
                base_ind = item.get("base_indent_cm", 2.54)
                txt_ind = item.get("text_indent_cm", 0.75)
                first_ind = item.get("first_line_indent_cm", 0)
                docx_styler.add_list_body_paragraph(doc, text, level=level, base_indent_cm=base_ind, text_indent_cm=txt_ind, first_line_indent_cm=first_ind)
            elif item_type == "placeholder":
                indent = item.get("left_indent_cm", current_content_indent_cm)
                docx_styler.add_image_placeholder(doc, text, left_indent_cm=indent)
            elif item_type == "page_break":
                docx_styler.add_page_break(doc)
            else:
                indent = item.get("left_indent_cm", current_content_indent_cm)
                docx_styler.add_body_paragraph(doc, text, first_line_indent_cm=0, left_indent_cm=indent)

    # 6. BAB V Penutup (Halaman Baru)
    docx_styler.add_page_break(doc)
    bab5 = data.get("bab5", {})
    docx_styler.add_heading_1(doc, "BAB V \nPENUTUP")
    docx_styler.add_heading_2(doc, "5.1\tKesimpulan")

    # Ambil kesimpulan kumulatif modul-modul sebelumnya
    historical = bab5.get("historical_conclusions", [])
    if not historical and scan_dir:
        historical = scan_historical_conclusions(scan_dir, module_num_int)

    for prev in historical:
        num = prev.get("number", "").strip()
        title = prev.get("title", "").strip()
        text = prev.get("text", "").strip()
        # Awali dengan \t agar nomor 5.1.N lurus sejajar dengan huruf K pada '5.1\tKesimpulan'
        docx_styler.add_heading_3(doc, f"\t{num}\t{title}")
        docx_styler.add_conclusion_paragraph(doc, text)

    # Kesimpulan modul saat ini
    current_conclusion = bab5.get("current_conclusion", "").strip()
    current_title = bab5.get("current_title", module_title).strip()
    docx_styler.add_heading_3(doc, f"\t5.1.{module_num_int}\t{current_title}")
    docx_styler.add_conclusion_paragraph(doc, current_conclusion)

    # 7. DAFTAR PUSTAKA (Halaman Baru)
    daftar_pustaka = data.get("daftar_pustaka", [])
    if daftar_pustaka:
        docx_styler.add_page_break(doc)
        docx_styler.add_heading_1(doc, "DAFTAR PUSTAKA")
        for ref in daftar_pustaka:
            docx_styler.add_bibliography_entry(doc, ref)

    # 8. Simpan Dokumen
    out_dir = os.path.dirname(output_path)
    if out_dir and not os.path.exists(out_dir):
        os.makedirs(out_dir, exist_ok=True)

    try:
        doc.save(output_path)
        print(f"[Sukses] Dokumen laporan praktikum berhasil dibuat di: {output_path}")
    except PermissionError:
        base, ext = os.path.splitext(output_path)
        fallback_path = f"{base}_revisi{ext}"
        doc.save(fallback_path)
        print(f"[Peringatan] Berkas '{output_path}' sedang dibuka/dikunci di Microsoft Word.")
        print(f"[Sukses Cadangan] Dokumen hasil revisi berhasil disimpan di: {fallback_path}")

def main():
    parser = argparse.ArgumentParser(description="Generator Laporan Praktikum .docx Standar Akademik")
    parser.add_argument("--input", "-i", required=True, help="Path ke file data laporan (JSON)")
    parser.add_argument("--output", "-o", help="Path tujuan file .docx luaran")
    parser.add_argument("--scan-dir", "-s", help="Direktori untuk memindai riwayat kesimpulan laporan sebelumnya")
    parser.add_argument("--template", "-t", help="Path ke file base_template.docx")

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"Error: File input tidak ditemukan: {args.input}")
        sys.exit(1)

    with open(args.input, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Tentukan output path
    output_path = args.output
    if not output_path:
        output_path = data.get("output_path")
    if not output_path:
        mod_num = data.get("module_number", 1)
        nim = data.get("student_nim", "f5512520089").lower()
        output_path = f"{nim}_prak_gis_{mod_num}.docx"

    # Tentukan scan dir
    scan_dir = args.scan_dir
    if not scan_dir:
        scan_dir = data.get("scan_dir")
    if not scan_dir:
        scan_dir = os.path.dirname(os.path.abspath(output_path))

    build_report_docx(data, output_path, template_path=args.template, scan_dir=scan_dir)

if __name__ == "__main__":
    main()
