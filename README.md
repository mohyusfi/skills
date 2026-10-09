# Yuss Males Laprak 📝

[![skills.sh](https://skills.sh/b/mohyusfi/yuss-males-laprak)](https://skills.sh/mohyusfi/yuss-males-laprak)

Skill automasi laporan praktikum akademik siap kumpul berformat Microsoft Word (`.docx`) untuk AI Coding Agents (**Antigravity**, **Claude Code**, **Cursor**, **Windsurf**, dll.).

Dirancang khusus untuk mahasiswa dan praktikan yang ingin menghasilkan laporan praktikum dengan presisi tipografi akademik tinggi sesuai standar perguruan tinggi.

---

## ✨ Fitur Utama

- **Standar Tata Tulis Akademik Resmi**: Margin 4-4-3-3 cm (A4), font Times New Roman 12 pt, 1.5 spasi, perataan *Justified*.
- **Integrasi Heading Native Word**: Heading 1, 2, dan 3 berbasis OpenXML terintegrasi langsung dengan Microsoft Word Navigation Pane.
- **Otomasi Cetak Miring Non-KBBI**: Mendeteksi dan memiringkan (*italic*) istilah asing/teknis non-KBBI secara otomatis menggunakan `PySastrawi`.
- **Sanitasi Tangkapan Layar Asdos**: Menghapus tangkapan layar modul asisten laboratorium dan menggantikannya dengan penanda gambar `[Sisipkan Gambar: ...]` dengan posisi perataan rapi.
- **Hierarki Penomoran Native**: Penomoran daftar otomatis (`1.`, `a.`, `1)`) menggunakan `<w:numPr>`.
- **Page Break Mutlak**: Menjamin setiap bab berdiri sendiri pada halaman baru (Halaman Sampul, Bab II, Bab III, Bab IV, Bab V, Daftar Pustaka).
- **Bab V Kumulatif**: Otomatis memindai riwayat kesimpulan dari dokumen laporan sebelumnya di folder kerja.
- **Daftar Pustaka APA 7th Edition**: Menghubungkan sitasi ilmiah *open-access* ($\ge 2020$) secara otomatis ke daftar pustaka berformat hanging indent.

---

## 🚀 Cara Instalasi

Pasang skill ini ke agent AI Anda menggunakan CLI resmi [skills.sh](https://skills.sh/):

```bash
npx skills add mohyusfi/yuss-males-laprak
```

Untuk melihat daftar skill yang tersedia di repositori sebelum instalasi:
```bash
npx skills add mohyusfi/yuss-males-laprak --list
```

---

## 📋 Prasyarat Sistem

- **Python 3.10+**
- **[uv](https://docs.astral.sh/uv/)** (direkomendasikan untuk eksekusi dependensi cepat)
- Dependensi Python:
  - `python-docx`
  - `PySastrawi`

---

## 📂 Struktur Repositori

```text
yuss-males-laprak/
├── SKILL.md                          # Definisi instruksi skill untuk AI agent
├── README.md                         # Dokumentasi repositori
├── .gitignore                        # Filter cache dan file luaran
├── references/
│   └── report_guidelines.md          # Panduan tata tulis & hierarki dokumen
├── scripts/
│   ├── docx_styler.py                # Engine styling dokumen docx & OpenXML
│   ├── kbbi_italicizer.py            # Modul cetak miring kata non-KBBI
│   └── generate_report.py            # CLI generator laporan praktikum
└── templates/
    └── base_template.docx            # Template dasar Microsoft Word
```

---

## 👤 Pembuat & Lisensi

- **Nama**: MOH. YUSFI LAKHAFIDUN
- **GitHub**: [@mohyusfi](https://github.com/mohyusfi)
- **Lisensi**: MIT License
