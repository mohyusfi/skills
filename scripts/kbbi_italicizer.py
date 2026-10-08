"""
kbbi_italicizer.py
Modul pendeteksi istilah non-KBBI / bahasa asing dan pemformat run cetak miring (italic)
untuk laporan praktikum akademik berstandar Microsoft Word (.docx).
"""

import re

# Kata serapan / istilah bahasa Indonesia yang sah di KBBI (jangan di-italic)
KBBI_ALLOWLIST = {
    # Kosakata umum
    "dan", "yang", "di", "ke", "dari", "pada", "untuk", "dengan", "adalah", "ini", "itu",
    "sebagai", "dalam", "akan", "bisa", "dapat", "oleh", "karena", "maka", "jika", "bila",
    "saat", "waktu", "proses", "data", "sistem", "informasi", "praktikum", "modul", "mahasiswa",
    "langkah", "kerja", "hasil", "pembahasan", "kesimpulan", "penutup", "tabel", "peta",
    "wilayah", "luas", "panjang", "keliling", "angka", "nilai", "fungsi", "kolom", "baris",
    "fitur", "jarak", "jalan", "sungai", "titik", "garis", "pengelolaan", "mempercepat",
    "mempermudah", "pengolahan", "spasial", "terstruktur", "perhitungan", "pembaruan",
    "teks", "analisis", "secara", "otomatis", "demikian", "penggunaan",
    "mengurangi", "kesalahan", "menjaga", "konsistensi", "meningkatkan", "efisiensi",
    "persiapan", "membuka", "aplikasi", "muat", "berkas", "sebelumnya", "pastikan",
    "administrasi", "telah", "tampil", "benar", "atas", "panel", "klik", "kanan",
    "lalu", "pilih", "menu", "atau", "tekan", "tombol", "pintas", "jendela", "buka",
    "gambar", "pensil", "pojok", "kiri", "mengaktifkan", "kalkulator", "kombinasi",
    "terbuka", "atur", "parameter", "bagian", "seperti", "tengah", "bentangkan",
    "kelompok", "cari", "dua", "kali", "sehingga", "kode", "berpindah", "lengkapi",
    "ekspresi", "menambahkan", "operator", "pembagian", "konversi", "perhatikan",
    "bawah", "kalkulasi", "desimal", "terbentuk", "masih", "keadaan",
    "aktif", "konfigurasi", "penambahan", "kotak", "masukkan", "rumus", "amati",
    "simpan", "perubahan", "matikan", "manipulasi", "penggabungan", "kembali",
    "tampilan", "beserta", "digunakan", "menghitung", "mengonversinya",
    "satuan", "meter", "persegi", "kilometer", "bawaan", "membaca", "mengikuti",
    "pengaturan", "proyek", "langsung", "menghasilkan", "sekitar", "tampak",
    "mencakup", "tingkat", "kota", "tercatat", "terluas", "ukuran", "sementara",
    "menjadi", "terkecil", "keberadaan", "menggabungkan", "pembulatan", "ditujukan",
    "memudahkan", "penyajian", "ringkas", "terutama", "mengatur", "pelabelan",
    "merangkum", "segmen", "persimpangan", "berhasil", "berkisar",
    "antara", "hingga", "pengelompokan", "berdasarkan", "nama", "simpang",
    "identifikasi", "karakteristik", "ruas", "penghubung", "serta", "inventarisasi",
    "jaringan", "tiap", "pertemuan", "lintas", "memanfaatkan", "mengintegrasikan",
    "identitas", "terpadu", "menerapkan", "teknik", "merangkai", "statis", "bersama",
    "dinamis", "rapi", "terlebih", "dahulu", "diformat", "membatasi", "presisi",
    "belakang", "koma", "kemudian", "dikonversi", "eksplisit", "mencegah", "kendala",
    "ketidaksesuaian", "eksekusi", "formula", "konsisten", "setiap", "jauh",
    "lebih", "komunikatif", "siap", "dimanfaatkan", "kebutuhan", "visualisasi",
    "maupun", "tanpa", "perlu", "melakukan", "berulang", "handal", "gratis",
    "alternatif", "terbaik", "perangkat", "lunak", "komersial", "dukungan", "komunitas",
    "krusial", "menambah", "memberikan", "konteks", "geografis", "akurat", "mendetail",
    "referensi", "koordinat", "posisi", "sebenarnya", "permukaan",
    "bumi", "membutuhkan", "pemahaman", "tersebar", "merata", "pemilihan", "metode",
    "ketelitian", "dinilai", "melalui", "semakin", "kecil", "baik",
    "kesesuaian", "pengukuran", "berbagai", "lainnya", "mempraktikkan", "pembuatan",
    "disempurnakan", "koreksi", "memakai", "simbol", "piktorial", "selain",
    "integrasi", "terbukti", "efektif", "menghadirkan", "interaktif", "berupa",
    "bangunan", "kursor", "diarahkan", "objek", "representatif", "informatif",
    "mudah", "dipahami", "membagi", "beberapa", "baru", "penerapan", "ketepatan",
    "memuat", "daftar", "pustaka", "universitas", "fakultas", "teknik", "jurusan",
    "poligon", "dokumen", "komputer", "digital", "geometri", "matematika", "aritmatika",
    "transformasi", "logika", "numerik", "komputasi", "variabel", "komprehensif", "visual",
    "klaster", "risiko", "ancaman", "mitigasi", "tanggul", "jalur", "evakuasi", "warga",
    "tumpang", "susun", "irisan", "penyangga", "sempadan", "banjir", "bandang",
    "serapan", "rujukan", "penelitian", "metodologi", "keseluruhan", "ruang", "kerangka"
}

# Frasa teknis multi-kata bahasa asing (di-italic secara utuh)
TECH_PHRASES = [
    "field calculator", "open field calculator", "attribute table", "open attribute table",
    "toggle editing mode", "toggle editing", "save edits", "functions group", "geometry functions",
    "basemap", "base map", "input layer", "shapefile", "geopackage",
    "svg marker", "proximity analysis", "overlay analysis", "split features",
    "vertex tool", "show map tips", "map tips", "quickmapservices",
    "nextgis quickmapservices", "esri standard", "ground control points",
    "open source", "string concatenation", "type mismatch", "whole number",
    "decimal number", "cookie cutter", "double click", "single click",
    "geoprocessing tools", "vector layer", "raster layer", "spatial join",
    "attribute table", "field calculator", "snapping tolerance", "topological editing",
    "linear gradient", "radial gradient", "angular gradient", "diamond gradient",
    "color scheme", "solid color", "image fill", "multiple fill", "color picker",
    "user interface", "design file", "frame design", "group selection",
    "drop shadow", "inner shadow", "layer blur", "background blur",
    "vector networks", "auto layout", "focal point", "color stops"
]

# Kata tunggal bahasa Inggris / istilah non-KBBI yang sering muncul
TECH_WORDS = {
    "field", "calculator", "attribute", "attributes", "table", "toggle", "editing", "mode", "save", "edits",
    "function", "functions", "group", "geometry", "basemap", "layer", "layers", "shapefile",
    "geopackage", "point", "polyline", "linestring", "marker", "svg", "buffer",
    "buffering", "clip", "intersect", "union", "dissolve", "proximity", "overlay",
    "split", "features", "vertex", "tool", "tools", "show", "tips", "quickmapservices",
    "resampling", "snapping", "canvas", "record", "ellipsoid", "source", "desktop",
    "software", "hardware", "real", "time", "string", "concatenation", "mismatch",
    "labeling", "whole", "number", "integer", "decimal", "date", "precision", "preview",
    "cookie", "cutter", "merge", "adjacent", "project", "shortcut", "plugin", "plugins",
    "click", "query", "update", "online", "offline", "output", "input", "import", "export",
    "file", "files", "folder", "default", "database", "layout", "render", "rendering",
    "open", "null", "boolean", "float", "double", "char", "varchar", "primary", "key",
    "foreign", "join", "spatial", "geoprocessing", "dock", "toolbox", "pop-up", "tips",
    "shapes", "shape", "figma", "frame", "frames", "fill", "fills", "stroke", "strokes",
    "rectangle", "arrow", "line", "ellipse", "polygon", "star", "gradient", "gradients",
    "angular", "diamond", "radial", "linear", "solid", "image", "images", "coloring",
    "contrast", "opacity", "artboard", "device", "navbar", "toolbar", "sidebar", "profile",
    "iphone", "android", "mobile", "tablet", "drag", "drop", "zoom"
}

# Inisialisasi Sastrawi jika tersedia
_stemmer = None
_sastrawi_words = set()
try:
    from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
    _stemmer = StemmerFactory().create_stemmer()
    _sastrawi_words = set(StemmerFactory().get_words())
except Exception:
    pass

def is_kbbi(word: str) -> bool:
    """Mengecek apakah kata termasuk kata dalam bahasa Indonesia/KBBI."""
    w = word.lower().strip(",.?!;:()\"'[]{}*`")
    if not w or w.isdigit():
        return True
    # Singkatan / akronim kapital (contoh: QGIS, SIG, CRS, UTM, WGS, BPBD, ID, HTML, dsb)
    if word.isupper() and len(word) >= 2:
        return True
    if w in KBBI_ALLOWLIST:
        return True
    if w in _sastrawi_words:
        return True
    if _stemmer:
        stemmed = _stemmer.stem(w)
        if stemmed in _sastrawi_words or stemmed in KBBI_ALLOWLIST:
            return True
    return False

def is_foreign_or_tech(word: str) -> bool:
    """Mengecek apakah kata merupakan istilah asing/teknis yang wajib dicetak miring."""
    w = word.lower().strip(",.?!;:()\"'[]{}*`")
    if not w:
        return False
    if w in TECH_WORDS:
        return True
    # Formula seperti $area, $length, $perimeter
    if w.startswith('$'):
        return True
    return not is_kbbi(w)

def parse_runs_for_italics(text: str):
    """
    Memecah teks menjadi potongan run berformat (text, is_bold, is_italic).
    Mendukung markdown manual (**bold**, *italic*, `code`) dan deteksi otomatis
    kata non-KBBI / frasa asing.
    """
    if not text:
        return []

    # 1. Pisahkan berdasarkan markdown syntax eksplisit
    pattern = re.compile(r'(\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*|`.*?`)')
    parts = pattern.split(text)
    
    sorted_phrases = sorted(TECH_PHRASES, key=lambda x: len(x), reverse=True)
    runs = []

    for part in parts:
        if not part:
            continue
        if part.startswith('***') and part.endswith('***') and len(part) >= 6:
            runs.append((part[3:-3], True, True))
        elif part.startswith('**') and part.endswith('**') and len(part) >= 4:
            # Bold teks, deteksi frasa non-KBBI di dalamnya
            subparts = _detect_spans(part[2:-2], sorted_phrases)
            for sub_text, sub_italic in subparts:
                runs.append((sub_text, True, sub_italic))
        elif part.startswith('*') and part.endswith('*') and len(part) >= 2:
            runs.append((part[1:-1], False, True))
        elif part.startswith('`') and part.endswith('`') and len(part) >= 2:
            runs.append((part[1:-1], False, True))
        else:
            # Plain text - deteksi otomatis kata/frasa asing
            subparts = _detect_spans(part, sorted_phrases)
            for sub_text, sub_italic in subparts:
                runs.append((sub_text, False, sub_italic))

    return runs

def _detect_spans(text: str, sorted_phrases):
    """Mendeteksi span teks yang harus di-italic karena frasa teknis atau kata asing."""
    if not text:
        return []
    
    lower_text = text.lower()
    matched_ranges = []

    # A. Cek frasa multi-kata terlebih dahulu
    for phrase in sorted_phrases:
        pattern = r'\b' + re.escape(phrase) + r'\b'
        for m in re.finditer(pattern, lower_text):
            m_start, m_end = m.start(), m.end()
            conflict = False
            for s, e in matched_ranges:
                if max(s, m_start) < min(e, m_end):
                    conflict = True
                    break
            if not conflict:
                matched_ranges.append((m_start, m_end))

    # B. Cek kata tunggal
    word_pattern = re.compile(r'(\$[a-zA-Z0-9_]+|[a-zA-Z0-9_]+)')
    for m in word_pattern.finditer(text):
        m_start, m_end = m.start(), m.end()
        # Lewati jika sudah termasuk dalam frasa
        if any(s <= m_start and m_end <= e for s, e in matched_ranges):
            continue
        word = m.group()
        w_lower = word.lower()
        if word.isdigit():
            continue
        if word.isupper() and len(word) >= 2:
            continue
        if word.startswith('$'):
            matched_ranges.append((m_start, m_end))
            continue
        if is_foreign_or_tech(word):
            matched_ranges.append((m_start, m_end))

    # C. Urutkan dan gabungkan range yang bertetangga/tumpang tindih
    matched_ranges.sort()
    merged_ranges = []
    for s, e in matched_ranges:
        if not merged_ranges:
            merged_ranges.append((s, e))
        else:
            prev_s, prev_e = merged_ranges[-1]
            if s <= prev_e:
                merged_ranges[-1] = (prev_s, max(prev_e, e))
            else:
                merged_ranges.append((s, e))

    # D. Potong teks menjadi run berurutan
    result = []
    curr = 0
    for s, e in merged_ranges:
        if s > curr:
            result.append((text[curr:s], False))
        result.append((text[s:e], True))
        curr = e
    if curr < len(text):
        result.append((text[curr:], False))

    return result
