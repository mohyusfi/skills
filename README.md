# Yusfi's Agent Skills 🚀

Kumpulan custom agent skills untuk Google Antigravity, Claude Code, GitHub Copilot, dan Cursor.

Repository ini menggunakan pola monorepo terpusat untuk mempermudah pemeliharaan, pengembangan, dan distribusi skills.

---

## 📦 Katalog Skills

| Skill | Deskripsi | Perintah Instalasi |
| :--- | :--- | :--- |
| **[`yuss-males-laprak`](./skills/yuss-males-laprak)** | Pembuat laporan praktikum Word otomatis sesuai format akademik (margin 4-4-3-3 cm, Times New Roman 12 pt, 1.5 spasi, heading terintegrasi Navigation Pane, pembersihan foto modul asdos). | `npx skills add mohyusfi/skills --skill yuss-males-laprak` |

---

## 🚀 Cara Instalasi

### 1. Menginstal Skill Tertentu
Gunakan opsi `--skill` untuk memasang satu skill spesifik:
```bash
npx skills add mohyusfi/skills --skill yuss-males-laprak
```
*Atau menggunakan notasi `@`:*
```bash
npx skills add mohyusfi/skills@yuss-males-laprak
```

### 2. Menginstal Semua Skill Sekaligus
```bash
npx skills add mohyusfi/skills --all
```

### 3. Memperbarui Skills ke Versi Terbaru
```bash
npx skills update
```

---

## 🛠️ Pengembangan (Untuk Kontributor / Pemilik)

1. Clone repositori ini:
   ```bash
   git clone https://github.com/mohyusfi/skills.git
   ```
2. Buat folder skill baru di dalam folder `skills/<nama-skill>/` lengkap dengan `SKILL.md`.
3. Commit dan push ke branch `main`:
   ```bash
   git add .
   git commit -m "feat: tambah skill baru"
   git push origin main
   ```
4. Skill baru akan langsung tersedia dan bisa diinstal melalui `npx skills add mohyusfi/skills --skill <nama-skill>`.

---

## 📄 Lisensi
[MIT License](./LICENSE)
