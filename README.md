# Kontrak Perkuliahan OBE - Antigravity Skill

Modul kecerdasan buatan (*custom skill*) untuk Google Antigravity guna menyusun, menata, dan menerbitkan dokumen **Kontrak Perkuliahan** resmi berstandar kurikulum **OBE (*Outcome-Based Education*)** perguruan tinggi Indonesia (spesifik standar mutu Universitas PGRI Ronggolawe Tuban) secara otomatis dari dokumen RPS (*Rencana Pembelajaran Semester*).

---

## 🌟 Fitur Utama

1. **Universal RPS Parser**:
   - Mendukung input dokumen RPS baik format Word (`.docx`) maupun Markdown (`.md`).
   - Ekstraksi otomatis identitas mata kuliah, dosen pengampu, koordinator RMK, kaprodi, deskripsi, CPMK 1–4, silabus 16 minggu, beban belajar SN-Dikti, dan pustaka rujukan.

2. **Penataan Tabel Word XML Mutakhir (Advanced Formatting)**:
   - **Lebar Kolom Absolut (`w:tblGrid` & `w:tcW w:type="dxa"`)**: Terkunci presisi 6.10 inci sesuai batas cetak A4 Portrait.
   - **Anti-Split Baris (`<w:cantSplit/>`)**: Mencegah baris tabel terpotong canggung saat pergantian halaman.
   - **Repeating Header (`<w:tblHeader/>`)**: Baris judul tabel silabus 16 minggu berulang otomatis di halaman kedua.
   - **Perataan Vertikal Sel (`<w:vAlign/>`)**: Rata tengah vertikal untuk nomor, kode, skor, bobot %, dan rata atas untuk materi/tugas multiline.
   - **Paragraph Spacing Bersih**: `space_before = Pt(0)` dan `space_after = Pt(0)` dengan `line_spacing = 1.05`.

3. **Kepatuhan Regulasi & Standar Kampus**:
   - Formula Nilai Akhir: $\text{NA} = (\text{P} + 2\cdot\text{TGS} + 3\cdot\text{UTS} + 4\cdot\text{UAS}) / 10$.
   - Matriks silabus 16 minggu: Minggu 8 = UTS (30%), Minggu 16 = UAS (40%), total bobot tepat 100%.
   - Standar Penilaian Skala Resmi UNIROW: A (4), AB (3,5), B (3), BC (2,5), C (2, batas lulus), D (1, tidak lulus), E (0, tidak lulus).
   - Klausul kelulusan dan aturan mengulang perkuliahan resmi Pedoman Akademik UNIROW.
   - Pengesahan Tripartit resmi (Komti Mahasiswa, Dosen Pengampu, dan Ketua Program Studi).

---

## 📁 Struktur Berkas

```text
kontrak-kuliah-rps/
├── SKILL.md                          # Panduan & metadata agen Antigravity
├── README.md                         # Dokumentasi repository
├── assets/
│   └── logo_unirow.png               # Logo resmi institusi untuk KOP Word
├── references/
│   ├── template-kontrak.md           # Acuan anatomi kontrak perkuliahan
│   └── cek-konsistensi.md            # Panduan QA & checklist validasi
└── scripts/
    ├── build_kontrak.py              # Generator per berkas RPS tunggal (.docx & .md)
    └── batch_generate_all.py         # Batch generator untuk seluruh semester
```

---

## 🚀 Prasyarat & Instalasi

### 1. Prasyarat Sistem
- Python 3.8 atau lebih baru.
- Library `python-docx`:
  ```bash
  pip install python-docx
  ```

### 2. Cara Pemasangan di Antigravity

#### Opsi A: Pasang pada Proyek Tertentu (Workspace Specific)
Buka terminal di root proyek Anda, lalu clone ke folder `.agents/skills/`:
```bash
git clone https://github.com/mariofahmi/skillISO.git .agents/skills/kontrak-kuliah-rps
```

#### Opsi B: Pasang Secara Global (Berlaku untuk Semua Proyek)
Clone ke folder konfigurasi global Antigravity:
* **Windows**:
  ```powershell
  git clone https://github.com/mariofahmi/skillISO.git "$env:USERPROFILE\.gemini\config\skills\kontrak-kuliah-rps"
  ```
* **Linux / macOS**:
  ```bash
  git clone https://github.com/mariofahmi/skillISO.git ~/.gemini/config/skills/kontrak-kuliah-rps
  ```

---

## 💻 Cara Penggunaan

### 1. Melalui Chat AI Antigravity
Cukup instruksikan agen di obrolan Antigravity:
> *"Tolong susun kontrak perkuliahan dari dokumen RPS mata kuliah Pancasila menggunakan skill /kontrak-kuliah-rps"*

### 2. Melalui Baris Perintah (CLI Langsung)
```powershell
# Hasilkan berkas Word (.docx) dan Markdown (.md) dari RPS:
py ".agents\skills\kontrak-kuliah-rps\scripts\build_kontrak.py" "path/ke/RPS.docx" --both

# Hasilkan kontrak perkuliahan untuk seluruh RPS di folder kurikulum:
py ".agents\skills\kontrak-kuliah-rps\scripts\batch_generate_all.py"
```

---

## 📄 Lisensi & Kontributor
- **Pengembang**: Mario Fahmi Syahrial, M.Pd. (Program Studi PPKn FKIP UNIROW Tuban)
- **Kompatibilitas**: Google Antigravity & Antigravity 2.0
