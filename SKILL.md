---
name: kontrak-kuliah-rps
description: >
  Skill untuk menyusun, memeriksa, memformat, dan menghasilkan dokumen Kontrak
  Perkuliahan (kontrak kuliah / kesepakatan belajar) resmi siap pakai berbasis
  dokumen RPS (Rencana Pembelajaran Semester) perguruan tinggi kurikulum OBE
  (Outcome-Based Education) di lingkungan Universitas PGRI Ronggolawe (UNIROW)
  Tuban maupun perguruan tinggi Indonesia. Gunakan skill ini setiap kali pengguna
  menyebut "kontrak kuliah", "kontrak perkuliahan", "kesepakatan perkuliahan",
  "kontrak belajar", "buat kontrak kuliah dari RPS", atau meminta berkas untuk
  perkuliahan tatap muka pertama. Menyediakan script generator otomatis
  (scripts/build_kontrak.py) untuk menghasilkan file dokumen Word (.docx) ber-KOP
  resmi dan Markdown (.md) yang selaras presisi dengan RPS.
---

# KONTRAK KULIAH DARI RPS — PANDUAN RESMI PENYUSUNAN

Kontrak Perkuliahan adalah dokumen kesepakatan akademik antara Dosen Pengampu dan Mahasiswa yang disahkan pada pertemuan tatap muka pertama. Seluruh informasi inti (bobot nilai, jadwal 16 minggu, batas kelulusan, dan pustaka) **wajib diselaraskan secara akurat dengan dokumen RPS (Rencana Pembelajaran Semester)** agar tidak terjadi sengketa penilaian.

---

## 1. Spesifikasi Format & Tipografi Dokumen Resmi

| Parameter | Ketentuan Baku Standar Kampus |
|---|---|
| **Ukuran & Orientasi Kertas** | **A4 Portrait** (`21.0 × 29.7 cm` / `8.27 × 11.69 inch`) |
| **Batas Margin Halaman** | Atas: `2.5 cm` (0.98"), Bawah: `2.5 cm` (0.98"), Kiri: `3.0 cm` (1.18"), Kanan: `2.5 cm` (0.98") |
| **Lebar Area Cetak Efektif** | **6.10 inci** (8.784 dxa / 439.2 pt) |
| **Tipografi Utama** | **Cambria** (Isi & Tabel: 8.5–9.5 pt, Sub-judul: 11 pt bold, Judul Dokumen: 14 pt bold) |
| **KOP Resmi Lembaga** | Memuat logo resmi UNIROW (`assets/logo_unirow.png`), identitas institusi, dan garis pembatas ganda |
| **Garis Tabel (Borders)** | Single line warna `#B0B5B3`, ketebalan `sz="4"` (0.5 pt) |
| **Bantalan Sel (Padding)**| Top/Bottom: `50–60 dxa`, Left/Right: `70–100 dxa` |
| **Palet Warna Shading** | Header Tabel: `#EAECEE`, Baris Ujian (UTS/UAS): `#F4F6F7`, Label Kolom: `#F4F6F7` |

---

## 2. Standar Penataan Tabel Presisi (Word XML & Markdown)

Setiap tabel di dalam dokumen Kontrak Perkuliahan wajib mengikuti tata kelola XML dan Markdown berikut:

### A. Penguncian Grid Kolom & Lebar Mutlak (`w:tblGrid` & `w:tcW w:type="dxa"`)
Word secara bawaan dapat mengubah proporsi kolom secara acak jika hanya menggunakan properti sel biasa. Oleh karena itu, tabel wajib mengunci lebar kolom pada level `w:tblGrid` dan `w:tcW` (satuan *dxa*, 1 inci = 1.440 dxa):
1. **Tabel KOP Institusi** (2 kolom, total 6.10"):
   - Kolom 0 (Logo Kampus): `1.20"` (1.728 dxa)
   - Kolom 1 (Teks Institusi): `4.90"` (7.056 dxa)
2. **Tabel Identitas Mata Kuliah** (2 kolom, total 6.10"):
   - Kolom 0 (Label Identitas, Shading `#F4F6F7`): `2.20"` (3.168 dxa)
   - Kolom 1 (Nilai Identitas): `3.90"` (5.616 dxa)
3. **Tabel Jadwal Perkuliahan 16 Minggu** (4 kolom, total 6.10"):
   - Kolom 0 (Minggu): `0.65"` (936 dxa) — perataan tengah (`center`)
   - Kolom 1 (Materi Pokok / Bahasan): `2.75"` (3.960 dxa) — perataan atas (`top`)
   - Kolom 2 (Bentuk Pembelajaran & Penugasan): `2.10"` (3.024 dxa) — perataan atas (`top`)
   - Kolom 3 (Bobot): `0.60"` (864 dxa) — perataan tengah (`center`)
4. **Tabel Evaluasi & Komposisi Nilai** (3 kolom, total 6.10"):
   - Kolom 0 (No): `0.55"` (792 dxa) — perataan tengah
   - Kolom 1 (Komponen Evaluasi): `4.75"` (6.840 dxa) — perataan tengah vertikal
   - Kolom 2 (Bobot %): `0.80"` (1.152 dxa) — perataan tengah
5. **Tabel Standar Konversi Skala 7** (5 kolom, total 6.10"):
   - Kolom 0 (Rentang Skor): `1.10"` (1.584 dxa)
   - Kolom 1 (Nilai Huruf): `0.85"` (1.224 dxa)
   - Kolom 2 (Bobot Mutu): `0.85"` (1.224 dxa)
   - Kolom 3 (Kategori Capaian): `2.30"` (3.312 dxa)
   - Kolom 4 (Status Kelulusan): `1.00"` (1.440 dxa)
6. **Tabel Pengesahan & Tanda Tangan** (2 kolom, total 6.10"):
   - Kolom 0 (Perwakilan Mahasiswa / Komti): `3.05"` (4.392 dxa)
   - Kolom 1 (Dosen Pengampu MK): `3.05"` (4.392 dxa)

### B. Proteksi Pergantian Halaman (`w:cantSplit`)
Seluruh baris tabel (`<w:tr>`) wajib disematkan tag XML `<w:cantSplit/>` pada `w:trPr`. Hal ini mencegah baris terpotong secara canggung menjadi dua bagian di perbatasan halaman.

### C. Baris Header Berulang di Tiap Halaman (`w:tblHeader`)
Tabel multi-halaman (khususnya Silabus 16 Minggu) wajib menyematkan `<w:tblHeader/>` pada baris header (`hdr_row._tr.get_or_add_trPr()`) sehingga judul kolom otomatis muncul kembali di bagian atas halaman berikutnya.

### D. Perataan Vertikal Sel (`w:vAlign`)
- Sel nomor minggu, kode, bobot, dan header menggunakan `<w:vAlign w:val="center"/>`.
- Sel deskripsi materi dan tugas yang memiliki teks panjang menggunakan `<w:vAlign w:val="top"/>`.

### E. Kebersihan Paragraf Sel
Setiap paragraf di dalam sel tabel wajib disetel `space_before = Pt(0)` dan `space_after = Pt(0)` dengan `line_spacing = 1.05` agar bantalan sel terlihat proporsional dan tidak terjadi lonjakan tinggi baris.

### F. Format Penataan Markdown Tables
Di dalam berkas Markdown (`.md`), setiap kolom tabel wajib memiliki format *header alignment separator*:
- Kolom teks: `:---|` (rata kiri).
- Kolom angka, bobot, nilai huruf, nomor minggu: `:---:|` (rata tengah).

---

## 3. Struktur Anatomi Kontrak Perkuliahan

Setiap dokumen Kontrak Perkuliahan wajib memuat komponen berurutan sebagai berikut:

1. **KOP Resmi Institusi**:
   - Logo Universitas PGRI Ronggolawe Tuban (kiri).
   - Teks nama Universitas, Fakultas Keguruan dan Ilmu Pendidikan, serta Program Studi PPKn (tengah).
   - Alamat kampus, nomor telepon, dan email resmi.
   - Garis pembatas ganda (*double horizontal line divider*).
2. **Judul Dokumen**:
   - `KONTRAK PERKULIAHAN`.
   - Nama Mata Kuliah, Kode Mata Kuliah, Bobot SKS, Semester, dan Tahun Akademik.
3. **Bagian A: Identitas Mata Kuliah**:
   - Tabel terstruktur memuat Nama MK, Kode/Bobot, Rumpun MK, Semester/TA, Dosen Pengampu, Koordinator RMK, Ketua Program Studi, Jadwal/Ruang (titik-titik untuk diisi saat kuliah pertama), dan Prasyarat MK.
4. **Bagian B: Deskripsi Singkat Mata Kuliah**:
   - Narasi ringkas profil kompetensi dan substansi kajian mata kuliah yang dipadatkan dari RPS.
5. **Bagian C: Capaian Pembelajaran Mata Kuliah (CPMK)**:
   - Rumusan kompetensi akhir mahasiswa bernomor urut (CPMK 1 s.d. CPMK 4).
6. **Bagian D: Materi Pokok dan Jadwal Perkuliahan 16 Minggu**:
   - Tabel matriks 4 kolom: `Minggu` \| `Materi Pokok / Bahasan` \| `Bentuk Pembelajaran & Penugasan` \| `Bobot`.
   - **Minggu 8** secara baku adalah **EVALUASI TENGAH SEMESTER (UTS)** (bobot 25%, shading `#F4F6F7`).
   - **Minggu 16** secara baku adalah **EVALUASI AKHIR SEMESTER (UAS)** (bobot 25%, shading `#F4F6F7`).
7. **Bagian E: Metode Pembelajaran dan Alokasi Beban Belajar (SN-Dikti)**:
   - Sesuai Permendikbudristek No. 53 Tahun 2023:
     - Tatap Muka / PB: $sks \times 50$ menit/minggu (2 SKS = 100 menit).
     - Terstruktur / PT: $sks \times 60$ menit/minggu (2 SKS = 120 menit).
     - Mandiri / KM: $sks \times 60$ menit/minggu (2 SKS = 120 menit).
     - Total Beban Belajar: 340 menit (5,67 jam/minggu untuk 2 SKS).
8. **Bagian F: Sistem Evaluasi, Bobot, dan Konversi Nilai Akhir**:
   - Tabel Komposisi Nilai: Aktivitas Partisipatif (15%), Tugas Terstruktur & Proyek (20%), Kuis (15%), UTS (25%), UAS (25%) = Total 100%.
   - Tabel Konversi Skala 7 Resmi UNIROW Tuban: A (85–100), B+ (75–84), B (68–74), C+ (65–67), C (56–64, batas lulus), D (40–55), E (0–39).
9. **Bagian G: Tata Tertib dan Kesepakatan Perkuliahan**:
   - Klausul terstruktur: Kehadiran minimal 75%, toleransi keterlambatan 15 menit, etika akademik anti-bullying, sanksi keterlambatan tugas 10%/hari, integritas bebas plagiarisme, dan syarat ujian susulan sah.
10. **Bagian H: Pustaka Rujukan**:
    - Pustaka Utama dan Pustaka Pendukung yang disalin persis dari RPS.
11. **Bagian I: Pernyataan Kesepakatan dan Pengesahan**:
    - Kalimat kesepakatan sadar dan sukarela, tempat dan tanggal penetapan.
    - Kolom tanda tangan simetris: **Perwakilan Mahasiswa (Ketua Tingkat / Komti)** dan **Dosen Pengampu MK**.
    - Mengetahui: **Ketua Program Studi PPKn (Mario Fahmi Syahrial, M.Pd.)**.

---

## 4. Alur Kerja dan Eksekusi Generator Otomatis

Script generator terintegrasi:
- `scripts/build_kontrak.py`: Generator per berkas RPS tunggal.
- `scripts/batch_generate_all.py`: Batch generator untuk seluruh RPS (Semester 1 s.d. 7).
- Aset logo resmi institusi: `assets/logo_unirow.png`

Script ini mendukung input berkas RPS langsung baik dalam format **Word (`.docx`)** maupun **Markdown (`.md`)**, serta mampu menghasilkan dokumen Word (`.docx`) ber-KOP resmi dan Markdown (`.md`) secara simultan.

### Sintaks Perintah CLI:

```powershell
# 1. Menghasilkan berkas .docx dan .md sekaligus dari RPS Word (.docx):
py ".agents\skills\kontrak-kuliah-rps\scripts\build_kontrak.py" "<path_ke_rps.docx>" --both

# 2. Menghasilkan berkas .docx dan .md sekaligus dari RPS Markdown (.md):
py ".agents\skills\kontrak-kuliah-rps\scripts\build_kontrak.py" "<path_ke_rps.md>" --both

# 3. Menghasilkan Kontrak Perkuliahan untuk SEMUA RPS sekaligus:
py ".agents\skills\kontrak-kuliah-rps\scripts\batch_generate_all.py"
```

### Opsi CLI yang Didukung:
- `rps_file` : Argumen posisional path ke berkas RPS (`.docx` atau `.md`).
- `--rps`     : Opsi path ke berkas RPS.
- `--output`, `-o` : Path output berkas Word (`.docx`).
- `--output-md`    : Path output berkas Markdown (`.md`).
- `--both`         : Flag otomatis untuk menghasilkan berkas `.docx` sekaligus `.md`.

### Fitur Otomasi Script:
1. **Universal Parser**: Menguraikan identitas MK, dosen, koordinator RMK, Kaprodi Mario Fahmi Syahrial, M.Pd., CPMK 1–4, silabus 16 minggu, beban belajar SN-Dikti, dan pustaka tanpa distorsi formatting.
2. **Presisi Penataan Tabel Word XML**:
   - Mengunci lebar kolom pada level `w:tblGrid` dan `w:tcW w:type="dxa"` (total area cetak 6.10 inci).
   - Menyematkan `<w:cantSplit/>` pada seluruh baris tabel agar tidak terpotong canggung saat ganti halaman.
   - Menyematkan `<w:tblHeader/>` pada baris judul tabel agar otomatis berulang di halaman lanjutan.
   - Mengatur perataan vertikal `<w:vAlign/>` (center untuk angka/bobot, top untuk uraian materi & tugas).
   - Menghilangkan *spacing* ekstra paragraf sel (`space_before=0`, `space_after=0`).
3. **Smart Naming & Routing**: Otomatis mendeteksi nomor semester mata kuliah dan menempatkan hasil di subfolder `KONTRAK_KULIAH`.
4. **Clean Sanitization**: Membersihkan seluruh tanda bintang markdown mentah (`**`) dan karakter kontrol biner agar berkas Word 100% siap cetak dan tandatangan.

---

## 5. Daftar Cek Kualitas Kontrak Kuliah (QA Checklist)

Sebelum menyerahkan berkas kontrak perkuliahan kepada dosen atau pengguna, pastikan seluruh kriteria berikut terpenuhi:

- [ ] **Orientasi & Margin**: A4 Portrait (`8.27 × 11.69 inci`), margin kiri 3.0 cm, atas/bawah/kanan 2.5 cm (area cetak 6.10 inci).
- [ ] **Tipografi**: Konsisten menggunakan font **Cambria** di seluruh heading, paragraf, dan tabel.
- [ ] **KOP Resmi Institusi**: Memuat logo resmi UNIROW Tuban (`logo_unirow.png`) dan garis pembatas ganda (*double bottom border*).
- [ ] **Kerapian Tabel & Baris**:
  - Kolom terkunci dengan lebar absolut `dxa` (tidak melar atau geser).
  - Baris tabel tidak terpotong di perbatasan halaman (`<w:cantSplit/>`).
  - Baris judul tabel silabus 16 minggu berulang di halaman kedua (`<w:tblHeader/>`).
  - Nomor minggu, skor, bobot, dan status rata tengah vertikal (`<w:vAlign w:val="center"/>`).
- [ ] **Silabus 16 Minggu Lengkap**: Matriks perkuliahan memuat Minggu 1 s.d. 16.
- [ ] **Jadwal Evaluasi Baku**: Minggu ke-8 adalah **UTS** (bobot 25%) dan Minggu ke-16 adalah **UAS** (bobot 25%).
- [ ] **Total Bobot Kumulatif**: Tepat berjumlah **100%**.
- [ ] **Alokasi Waktu SN-Dikti**: Menghitung beban belajar berdasarkan SKS ($sks \times 170$ menit/minggu: 50 menit tatap muka, 60 menit terstruktur, 60 menit mandiri).
- [ ] **Tabel Skala 7 UNIROW**: Mengadopsi standar penilaian Skala 7 (A, B+, B, C+, C, D, E) dengan batas minimal kelulusan 56.00 (C).
- [ ] **Tata Tertib Akademik**: Memuat klausul presensi minimal 75%, toleransi keterlambatan 15 menit, dan sanksi plagiarisme.
- [ ] **Lembar Pengesahan Resmi**: Memuat kolom Komti Mahasiswa (kiri), Dosen Pengampu (kanan), dan Mengetahui Ketua Program Studi PPKn **Mario Fahmi Syahrial, M.Pd.** (bawah tengah).
- [ ] **Kebersihan Teks**: Bebas dari karakter bintang markdown mentah (`**`) dan karakter biner tidak terduga.

---

## 6. Berkas Referensi Terkait
- [Template Struktur Kontrak Kuliah](file:///d:/ANTY%20GRAVITY/OBE%202026/.agents/skills/kontrak-kuliah-rps/references/template-kontrak.md)
- [Panduan Cek Konsistensi RPS](file:///d:/ANTY%20GRAVITY/OBE%202026/.agents/skills/kontrak-kuliah-rps/references/cek-konsistensi.md)
- [Script Generator build_kontrak.py](file:///d:/ANTY%20GRAVITY/OBE%202026/.agents/skills/kontrak-kuliah-rps/scripts/build_kontrak.py)
- [Script Batch batch_generate_all.py](file:///d:/ANTY%20GRAVITY/OBE%202026/.agents/skills/kontrak-kuliah-rps/scripts/batch_generate_all.py)
- [Aset Logo Resmi UNIROW](file:///d:/ANTY%20GRAVITY/OBE%202026/.agents/skills/kontrak-kuliah-rps/assets/logo_unirow.png)
