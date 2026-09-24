#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_kontrak.py - Generator Dokumen Kontrak Perkuliahan Standar Resmi UNIROW Tuban
Mengekstrak data dari RPS (baik berkas .docx maupun .md) dan menyusun berkas
Kontrak Perkuliahan lengkap berformat .docx (A4 Portrait, Tipografi Cambria,
KOP Resmi berlogo, Border #B0B5B3) dan .md (Markdown GFM rapi).
"""

import os
import sys
import re
import argparse
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(SCRIPT_DIR)
ASSETS_DIR = os.path.join(SKILL_DIR, 'assets')
LOGO_FILE = os.path.join(ASSETS_DIR, 'logo_unirow.png')

# ---------------------------------------------------------
# XML & Styling Helpers
# ---------------------------------------------------------

def clean_str(text):
    """Membersihkan karakter kontrol XML yang tidak diizinkan."""
    if text is None:
        return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(text))

def clean_name(name):
    """Membersihkan nama dosen/pejabat dari tanda bintang dan placeholder NIDN."""
    if not name:
        return ""
    n = clean_str(name).replace('**', '').replace('*', '').replace('<br>', ' ').strip()
    n = re.sub(r'\s*NIDN\b.*$', '', n, flags=re.IGNORECASE)
    return n.strip()

def set_cell_shading(cell, color_hex):
    """Menyetel warna latar sel tabel."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('shd'):
            tcPr.remove(child)
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_padding(cell, top=70, bottom=70, left=120, right=120):
    """Menyetel bantalan (margin dalam) sel tabel dalam satuan dxa."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{edge}')
        node.set(qn('w:w'), str(val))
        tcMar.append(node)
    tcPr.append(tcMar)

def set_cell_borders(cell, top='single', bottom='single', left='single', right='single', color='B0B5B3', sz='4'):
    """Menyetel garis batas sel tabel standar UNIROW (#B0B5B3, sz 4 = 0.5pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        if val:
            b_elem = OxmlElement(f'w:{edge}')
            b_elem.set(qn('w:val'), val)
            b_elem.set(qn('w:sz'), sz)
            b_elem.set(qn('w:space'), '0')
            b_elem.set(qn('w:color'), color)
            tcBorders.append(b_elem)
        else:
            b_elem = OxmlElement(f'w:{edge}')
            b_elem.set(qn('w:val'), 'none')
            tcBorders.append(b_elem)
    tcPr.append(tcBorders)

def set_row_cant_split(row):
    """Mencegah baris tabel terpotong di batas pergantian halaman (cantSplit)."""
    trPr = row._tr.get_or_add_trPr()
    if not trPr.xpath('w:cantSplit'):
        cantSplit = parse_xml(f'<w:cantSplit {nsdecls("w")}/>')
        trPr.append(cantSplit)

def set_repeat_header(row):
    """Mengulang baris judul/header tabel secara otomatis saat berpindah halaman (tblHeader)."""
    trPr = row._tr.get_or_add_trPr()
    if not trPr.xpath('w:tblHeader'):
        tblHeader = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
        trPr.append(tblHeader)

def set_cell_valign(cell, val="center"):
    """Menyetel perataan vertikal teks dalam sel tabel (center/top/bottom)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith('vAlign'):
            tcPr.remove(child)
    vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{val}"/>')
    tcPr.append(vAlign)

def set_table_col_widths(table, col_widths_in_inches):
    """Mengunci lebar kolom tabel secara presisi pada level tblGrid dan tcW dxa."""
    table.autofit = False
    tblGrid = table._tbl.tblGrid
    for c in list(tblGrid):
        tblGrid.remove(c)
    for w_in in col_widths_in_inches:
        w_dxa = int(w_in * 1440)
        gridCol = parse_xml(f'<w:gridCol {nsdecls("w")} w:w="{w_dxa}"/>')
        tblGrid.append(gridCol)
    for row in table.rows:
        for i, w_in in enumerate(col_widths_in_inches):
            if i < len(row.cells):
                cell = row.cells[i]
                w_dxa = int(w_in * 1440)
                tcPr = cell._tc.get_or_add_tcPr()
                for child in list(tcPr):
                    if child.tag.endswith('tcW'):
                        tcPr.remove(child)
                tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{w_dxa}" w:type="dxa"/>')
                tcPr.append(tcW)

def add_paragraph_run(p, text, font_size=10, bold=False, italic=False, color_rgb=None):
    """Menambahkan run dengan font Cambria dan format yang konsisten."""
    clean = clean_str(text)
    run = p.add_run(clean)
    run.font.name = 'Cambria'
    run.font.size = Pt(font_size)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = color_rgb
    return run

def add_double_bottom_divider(paragraph, color='003366', sz='12', space='8'):
    """Menambahkan garis pembatas ganda (double line) khas kop resmi institusi."""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="{sz}" w:space="{space}" w:color="{color}"/></w:pBdr>')
    pPr.append(pBdr)

# ---------------------------------------------------------
# Data Extraction: Markdown (.md)
# ---------------------------------------------------------

def parse_rps_md(md_path):
    """
    Ekstraksi data terstruktur dari berkas RPS berformat Markdown (.md).
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        text = clean_str(f.read())

    data = {
        'nama_mk': '',
        'kode_mk': '',
        'sks': '2',
        'semester': '',
        'rumpun_mk': 'Mata Kuliah Keilmuan dan Keterampilan (MKK PPKn)',
        'dosen_pengampu': 'Tim Dosen Pengampu Program Studi PPKn UNIROW Tuban',
        'koordinator_rmk': 'Dwi Wahyu Kartikasari, M.Pd.',
        'ka_prodi': 'Mario Fahmi Syahrial, M.Pd.',
        'prasyarat': 'Tidak ada',
        'deskripsi': '',
        'cpmk': [],
        'jadwal': [],
        'pustaka_utama': [],
        'pustaka_pendukung': []
    }

    # 1. Identitas Mata Kuliah
    mk_m = re.search(r'\|\s*\*{0,2}MATA\s+KULIAH\s*(?:\(MK\))?\*{0,2}\s*\|\s*\*{0,2}([^\*\n|]+?)\*{0,2}\s*\|', text, re.IGNORECASE)
    if mk_m:
        data['nama_mk'] = mk_m.group(1).strip()
    else:
        t_m = re.search(r'#\s*RENCANA\s+PEMBELAJARAN\s+SEMESTER.*?##\s*MATA\s+KULIAH\s*:\s*([^\n]+)', text, re.IGNORECASE)
        if t_m:
            data['nama_mk'] = t_m.group(1).strip()

    kd_m = re.search(r'\|\s*\*{0,2}KODE(?:\s+MK)?\*{0,2}\s*\|\s*\*{0,2}([^\*\n|]+?)\*{0,2}\s*\|', text, re.IGNORECASE)
    if kd_m:
        data['kode_mk'] = kd_m.group(1).strip()

    sks_m = re.search(r'\|\s*\*{0,2}BOBOT\s*(?:\(sks\))?\*{0,2}\s*\|\s*\*{0,2}(\d+)', text, re.IGNORECASE)
    if sks_m:
        data['sks'] = sks_m.group(1).strip()

    sem_m = re.search(r'\|\s*\*{0,2}SEMESTER\*{0,2}\s*\|\s*\*{0,2}([^\*\n|]+?)\*{0,2}\s*\|', text, re.IGNORECASE)
    if sem_m:
        data['semester'] = sem_m.group(1).strip()

    rmk_m = re.search(r'\|\s*\*{0,2}Rumpun\s+MK\*{0,2}\s*\|\s*\*{0,2}([^\*\n|]+?)\*{0,2}\s*\|', text, re.IGNORECASE)
    if rmk_m:
        data['rumpun_mk'] = rmk_m.group(1).strip()

    # Otorisasi Dosen, RMK, Ka Prodi
    oto_block = re.search(r'##\s*OTORISASI[^\n]*\n+(.+?)(?=\n##|\Z)', text, re.DOTALL | re.IGNORECASE)
    if oto_block:
        lines = [l.strip() for l in oto_block.group(1).split('\n') if l.strip() and '|' in l and not re.match(r'^[|\s:-]+$', l)]
        if lines:
            last_cols = [clean_name(c) for c in lines[-1].split('|')[1:-1]]
            if len(last_cols) >= 3:
                if last_cols[0]: data['dosen_pengampu'] = last_cols[0]
                if last_cols[1]: data['koordinator_rmk'] = last_cols[1]
                if last_cols[2]: data['ka_prodi'] = last_cols[2]

    dp_m = re.search(r'###?\s*Dosen\s+Pengampu[^\n]*\n+([^\n#]+)', text, re.IGNORECASE)
    if dp_m:
        val = clean_name(dp_m.group(1).lstrip('-* '))
        if val: data['dosen_pengampu'] = val

    ms_m = re.search(r'###?\s*Mata\s*kuliah\s+Syarat[^\n]*\n+([^\n#]+)', text, re.IGNORECASE)
    if ms_m:
        data['prasyarat'] = ms_m.group(1).strip().lstrip('-* ').strip()

    # 2. Deskripsi Singkat MK
    desc_m = re.search(r'(?:###?\s*Deskripsi\s+Singkat[^\n]*|\*{1,2}Deskripsi\s+Singkat[^\n]*\*{0,2}:?)[ \t]*\n+(.+?)(?=\n\s*(?:###|##|\*{1,2}(?:Bahan|BK\d+|Materi|Pustaka)|\bBK\d+)\b|\Z)', text, re.DOTALL | re.IGNORECASE)
    if desc_m:
        data['deskripsi'] = ' '.join([l.strip() for l in desc_m.group(1).strip().split('\n') if l.strip()])

    # 3. CPMK
    cpmk_block = re.search(r'###?\s*Capaian\s+Pembelajaran\s+Mata\s+Kuliah\s*\(CPMK\)[^\n]*\n+(.+?)(?=\n###|\n##|\Z)', text, re.DOTALL | re.IGNORECASE)
    if cpmk_block:
        for line in cpmk_block.group(1).split('\n'):
            line = line.strip().lstrip('-* ')
            if line.startswith('CPMK') or re.match(r'^\d+\.', line):
                parts = re.split(r'[:\t]\s*', line, maxsplit=1)
                if len(parts) == 2:
                    code = parts[0].replace('**', '').strip()
                    desc = parts[1].replace('**', '').strip()
                    data['cpmk'].append((code, desc))
                else:
                    data['cpmk'].append(('CPMK', line.replace('**', '').strip()))

    # 4. Jadwal Mingguan 1-16
    mg_block = re.search(r'##\s*RENCANA\s+(?:KEGIATAN\s+PEMBELAJARAN|PEMBELAJARAN\s+MINGGUAN).*?\n(.*?)(?=\n##\s+[A-Z]|\Z)', text, re.DOTALL | re.IGNORECASE)
    if mg_block:
        lines = [l.strip() for l in mg_block.group(1).split('\n') if l.strip() and '|' in l and not re.match(r'^[|\s:-]+$', l)]
        header_cols = []
        if lines:
            header_cols = [c.strip().replace('**', '').lower() for c in lines[0].split('|')[1:-1]]

        idx_bentuk = -1
        idx_materi = -1
        for i, h in enumerate(header_cols):
            if any(k in h for k in ['bentuk', 'metode', 'aktivitas']):
                if idx_bentuk == -1: idx_bentuk = i
            if any(k in h for k in ['materi pembelajaran', 'bahan kajian', 'pokok bahasan', 'materi']):
                idx_materi = i

        seen_weeks = set()
        for l in lines[1:]:
            cols = [c.strip() for c in l.split('|')[1:-1]]
            if not cols: continue
            col0_clean = cols[0].replace('**', '').replace('*', '').strip()
            if col0_clean.startswith('(') or 'MG KE' in col0_clean.upper() or not col0_clean:
                continue
            mg_match = re.search(r'\b(\d{1,2})\b', col0_clean)
            if not mg_match:
                if re.search(r'\bUTS\b', col0_clean.upper()) or 'TENGAH SEMESTER' in col0_clean.upper():
                    mg_no = 8
                elif re.search(r'\bUAS\b', col0_clean.upper()) or 'AKHIR SEMESTER' in col0_clean.upper():
                    mg_no = 16
                else:
                    continue
            else:
                mg_no = int(mg_match.group())

            if mg_no < 1 or mg_no > 16:
                continue
            if mg_no in seen_weeks:
                continue
            seen_weeks.add(mg_no)

            bobot_str = cols[-1].replace('**', '').replace('%', '').strip()
            bobot_val = f"{bobot_str}%" if bobot_str else ""

            is_uts = (mg_no == 8)
            is_uas = (mg_no == 16)

            if is_uts:
                data['jadwal'].append({
                    'minggu': '8',
                    'topik': 'EVALUASI TENGAH SEMESTER (UTS)',
                    'bentuk_tugas': 'Ujian Tertulis Teoretis & Analisis Instrumen (Komprehensif Minggu 1 s.d. 7)',
                    'bobot': bobot_val or '25%'
                })
            elif is_uas:
                data['jadwal'].append({
                    'minggu': '16',
                    'topik': 'EVALUASI AKHIR SEMESTER (UAS)',
                    'bentuk_tugas': 'Ujian Kasus Terjadwal Komprehensif & Presentasi Portofolio Proyek',
                    'bobot': bobot_val or '25%'
                })
            else:
                sub_txt = cols[1].replace('**', '').strip() if len(cols) > 1 else ""
                clean_sub = re.sub(r'^Sub-CPMK\s*\d+\s*(?:\[[^\]]*\])?\s*:\s*', '', sub_txt, flags=re.IGNORECASE)

                materi_txt = ""
                if idx_materi != -1 and idx_materi < len(cols):
                    materi_txt = cols[idx_materi].replace('**', '').replace('<br>', '; ').strip()
                    materi_txt = re.sub(r'\[Pustaka[^\]]*\]', '', materi_txt, flags=re.IGNORECASE).strip()
                elif len(cols) >= 5:
                    materi_txt = cols[-2].replace('**', '').replace('<br>', '; ').strip()
                    materi_txt = re.sub(r'\[Pustaka[^\]]*\]', '', materi_txt, flags=re.IGNORECASE).strip()

                if materi_txt and materi_txt != '-' and len(materi_txt) > 3:
                    clean_topik = materi_txt
                else:
                    clean_topik = clean_sub

                bentuk_tugas = ""
                if idx_bentuk != -1 and idx_bentuk < len(cols):
                    b_parts = [cols[idx_bentuk]]
                    if idx_bentuk + 1 < len(cols) and (idx_materi == -1 or idx_bentuk + 1 != idx_materi) and idx_bentuk + 1 != len(cols) - 1:
                        if 'daring' in (header_cols[idx_bentuk+1] if idx_bentuk+1 < len(header_cols) else ''):
                            b_parts.append(cols[idx_bentuk+1])
                    bentuk_tugas = '; '.join([p.replace('<br>', '; ').replace('**', '').strip() for p in b_parts if p.strip()])
                elif len(cols) >= 6:
                    bentuk_tugas = cols[3].replace('<br>', '; ').replace('**', '').strip()
                elif len(cols) >= 5:
                    bentuk_tugas = cols[2].replace('<br>', '; ').replace('**', '').strip()
                else:
                    bentuk_tugas = "Kuliah Tatap Muka & Diskusi Terstruktur"

                data['jadwal'].append({
                    'minggu': str(mg_no),
                    'topik': clean_topik,
                    'bentuk_tugas': bentuk_tugas,
                    'bobot': bobot_val or '3%'
                })

    data['jadwal'] = sorted(data['jadwal'], key=lambda x: int(x['minggu']))

    # 5. Pustaka
    pu_m = re.search(r'(?:####?\s*Pustaka\s+Utama|\*\*Pustaka\s+Utama\*\*)[^\n]*\n*(.+?)(?=(?:####?\s*Pustaka\s+Pendukung|\*\*Pustaka\s+Pendukung\*\*)|\Z)', text, re.DOTALL | re.IGNORECASE)
    if pu_m:
        data['pustaka_utama'] = [l.strip().lstrip('-* ') for l in pu_m.group(1).split('\n') if l.strip() and not l.strip().startswith('#')]

    pp_m = re.search(r'(?:####?\s*Pustaka\s+Pendukung|\*\*Pustaka\s+Pendukung\*\*)[^\n]*\n*(.+?)(?=(?:###?\s*Dosen\s+Pengampu|\*\*Dosen\s+Pengampu\*\*)|\Z)', text, re.DOTALL | re.IGNORECASE)
    if pp_m:
        data['pustaka_pendukung'] = [l.strip().lstrip('-* ') for l in pp_m.group(1).split('\n') if l.strip() and not l.strip().startswith('#')]

    return data

# ---------------------------------------------------------
# Data Extraction: Word (.docx)
# ---------------------------------------------------------

def parse_rps_docx(docx_path):
    """
    Ekstraksi data terstruktur langsung dari dokumen RPS Word (.docx) resmi UNIROW Tuban.
    """
    doc = docx.Document(docx_path)
    if len(doc.tables) < 2:
        raise ValueError(f"Format RPS tidak standar, tabel terdeteksi: {len(doc.tables)} (minimal butuh 2 tabel).")

    data = {
        'nama_mk': '',
        'kode_mk': '',
        'sks': '2',
        'semester': '',
        'rumpun_mk': 'Mata Kuliah Keilmuan dan Keterampilan (MKK PPKn)',
        'dosen_pengampu': 'Tim Dosen Pengampu Program Studi PPKn UNIROW Tuban',
        'koordinator_rmk': 'Dwi Wahyu Kartikasari, M.Pd.',
        'ka_prodi': 'Mario Fahmi Syahrial, M.Pd.',
        'prasyarat': 'Tidak ada',
        'deskripsi': '',
        'cpmk': [],
        'jadwal': [],
        'pustaka_utama': [],
        'pustaka_pendukung': []
    }

    t0 = doc.tables[0]
    # Row 3: Identitas
    if len(t0.rows) > 3:
        r3 = [c.text.strip() for c in t0.rows[3].cells]
        if len(r3) >= 7:
            data['nama_mk'] = r3[0]
            data['kode_mk'] = r3[2]
            data['rumpun_mk'] = r3[3]
            sks_m = re.search(r'(\d+)', r3[5])
            if sks_m: data['sks'] = sks_m.group(1)
            data['semester'] = r3[6]

    # Row 5: Otorisasi
    if len(t0.rows) > 5:
        r5 = [clean_name(c.text) for c in t0.rows[5].cells]
        if len(r5) >= 7:
            if r5[2]: data['dosen_pengampu'] = r5[2]
            if r5[4]: data['koordinator_rmk'] = r5[4]
            if r5[6]: data['ka_prodi'] = r5[6]

    # Row 9: CPMK
    if len(t0.rows) > 9:
        cpmk_cell_text = t0.rows[9].cells[1].text.strip()
        lines = re.split(r'[\r\n]+', cpmk_cell_text)
        for line in lines:
            line = line.strip()
            if not line: continue
            # CPMK format: CPMK1 \t Text
            parts = re.split(r'[:\t]\s*', line, maxsplit=1)
            if len(parts) == 2:
                data['cpmk'].append((parts[0].strip(), parts[1].strip()))
            elif line.startswith('CPMK'):
                m_c = re.match(r'^(CPMK\s*\d+)\s*(.*)$', line)
                if m_c:
                    data['cpmk'].append((m_c.group(1).strip(), m_c.group(2).strip()))
                else:
                    data['cpmk'].append(('CPMK', line))

    # Row 14: Deskripsi
    if len(t0.rows) > 14:
        data['deskripsi'] = t0.rows[14].cells[1].text.strip()

    # Row 16 & 17: Pustaka
    if len(t0.rows) > 16:
        # Pustaka Utama biasanya di cell ke-2
        c_pu = t0.rows[16].cells[2].text.strip() if len(t0.rows[16].cells) > 2 else t0.rows[16].cells[1].text.strip()
        data['pustaka_utama'] = [l.strip().lstrip('-* ') for l in c_pu.split('\n') if l.strip() and not l.strip().lower().startswith('utama')]

    if len(t0.rows) > 17:
        c_pp = t0.rows[17].cells[2].text.strip() if len(t0.rows[17].cells) > 2 else t0.rows[17].cells[1].text.strip()
        data['pustaka_pendukung'] = [l.strip().lstrip('-* ') for l in c_pp.split('\n') if l.strip() and not l.strip().lower().startswith('pendukung')]

    # Row 18: Dosen Pengampu & Row 19: Prasyarat
    if len(t0.rows) > 18:
        dp_cell = t0.rows[18].cells[1].text.strip()
        if dp_cell: data['dosen_pengampu'] = clean_name(dp_cell)

    if len(t0.rows) > 19:
        ms_cell = t0.rows[19].cells[1].text.strip()
        if ms_cell: data['prasyarat'] = ms_cell

    # Table 1: Jadwal Mingguan
    t1 = doc.tables[1]
    for r in range(2, len(t1.rows)):
        c = [cell.text.strip().replace('\n', ' ') for cell in t1.rows[r].cells]
        if not c: continue
        mg_str = c[0].strip()
        if mg_str.startswith('(') or 'MG' in mg_str.upper() or not mg_str:
            continue
        mg_m = re.search(r'^\s*(\d{1,2})\b', mg_str)
        if not mg_m:
            if 'UTS' in mg_str.upper() or 'TENGAH' in mg_str.upper():
                mg_no = 8
            elif 'UAS' in mg_str.upper() or 'AKHIR' in mg_str.upper():
                mg_no = 16
            else:
                continue
        else:
            mg_no = int(mg_m.group(1))

        if mg_no < 1 or mg_no > 16:
            continue

        bobot_val = c[-1].replace('%', '').strip()
        bobot_str = f"{bobot_val}%" if bobot_val else ("25%" if mg_no in (8, 16) else "3%")

        if mg_no == 8:
            data['jadwal'].append({
                'minggu': '8',
                'topik': 'EVALUASI TENGAH SEMESTER (UTS)',
                'bentuk_tugas': 'Ujian Tertulis Teoretis & Analisis Instrumen (Komprehensif Minggu 1 s.d. 7)',
                'bobot': bobot_str
            })
        elif mg_no == 16:
            data['jadwal'].append({
                'minggu': '16',
                'topik': 'EVALUASI AKHIR SEMESTER (UAS)',
                'bentuk_tugas': 'Ujian Kasus Terjadwal Komprehensif & Presentasi Portofolio Proyek',
                'bobot': bobot_str
            })
        else:
            topik = c[1]
            topik_clean = re.sub(r'^Sub-CPMK\s*\d+\s*:\s*', '', topik, flags=re.IGNORECASE)
            tugas = c[4] if len(c) > 4 else (c[2] if len(c) > 2 else "Kuliah Tatap Muka & Diskusi Terstruktur")
            data['jadwal'].append({
                'minggu': str(mg_no),
                'topik': topik_clean,
                'bentuk_tugas': tugas,
                'bobot': bobot_str
            })

    data['jadwal'] = sorted(data['jadwal'], key=lambda x: int(x['minggu']))
    return data

def parse_rps_data_for_kontrak(source_path):
    """
    Fungsi universal untuk mem-parsing data RPS baik dari berkas .docx maupun .md.
    """
    ext = os.path.splitext(source_path)[1].lower()
    if ext == '.docx':
        return parse_rps_docx(source_path)
    elif ext == '.md':
        return parse_rps_md(source_path)
    else:
        # Coba periksa apakah berkas .docx atau .md dengan nama sama ada
        if os.path.exists(source_path + '.docx'):
            return parse_rps_docx(source_path + '.docx')
        elif os.path.exists(source_path + '.md'):
            return parse_rps_md(source_path + '.md')
        else:
            raise ValueError(f"Format berkas tidak dikenali ({source_path}). Gunakan berkas .docx atau .md.")

# ---------------------------------------------------------
# Document Generator: Word (.docx)
# ---------------------------------------------------------

def build_kontrak_docx(data, output_path):
    """
    Membangun file .docx Kontrak Kuliah resmi standar UNIROW Tuban berorientasi A4 Portrait.
    """
    doc = docx.Document()

    # 1. Page Setup: A4 Portrait, Margins 2.5cm (atas, bawah, kanan) & 3.0cm (kiri)
    sec = doc.sections[0]
    sec.page_width = Inches(8.27)
    sec.page_height = Inches(11.69)
    sec.top_margin = Inches(0.98)     # 2.5 cm
    sec.bottom_margin = Inches(0.98)  # 2.5 cm
    sec.left_margin = Inches(1.18)    # 3.0 cm
    sec.right_margin = Inches(0.98)   # 2.5 cm

    # 2. KOP Table
    kop_tbl = doc.add_table(rows=1, cols=2)
    kop_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_col_widths(kop_tbl, [1.2, 4.9])
    set_row_cant_split(kop_tbl.rows[0])

    c0 = kop_tbl.rows[0].cells[0]
    set_cell_valign(c0, "center")
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(0)
    if os.path.exists(LOGO_FILE):
        r_img = p0.add_run()
        r_img.add_picture(LOGO_FILE, width=Inches(1.0))

    c1 = kop_tbl.rows[0].cells[1]
    set_cell_valign(c1, "center")
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.line_spacing = 1.05
    p1.paragraph_format.space_after = Pt(2)
    add_paragraph_run(p1, "UNIVERSITAS PGRI RONGGOLAWE (UNIROW) TUBAN\n", font_size=11.5, bold=True)
    add_paragraph_run(p1, "FAKULTAS KEGURUAN DAN ILMU PENDIDIKAN\n", font_size=10.5, bold=True)
    add_paragraph_run(p1, "PROGRAM STUDI PENDIDIKAN PANCASILA DAN KEWARGANEGARAAN\n", font_size=10.5, bold=True)
    add_paragraph_run(p1, "Jl. Manunggal No. 61 Tuban 62319, Jawa Timur - Telp. (0356) 322233\nWebsite: www.unirow.ac.id | Email: ppkn@unirow.ac.id", font_size=8.5, italic=True)

    for cell in kop_tbl.rows[0].cells:
        set_cell_borders(cell, None, None, None, None)

    # Native Double Line Divider khas Word KOP
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(2)
    p_div.paragraph_format.space_after = Pt(12)
    add_double_bottom_divider(p_div, color='003366', sz='12', space='6')

    # 3. Document Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    add_paragraph_run(p_title, "KONTRAK PERKULIAHAN", font_size=14, bold=True, color_rgb=RGBColor(26, 37, 44))

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(14)
    sub_title_text = f"MATA KULIAH: {data.get('nama_mk', 'NAMA MK').upper()} (KODE: {data.get('kode_mk', '-')}) — BOBOT: {data.get('sks', '2')} SKS\nSEMESTER {data.get('semester', '4')} TAHUN AKADEMIK 2026/2027"
    add_paragraph_run(p_sub, sub_title_text, font_size=10.5, bold=True, color_rgb=RGBColor(80, 90, 95))

    # 4. Bagian A: Identitas MK
    p_secA = doc.add_paragraph()
    p_secA.paragraph_format.space_before = Pt(8)
    p_secA.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_secA, "A. IDENTITAS MATA KULIAH", font_size=11, bold=True)

    identitas_data = [
        ("Nama Mata Kuliah", data.get('nama_mk', '-')),
        ("Kode Mata Kuliah / Bobot", f"{data.get('kode_mk', '-')} / {data.get('sks', '2')} SKS (Teori)"),
        ("Rumpun Mata Kuliah", data.get('rumpun_mk', 'Mata Kuliah Keahlian (MKK PPKn)')),
        ("Semester / Tahun Akademik", f"{data.get('semester', '4')} / Tahun Akademik 2026/2027"),
        ("Dosen Pengampu", data.get('dosen_pengampu', 'Tim Dosen Keahlian PPKn UNIROW Tuban')),
        ("Koordinator RMK", data.get('koordinator_rmk', 'Dwi Wahyu Kartikasari, M.Pd.')),
        ("Ketua Program Studi", data.get('ka_prodi', 'Mario Fahmi Syahrial, M.Pd.')),
        ("Hari / Jam / Ruang", "................................ / ................ WIB / Ruang: ........"),
        ("Mata Kuliah Prasyarat", data.get('prasyarat', 'Tidak ada'))
    ]

    id_tbl = doc.add_table(rows=len(identitas_data), cols=2)
    id_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_col_widths(id_tbl, [2.2, 3.9])

    for idx, (label, val) in enumerate(identitas_data):
        row = id_tbl.rows[idx]
        set_row_cant_split(row)
        c_lbl = row.cells[0]
        c_val = row.cells[1]
        set_cell_padding(c_lbl, 55, 55, 100, 100)
        set_cell_padding(c_val, 55, 55, 100, 100)
        set_cell_borders(c_lbl, 'single', 'single', 'single', 'single')
        set_cell_borders(c_val, 'single', 'single', 'single', 'single')
        set_cell_shading(c_lbl, 'F4F6F7')
        set_cell_valign(c_lbl, 'center')
        set_cell_valign(c_val, 'center')

        p_l = c_lbl.paragraphs[0]
        p_l.paragraph_format.space_before = Pt(0)
        p_l.paragraph_format.line_spacing = 1.05
        p_l.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p_l, label, font_size=9.5, bold=True)

        p_v = c_val.paragraphs[0]
        p_v.paragraph_format.space_before = Pt(0)
        p_v.paragraph_format.line_spacing = 1.05
        p_v.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p_v, val, font_size=9.5)

    # 5. Bagian B: Deskripsi MK
    p_secB = doc.add_paragraph()
    p_secB.paragraph_format.space_before = Pt(12)
    p_secB.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_secB, "B. DESKRIPSI SINGKAT MATA KULIAH", font_size=11, bold=True)

    p_desc = doc.add_paragraph()
    p_desc.paragraph_format.line_spacing = 1.15
    p_desc.paragraph_format.space_after = Pt(8)
    p_desc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_paragraph_run(p_desc, data.get('deskripsi', 'Mata kuliah ini membekali mahasiswa dengan kompetensi keilmuan sesuai capaian pembelajaran kurikulum OBE UNIROW Tuban.'), font_size=9.5)

    # 6. Bagian C: CPMK
    p_secC = doc.add_paragraph()
    p_secC.paragraph_format.space_before = Pt(10)
    p_secC.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_secC, "C. CAPAIAN PEMBELAJARAN MATA KULIAH (CPMK)", font_size=11, bold=True)

    p_cp_intro = doc.add_paragraph()
    p_cp_intro.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_cp_intro, "Setelah menyelesaikan perkuliahan ini, mahasiswa diharapkan mampu:", font_size=9.5, italic=True)

    if data.get('cpmk'):
        for code, desc in data['cpmk']:
            p_cp = doc.add_paragraph()
            p_cp.paragraph_format.left_indent = Inches(0.2)
            p_cp.paragraph_format.line_spacing = 1.15
            p_cp.paragraph_format.space_after = Pt(2)
            add_paragraph_run(p_cp, f"{code}: ", font_size=9.5, bold=True)
            add_paragraph_run(p_cp, desc, font_size=9.5)
    else:
        default_cpmk = [
            ("CPMK 1", "Mampu memahami dan menjelaskan teori dasar dan konsep keilmuan secara komprehensif."),
            ("CPMK 2", "Mampu menganalisis fenomena dan regulasi terkait sesuai konteks keindonesiaan."),
            ("CPMK 3", "Mampu mengevaluasi permasalahan praksis di bidang kajian dengan pendekatan kritis."),
            ("CPMK 4", "Mampu merancang modul ajar dan luaran proyek inovatif berbasis masalah riil.")
        ]
        for code, desc in default_cpmk:
            p_cp = doc.add_paragraph()
            p_cp.paragraph_format.left_indent = Inches(0.2)
            p_cp.paragraph_format.line_spacing = 1.15
            p_cp.paragraph_format.space_after = Pt(2)
            add_paragraph_run(p_cp, f"{code}: ", font_size=9.5, bold=True)
            add_paragraph_run(p_cp, desc, font_size=9.5)

    # 7. Bagian D: Jadwal Perkuliahan 16 Minggu
    p_secD = doc.add_paragraph()
    p_secD.paragraph_format.space_before = Pt(12)
    p_secD.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_secD, "D. JADWAL, MATERI POKOK, DAN PENUGASAN (16 MINGGU)", font_size=11, bold=True)

    jadwal_list = data.get('jadwal', [])
    tbl_jadwal = doc.add_table(rows=len(jadwal_list) + 2, cols=4)
    tbl_jadwal.alignment = WD_TABLE_ALIGNMENT.CENTER
    jadwal_col_widths = [0.65, 2.75, 2.10, 0.60]
    set_table_col_widths(tbl_jadwal, jadwal_col_widths)

    # Header Row (mengulang otomatis di tiap halaman / tblHeader & cantSplit)
    hdr_row = tbl_jadwal.rows[0]
    set_repeat_header(hdr_row)
    set_row_cant_split(hdr_row)

    headers = ["Minggu", "Materi Pokok / Bahasan", "Bentuk Pembelajaran & Penugasan Terstruktur", "Bobot"]
    for i, h_text in enumerate(headers):
        cell = hdr_row.cells[i]
        set_cell_padding(cell, 60, 60, 80, 80)
        set_cell_borders(cell, 'single', 'single', 'single', 'single')
        set_cell_shading(cell, 'EAECEE')
        set_cell_valign(cell, 'center')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p, h_text, font_size=9, bold=True)

    # Rows 1 to 16
    for idx, jw in enumerate(jadwal_list):
        row = tbl_jadwal.rows[idx + 1]
        set_row_cant_split(row)
        mg_no = jw.get('minggu', str(idx + 1))
        is_exam = mg_no in ('8', '16') or 'UTS' in jw.get('topik', '') or 'UAS' in jw.get('topik', '')

        # Col 0: Minggu
        c0 = row.cells[0]
        set_cell_padding(c0, 50, 50, 60, 60)
        set_cell_borders(c0, 'single', 'single', 'single', 'single')
        if is_exam: set_cell_shading(c0, 'F4F6F7')
        set_cell_valign(c0, 'center')
        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p0, mg_no, font_size=9, bold=is_exam)

        # Col 1: Materi Pokok
        c1 = row.cells[1]
        set_cell_padding(c1, 50, 50, 80, 80)
        set_cell_borders(c1, 'single', 'single', 'single', 'single')
        if is_exam: set_cell_shading(c1, 'F4F6F7')
        set_cell_valign(c1, 'top')
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.line_spacing = 1.05
        p1.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p1, jw.get('topik', ''), font_size=9, bold=is_exam)

        # Col 2: Bentuk Pembelajaran & Penugasan
        c2 = row.cells[2]
        set_cell_padding(c2, 50, 50, 80, 80)
        set_cell_borders(c2, 'single', 'single', 'single', 'single')
        if is_exam: set_cell_shading(c2, 'F4F6F7')
        set_cell_valign(c2, 'top')
        p2 = c2.paragraphs[0]
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.line_spacing = 1.05
        p2.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p2, jw.get('bentuk_tugas', ''), font_size=8.5, bold=is_exam)

        # Col 3: Bobot
        c3 = row.cells[3]
        set_cell_padding(c3, 50, 50, 60, 60)
        set_cell_borders(c3, 'single', 'single', 'single', 'single')
        if is_exam: set_cell_shading(c3, 'F4F6F7')
        set_cell_valign(c3, 'center')
        p3 = c3.paragraphs[0]
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p3, jw.get('bobot', '3%'), font_size=9, bold=is_exam)

    # Footer Row: Total Bobot 100%
    tot_row = tbl_jadwal.rows[-1]
    set_row_cant_split(tot_row)
    c_tot_lbl = tot_row.cells[0]
    set_cell_padding(c_tot_lbl, 55, 55, 80, 80)
    set_cell_borders(c_tot_lbl, 'single', 'single', 'single', 'single')
    set_cell_shading(c_tot_lbl, 'EAECEE')
    set_cell_valign(c_tot_lbl, 'center')
    p_tl = c_tot_lbl.paragraphs[0]
    p_tl.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tl.paragraph_format.space_before = Pt(0)
    p_tl.paragraph_format.space_after = Pt(0)
    add_paragraph_run(p_tl, "TOTAL BOBOT KUMULATIF PERKULIAHAN", font_size=9, bold=True)

    # Merge cols 0, 1, 2
    c_tot_lbl.merge(tot_row.cells[1])
    c_tot_lbl.merge(tot_row.cells[2])

    c_tot_val = tot_row.cells[3]
    set_cell_padding(c_tot_val, 55, 55, 60, 60)
    set_cell_borders(c_tot_val, 'single', 'single', 'single', 'single')
    set_cell_shading(c_tot_val, 'EAECEE')
    set_cell_valign(c_tot_val, 'center')
    p_tv = c_tot_val.paragraphs[0]
    p_tv.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_tv.paragraph_format.space_before = Pt(0)
    p_tv.paragraph_format.space_after = Pt(0)
    add_paragraph_run(p_tv, "100%", font_size=9, bold=True)

    # 8. Bagian E: Sistem Evaluasi, Bobot, dan Konversi Nilai Akhir
    p_secE = doc.add_paragraph()
    p_secE.paragraph_format.space_before = Pt(12)
    p_secE.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_secE, "E. SISTEM EVALUASI, BOBOT, DAN KONVERSI NILAI AKHIR", font_size=11, bold=True)

    # 1. Nilai Akhir
    p_na_title = doc.add_paragraph()
    p_na_title.paragraph_format.space_after = Pt(2)
    add_paragraph_run(p_na_title, "1. Nilai Akhir (NA)", font_size=9.5, bold=True)

    p_na_desc = doc.add_paragraph()
    p_na_desc.paragraph_format.line_spacing = 1.15
    p_na_desc.paragraph_format.space_after = Pt(4)
    p_na_desc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_paragraph_run(p_na_desc,
        "Nilai Akhir (NA) diperoleh dari Penilaian Presensi (P), Tugas/praktikum (TGS), Ujian Tengah Semester (UTS), dan Ujian Akhir Semester (UAS). "
        "Ketentuan perhitungan nilai akhir adalah sebagai berikut:\n"
        "NA = (P + 2·TGS + 3·UTS + 4·UAS) / 10",
        font_size=9.5
    )

    komponen_nilai = [
        ("1", "Penilaian Presensi (P)", "10%"),
        ("2", "Tugas / Praktikum (TGS)", "20%"),
        ("3", "Ujian Tengah Semester (UTS)", "30%"),
        ("4", "Ujian Akhir Semester (UAS)", "40%"),
        ("", "TOTAL KOMULATIF BOBOT EVALUASI", "100%")
    ]
    tbl_eval = doc.add_table(rows=len(komponen_nilai) + 1, cols=3)
    tbl_eval.alignment = WD_TABLE_ALIGNMENT.CENTER
    eval_col_widths = [0.55, 4.55, 1.00]
    set_table_col_widths(tbl_eval, eval_col_widths)

    set_repeat_header(tbl_eval.rows[0])
    set_row_cant_split(tbl_eval.rows[0])

    h_eval = ["No", "Komponen Evaluasi", "Bobot"]
    for i, h in enumerate(h_eval):
        cell = tbl_eval.rows[0].cells[i]
        set_cell_padding(cell, 55, 55, 70, 70)
        set_cell_borders(cell, 'single', 'single', 'single', 'single')
        set_cell_shading(cell, 'EAECEE')
        set_cell_valign(cell, 'center')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p, h, font_size=9, bold=True)

    for idx, (no, komp, bbt) in enumerate(komponen_nilai):
        row = tbl_eval.rows[idx + 1]
        set_row_cant_split(row)
        is_tot = (no == "")
        c0 = row.cells[0]
        c1 = row.cells[1]
        c2 = row.cells[2]
        for c in (c0, c1, c2):
            set_cell_padding(c, 45, 45, 70, 70)
            set_cell_borders(c, 'single', 'single', 'single', 'single')
            set_cell_valign(c, 'center')
            if is_tot: set_cell_shading(c, 'EAECEE')

        p0 = c0.paragraphs[0]
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p0, no, font_size=9, bold=is_tot)

        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p1, komp, font_size=9, bold=is_tot)

        p2 = c2.paragraphs[0]
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p2, bbt, font_size=9, bold=is_tot)

    # Distribusi Nilai Huruf & Nilai Mutu (Pedoman Akademik UNIROW)
    p_f2 = doc.add_paragraph()
    p_f2.paragraph_format.space_before = Pt(8)
    p_f2.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_f2, "Adapun nilai hasil belajar mahasiswa dinyatakan dengan Nilai Huruf yang didistribusikan sebagai berikut:", font_size=9.5)

    skala_data = [
        ("85 < NA ≤ 100", "A", "4"),
        ("77,5 < NA ≤ 85", "AB", "3,5"),
        ("70 < NA ≤ 77,5", "B", "3"),
        ("65 < NA ≤ 70", "BC", "2,5"),
        ("55 < NA ≤ 65", "C", "2"),
        ("45 < NA ≤ 55", "D", "1"),
        ("0 < NA ≤ 45", "E", "0")
    ]

    tbl_skala = doc.add_table(rows=len(skala_data) + 1, cols=3)
    tbl_skala.alignment = WD_TABLE_ALIGNMENT.CENTER
    skala_col_widths = [2.30, 1.90, 1.90]
    set_table_col_widths(tbl_skala, skala_col_widths)

    set_repeat_header(tbl_skala.rows[0])
    set_row_cant_split(tbl_skala.rows[0])

    h_skala = ["Interval Nilai", "Nilai Huruf", "Nilai Mutu"]
    for i, h in enumerate(h_skala):
        cell = tbl_skala.rows[0].cells[i]
        set_cell_padding(cell, 55, 55, 70, 70)
        set_cell_borders(cell, 'single', 'single', 'single', 'single')
        set_cell_shading(cell, 'EAECEE')
        set_cell_valign(cell, 'center')
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        add_paragraph_run(p, h, font_size=9, bold=True)

    for idx, (skor, huruf, mutu) in enumerate(skala_data):
        row = tbl_skala.rows[idx + 1]
        set_row_cant_split(row)
        for i_col, val in enumerate((skor, huruf, mutu)):
            cell = row.cells[i_col]
            set_cell_padding(cell, 40, 40, 60, 60)
            set_cell_borders(cell, 'single', 'single', 'single', 'single')
            set_cell_valign(cell, 'center')
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            is_c = (huruf == "C")
            if is_c: set_cell_shading(cell, 'F4F6F7')
            add_paragraph_run(p, val, font_size=9, bold=(i_col == 1))

    # 2. Kelulusan Matakuliah
    p_lulus_title = doc.add_paragraph()
    p_lulus_title.paragraph_format.space_before = Pt(8)
    p_lulus_title.paragraph_format.space_after = Pt(3)
    add_paragraph_run(p_lulus_title, "2. Kelulusan matakuliah", font_size=9.5, bold=True)

    kelulusan_poin = [
        "a. Mahasiswa dinyatakan lulus matakuliah jika mendapat nilai A, AB, B, BC, C.",
        "b. Nilai D dinyatakan tidak lulus. Mahasiswa dengan nilai D dapat mengulang perkuliahan dengan kehadiran minimal 50% dan mengikuti ujian pada Ujian Akhir Semester (UAS) sesuai dengan ketentuan.",
        "c. Nilai E dinyatakan tidak lulus, dan mahasiswa wajib mengikuti perkuliahan pada semester berikutnya sesuai ketentuan."
    ]
    for kp in kelulusan_poin:
        p_kp = doc.add_paragraph()
        p_kp.paragraph_format.left_indent = Inches(0.2)
        p_kp.paragraph_format.line_spacing = 1.15
        p_kp.paragraph_format.space_after = Pt(2)
        add_paragraph_run(p_kp, kp, font_size=9.5)

    # 9. Bagian F: Tata Tertib Perkuliahan
    p_secF = doc.add_paragraph()
    p_secF.paragraph_format.space_before = Pt(12)
    p_secF.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_secF, "F. TATA TERTIB DAN KESEPAKATAN PERKULIAHAN", font_size=11, bold=True)

    aturan_list = [
        "1. Mahasiswa wajib hadir tepat waktu. Toleransi keterlambatan maksimal adalah 15 menit setelah perkuliahan dimulai.",
        "2. Kehadiran tatap muka minimal adalah 75% dari total 16 pertemuan sebagai prasyarat mutlak mengikuti Evaluasi Akhir Semester (UAS).",
        "3. Berpakaian sopan, rapi, berkerah, dan memakai sepatu formal sesuai Kode Etik Sivitas Akademika Universitas PGRI Ronggolawe Tuban.",
        "4. Menjaga ketertiban kelas, etika berkomunikasi, kesantunan akademik, dan saling menghormati perbedaan pandangan ilmiah.",
        "5. Menggunakan perangkat gawai (smartphone/laptop) secara bijak hanya untuk menunjang aktivitas perkuliahan dan riset materi.",
        "6. Tugas terstruktur wajib dikumpulkan tepat waktu sesuai instruksi. Keterlambatan tanpa alasan sah dikenakan pemotongan nilai 10% per hari.",
        "7. Menjunjung tinggi kebebasan mimbar akademik yang inklusif serta menolak segala bentuk perundungan (anti-bullying) dan kekerasan seksual.",
        "8. Seluruh penugasan wajib menjunjung tinggi integritas akademik. Tindakan plagiarisme dan kecurangan saat ujian dikenakan sanksi nilai 0 (E).",
        "9. Ujian susulan (UTS/UAS) hanya diberikan bagi mahasiswa yang memiliki alasan sah (sakit rawat inap atau tugas dinas kampus) dengan bukti tertulis lengkap."
    ]
    for atr in aturan_list:
        p_a = doc.add_paragraph()
        p_a.paragraph_format.left_indent = Inches(0.2)
        p_a.paragraph_format.line_spacing = 1.15
        p_a.paragraph_format.space_after = Pt(3)
        add_paragraph_run(p_a, atr, font_size=9.5)

    # 10. Bagian G: Pustaka
    p_secG = doc.add_paragraph()
    p_secG.paragraph_format.space_before = Pt(12)
    p_secG.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_secG, "G. PUSTAKA RUJUKAN", font_size=11, bold=True)

    p_pu_title = doc.add_paragraph()
    p_pu_title.paragraph_format.space_before = Pt(2)
    p_pu_title.paragraph_format.space_after = Pt(2)
    add_paragraph_run(p_pu_title, "Pustaka Utama:", font_size=9.5, bold=True)
    pu_clean = [re.sub(r'^\d+[\.\)]\s*', '', p) for p in data.get('pustaka_utama', []) if p.strip()]
    if not pu_clean:
        pu_clean = ["Buku Referensi Utama Program Studi PPKn UNIROW Tuban."]
    for i, p in enumerate(pu_clean):
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Inches(0.2)
        p_item.paragraph_format.space_after = Pt(2)
        p_item.paragraph_format.line_spacing = 1.15
        add_paragraph_run(p_item, f"{i+1}. {p}", font_size=9)

    p_pp_title = doc.add_paragraph()
    p_pp_title.paragraph_format.space_before = Pt(6)
    p_pp_title.paragraph_format.space_after = Pt(2)
    add_paragraph_run(p_pp_title, "Pustaka Pendukung:", font_size=9.5, bold=True)
    pp_clean = [re.sub(r'^\d+[\.\)]\s*', '', p) for p in data.get('pustaka_pendukung', []) if p.strip()]
    if not pp_clean:
        pp_clean = ["Peraturan Perundang-undangan dan Jurnal Ilmiah Terkait."]
    for i, p in enumerate(pp_clean):
        p_item = doc.add_paragraph()
        p_item.paragraph_format.left_indent = Inches(0.2)
        p_item.paragraph_format.space_after = Pt(2)
        p_item.paragraph_format.line_spacing = 1.15
        add_paragraph_run(p_item, f"{i+1}. {p}", font_size=9)

    # 11. Bagian H: Pengesahan & Tanda Tangan
    p_secH = doc.add_paragraph()
    p_secH.paragraph_format.space_before = Pt(14)
    p_secH.paragraph_format.space_after = Pt(4)
    add_paragraph_run(p_secH, "H. PERNYATAAN KESEPAKATAN DAN PENGESAHAN", font_size=11, bold=True)

    p_setuju = doc.add_paragraph()
    p_setuju.paragraph_format.line_spacing = 1.15
    p_setuju.paragraph_format.space_after = Pt(12)
    p_setuju.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_paragraph_run(p_setuju,
        f"Kontrak Perkuliahan ini disepakati secara sadar dan sukarela oleh Dosen Pengampu Mata Kuliah dan seluruh Mahasiswa "
        f"peserta mata kuliah {data.get('nama_mk', '')} ({data.get('kode_mk', '')}) Program Studi PPKn FKIP Universitas PGRI Ronggolawe Tuban untuk dijadikan "
        f"pedoman bersama selama perkuliahan Semester {data.get('semester', '4')} Tahun Akademik 2026/2027.",
        font_size=9.5
    )

    p_tgl = doc.add_paragraph()
    p_tgl.paragraph_format.space_after = Pt(8)
    p_tgl.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_paragraph_run(p_tgl, "Tuban, ........................................ 2026", font_size=9.5, italic=True)

    ttd_tbl = doc.add_table(rows=1, cols=2)
    ttd_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    ttd_col_widths = [3.05, 3.05]
    set_table_col_widths(ttd_tbl, ttd_col_widths)
    set_row_cant_split(ttd_tbl.rows[0])

    # Kolom Mahasiswa
    c_mhs = ttd_tbl.rows[0].cells[0]
    set_cell_borders(c_mhs, None, None, None, None)
    set_cell_padding(c_mhs, 60, 60, 80, 80)
    set_cell_valign(c_mhs, 'top')
    p_m = c_mhs.paragraphs[0]
    p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_m.paragraph_format.space_before = Pt(0)
    p_m.paragraph_format.space_after = Pt(0)
    add_paragraph_run(p_m, "Perwakilan Mahasiswa / Komti,\n\n\n\n\n", font_size=9.5)
    add_paragraph_run(p_m, "( .................................................... )\n", font_size=9.5, bold=True)
    add_paragraph_run(p_m, "NIM. .................................................", font_size=9)

    # Kolom Dosen
    c_dos = ttd_tbl.rows[0].cells[1]
    set_cell_borders(c_dos, None, None, None, None)
    set_cell_padding(c_dos, 60, 60, 80, 80)
    set_cell_valign(c_dos, 'top')
    p_d = c_dos.paragraphs[0]
    p_d.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_d.paragraph_format.space_before = Pt(0)
    p_d.paragraph_format.space_after = Pt(0)
    add_paragraph_run(p_d, "Dosen Pengampu Mata Kuliah,\n\n\n\n\n", font_size=9.5)
    add_paragraph_run(p_d, f"{data.get('dosen_pengampu', 'Tim Dosen Pengampu')}\n", font_size=9.5, bold=True)
    add_paragraph_run(p_d, "NIDN. .................................................", font_size=9)

    # Mengetahui Kaprodi
    p_kpd = doc.add_paragraph()
    p_kpd.paragraph_format.space_before = Pt(14)
    p_kpd.paragraph_format.space_after = Pt(2)
    p_kpd.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_paragraph_run(p_kpd, "Mengetahui,\nKetua Program Studi Pendidikan Pancasila dan Kewarganegaraan\nFKIP UNIROW Tuban\n\n\n\n\n", font_size=9.5)
    add_paragraph_run(p_kpd, f"{data.get('ka_prodi', 'Mario Fahmi Syahrial, M.Pd.')}\n", font_size=10, bold=True)
    add_paragraph_run(p_kpd, "NIDN. .................................................", font_size=9)

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc.save(output_path)
    print(f"[OK] Kontrak Perkuliahan Word (.docx) berhasil dibuat: {output_path}")

# ---------------------------------------------------------
# Document Generator: Markdown (.md)
# ---------------------------------------------------------

def build_kontrak_md(data, output_md_path):
    """
    Membangun file .md Kontrak Kuliah resmi berformat GitHub-Flavored Markdown.
    """
    sks_val = int(data.get('sks', '2')) if str(data.get('sks', '2')).isdigit() else 2

    lines = []
    lines.append("# KONTRAK PERKULIAHAN\n")
    lines.append(f"**Mata Kuliah : {data.get('nama_mk', '')}**  ")
    lines.append(f"**Kode Mata Kuliah : {data.get('kode_mk', '')}**  ")
    lines.append("Program Studi : Pendidikan Pancasila dan Kewarganegaraan (PPKn)  ")
    lines.append("Fakultas : Keguruan dan Ilmu Pendidikan (FKIP)  ")
    lines.append("Perguruan Tinggi : Universitas PGRI Ronggolawe (UNIROW) Tuban  ")
    lines.append(f"Semester / Bobot : {data.get('semester', '4')} / {sks_val} SKS (Teori)  ")
    lines.append("Tahun Akademik : 2026/2027  \n")
    lines.append("---\n")

    # A. Identitas
    lines.append("## A. Identitas Mata Kuliah\n")
    lines.append("| Komponen Identitas | Keterangan |")
    lines.append("|:---|:---|")
    lines.append(f"| **Nama Mata Kuliah** | {data.get('nama_mk', '')} |")
    lines.append(f"| **Kode / Bobot SKS** | {data.get('kode_mk', '')} / {sks_val} SKS (Teori) |")
    lines.append(f"| **Rumpun Mata Kuliah** | {data.get('rumpun_mk', '')} |")
    lines.append(f"| **Semester / Tahun Akademik** | {data.get('semester', '')} / Tahun Akademik 2026/2027 |")
    lines.append(f"| **Dosen Pengampu** | {data.get('dosen_pengampu', '')} |")
    lines.append(f"| **Koordinator RMK** | {data.get('koordinator_rmk', '')} |")
    lines.append(f"| **Ketua Program Studi** | {data.get('ka_prodi', 'Mario Fahmi Syahrial, M.Pd.')} |")
    lines.append("| **Hari / Jam / Ruang** | ........................................ / ................ WIB / Ruang: ............ |")
    lines.append(f"| **Mata Kuliah Prasyarat** | {data.get('prasyarat', 'Tidak ada')} |\n")
    lines.append("---\n")

    # B. Deskripsi
    lines.append("## B. Deskripsi Singkat Mata Kuliah\n")
    lines.append(f"{data.get('deskripsi', '')}\n")
    lines.append("---\n")

    # C. CPMK
    lines.append("## C. Capaian Pembelajaran Mata Kuliah (CPMK)\n")
    lines.append("Setelah menyelesaikan perkuliahan ini, mahasiswa diharapkan mampu:")
    if data.get('cpmk'):
        for i, (code, desc) in enumerate(data['cpmk']):
            lines.append(f"{i+1}. **{code}**: {desc}")
    else:
        lines.append("1. **CPMK 1**: Menguasai konsep dasar teoretis dan metodologis kajian secara komprehensif.")
        lines.append("2. **CPMK 2**: Menganalisis fenomena dan regulasi terkait sesuai konteks keindonesiaan.")
        lines.append("3. **CPMK 3**: Mengevaluasi permasalahan praksis di bidang kajian dengan pendekatan kritis.")
        lines.append("4. **CPMK 4**: Merancang modul ajar dan luaran proyek inovatif berbasis masalah riil.")
    lines.append("\n---\n")

    # D. Jadwal Mingguan
    lines.append("## D. Jadwal, Materi Pokok, dan Penugasan (16 Minggu)\n")
    lines.append("| Mg | Topik / Pokok Bahasan | Bentuk Pembelajaran & Penugasan Terstruktur | Bobot |")
    lines.append("|:---:|:---|:---|:---:|")
    for jw in data.get('jadwal', []):
        mg = jw.get('minggu', '')
        topik = jw.get('topik', '')
        tugas = jw.get('bentuk_tugas', '')
        bobot = jw.get('bobot', '')
        if mg in ('8', '16'):
            lines.append(f"| **{mg}** | **{topik}** | **{tugas}** | **{bobot}** |")
        else:
            lines.append(f"| **{mg}** | {topik} | {tugas} | {bobot} |")
    lines.append("| | **TOTAL BOBOT KUMULATIF PERKULIAHAN** | | **100%** |\n")
    lines.append("---\n")

    # E. Evaluasi & Konversi Nilai Akhir
    lines.append("## E. Sistem Evaluasi, Bobot, dan Konversi Nilai Akhir\n")
    lines.append("### 1. Nilai Akhir (NA)")
    lines.append("Nilai Akhir (NA) diperoleh dari Penilaian Presensi (P), Tugas/praktikum (TGS), Ujian Tengah Semester (UTS), dan Ujian Akhir Semester (UAS). Ketentuan perhitungan nilai akhir adalah sebagai berikut:\n")
    lines.append("$$\\text{NA} = \\frac{\\text{P} + 2(\\text{TGS}) + 3(\\text{UTS}) + 4(\\text{UAS})}{10}$$\n")
    lines.append("| No | Komponen Evaluasi | Simbol | Bobot (%) |")
    lines.append("|:---:|:---|:---:|:---:|")
    lines.append("| 1 | **Penilaian Presensi** | **P** | **10%** |")
    lines.append("| 2 | **Tugas / Praktikum** | **TGS** | **20%** |")
    lines.append("| 3 | **Ujian Tengah Semester (UTS)** | **UTS** | **30%** |")
    lines.append("| 4 | **Ujian Akhir Semester (UAS)** | **UAS** | **40%** |")
    lines.append("| | **TOTAL** | | **100%** |\n")

    lines.append("Adapun nilai hasil belajar mahasiswa dinyatakan dengan Nilai Huruf yang didistribusikan sebagai berikut:\n")
    lines.append("| Interval Nilai | Nilai Huruf | Nilai Mutu |")
    lines.append("|:---:|:---:|:---:|")
    lines.append("| 85 < NA ≤ 100 | **A** | 4 |")
    lines.append("| 77,5 < NA ≤ 85 | **AB** | 3,5 |")
    lines.append("| 70 < NA ≤ 77,5 | **B** | 3 |")
    lines.append("| 65 < NA ≤ 70 | **BC** | 2,5 |")
    lines.append("| 55 < NA ≤ 65 | **C** | 2 |")
    lines.append("| 45 < NA ≤ 55 | **D** | 1 |")
    lines.append("| 0 < NA ≤ 45 | **E** | 0 |\n")

    lines.append("### 2. Kelulusan matakuliah")
    lines.append("a. Mahasiswa dinyatakan lulus matakuliah jika mendapat nilai **A, AB, B, BC, C**.")
    lines.append("b. Nilai **D** dinyatakan **tidak lulus**. Mahasiswa dengan nilai D dapat mengulang perkuliahan dengan kehadiran minimal 50% dan mengikuti ujian pada Ujian Akhir Semester (UAS) sesuai dengan ketentuan.")
    lines.append("c. Nilai **E** dinyatakan **tidak lulus**, dan mahasiswa wajib mengikuti perkuliahan pada semester berikutnya sesuai ketentuan.\n")
    lines.append("---\n")

    # F. Tata Tertib
    lines.append("## F. Tata Tertib dan Kesepakatan Perkuliahan\n")
    lines.append("1. **Kehadiran**: Minimal 75% dari total 16 pertemuan tatap muka sebagai syarat mutlak mengikuti Evaluasi Akhir Semester (UAS).")
    lines.append("2. **Keterlambatan**: Toleransi maksimal 15 menit. Keterlambatan lebih dari 15 menit dicatat alpa kecuali memiliki izin sah tertulis.")
    lines.append("3. **Izin/Sakit**: Surat keterangan dokter atau tugas dinas kampus wajib diserahkan maksimal 3 hari kerja setelah pertemuan.")
    lines.append("4. **Etika Busana**: Berpakaian sopan, rapi, berkerah, dan bersepatu formal sesuai kode etik kampus UNIROW Tuban.")
    lines.append("5. **Gawai Elektronik**: Penggunaan ponsel dan laptop di ruang kelas hanya untuk kepentingan aktivitas akademik.")
    lines.append("6. **Integritas Akademik**: Menjunjung tinggi kebebasan mimbar akademik, bersikap inklusif, bebas perundungan (*anti-bullying*), serta sanksi nilai 0 (E) untuk plagiarisme atau kecurangan ujian.")
    lines.append("7. **Pengumpulan Tugas**: Tepat waktu. Keterlambatan tanpa konfirmasi sah dikenakan penalti pemotongan nilai 10% per hari keterlambatan.")
    lines.append("8. **Ujian Susulan**: Hanya diberikan bagi mahasiswa dengan alasan darurat (rawat inap/dinas) dengan bukti sah tertulis.\n")
    lines.append("---\n")

    # G. Pustaka
    lines.append("## G. Pustaka Rujukan\n")
    lines.append("### Pustaka Utama:")
    if data.get('pustaka_utama'):
        for i, p in enumerate(data['pustaka_utama']):
            p_cl = re.sub(r'^\d+[\.\)]\s*', '', p)
            lines.append(f"{i+1}. {p_cl}")
    else:
        lines.append("1. Buku Referensi Utama Program Studi PPKn UNIROW Tuban.")
    lines.append("\n### Pustaka Pendukung:")
    if data.get('pustaka_pendukung'):
        for i, p in enumerate(data['pustaka_pendukung']):
            p_cl = re.sub(r'^\d+[\.\)]\s*', '', p)
            lines.append(f"{i+1}. {p_cl}")
    else:
        lines.append("1. Peraturan Perundang-undangan dan Jurnal Ilmiah Terkait.")
    lines.append("\n---\n")

    # H. Pengesahan
    lines.append("## H. Pernyataan Kesepakatan dan Pengesahan\n")
    lines.append(f"Kontrak Perkuliahan ini disepakati secara sadar dan sukarela oleh Dosen Pengampu Mata Kuliah dan seluruh Mahasiswa peserta mata kuliah {data.get('nama_mk', '')} ({data.get('kode_mk', '')}) Program Studi PPKn FKIP Universitas PGRI Ronggolawe Tuban untuk ditaati bersama selama perkuliahan Semester {data.get('semester', '4')} Tahun Akademik 2026/2027.\n")
    lines.append("Dibuat dan disahkan di: **Tuban**  ")
    lines.append("Pada tanggal: ........................................ 2026  \n")
    lines.append("| Perwakilan Mahasiswa (Ketua Tingkat), | Dosen Pengampu Mata Kuliah, |")
    lines.append("|:---:|:---:|")
    lines.append(f"| <br><br><br><br>**( .................................................... )**<br>NIM. ................................................. | <br><br><br><br>**{data.get('dosen_pengampu', 'Tim Dosen Pengampu')}**<br>NIDN. ................................................. |\n")
    lines.append("Mengetahui,<br>")
    lines.append("**Ketua Program Studi PPKn FKIP UNIROW Tuban**\n")
    lines.append("<br><br><br>")
    lines.append(f"**{data.get('ka_prodi', 'Mario Fahmi Syahrial, M.Pd.')}**  ")
    lines.append("NIDN. .................................................\n")

    os.makedirs(os.path.dirname(os.path.abspath(output_md_path)), exist_ok=True)
    with open(output_md_path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(lines))
    print(f"[OK] Kontrak Perkuliahan Markdown (.md) berhasil dibuat: {output_md_path}")

# ---------------------------------------------------------
# Main CLI Interface
# ---------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generator Kontrak Perkuliahan Standar Resmi UNIROW Tuban Berbasis Dokumen RPS (OBE 2026)"
    )
    parser.add_argument('rps_file', nargs='?', default=None, help="Path ke berkas RPS (format .docx atau .md)")
    parser.add_argument('--rps', dest='rps_opt', default=None, help="Path ke berkas RPS (.docx / .md)")
    parser.add_argument('--md', dest='md_opt', default=None, help="Alias untuk berkas RPS .md")
    parser.add_argument('--docx', dest='docx_opt', default=None, help="Alias untuk berkas RPS .docx")
    parser.add_argument('--output', '-o', default=None, help="Path output berkas .docx")
    parser.add_argument('--output-md', default=None, help="Path output berkas .md")
    parser.add_argument('--both', action='store_true', help="Hasilkan berkas .docx sekaligus .md")

    args = parser.parse_args()
    input_path = args.rps_file or args.rps_opt or args.md_opt or args.docx_opt

    if not input_path:
        parser.print_help()
        sys.exit(1)

    if not os.path.exists(input_path):
        print(f"[ERROR] Berkas RPS tidak ditemukan: {input_path}")
        sys.exit(1)

    print(f"[*] Membaca data RPS dari: {input_path}")
    data = parse_rps_data_for_kontrak(input_path)
    print(f"[*] Ekstraksi berhasil: {data.get('nama_mk')} ({data.get('kode_mk')}) - {len(data.get('jadwal', []))} minggu silabus.")

    # Tentukan nama dan path output default
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    m_p = re.match(r'^(\d+_\d+_)', base_name)
    prefix = m_p.group(1) if m_p else ''
    tail = base_name[len(prefix):]
    tail = re.sub(r'^RPS_', '', tail)
    tail = re.sub(r'_(?:FINAL|OBE)$', '', tail)
    kontrak_name = f"{prefix}Kontrak_Kuliah_{tail}_OBE"

    default_dir = os.path.dirname(os.path.abspath(input_path))
    if 'RPS_OBE' in default_dir:
        # Tempatkan di folder KONTRAK_KULIAH jika ada
        p_up = default_dir
        while p_up and os.path.basename(p_up) not in ('RPS OBE 2026', ''):
            if os.path.exists(os.path.join(p_up, 'KONTRAK_KULIAH')):
                default_dir = os.path.join(p_up, 'KONTRAK_KULIAH')
                break
            p_up = os.path.dirname(p_up)
        else:
            if 'RPS OBE 2026' in default_dir:
                root_part = default_dir.split('RPS OBE 2026')[0] + 'RPS OBE 2026'
                default_dir = os.path.join(root_part, 'KONTRAK_KULIAH')

    docx_out = args.output
    if not docx_out:
        docx_out = os.path.join(default_dir, f"{kontrak_name}.docx")

    md_out = args.output_md
    if not md_out and args.both:
        md_out = os.path.splitext(docx_out)[0] + '.md'

    # Generate Word (.docx)
    build_kontrak_docx(data, docx_out)

    # Generate Markdown (.md) jika diminta atau --both
    if md_out:
        build_kontrak_md(data, md_out)

if __name__ == '__main__':
    main()
