#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
batch_generate_all.py - Batch Generator Kontrak Perkuliahan Seluruh RPS OBE UNIROW Tuban
Menemukan seluruh berkas RPS resmi (Semester 1 s.d. 7) dan menghasilkan dokumen Kontrak Perkuliahan
lengkap berformat Word (.docx) ber-KOP resmi dan Markdown (.md) sesuai standar baku mutu kampus.
"""

import os
import sys
import re
import shutil
import time

# Pastikan import build_kontrak dapat diakses
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
if SCRIPT_DIR not in sys.path:
    sys.path.append(SCRIPT_DIR)

from build_kontrak import parse_rps_data_for_kontrak, build_kontrak_docx, build_kontrak_md

WORKSPACE_ROOT = os.path.abspath(os.path.join(SCRIPT_DIR, '..', '..', '..', '..'))
BASE_RPS_DIR = os.path.join(WORKSPACE_ROOT, '01_FIK', 'RPS OBE 2026', 'RPS_OBE_FINAL')
TARGET_DIR_FINAL = os.path.join(BASE_RPS_DIR, 'KONTRAK_KULIAH')
TARGET_DIR_OBE = os.path.join(WORKSPACE_ROOT, '01_FIK', 'RPS OBE 2026', 'KONTRAK_KULIAH')

os.makedirs(TARGET_DIR_FINAL, exist_ok=True)
os.makedirs(TARGET_DIR_OBE, exist_ok=True)

def find_all_rps_courses():
    courses = {}
    for sem in range(1, 8):
        rev_dir = os.path.join(BASE_RPS_DIR, f"{sem:02d}_RPS_OBE_HASIL_REVISI")
        orig_dir = os.path.join(BASE_RPS_DIR, f"{sem:02d}_RPS_OBE")
        
        # Cari file di rev_dir terlebih dahulu (prioritas tertinggi)
        if os.path.exists(rev_dir):
            for f in sorted(os.listdir(rev_dir)):
                if f.endswith('.md') and '_RPS_' in f and '_FINAL' in f:
                    m = re.match(r'^(\d{2}_\d+)_RPS_(.+?)_FINAL\.md$', f, re.IGNORECASE)
                    if m:
                        c_id = m.group(1)
                        c_name = m.group(2)
                        courses[c_id] = {
                            'sem': sem,
                            'id': c_id,
                            'name': c_name,
                            'source': os.path.join(rev_dir, f),
                            'type': 'md'
                        }
                elif f.endswith('.docx') and '_RPS_' in f:
                    m = re.match(r'^(\d{2}_\d+)_RPS_(.+?)(?:_OBE|_FINAL)?\.docx$', f, re.IGNORECASE)
                    if m:
                        c_id = m.group(1)
                        if c_id not in courses:
                            courses[c_id] = {
                                'sem': sem,
                                'id': c_id,
                                'name': m.group(2),
                                'source': os.path.join(rev_dir, f),
                                'type': 'docx'
                            }

        # Fallback ke orig_dir jika belum ada
        if os.path.exists(orig_dir):
            for f in sorted(os.listdir(orig_dir)):
                if f.endswith('.docx') and '_RPS_' in f:
                    m = re.match(r'^(\d{2}_\d+)_RPS_(.+?)(?:_OBE|_COMPLETE)?\.docx$', f, re.IGNORECASE)
                    if m:
                        c_id = m.group(1)
                        if c_id not in courses:
                            courses[c_id] = {
                                'sem': sem,
                                'id': c_id,
                                'name': m.group(2),
                                'source': os.path.join(orig_dir, f),
                                'type': 'docx'
                            }
    return courses

def main():
    print("=" * 75)
    print(" BATCH GENERATOR KONTRAK PERKULIAHAN RPS OBE 2026 - UNIROW TUBAN")
    print("=" * 75)

    courses = find_all_rps_courses()
    sorted_keys = sorted(courses.keys(), key=lambda x: [int(p) for p in x.split('_')])
    total_courses = len(sorted_keys)
    print(f"[*] Ditemukan total {total_courses} mata kuliah RPS terdaftar dari Semester 1 s.d. 7.\n")

    results = []
    success_count = 0
    start_time = time.time()

    for idx, c_id in enumerate(sorted_keys, 1):
        info = courses[c_id]
        src_path = info['source']
        tail_name = info['name'].replace(' ', '_')
        clean_tail = re.sub(r'_(?:FINAL|OBE)$', '', tail_name)
        out_base = f"{c_id}_Kontrak_Kuliah_{clean_tail}_OBE"
        
        docx_final_path = os.path.join(TARGET_DIR_FINAL, f"{out_base}.docx")
        md_final_path = os.path.join(TARGET_DIR_FINAL, f"{out_base}.md")
        
        docx_obe_path = os.path.join(TARGET_DIR_OBE, f"{out_base}.docx")
        md_obe_path = os.path.join(TARGET_DIR_OBE, f"{out_base}.md")

        print(f"[{idx:02d}/{total_courses:02d}] Memproses: {c_id} - {clean_tail} (Sem {info['sem']})...", end=" ", flush=True)

        try:
            data = parse_rps_data_for_kontrak(src_path)
            
            # Validasi minimal
            if not data.get('semester'):
                data['semester'] = str(info['sem'])
            if not data.get('dosen_pengampu') or data.get('dosen_pengampu') == 'Tim Dosen Pengampu Program Studi PPKn UNIROW Tuban':
                data['dosen_pengampu'] = 'Mario Fahmi Syahrial, M.Pd. / Tim Dosen PPKn'

            # 1. Bangun Word (.docx)
            build_kontrak_docx(data, docx_final_path)
            # Salin ke direktori OBE sekunder
            shutil.copy2(docx_final_path, docx_obe_path)

            # 2. Bangun Markdown (.md)
            build_kontrak_md(data, md_final_path)
            # Salin ke direktori OBE sekunder
            shutil.copy2(md_final_path, md_obe_path)

            mk_name = data.get('nama_mk', clean_tail)
            mk_code = data.get('kode_mk', '-')
            sks_val = data.get('sks', '2')
            n_mg = len(data.get('jadwal', []))
            n_cpmk = len(data.get('cpmk', []))

            results.append({
                'no': idx,
                'id': c_id,
                'sem': info['sem'],
                'mk': mk_name,
                'kode': mk_code,
                'sks': sks_val,
                'jadwal': n_mg,
                'cpmk': n_cpmk,
                'file_docx': os.path.basename(docx_final_path),
                'status': 'SUKSES'
            })
            success_count += 1
            print(f"[OK] -> {mk_name} ({mk_code}, {sks_val} SKS) - {n_mg} mg")
        except Exception as e:
            print(f"[FAIL] Error: {e}")
            results.append({
                'no': idx,
                'id': c_id,
                'sem': info['sem'],
                'mk': clean_tail,
                'kode': '-',
                'sks': '-',
                'jadwal': 0,
                'cpmk': 0,
                'file_docx': f"{out_base}.docx",
                'status': f'ERROR: {e}'
            })

    elapsed = time.time() - start_time
    print("\n" + "=" * 75)
    print(f" GENERASI SELESAI: {success_count}/{total_courses} Berhasil dalam {elapsed:.1f} detik.")
    print(f" Dokumen disimpan di:")
    print(f" 1. {TARGET_DIR_FINAL}")
    print(f" 2. {TARGET_DIR_OBE}")
    print("=" * 75)

    # Tulis laporan hasil batch
    report_path = os.path.join(TARGET_DIR_FINAL, "00_REKAPITULASI_KONTRAK_KULIAH_SEMUA_RPS.md")
    with open(report_path, 'w', encoding='utf-8') as f:
        f.write("# REKAPITULASI KONTRAK PERKULIAHAN PRODI PPKN OBE 2026\n\n")
        f.write("## UNIVERSITAS PGRI RONGGOLAWE (UNIROW) TUBAN\n\n")
        f.write(f"Total Dokumen Dihasilkan: **{success_count} / {total_courses} Mata Kuliah**  \n")
        f.write(f"Format: Word (`.docx`) A4 Portrait Ber-KOP Resmi & Markdown (`.md`)  \n")
        f.write(f"Waktu Pembuatan: {time.strftime('%d %B %Y, %H:%M:%S WIB')}  \n\n")
        f.write("| No | ID MK | Sem | Nama Mata Kuliah | Kode MK | SKS | Silabus Mg | CPMK | Status | Berkas Word (.docx) |\n")
        f.write("|:---:|:---:|:---:|---|:---:|:---:|:---:|:---:|:---:|---|\n")
        for r in results:
            f.write(f"| {r['no']} | {r['id']} | {r['sem']} | {r['mk']} | {r['kode']} | {r['sks']} | {r['jadwal']} | {r['cpmk']} | **{r['status']}** | `{r['file_docx']}` |\n")

    print(f"[OK] Berkas rekapitulasi tersimpan di: {report_path}")

if __name__ == '__main__':
    main()
