// Data 63 Mata Kuliah PPKN OBE 2026 UNIROW
const COURSES_DATA = [
  // SEMESTER 1
  { no: 1, id: "01_1", sem: 1, name: "Antropologi Budaya", code: "PKN3202", sks: 2, cpmk: 4, file: "01_1_Kontrak_Kuliah_Antropologi_Budaya_OBE.docx" },
  { no: 2, id: "01_2", sem: 1, name: "Ilmu Budaya Dasar", code: "KIP1105", sks: 2, cpmk: 6, file: "01_2_Kontrak_Kuliah_Ilmu_Budaya_Dasar_OBE.docx" },
  { no: 3, id: "01_3", sem: 1, name: "Ilmu Kewarganegaraan", code: "PKN3108", sks: 3, cpmk: 6, file: "01_3_Kontrak_Kuliah_Ilmu_Kewarganegaraan_OBE.docx" },
  { no: 4, id: "01_4", sem: 1, name: "Pancasila", code: "UNV1108", sks: 2, cpmk: 6, file: "01_4_Kontrak_Kuliah_Pancasila_OBE.docx" },
  { no: 5, id: "01_5", sem: 1, name: "Pengantar Ilmu Politik", code: "PKN3106", sks: 3, cpmk: 4, file: "01_5_Kontrak_Kuliah_Pengantar_Ilmu_Politik_OBE.docx" },
  { no: 6, id: "01_6", sem: 1, name: "Pengantar Pendidikan", code: "KIP2101", sks: 2, cpmk: 3, file: "01_6_Kontrak_Kuliah_Pengantar_Pendidikan_OBE.docx" },
  { no: 7, id: "01_7", sem: 1, name: "Perkembangan Peserta Didik", code: "KIP2102", sks: 2, cpmk: 3, file: "01_7_Kontrak_Kuliah_Perkembangan_Peserta_Didik_OBE.docx" },
  { no: 8, id: "01_8", sem: 1, name: "Pengantar Ilmu Hukum dan Pengantar Hukum Indonesia (PIH–PHI)", code: "PKN3110", sks: 4, cpmk: 6, file: "01_8_Kontrak_Kuliah_PIH_PHI_OBE.docx" },
  { no: 9, id: "01_9", sem: 1, name: "Ilmu Negara", code: "PKN3123", sks: 2, cpmk: 4, file: "01_9_Kontrak_Kuliah_Ilmu_Negara_OBE.docx" },

  // SEMESTER 2
  { no: 10, id: "02_1", sem: 2, name: "Pendidikan Agama", code: "UNV1101", sks: 2, cpmk: 4, file: "02_1_Kontrak_Kuliah_Agama_OBE.docx" },
  { no: 11, id: "02_2", sem: 2, name: "Bahasa Indonesia", code: "UNV1106", sks: 2, cpmk: 4, file: "02_2_Kontrak_Kuliah_Bahasa_Indonesia_OBE.docx" },
  { no: 12, id: "02_3", sem: 2, name: "Belajar dan Pembelajaran", code: "KIP2207", sks: 2, cpmk: 4, file: "02_3_Kontrak_Kuliah_Belajar_dan_Pembelajaran_OBE.docx" },
  { no: 13, id: "02_4", sem: 2, name: "Dasar dan Konsep Pendidikan Moral", code: "PKN3105", sks: 2, cpmk: 4, file: "02_4_Kontrak_Kuliah_Dasar_Konsep_Pendidikan_Moral_OBE.docx" },
  { no: 14, id: "02_5", sem: 2, name: "Filsafat Pancasila", code: "PKN3135", sks: 4, cpmk: 4, file: "02_5_Kontrak_Kuliah_Filsafat_Pancasila_OBE.docx" },
  { no: 15, id: "02_6", sem: 2, name: "Filsafat Pendidikan", code: "KIP2208", sks: 2, cpmk: 4, file: "02_6_Kontrak_Kuliah_Filsafat_Pendidikan_OBE.docx" },
  { no: 16, id: "02_7", sem: 2, name: "Hukum Islam", code: "PKN3126", sks: 2, cpmk: 4, file: "02_7_Kontrak_Kuliah_Hukum_Islam_OBE.docx" },
  { no: 17, id: "02_8", sem: 2, name: "Kewarganegaraan", code: "UNV1109", sks: 2, cpmk: 4, file: "02_8_Kontrak_Kuliah_Kewarganegaraan_OBE.docx" },
  { no: 18, id: "02_9", sem: 2, name: "Pengantar Ilmu Sosial", code: "PKN3107", sks: 2, cpmk: 4, file: "02_9_Kontrak_Kuliah_Pengantar_Ilmu_Sosial_OBE.docx" },

  // SEMESTER 3
  { no: 19, id: "03_1", sem: 3, name: "Bahasa Inggris", code: "UNV1105", sks: 2, cpmk: 4, file: "03_1_Kontrak_Kuliah_Bahasa_Inggris_OBE.docx" },
  { no: 20, id: "03_2", sem: 3, name: "Gender dan Pendidikan", code: "PKN3231", sks: 2, cpmk: 4, file: "03_2_Kontrak_Kuliah_Gender_dan_Pendidikan_OBE.docx" },
  { no: 21, id: "03_3", sem: 3, name: "Hukum Perdata", code: "PKN3141", sks: 2, cpmk: 4, file: "03_3_Kontrak_Kuliah_Hukum_Perdata_OBE.docx" },
  { no: 22, id: "03_4", sem: 3, name: "Hukum Tata Negara", code: "PKN3116", sks: 4, cpmk: 4, file: "03_4_Kontrak_Kuliah_Hukum_Tata_Negara_OBE.docx" },
  { no: 23, id: "03_5", sem: 3, name: "Kriminologi", code: "PKN3144", sks: 2, cpmk: 4, file: "03_5_Kontrak_Kuliah_Kriminologi_OBE.docx" },
  { no: 24, id: "03_6", sem: 3, name: "Pengantar Filsafat", code: "PKN3134", sks: 2, cpmk: 4, file: "03_6_Kontrak_Kuliah_Pengantar_Filsafat_OBE.docx" },
  { no: 25, id: "03_7", sem: 3, name: "Profesi dan Kebijakan Pendidikan", code: "KIP2301", sks: 2, cpmk: 4, file: "03_7_Kontrak_Kuliah_Profesi_Kependidikan_OBE.docx" },
  { no: 26, id: "03_8", sem: 3, name: "Sosiologi Politik", code: "PKN3115", sks: 2, cpmk: 4, file: "03_8_Kontrak_Kuliah_Sosiologi_Politik_OBE.docx" },
  { no: 27, id: "03_9", sem: 3, name: "Hukum Pidana", code: "PKN3117", sks: 3, cpmk: 4, file: "03_9_Kontrak_Kuliah_Hukum_Pidana_OBE.docx" },

  // SEMESTER 4
  { no: 28, id: "04_1", sem: 4, name: "Hak Asasi Manusia", code: "PKN3132", sks: 2, cpmk: 4, file: "04_1_Kontrak_Kuliah_Hak_Asasi_Manusia_OBE.docx" },
  { no: 29, id: "04_2", sem: 4, name: "Hubungan Internasional", code: "PKN3138", sks: 2, cpmk: 4, file: "04_2_Kontrak_Kuliah_Hubungan_Internasional_OBE.docx" },
  { no: 30, id: "04_3", sem: 4, name: "Hukum Dagang", code: "PKN3146", sks: 2, cpmk: 4, file: "04_3_Kontrak_Kuliah_Hukum_Dagang_OBE.docx" },
  { no: 31, id: "04_4", sem: 4, name: "Kewarganegaraan Digital", code: "PKN2182", sks: 2, cpmk: 4, file: "04_4_Kontrak_Kuliah_Kewarganegaraan_Digital_OBE.docx" },
  { no: 32, id: "04_5", sem: 4, name: "Pendidikan Ilmu Sosial", code: "PKN2108", sks: 2, cpmk: 4, file: "04_5_Kontrak_Kuliah_Pendidikan_Ilmu_Sosial_OBE.docx" },
  { no: 33, id: "04_6", sem: 4, name: "Pengenalan PGRI (Ke-PGRI-an)", code: "UNV1106", sks: 2, cpmk: 4, file: "04_6_Kontrak_Kuliah_Pengenalan_PGRI_OBE.docx" },
  { no: 34, id: "04_7", sem: 4, name: "Sosiologi Indonesia", code: "PKN5103", sks: 2, cpmk: 4, file: "04_7_Kontrak_Kuliah_Sosiologi_Indonesia_OBE.docx" },
  { no: 35, id: "04_8", sem: 4, name: "Statistik Pendidikan", code: "PKN4116", sks: 2, cpmk: 4, file: "04_8_Kontrak_Kuliah_STATISTIK_OBE.docx" },
  { no: 36, id: "04_9", sem: 4, name: "Teori dan Hukum Konstitusi", code: "PKN3130", sks: 3, cpmk: 4, file: "04_9_Kontrak_Kuliah_Teori_Hukum_Konstitusi_OBE.docx" },

  // SEMESTER 5
  { no: 37, id: "05_1", sem: 5, name: "Etika Digital", code: "PKN2202", sks: 2, cpmk: 4, file: "05_1_Kontrak_Kuliah_Etika_Digital_OBE.docx" },
  { no: 38, id: "05_2", sem: 5, name: "Hukum Acara Pidana", code: "PKN3125", sks: 2, cpmk: 4, file: "05_2_Kontrak_Kuliah_Hukum_Acara_Pidana_OBE.docx" },
  { no: 39, id: "05_3", sem: 5, name: "Hukum Adat", code: "PKN3119", sks: 2, cpmk: 4, file: "05_3_Kontrak_Kuliah_Hukum_Adat_OBE.docx" },
  { no: 40, id: "05_4", sem: 5, name: "Hukum Internasional", code: "PKN3122", sks: 2, cpmk: 4, file: "05_4_Kontrak_Kuliah_Hukum_Internasional_OBE.docx" },
  { no: 41, id: "05_5", sem: 5, name: "Hukum Pajak", code: "PKN3112", sks: 2, cpmk: 4, file: "05_5_Kontrak_Kuliah_Hukum_Pajak_OBE.docx" },
  { no: 42, id: "05_6", sem: 5, name: "Kewirausahaan", code: "PKN5110", sks: 2, cpmk: 4, file: "05_6_Kontrak_Kuliah_Kewirausahaan_OBE.docx" },
  { no: 43, id: "05_7", sem: 5, name: "Logika", code: "PKN4101", sks: 2, cpmk: 4, file: "05_7_Kontrak_Kuliah_Logika_OBE.docx" },
  { no: 44, id: "05_8", sem: 5, name: "Politik Hukum", code: "PKN3118", sks: 2, cpmk: 4, file: "05_8_Kontrak_Kuliah_Politik_Hukum_OBE.docx" },
  { no: 45, id: "05_9", sem: 5, name: "Metodologi Penelitian", code: "KIP2401", sks: 2, cpmk: 4, file: "05_9_Kontrak_Kuliah_Metodologi_Penelitian_OBE.docx" },
  { no: 46, id: "05_10", sem: 5, name: "Sistem Politik Indonesia", code: "PKN3114", sks: 2, cpmk: 4, file: "05_10_Kontrak_Kuliah_Sistem_Politik_Indonesia_OBE.docx" },

  // SEMESTER 6
  { no: 47, id: "06_1", sem: 6, name: "Hukum Acara Tata Usaha Negara (HTUN)", code: "PKN3129", sks: 2, cpmk: 4, file: "06_1_Kontrak_Kuliah_HTUN_OBE.docx" },
  { no: 48, id: "06_2", sem: 6, name: "Hukum Agraria", code: "PKN3111", sks: 2, cpmk: 4, file: "06_2_Kontrak_Kuliah_Hukum_Agraria_OBE.docx" },
  { no: 49, id: "06_3", sem: 6, name: "Ilmu Alamiah Dasar (IAD)", code: "KIP1104", sks: 2, cpmk: 4, file: "06_3_Kontrak_Kuliah_Ilmu_Alamiah_Dasar_OBE.docx" },
  { no: 50, id: "06_4", sem: 6, name: "Pemerintahan Daerah", code: "PKN3109", sks: 2, cpmk: 4, file: "06_4_Kontrak_Kuliah_Pemerintahan_Daerah_OBE.docx" },
  { no: 51, id: "06_5", sem: 6, name: "Pendidikan Anti Korupsi", code: "PKN3136", sks: 2, cpmk: 4, file: "06_5_Kontrak_Kuliah_Pendidikan_Anti_Korupsi_OBE.docx" },
  { no: 52, id: "06_6", sem: 6, name: "Penelitian PPKn", code: "PKN4103", sks: 2, cpmk: 4, file: "06_6_Kontrak_Kuliah_Penelitian_PPKn_OBE.docx" },
  { no: 53, id: "06_7", sem: 6, name: "Psikologi Pendidikan dan Dinamika Kelompok", code: "PKN2184", sks: 2, cpmk: 4, file: "06_7_Kontrak_Kuliah_Psi_Pend_dan_Dinamika_Kelompok_OBE.docx" },
  { no: 54, id: "06_8", sem: 6, name: "Hukum Acara Perdata", code: "PKN3128", sks: 2, cpmk: 4, file: "06_8_Kontrak_Kuliah_Hukum_Acara_Perdata_OBE.docx" },

  // SEMESTER 7
  { no: 55, id: "07_1", sem: 7, name: "PLP - Asesmen Pembelajaran", code: "KIP2701", sks: 3, cpmk: 4, file: "07_1_Kontrak_Kuliah_PLP_Asesmen_Pembelajaran_OBE.docx" },
  { no: 56, id: "07_2", sem: 7, name: "PLP - Pengembangan Strategi Pembelajaran", code: "KIP2708", sks: 2, cpmk: 4, file: "07_2_Kontrak_Kuliah_PLP_Pengembangan_Strategi_Pembelajaran_OBE.docx" },
  { no: 57, id: "07_3", sem: 7, name: "PLP - Manajemen Pengembangan Program Sekolah", code: "KIP2702", sks: 2, cpmk: 4, file: "07_3_Kontrak_Kuliah_PLP_Manajemen_Pengembangan_Program_Sekolah_OBE.docx" },
  { no: 58, id: "07_4", sem: 7, name: "PLP - Praktik Mengajar", code: "KIP2707", sks: 4, cpmk: 4, file: "07_4_Kontrak_Kuliah_PLP_Praktik_Mengajar_OBE.docx" },
  { no: 59, id: "07_5", sem: 7, name: "PLP - Telaah Kurikulum", code: "KIP2706", sks: 2, cpmk: 4, file: "07_5_Kontrak_Kuliah_PLP_Telaah_Kurikulum_OBE.docx" },
  { no: 60, id: "07_6", sem: 7, name: "PLP - Pengembangan Media Pembelajaran", code: "KIP2704", sks: 3, cpmk: 4, file: "07_6_Kontrak_Kuliah_PLP_Pengembangan_Media_Pembelajaran_OBE.docx" },
  { no: 61, id: "07_7", sem: 7, name: "PLP - Perencanaan Pembelajaran", code: "KIP2705", sks: 2, cpmk: 4, file: "07_7_Kontrak_Kuliah_PLP_Perencanaan_Pembelajaran_OBE.docx" },
  { no: 62, id: "07_8", sem: 7, name: "Seminar Bidang Studi", code: "PKN4108", sks: 3, cpmk: 4, file: "07_8_Kontrak_Kuliah_Seminar_Bidang_Studi_OBE.docx" },
  { no: 63, id: "07_9", sem: 7, name: "PLP - Pengembangan Bahan Ajar", code: "KIP2703", sks: 2, cpmk: 4, file: "07_9_Kontrak_Kuliah_PLP_Pengembangan_Bahan_Ajar_OBE.docx" }
];

// State
let currentSemester = 'all';
let searchQuery = '';

// DOM Elements
const courseGrid = document.getElementById('courseGrid');
const searchInput = document.getElementById('searchInput');
const tabButtons = document.querySelectorAll('.tab-btn');
const modalOverlay = document.getElementById('modalOverlay');
const modalContent = document.getElementById('modalDetail');
const themeToggleBtn = document.getElementById('themeToggle');

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  renderCourses();
  setupEventListeners();
  initTheme();
});

// Render Course Cards
function renderCourses() {
  if (!courseGrid) return;
  
  const filtered = COURSES_DATA.filter(course => {
    const matchSem = (currentSemester === 'all') || (course.sem === parseInt(currentSemester, 10));
    const query = searchQuery.toLowerCase();
    const matchSearch = course.name.toLowerCase().includes(query) ||
                        course.code.toLowerCase().includes(query) ||
                        course.id.toLowerCase().includes(query);
    return matchSem && matchSearch;
  });

  if (filtered.length === 0) {
    courseGrid.innerHTML = `
      <div style="grid-column: 1/-1; text-align: center; padding: 48px; color: var(--text-muted);">
        <p style="font-size: 1.2rem; margin-bottom: 8px;">Tidak ada mata kuliah yang cocok.</p>
        <p style="font-size: 0.9rem;">Coba gunakan kata kunci pencarian yang lain.</p>
      </div>
    `;
    return;
  }

  courseGrid.innerHTML = filtered.map(c => `
    <div class="course-card" onclick="openCourseModal('${c.id}')">
      <div>
        <div class="course-header">
          <span class="course-sem-badge">Semester ${c.sem}</span>
          <span class="course-code">${c.code}</span>
        </div>
        <h4 class="course-title">${c.name}</h4>
      </div>
      <div>
        <div class="course-meta">
          <div>
            <div class="meta-item-val">${c.sks} SKS</div>
            <div class="meta-item-lbl">Bobot SKS</div>
          </div>
          <div>
            <div class="meta-item-val">16 Mg</div>
            <div class="meta-item-lbl">Silabus</div>
          </div>
          <div>
            <div class="meta-item-val">${c.cpmk} CPMK</div>
            <div class="meta-item-lbl">Capaian</div>
          </div>
        </div>
        <div class="course-eval-bar">
          <span class="course-status-tag">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>
            Tersinkronisasi ISO
          </span>
          <button class="btn-card-action" type="button">
            Lihat Detail
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14"></path><path d="m12 5 7 7-7 7"></path></svg>
          </button>
        </div>
      </div>
    </div>
  `).join('');
}

// Setup Event Listeners
function setupEventListeners() {
  if (searchInput) {
    searchInput.addEventListener('input', (e) => {
      searchQuery = e.target.value.trim();
      renderCourses();
    });
  }

  tabButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      tabButtons.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      currentSemester = btn.getAttribute('data-sem');
      renderCourses();
    });
  });

  // Guide Tabs
  const guideBtns = document.querySelectorAll('.guide-tab-btn');
  guideBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      guideBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const target = btn.getAttribute('data-target');
      document.querySelectorAll('.guide-content').forEach(c => c.style.display = 'none');
      const el = document.getElementById(target);
      if (el) el.style.display = 'block';
    });
  });
}

// Modal Details
window.openCourseModal = function(id) {
  const c = COURSES_DATA.find(item => item.id === id);
  if (!c || !modalOverlay || !modalContent) return;

  modalContent.innerHTML = `
    <div style="margin-bottom: 20px;">
      <span class="course-sem-badge" style="display:inline-block; margin-bottom: 8px;">Semester ${c.sem} • Kurikulum OBE 2026</span>
      <h2 style="font-size: 1.6rem; font-weight: 800; margin-bottom: 6px;">${c.name}</h2>
      <p style="color: var(--text-muted); font-family: var(--font-mono); font-size: 0.9rem;">Kode MK: ${c.code} • Bobot: ${c.sks} SKS • File: ${c.file}</p>
    </div>

    <div style="background: rgba(99, 102, 241, 0.08); border: 1px solid var(--border-glow); border-radius: var(--radius-md); padding: 18px; margin-bottom: 24px;">
      <h4 style="font-size: 1rem; font-weight: 700; color: #a5b4fc; margin-bottom: 8px;">Struktur Evaluasi & Penilaian Resmi (Total: 100%)</h4>
      <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(110px, 1fr)); gap: 12px; text-align: center;">
        <div style="background: var(--bg-card); padding: 10px; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
          <div style="font-weight: 800; font-size: 1.1rem; color: var(--accent-primary);">10%</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">Presensi (>=75%)</div>
        </div>
        <div style="background: var(--bg-card); padding: 10px; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
          <div style="font-weight: 800; font-size: 1.1rem; color: var(--accent-primary);">20%</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">Tugas Terstruktur</div>
        </div>
        <div style="background: var(--bg-card); padding: 10px; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
          <div style="font-weight: 800; font-size: 1.1rem; color: var(--accent-primary);">10%</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">Kuis / Formatif</div>
        </div>
        <div style="background: var(--bg-card); padding: 10px; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
          <div style="font-weight: 800; font-size: 1.1rem; color: #ec4899;">30%</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">UTS (Mg 8)</div>
        </div>
        <div style="background: var(--bg-card); padding: 10px; border-radius: var(--radius-sm); border: 1px solid var(--border-color);">
          <div style="font-weight: 800; font-size: 1.1rem; color: #ec4899;">30%</div>
          <div style="font-size: 0.75rem; color: var(--text-muted);">UAS (Mg 16)</div>
        </div>
      </div>
    </div>

    <div style="margin-bottom: 24px;">
      <h4 style="font-size: 1rem; font-weight: 700; margin-bottom: 10px;">Roadmap 16 Minggu Perkuliahan</h4>
      <div style="display: flex; gap: 8px; flex-wrap: wrap;">
        ${Array.from({length: 16}, (_, i) => i + 1).map(w => {
          const isUts = w === 8;
          const isUas = w === 16;
          const bg = isUts || isUas ? '#ec4899' : 'rgba(255, 255, 255, 0.08)';
          const label = isUts ? 'UTS' : (isUas ? 'UAS' : `Mg ${w}`);
          return `<span style="background: ${bg}; padding: 6px 12px; border-radius: 6px; font-size: 0.78rem; font-weight: 600; color: #fff;">${label}</span>`;
        }).join('')}
      </div>
    </div>

    <div style="display: flex; gap: 12px; justify-content: flex-end;">
      <button class="btn-primary" onclick="copySkillCommand('${c.name}')" style="padding: 10px 20px; font-size: 0.88rem;">
        Salin Perintah Generate AI
      </button>
    </div>
  `;

  modalOverlay.classList.add('active');
};

window.closeModal = function() {
  if (modalOverlay) {
    modalOverlay.classList.remove('active');
  }
};

window.copyToClipboard = function(text, btnElement) {
  navigator.clipboard.writeText(text).then(() => {
    const originalText = btnElement.innerHTML;
    btnElement.innerHTML = `
      <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>
      Tersalin!
    `;
    btnElement.style.background = 'var(--success)';
    setTimeout(() => {
      btnElement.innerHTML = originalText;
      btnElement.style.background = '';
    }, 2000);
  });
};

window.copySkillCommand = function(courseName) {
  const prompt = `/kontrak-kuliah-rps buatkan kontrak perkuliahan presisi ISO untuk mata kuliah ${courseName} sesuai RPS OBE`;
  navigator.clipboard.writeText(prompt).then(() => {
    alert(`Perintah prompt disalin ke clipboard:\n"${prompt}"\n\nTempelkan di Antigravity, Claude Code, atau Cursor!`);
  });
};

// Dark / Light Theme
function initTheme() {
  const saved = localStorage.getItem('theme') || 'dark';
  document.documentElement.setAttribute('data-theme', saved);
  updateThemeIcon(saved);
  
  if (themeToggleBtn) {
    themeToggleBtn.addEventListener('click', () => {
      const current = document.documentElement.getAttribute('data-theme') || 'dark';
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('theme', next);
      updateThemeIcon(next);
    });
  }
}

function updateThemeIcon(theme) {
  if (!themeToggleBtn) return;
  themeToggleBtn.innerHTML = theme === 'dark' 
    ? `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>`
    : `<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>`;
}
