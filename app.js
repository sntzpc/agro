(() => {
  'use strict';

  /* ============================================================
   * KLP1 AGRO — Single-Page App (dipisah ke file JS)
   * Struktur kode dibagi per BLOK FUNGSI dengan komentar detail:
   *   01) KONFIGURASI & STATE
   *   02) UTILITAS DOM & FORMAT
   *   03) PERSISTENSI (load/save localStorage)
   *   04) INISIALISASI UI (tabs, modal, listeners)
   *   05) FORM PERAWATAN (add row / save / reset / WA)
   *   06) FORM PANEN (save / reset / WA)
   *   07) ACTTYP MANAGEMENT (load, kelola, tambah)
   *   08) HISTORY & STATISTIK (render, view, edit, delete, filter)
   *   09) EXPORT/IMPORT EXCEL
   *   10) SINKRONISASI KE GOOGLE SHEETS (fetch + JSONP fallback)
   *   11) RESET APP
   *   12) APP INIT
   * ============================================================ */

  /* ========================= 01) KONFIGURASI & STATE ========================= */
  // Konfigurasi integrasi Google Apps Script
  const CONFIG = {
    scriptUrl: 'https://script.google.com/macros/s/AKfycbx_9N6wqjZrGttRcg9JPivPEjqDbDJ2TuPg8c1-Lud2X4h8cFCLQoi5aQ6iYapXbylI/exec', // ganti jika perlu
    sheetId:   '1w2aqNwImlU7jRM28Hauh76ySH18zvEFYkHt-GaXkG3g' // ganti jika perlu
  };

  // State aplikasi disatukan agar mudah dipantau & disimpan
  const state = {
    reports: [],
    userInfo: {},
    lastInputs: { perawatan: {}, panen: {} },
    actTyps:   { perawatan: [], panen: [] }
  };

  // Kunci localStorage (agar konsisten)
  const LS_KEYS = {
    reports:     'klp1AgroReports',
    userInfo:    'klp1AgroUserInfo',
    lastInputs:  'klp1AgroLastInputs',
    actTyps:     'klp1AgroActTyps'
  };

  /* ========================= 02) UTILITAS DOM & FORMAT ======================= */
  // Helper selektor cepat
  const $ = (sel, root = document) => root.querySelector(sel);
  const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  // Format tanggal tampilan dd/mm/yyyy (lokal Indonesia)
  const formatDateForDisplay = (dateStr) => {
    const d = new Date(dateStr);
    return d.toLocaleDateString('id-ID', { day: '2-digit', month: '2-digit', year: 'numeric' });
  };

  // Ambil laporan berdasarkan tanggal & tipe
  const getReportsByDateAndType = (date, type) =>
    state.reports.filter(r => r.tanggal === date && r.type === type);

  // Salin ke clipboard dengan fallback
  function copyToClipboardWithFallback(text) {
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text)
        .then(() => alert('Pesan WA telah disalin ke clipboard!'))
        .catch(() => fallbackCopyTextToClipboard(text));
    } else {
      fallbackCopyTextToClipboard(text);
    }
  }
  function fallbackCopyTextToClipboard(text) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    try {
      const ok = document.execCommand('copy');
      if (ok) alert('Pesan WA telah disalin ke clipboard!');
      else alert('Gagal menyalin otomatis. Salin manual dari preview.');
    } catch {
      alert('Gagal menyalin otomatis. Salin manual dari preview.');
    }
    document.body.removeChild(ta);
  }

  // Mini date picker modal (reusable)
  function showDatePickerModal(onPick) {
    const overlay = document.createElement('div');
    Object.assign(overlay.style, {
      position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.5)',
      display: 'flex', justifyContent: 'center', alignItems: 'center', zIndex: 2000
    });

    const box = document.createElement('div');
    Object.assign(box.style, {
      background: '#fff', padding: '20px', borderRadius: '5px', width: '300px'
    });

    const title = document.createElement('h3');
    title.textContent = 'Pilih Tanggal Laporan';
    const inp = document.createElement('input');
    inp.type = 'date';
    inp.style.width = '100%';
    inp.style.padding = '10px';
    inp.style.margin = '10px 0';
    inp.value = new Date().toISOString().split('T')[0];

    const rowBtn = document.createElement('div');
    Object.assign(rowBtn.style, { display: 'flex', justifyContent: 'flex-end', gap: '10px', marginTop: '15px' });

    const btnCancel = document.createElement('button');
    btnCancel.textContent = 'Batal';
    btnCancel.className = 'danger';
    btnCancel.onclick = () => document.body.removeChild(overlay);

    const btnOk = document.createElement('button');
    btnOk.textContent = 'OK';
    btnOk.onclick = () => { document.body.removeChild(overlay); onPick(inp.value); };

    rowBtn.append(btnCancel, btnOk);
    box.append(title, inp, rowBtn);
    overlay.append(box);
    document.body.appendChild(overlay);
  }

  /* ========================= 03) PERSISTENSI (localStorage) ================== */
  function loadData() {
    try {
      const rep = localStorage.getItem(LS_KEYS.reports);
      const usr = localStorage.getItem(LS_KEYS.userInfo);
      const last = localStorage.getItem(LS_KEYS.lastInputs);
      const act = localStorage.getItem(LS_KEYS.actTyps);
      if (rep)  state.reports   = JSON.parse(rep);
      if (usr)  state.userInfo  = JSON.parse(usr);
      if (last) state.lastInputs= JSON.parse(last);
      if (act)  state.actTyps   = JSON.parse(act);
    } catch (e) {
      console.warn('Gagal load localStorage:', e);
    }
  }
  function saveData() {
    localStorage.setItem(LS_KEYS.reports, JSON.stringify(state.reports));
    localStorage.setItem(LS_KEYS.userInfo, JSON.stringify(state.userInfo));
    localStorage.setItem(LS_KEYS.lastInputs, JSON.stringify(state.lastInputs));
    localStorage.setItem(LS_KEYS.actTyps, JSON.stringify(state.actTyps));
  }

  /* ========================= 04) INISIALISASI UI ============================= */
  // Switch tab (klik tab -> tampilkan konten terkait)
  function wireTabs() {
    $$('.tab').forEach(tab => {
      tab.addEventListener('click', function () {
        $$('.tab').forEach(t => t.classList.remove('active'));
        $$('.tab-content').forEach(c => c.classList.remove('active'));
        this.classList.add('active');
        const id = this.dataset.tab;
        $('#' + id).classList.add('active');
      });
    });
  }

  // Modal close (tombol X)
  function wireModals() {
    $$('#userInfoModal .close, #acttypModal .close').forEach(btn => {
      btn.addEventListener('click', function () {
        this.closest('.modal').style.display = 'none';
      });
    });
  }

  // Tampilkan nama mentee/mentor pada Settings
  function updateUserInfoDisplay() {
    $('#currentMentee').textContent = state.userInfo.menteeName || '-';
    $('#currentMentor').textContent = state.userInfo.mentorName || '-';
  }

  // Listener umum (klik tombol, change input, dll)
  function setupEventListeners() {
    // User info modal
    $('#saveUserInfo').addEventListener('click', saveUserInfo);
    $('#editUserInfo').addEventListener('click', editUserInfo);

    // Perawatan
    $('#savePerawatan').addEventListener('click', savePerawatanReport);
    $('#generatePerawatanWA').addEventListener('click', generatePerawatanWA);
    $('#resetPerawatanForm').addEventListener('click', () => resetPerawatanForm(false));
    $('#addPerawatanWorker').addEventListener('click', () => addPerawatanWorkerInput());
    $('#addPerawatanMaterial').addEventListener('click', () => addPerawatanMaterialInput());

    // Panen
    $('#savePanen').addEventListener('click', savePanenReport);
    $('#generatePanenWA').addEventListener('click', generatePanenWA);
    $('#resetPanenForm').addEventListener('click', () => resetPanenForm(false));

    // History
    $('#exportData').addEventListener('click', exportToExcel);
    $('#importData').addEventListener('click', () => $('#fileInput').click());
    $('#fileInput').addEventListener('change', importFromExcel);
    $('#syncData').addEventListener('click', syncWithGoogleSheet);
    $('#deleteAllData').addEventListener('click', confirmDeleteAllData);
    $('#applyFilter').addEventListener('click', displayHistory);
    $('#clearFilter').addEventListener('click', clearFilters);

    // Settings
    $('#manageActtyp').addEventListener('click', manageActtyp);
    $('#addActtyp').addEventListener('click', addActtyp);
    $('#resetApp').addEventListener('click', confirmResetApp);

    // Simpan last inputs bila ada perubahan field (perawatan/panen)
    $$('#perawatan input, #perawatan select, #perawatan textarea').forEach(el => {
      el.addEventListener('change', () => saveLastInputs('perawatan'));
    });
    $$('#panen input, #panen select, #panen textarea').forEach(el => {
      el.addEventListener('change', () => saveLastInputs('panen'));
    });
  }

  // Isi awal form dari lastInputs, dan set tanggal hari ini
  function initializeForms() {
    const t = new Date().toISOString().split('T')[0];
    $('#perawatanTanggal').value = t;
    $('#panenTanggal').value = t;
    $('#filterDate').value = t;

    const Lp = state.lastInputs.perawatan || {};
    const Ln = state.lastInputs.panen || {};
    if (Lp.estate)   $('#perawatanEstate').value = Lp.estate;
    if (Lp.divisi)   $('#perawatanDivisi').value = Lp.divisi;
    if (Lp.actTyp)   $('#perawatanActTyp').value = Lp.actTyp;
    if (Lp.pekerjaan)$('#perawatanPekerjaan').value = Lp.pekerjaan;
    if (Ln.estate)   $('#panenEstate').value = Ln.estate;
    if (Ln.divisi)   $('#panenDivisi').value = Ln.divisi;
    if (Ln.actTyp)   $('#panenActTyp').value = Ln.actTyp;
    if (Ln.pekerjaan)$('#panenPekerjaan').value = Ln.pekerjaan;
  }

  // Simpan nilai terakhir yang diisi pada form tertentu
  function saveLastInputs(formType) {
    const id = formType === 'perawatan' ? 'perawatan' : 'panen';
    const obj = {
      estate:    $(`#${id}Estate`).value,
      divisi:    $(`#${id}Divisi`).value,
      actTyp:    $(`#${id}ActTyp`).value,
      pekerjaan: $(`#${id}Pekerjaan`).value
    };
    state.lastInputs[formType] = obj;
    saveData();
  }

  /* ========================= 05) FORM PERAWATAN ============================== */
  // Tambah input baris Tenaga Kerja
  function addPerawatanWorkerInput(worker = {}) {
    const box = document.createElement('div');
    box.className = 'worker-item';
    box.innerHTML = `
      <input type="text" class="worker-type" placeholder="Jenis (Harian/Borongan/Mekanis)" value="${worker.type || ''}">
      <input type="number" class="worker-count" placeholder="Jumlah" min="0" value="${worker.count || ''}">
      <button type="button" class="remove-btn">×</button>
    `;
    $('#perawatanTenagaKerjaContainer').appendChild(box);
    box.querySelector('.remove-btn').addEventListener('click', () => box.remove());
  }

  // Tambah input baris Bahan
  function addPerawatanMaterialInput(material = {}) {
    const box = document.createElement('div');
    box.className = 'material-item';
    box.innerHTML = `
      <input type="text"  class="material-type"     placeholder="Jenis Bahan" value="${material.type || ''}">
      <input type="number"class="material-quantity" placeholder="Jumlah" min="0" step="0.01" value="${material.quantity || ''}">
      <input type="text"  class="material-unit"     placeholder="Satuan" value="${material.unit || ''}">
      <button type="button" class="remove-btn">×</button>
    `;
    $('#perawatanBahanContainer').appendChild(box);
    box.querySelector('.remove-btn').addEventListener('click', () => box.remove());
  }

  // Simpan laporan perawatan
  function savePerawatanReport() {
    const estate   = $('#perawatanEstate').value.trim();
    const divisi   = $('#perawatanDivisi').value;
    const blok     = $('#perawatanBlok').value.trim();
    const tanggal  = $('#perawatanTanggal').value;
    const actTyp   = $('#perawatanActTyp').value;
    const pekerjaan= $('#perawatanPekerjaan').value.trim();
    const rencanaHa= parseFloat($('#perawatanRencanaHa').value) || 0;
    const aktualHa = parseFloat($('#perawatanAktualHa').value) || 0;

    if (!estate || !divisi || !blok || !tanggal || !actTyp || !pekerjaan) {
      alert('Estate, Divisi, Blok, Tanggal, ActTyp dan Pekerjaan harus diisi!');
      return;
    }

    const workers = [];
    $$('#perawatanTenagaKerjaContainer .worker-item').forEach(item => {
      const type  = item.querySelector('.worker-type').value.trim();
      const count = parseInt(item.querySelector('.worker-count').value) || 0;
      if (type && count > 0) workers.push({ type, count });
    });

    const materials = [];
    $$('#perawatanBahanContainer .material-item').forEach(item => {
      const type     = item.querySelector('.material-type').value.trim();
      const quantity = parseFloat(item.querySelector('.material-quantity').value) || 0;
      const unit     = item.querySelector('.material-unit').value.trim();
      if (type && quantity > 0) materials.push({ type, quantity, unit });
    });

    const keterangan = $('#perawatanKeterangan').value.trim();

    const report = {
      id: Date.now(),
      type: 'perawatan',
      timestamp: new Date().toISOString(),
      estate,
      divisi: parseInt(divisi),
      blok, tanggal, actTyp, pekerjaan,
      rencanaHa, aktualHa,
      workers, materials,
      keterangan,
      synced: false
    };

    state.reports.push(report);
    saveData();
    displayHistory();
    alert('Laporan perawatan berhasil disimpan!');
    resetPerawatanForm(true);
  }

  // Reset form perawatan (preserveLastInputs = true agar tidak hapus lastInputs)
  function resetPerawatanForm(preserveLastInputs = false) {
    if (!preserveLastInputs) { state.lastInputs.perawatan = {}; saveData(); }
    $('#perawatanBlok').value = '';
    $('#perawatanRencanaHa').value = '';
    $('#perawatanAktualHa').value = '';
    $('#perawatanKeterangan').value = '';
    $('#perawatanTenagaKerjaContainer').innerHTML = '';
    $('#perawatanBahanContainer').innerHTML = '';
    addPerawatanWorkerInput();
    addPerawatanMaterialInput();
    $('#perawatanReportPreview').style.display = 'none';
  }

  // Generate WA perawatan berdasarkan tanggal & filter estate/divisi
  function generatePerawatanWA() {
    showDatePickerModal(dateInput => {
      const estate = $('#perawatanEstate').value.trim();
      const divisi = $('#perawatanDivisi').value;
      if (!estate || !divisi) { alert('Estate dan Divisi harus diisi untuk filter laporan!'); return; }

      const list = getReportsByDateAndType(dateInput, 'perawatan')
        .filter(r => r.estate === estate && String(r.divisi) === String(divisi));

      if (!list.length) { alert('Tidak ada laporan perawatan pada tanggal tersebut!'); return; }

      let full = '';
      list.forEach((report, idx) => {
        const formattedDate = formatDateForDisplay(report.tanggal);

        // Hitung akumulasi per ActTyp (SD Hi = akumulasi bulan berjalan s/d hari tsb, SD Bi = akumulasi tahun berjalan s/d hari tsb)
        const sdHiAktualHa = state.reports
          .filter(r => r.type === 'perawatan' &&
            r.estate === estate && String(r.divisi) === String(divisi) &&
            r.actTyp === report.actTyp &&
            new Date(r.tanggal).getMonth() === new Date(report.tanggal).getMonth() &&
            new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
            new Date(r.tanggal).getDate() <= new Date(report.tanggal).getDate())
          .reduce((s, r) => s + (r.aktualHa || 0), 0);

        const sdHiTenagaKerja = state.reports
          .filter(r => r.type === 'perawatan' &&
            r.estate === estate && String(r.divisi) === String(divisi) &&
            r.actTyp === report.actTyp &&
            new Date(r.tanggal).getMonth() === new Date(report.tanggal).getMonth() &&
            new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
            new Date(r.tanggal).getDate() <= new Date(report.tanggal).getDate())
          .reduce((s, r) => s + (r.workers?.reduce((ws, w) => ws + (w.count || 0), 0) || 0), 0);

        const sdBiAktualHa = state.reports
          .filter(r => r.type === 'perawatan' &&
            r.estate === estate && String(r.divisi) === String(divisi) &&
            r.actTyp === report.actTyp &&
            new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
            new Date(r.tanggal) <= new Date(report.tanggal))
          .reduce((s, r) => s + (r.aktualHa || 0), 0);

        const sdBiTenagaKerja = state.reports
          .filter(r => r.type === 'perawatan' &&
            r.estate === estate && String(r.divisi) === String(divisi) &&
            r.actTyp === report.actTyp &&
            new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
            new Date(r.tanggal) <= new Date(report.tanggal))
          .reduce((s, r) => s + (r.workers?.reduce((ws, w) => ws + (w.count || 0), 0) || 0), 0);

        const workersList = (report.workers || []).map(w => `${w.count} ${w.type}`);
        const materialsList = (report.materials || []).map(m => `${m.quantity} ${m.unit || ''} ${m.type}`);

        let msg = `*LAPORAN HARIAN PERAWATAN ${estate}, Divisi ${divisi}, Blok ${report.blok}, ${formattedDate}*\n\n`;
        msg += `Mentee: ${state.userInfo.menteeName || '-'}\n`;
        msg += `Mentor: ${state.userInfo.mentorName || '-'}\n\n`;
        msg += `${idx + 1}. Pekerjaan : ${report.pekerjaan}\n`;
        msg += `    ActTyp: ${report.actTyp}\n\n`;
        msg += `${idx + 2}. Rencana Ha: ${report.rencanaHa}\n\n`;
        msg += `${idx + 3}. Aktual Ha\n`;
        msg += `    Hi: ${report.aktualHa}\n`;
        msg += `    SD Hi: ${sdHiAktualHa}\n`;
        msg += `    SD Bi: ${sdBiAktualHa}\n\n`;
        msg += `${idx + 4}. Bahan\n`;
        msg += `    Hi: ${materialsList.join('\n    ') || '-'}\n\n`;
        msg += `${idx + 5}. Tenaga Kerja\n`;
        msg += `    Hi: ${workersList.join('\n    ') || '-'}\n`;
        msg += `    SD Hi: ${sdHiTenagaKerja}\n`;
        msg += `    SD Bi: ${sdBiTenagaKerja}\n\n`;
        msg += `${idx + 6}. Keterangan: ${report.keterangan || '-'}\n\n`;
        full += msg;
      });
      full += `*Demikian kami sampaikan dan terima kasih*`;

      const preview = $('#perawatanReportPreview');
      preview.textContent = full;
      preview.style.display = 'block';
      copyToClipboardWithFallback(full);
    });
  }

  /* ========================= 06) FORM PANEN ================================== */
  function savePanenReport() {
    const estate   = $('#panenEstate').value.trim();
    const divisi   = $('#panenDivisi').value;
    const blok     = $('#panenBlok').value.trim();
    const tanggal  = $('#panenTanggal').value;
    const actTyp   = $('#panenActTyp').value;
    const pekerjaan= $('#panenPekerjaan').value.trim();

    const rencanaHa= parseFloat($('#panenRencanaHa').value) || 0;
    const rencanaTon= parseFloat($('#panenRencanaTon').value) || 0;
    const aktualHa = parseFloat($('#panenAktualHa').value) || 0;
    const aktualTon= parseFloat($('#panenAktualTon').value) || 0;
    const tenagaKerja = parseInt($('#panenTenagaKerja').value) || 0;
    const kirimTon = parseFloat($('#panenKirimTon').value) || 0;
    const restan   = parseFloat($('#panenRestan').value) || 0;

    if (!estate || !divisi || !blok || !tanggal || !actTyp || !pekerjaan) {
      alert('Estate, Divisi, Blok, Tanggal, ActTyp dan Pekerjaan harus diisi!');
      return;
    }

    const pusingan = $('#panenPusingan').value.trim();
    const feeder   = $('#panenFeeder').value.trim();
    const truk     = $('#panenTruk').value.trim();
    const keterangan = $('#panenKeterangan').value.trim();

    const report = {
      id: Date.now(),
      type: 'panen',
      timestamp: new Date().toISOString(),
      estate,
      divisi: parseInt(divisi),
      blok, tanggal, actTyp, pekerjaan,
      rencanaHa, rencanaTon,
      aktualHa, aktualTon,
      tenagaKerja, kirimTon, restan,
      pusingan, feeder, truk,
      keterangan,
      synced: false
    };

    state.reports.push(report);
    saveData();
    displayHistory();
    alert('Laporan panen berhasil disimpan!');
    resetPanenForm(true);
  }

  function resetPanenForm(preserveLastInputs = false) {
    if (!preserveLastInputs) { state.lastInputs.panen = {}; saveData(); }
    $('#panenBlok').value = '';
    $('#panenRencanaHa').value = '';
    $('#panenRencanaTon').value = '';
    $('#panenAktualHa').value = '';
    $('#panenAktualTon').value = '';
    $('#panenTenagaKerja').value = '';
    $('#panenKirimTon').value = '';
    $('#panenRestan').value = '';
    $('#panenPusingan').value = '';
    $('#panenFeeder').value = '';
    $('#panenTruk').value = '';
    $('#panenKeterangan').value = '';
    $('#panenReportPreview').style.display = 'none';
  }

  function generatePanenWA() {
    showDatePickerModal(dateInput => {
      const estate = $('#panenEstate').value.trim();
      const divisi = $('#panenDivisi').value;
      if (!estate || !divisi) { alert('Estate dan Divisi harus diisi untuk filter laporan!'); return; }

      const list = getReportsByDateAndType(dateInput, 'panen')
        .filter(r => r.estate === estate && String(r.divisi) === String(divisi));

      if (!list.length) { alert('Tidak ada laporan panen pada tanggal tersebut!'); return; }

      let full = '';
      list.forEach((report, idx) => {
        const formattedDate = formatDateForDisplay(report.tanggal);

        const sdHiAktualHa = state.reports
          .filter(r => r.type === 'panen' && r.estate === estate && String(r.divisi) === String(divisi) &&
                       r.actTyp === report.actTyp &&
                       new Date(r.tanggal).getMonth() === new Date(report.tanggal).getMonth() &&
                       new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
                       new Date(r.tanggal).getDate() <= new Date(report.tanggal).getDate())
          .reduce((s, r) => s + (r.aktualHa || 0), 0);

        const sdHiAktualTon = state.reports
          .filter(r => r.type === 'panen' && r.estate === estate && String(r.divisi) === String(divisi) &&
                       r.actTyp === report.actTyp &&
                       new Date(r.tanggal).getMonth() === new Date(report.tanggal).getMonth() &&
                       new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
                       new Date(r.tanggal).getDate() <= new Date(report.tanggal).getDate())
          .reduce((s, r) => s + (r.aktualTon || 0), 0);

        const sdHiKirimTon = state.reports
          .filter(r => r.type === 'panen' && r.estate === estate && String(r.divisi) === String(divisi) &&
                       r.actTyp === report.actTyp &&
                       new Date(r.tanggal).getMonth() === new Date(report.tanggal).getMonth() &&
                       new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
                       new Date(r.tanggal).getDate() <= new Date(report.tanggal).getDate())
          .reduce((s, r) => s + (r.kirimTon || 0), 0);

        const sdHiTenagaKerja = state.reports
          .filter(r => r.type === 'panen' && r.estate === estate && String(r.divisi) === String(divisi) &&
                       r.actTyp === report.actTyp &&
                       new Date(r.tanggal).getMonth() === new Date(report.tanggal).getMonth() &&
                       new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
                       new Date(r.tanggal).getDate() <= new Date(report.tanggal).getDate())
          .reduce((s, r) => s + (r.tenagaKerja || 0), 0);

        const sdBiAktualHa = state.reports
          .filter(r => r.type === 'panen' && r.estate === estate && String(r.divisi) === String(divisi) &&
                       r.actTyp === report.actTyp &&
                       new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
                       new Date(r.tanggal) <= new Date(report.tanggal))
          .reduce((s, r) => s + (r.aktualHa || 0), 0);

        const sdBiAktualTon = state.reports
          .filter(r => r.type === 'panen' && r.estate === estate && String(r.divisi) === String(divisi) &&
                       r.actTyp === report.actTyp &&
                       new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
                       new Date(r.tanggal) <= new Date(report.tanggal))
          .reduce((s, r) => s + (r.aktualTon || 0), 0);

        const sdBiKirimTon = state.reports
          .filter(r => r.type === 'panen' && r.estate === estate && String(r.divisi) === String(divisi) &&
                       r.actTyp === report.actTyp &&
                       new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
                       new Date(r.tanggal) <= new Date(report.tanggal))
          .reduce((s, r) => s + (r.kirimTon || 0), 0);

        const sdBiTenagaKerja = state.reports
          .filter(r => r.type === 'panen' && r.estate === estate && String(r.divisi) === String(divisi) &&
                       r.actTyp === report.actTyp &&
                       new Date(r.tanggal).getFullYear() === new Date(report.tanggal).getFullYear() &&
                       new Date(r.tanggal) <= new Date(report.tanggal))
          .reduce((s, r) => s + (r.tenagaKerja || 0), 0);

        let msg = `*LAPORAN HARIAN PANEN ${estate}, Divisi ${divisi}, Blok ${report.blok}, ${formattedDate}*\n\n`;
        msg += `Mentee: ${state.userInfo.menteeName || '-'}\n`;
        msg += `Mentor: ${state.userInfo.mentorName || '-'}\n\n`;
        msg += `${idx + 1}. Pekerjaan: ${report.pekerjaan}\n`;
        msg += `    ActTyp: ${report.actTyp}\n\n`;
        msg += `${idx + 2}. Rencana Ton: ${report.rencanaTon}\n\n`;
        msg += `${idx + 3}. Rencana Ha: ${report.rencanaHa}\n\n`;
        msg += `${idx + 4}. Aktual Ha\n`;
        msg += `    Hi: ${report.aktualHa}\n`;
        msg += `    SD Hi: ${sdHiAktualHa}\n`;
        msg += `    SD Bi: ${sdBiAktualHa}\n\n`;
        msg += `${idx + 5}. Aktual Ton\n`;
        msg += `    Hi: ${report.aktualTon}\n`;
        msg += `    SD Hi: ${sdHiAktualTon}\n`;
        msg += `    SD Bi: ${sdBiAktualTon}\n\n`;
        msg += `${idx + 6}. Kirim Ton\n`;
        msg += `    Hi: ${report.kirimTon}\n`;
        msg += `    SD Hi: ${sdHiKirimTon}\n`;
        msg += `    SD Bi: ${sdBiKirimTon}\n\n`;
        msg += `${idx + 7}. Restan: ${report.restan}\n\n`;
        msg += `${idx + 8}. Tenaga Kerja\n`;
        msg += `    Hi: ${report.tenagaKerja}\n`;
        msg += `    SD Hi: ${sdHiTenagaKerja}\n`;
        msg += `    SD Bi: ${sdBiTenagaKerja}\n\n`;
        msg += `${idx + 9}. Pusingan: ${report.pusingan}\n\n`;
        msg += `${idx + 10}. Feeder: ${report.feeder}\n\n`;
        msg += `${idx + 11}. Truk: ${report.truk}\n\n`;
        msg += `${idx + 12}. Keterangan: ${report.keterangan || '-'}\n\n`;
        full += msg;
      });

      full += `*Demikian kami sampaikan dan terima kasih*`;
      const preview = $('#panenReportPreview');
      preview.textContent = full;
      preview.style.display = 'block';
      copyToClipboardWithFallback(full);
    });
  }

  /* ========================= 07) ACTTYP MANAGEMENT =========================== */
  // Load acttyp ke dropdown (perawatan & panen)
  function loadActTyps() {
    const perSel = $('#perawatanActTyp');
    const panSel = $('#panenActTyp');
    perSel.innerHTML = ''; panSel.innerHTML = '';
    perSel.add(new Option('Pilih ActTyp...', ''));
    panSel.add(new Option('Pilih ActTyp...', ''));
    state.actTyps.perawatan.forEach(a => perSel.add(new Option(`${a.code} - ${a.desc}`, a.code)));
    state.actTyps.panen.forEach(a => panSel.add(new Option(`${a.code} - ${a.desc}`, a.code)));

    // set last selected
    if (state.lastInputs.perawatan.actTyp) perSel.value = state.lastInputs.perawatan.actTyp;
    if (state.lastInputs.panen.actTyp)      panSel.value = state.lastInputs.panen.actTyp;
  }

  // Buka modal kelola acttyp + render list
  function manageActtyp() {
    const modal = $('#acttypModal');
    const list  = $('#acttypList');
    list.innerHTML = '';

    state.actTyps.perawatan.forEach(a => {
      const div = document.createElement('div');
      div.style.marginBottom = '10px';
      div.innerHTML = `<strong>Perawatan:</strong> ${a.code} - ${a.desc}
        <button class="remove-acttyp" data-type="perawatan" data-code="${a.code}" style="float:right;">Hapus</button>`;
      list.appendChild(div);
    });
    state.actTyps.panen.forEach(a => {
      const div = document.createElement('div');
      div.style.marginBottom = '10px';
      div.innerHTML = `<strong>Panen:</strong> ${a.code} - ${a.desc}
        <button class="remove-acttyp" data-type="panen" data-code="${a.code}" style="float:right;">Hapus</button>`;
      list.appendChild(div);
    });

    $$('.remove-acttyp', list).forEach(btn => {
      btn.addEventListener('click', function () {
        const type = this.dataset.type;
        const code = this.dataset.code;
        if (confirm(`Apakah Anda yakin ingin menghapus ActTyp ${code}?`)) {
          state.actTyps[type] = state.actTyps[type].filter(x => x.code !== code);
          saveData();
          loadActTyps();
          manageActtyp(); // refresh list
        }
      });
    });

    modal.style.display = 'block';
  }

  // Tambah acttyp baru
  function addActtyp() {
    const type = $('#acttypType').value;
    const code = $('#newActtypCode').value.trim().toUpperCase();
    const desc = $('#newActtypDesc').value.trim();
    if (!code || !desc) { alert('Kode dan Deskripsi ActTyp harus diisi!'); return; }
    if (state.actTyps[type].some(a => a.code === code)) {
      alert(`Kode ${code} sudah ada untuk jenis ${type}!`); return;
    }
    state.actTyps[type].push({ code, desc });
    saveData();
    loadActTyps();
    $('#newActtypCode').value = '';
    $('#newActtypDesc').value = '';
    manageActtyp();
    alert(`ActTyp ${code} - ${desc} berhasil ditambahkan!`);
  }

  /* ========================= 08) HISTORY & STATISTIK ========================= */
  // Render tabel history + tombol aksi
  function displayHistory() {
    const fDate = $('#filterDate').value;
    const fType = $('#filterType').value;
    const tbody = $('#historyTable tbody');
    tbody.innerHTML = '';

    let arr = [...state.reports];
    if (fDate) arr = arr.filter(r => r.tanggal === fDate);
    if (fType !== 'all') arr = arr.filter(r => r.type === fType);
    arr.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));

    arr.forEach((r, i) => {
      const tr = document.createElement('tr');
      const v = r.type === 'perawatan' ? `${r.aktualHa || 0} Ha` : `${r.aktualTon || 0} Ton`;
      tr.innerHTML = `
        <td>${i + 1}</td>
        <td>${formatDateForDisplay(r.tanggal)}</td>
        <td>${r.type === 'perawatan' ? 'Perawatan' : 'Panen'}</td>
        <td>${r.estate}</td>
        <td>${r.divisi}</td>
        <td>${r.blok}</td>
        <td>${r.pekerjaan}</td>
        <td>${v}</td>
        <td>
          <button class="view-btn" data-id="${r.id}">Lihat</button>
          <button class="edit-btn" data-id="${r.id}">Edit</button>
          <button class="delete-btn" data-id="${r.id}">Hapus</button>
        </td>
      `;
      tbody.appendChild(tr);
    });

    $$('.view-btn').forEach(b => b.addEventListener('click', () => viewReport(b.dataset.id)));
    $$('.edit-btn').forEach(b => b.addEventListener('click', () => editReport(b.dataset.id)));
    $$('.delete-btn').forEach(b => b.addEventListener('click', () => deleteReport(b.dataset.id)));

    updateStats();
  }

  // Update ringkasan statistik (total, perawatan, panen)
  function updateStats() {
    $('#totalReports').textContent   = state.reports.length;
    $('#totalPerawatan').textContent = state.reports.filter(r => r.type === 'perawatan').length;
    $('#totalPanen').textContent     = state.reports.filter(r => r.type === 'panen').length;
  }

  // Detail laporan (alert sederhana)
  function viewReport(id) {
    const r = state.reports.find(x => String(x.id) === String(id));
    if (!r) return;
    let msg = '';
    if (r.type === 'perawatan') {
      msg = `*LAPORAN PERAWATAN*\n\n` +
        `Estate: ${r.estate}\nDivisi: ${r.divisi}\nBlok: ${r.blok}\nTanggal: ${r.tanggal}\n` +
        `ActTyp: ${r.actTyp}\nPekerjaan: ${r.pekerjaan}\n` +
        `Rencana Ha: ${r.rencanaHa}\nAktual Ha: ${r.aktualHa}\n` +
        `Tenaga Kerja:\n` +
        ((r.workers && r.workers.length)
          ? r.workers.map(w => `- ${w.type}: ${w.count}`).join('\n')
          : '- Tidak ada data') + `\n` +
        `Bahan:\n` +
        ((r.materials && r.materials.length)
          ? r.materials.map(m => `- ${m.type}: ${m.quantity} ${m.unit || ''}`).join('\n')
          : '- Tidak ada data') + `\n` +
        `Keterangan: ${r.keterangan || '-'}\n` +
        `Status Sync: ${r.synced ? 'Tersinkronisasi' : 'Belum sinkron'}`;
    } else {
      msg = `*LAPORAN PANEN & TRANSPORT*\n\n` +
        `Estate: ${r.estate}\nDivisi: ${r.divisi}\nBlok: ${r.blok}\nTanggal: ${r.tanggal}\n` +
        `ActTyp: ${r.actTyp}\nPekerjaan: ${r.pekerjaan}\n` +
        `Rencana Ha: ${r.rencanaHa}\nRencana Ton: ${r.rencanaTon}\n` +
        `Aktual Ha: ${r.aktualHa}\nAktual Ton: ${r.aktualTon}\n` +
        `Tenaga Kerja: ${r.tenagaKerja}\nKirim Ton: ${r.kirimTon}\nRestan: ${r.restan}\n` +
        `Pusingan: ${r.pusingan}\nFeeder: ${r.feeder}\nTruk: ${r.truk}\n` +
        `Keterangan: ${r.keterangan || '-'}\n` +
        `Status Sync: ${r.synced ? 'Tersinkronisasi' : 'Belum sinkron'}`;
    }
    alert(msg);
  }

  // Muat laporan ke form untuk diedit (menghapus sementara dari list agar tidak dobel)
  function editReport(id) {
    const r = state.reports.find(x => String(x.id) === String(id));
    if (!r) return;

    if (r.type === 'perawatan') {
      $('.tab[data-tab="perawatan"]').click();
      $('#perawatanEstate').value = r.estate;
      $('#perawatanDivisi').value = r.divisi;
      $('#perawatanBlok').value = r.blok;
      $('#perawatanTanggal').value = r.tanggal;
      $('#perawatanActTyp').value = r.actTyp;
      $('#perawatanPekerjaan').value = r.pekerjaan;
      $('#perawatanRencanaHa').value = r.rencanaHa;
      $('#perawatanAktualHa').value = r.aktualHa;
      $('#perawatanKeterangan').value = r.keterangan;

      $('#perawatanTenagaKerjaContainer').innerHTML = '';
      $('#perawatanBahanContainer').innerHTML = '';

      if (r.workers?.length) r.workers.forEach(w => addPerawatanWorkerInput(w));
      else addPerawatanWorkerInput();

      if (r.materials?.length) r.materials.forEach(m => addPerawatanMaterialInput(m));
      else addPerawatanMaterialInput();

      state.reports = state.reports.filter(x => String(x.id) !== String(id));
      saveData();
    } else {
      $('.tab[data-tab="panen"]').click();
      $('#panenEstate').value = r.estate;
      $('#panenDivisi').value = r.divisi;
      $('#panenBlok').value = r.blok;
      $('#panenTanggal').value = r.tanggal;
      $('#panenActTyp').value = r.actTyp;
      $('#panenPekerjaan').value = r.pekerjaan;
      $('#panenRencanaHa').value = r.rencanaHa;
      $('#panenRencanaTon').value = r.rencanaTon;
      $('#panenAktualHa').value = r.aktualHa;
      $('#panenAktualTon').value = r.aktualTon;
      $('#panenTenagaKerja').value = r.tenagaKerja;
      $('#panenKirimTon').value = r.kirimTon;
      $('#panenRestan').value = r.restan;
      $('#panenPusingan').value = r.pusingan;
      $('#panenFeeder').value = r.feeder;
      $('#panenTruk').value = r.truk;
      $('#panenKeterangan').value = r.keterangan;

      state.reports = state.reports.filter(x => String(x.id) !== String(id));
      saveData();
    }
    alert('Laporan dimuat ke form untuk diedit. Silakan perbaiki data dan simpan kembali.');
  }

  // Hapus laporan
  function deleteReport(id) {
    if (!confirm('Apakah Anda yakin ingin menghapus laporan ini?')) return;
    state.reports = state.reports.filter(r => String(r.id) !== String(id));
    saveData();
    displayHistory();
  }

  // Hapus filter
  function clearFilters() {
    $('#filterDate').value = '';
    $('#filterType').value = 'all';
    displayHistory();
  }

  /* ========================= 09) EXPORT/IMPORT EXCEL ========================= */
  // Export semua laporan ke Excel
  function exportToExcel() {
    if (!state.reports.length) { alert('Tidak ada data untuk diexport!'); return; }

    const wb = XLSX.utils.book_new();
    const rows = state.reports.map(r => {
      const row = {
        ID: r.id,
        Tipe: r.type,
        Timestamp: r.timestamp,
        Estate: r.estate,
        Divisi: r.divisi,
        Blok: r.blok,
        Tanggal: r.tanggal,
        ActTyp: r.actTyp,
        Pekerjaan: r.pekerjaan,
        'Rencana Ha': r.rencanaHa || '',
        'Aktual Ha': r.aktualHa || '',
        Keterangan: r.keterangan || '',
        Synced: r.synced ? 'TRUE' : 'FALSE'
      };
      if (r.type === 'perawatan') {
        row['Tenaga Kerja'] = r.workers?.map(w => `${w.type}:${w.count}`).join(';') || '';
        row['Bahan'] = r.materials?.map(m => `${m.type}:${m.quantity}:${m.unit || ''}`).join(';') || '';
        row['Rencana Ton'] = '';
        row['Aktual Ton'] = '';
      } else {
        row['Rencana Ton'] = r.rencanaTon || '';
        row['Aktual Ton'] = r.aktualTon || '';
        row['Tenaga Kerja'] = r.tenagaKerja || '';
        row['Kirim Ton'] = r.kirimTon || '';
        row['Restan'] = r.restan || '';
        row['Pusingan'] = r.pusingan || '';
        row['Feeder'] = r.feeder || '';
        row['Truk'] = r.truk || '';
        row['Bahan'] = '';
      }
      return row;
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'Laporan');
    const dateStr = new Date().toISOString().slice(0,10);
    XLSX.writeFile(wb, `Laporan_KLP1_AGRO_${dateStr}.xlsx`);
  }

  // Import dari Excel -> merge (hindari duplikat ID)
  function importFromExcel(evt) {
    const file = evt.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws);
        if (!data.length) { alert('File Excel tidak mengandung data!'); return; }

        const imported = data.map(row => {
          const base = {
            id: row.ID || (Date.now() + Math.floor(Math.random()*1000)),
            type: row.Tipe,
            timestamp: row.Timestamp || new Date().toISOString(),
            estate: row.Estate, divisi: row.Divisi, blok: row.Blok,
            tanggal: row.Tanggal,
            actTyp: row.ActTyp, pekerjaan: row.Pekerjaan,
            rencanaHa: row['Rencana Ha'] || 0,
            aktualHa: row['Aktual Ha'] || 0,
            keterangan: row.Keterangan || '',
            synced: row.Synced === 'TRUE' || false
          };
          if (row.Tipe === 'perawatan') {
            base.workers = [];
            base.materials = [];
            if (row['Tenaga Kerja']) {
              row['Tenaga Kerja'].split(';').forEach(s => {
                const [type, cnt] = s.split(':');
                if (type && cnt) base.workers.push({ type: type.trim(), count: parseInt(cnt) });
              });
            }
            if (row['Bahan']) {
              row['Bahan'].split(';').forEach(s => {
                const [type, qty, unit] = s.split(':');
                if (type && qty) base.materials.push({ type: type.trim(), quantity: parseFloat(qty), unit: unit || '' });
              });
            }
          } else {
            base.rencanaTon = row['Rencana Ton'] || 0;
            base.aktualTon  = row['Aktual Ton'] || 0;
            base.tenagaKerja= row['Tenaga Kerja'] || 0;
            base.kirimTon   = row['Kirim Ton'] || 0;
            base.restan     = row.Restan || 0;
            base.pusingan   = row.Pusingan || '';
            base.feeder     = row.Feeder || '';
            base.truk       = row.Truk || '';
          }
          return base;
        });

        const existingIds = new Set(state.reports.map(r => String(r.id)));
        const newOnes = imported.filter(r => !existingIds.has(String(r.id)));
        if (!newOnes.length) { alert('Semua data dalam file sudah ada di sistem!'); return; }

        state.reports = [...state.reports, ...newOnes];
        saveData();
        displayHistory();
        alert(`Berhasil mengimpor ${newOnes.length} laporan baru!`);
      } catch (err) {
        console.error('Error importing data:', err);
        alert('Gagal mengimpor data. Pastikan format file sesuai!');
      }
    };
    reader.readAsBinaryString(file);
    evt.target.value = '';
  }

  /* ========================= 10) SYNC GOOGLE SHEETS ========================== */
  // Sinkronisasi hanya data yang belum synced
  function syncWithGoogleSheet() {
    if (!state.reports.length) { alert('Tidak ada data untuk disinkronisasi!'); return; }
    const unsynced = state.reports.filter(r => !r.synced);
    if (!unsynced.length) { alert('Semua data sudah tersinkronisasi!'); return; }

    const status = $('#syncStatus');
    status.textContent = 'Menyinkronisasi data...';
    status.className = 'sync-status sync-warning';
    status.style.display = 'block';

    const payload = unsynced.map(r => ({
      id: r.id,
      type: r.type,
      timestamp: r.timestamp,
      estate: r.estate,
      divisi: r.divisi,
      blok: r.blok || '',
      tanggal: r.tanggal,
      actTyp: r.actTyp || '',
      pekerjaan: r.pekerjaan || '',
      rencanaHa: r.rencanaHa || 0,
      aktualHa: r.aktualHa || 0,
      rencanaTon: r.type === 'panen' ? (r.rencanaTon || 0) : undefined,
      aktualTon: r.type === 'panen' ? (r.aktualTon || 0) : undefined,
      tenagaKerja: r.tenagaKerja || 0,
      kirimTon: r.type === 'panen' ? (r.kirimTon || 0) : undefined,
      restan: r.type === 'panen' ? (r.restan || 0) : undefined,
      pusingan: r.type === 'panen' ? (r.pusingan || '') : undefined,
      feeder: r.type === 'panen' ? (r.feeder || '') : undefined,
      truk: r.type === 'panen' ? (r.truk || '') : undefined,
      keterangan: r.keterangan || '',
      menteeName: state.userInfo.menteeName || '',
      mentorName: state.userInfo.mentorName || '',
      workers: r.workers || [],
      materials: r.materials || []
    }));

    fetch(CONFIG.scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'appendData',
        sheetId: CONFIG.sheetId,
        timestamp: new Date().toISOString(),
        data: payload
      })
    })
      .then(res => { if (!res.ok) throw new Error(`HTTP ${res.status}`); return res.json(); })
      .then(data => {
        if (!data.success || !Array.isArray(data.syncedIds)) throw new Error(data.message || 'Response tidak valid');

        data.syncedIds.forEach(id => {
          const r = state.reports.find(x => String(x.id) === String(id));
          if (r) r.synced = true;
        });
        saveData();
        displayHistory();

        status.textContent = `Berhasil sinkronisasi ${data.syncedIds.length} laporan!`;
        status.className = 'sync-status sync-success';
      })
      .catch(err => {
        console.error('Sync error:', err);
        status.textContent = 'Gagal sinkronisasi: ' + (err.message || 'Unknown error');
        status.className = 'sync-status sync-error';
        // fallback JSONP
        tryJsonpSync(payload);
      });
  }

  // Fallback JSONP bila fetch gagal (mis. CORS)
  function tryJsonpSync(dataToSync) {
    return new Promise((resolve, reject) => {
      const cb = 'jsonpCallback_' + Math.round(100000 * Math.random());
      const status = $('#syncStatus');

      window[cb] = function (data) {
        delete window[cb];
        if (data.success) {
          data.syncedIds.forEach(id => {
            const r = state.reports.find(x => String(x.id) === String(id));
            if (r) r.synced = true;
          });
          saveData();
          displayHistory();

          status.textContent = `Berhasil sinkronisasi ${data.syncedIds.length} laporan (JSONP)!`;
          status.className = 'sync-status sync-success';
          resolve(data);
        } else {
          status.textContent = 'Gagal sinkronisasi: ' + (data.message || 'JSONP error');
          status.className = 'sync-status sync-error';
          reject(new Error(data.message || 'JSONP error'));
        }
      };

      const s = document.createElement('script');
      s.src = `${CONFIG.scriptUrl}?data=${encodeURIComponent(JSON.stringify({
        action: 'appendData',
        sheetId: CONFIG.sheetId,
        timestamp: new Date().toISOString(),
        data: dataToSync,
        callback: cb
      }))}`;
      s.onerror = () => {
        delete window[cb];
        status.textContent = 'Gagal sinkronisasi: Koneksi error';
        status.className = 'sync-status sync-error';
        reject(new Error('JSONP request failed'));
      };
      document.body.appendChild(s);
    });
  }

  /* ========================= 11) RESET APP =================================== */
  function confirmDeleteAllData() {
    if (!confirm('Apakah Anda yakin ingin menghapus SEMUA data laporan? Tindakan ini tidak dapat dibatalkan!')) return;
    state.reports = [];
    saveData();
    displayHistory();
    alert('Semua data laporan telah dihapus!');
  }

  function confirmResetApp() {
    if (!confirm('Apakah Anda yakin ingin RESET APLIKASI? Semua data termasuk laporan, pengguna, dan pengaturan akan dihapus!')) return;
    localStorage.clear();
    state.reports = [];
    state.userInfo = {};
    state.lastInputs = { perawatan: {}, panen: {} };
    state.actTyps = {
      perawatan: [
        { code: 'TMPM01', desc: 'Pemupukan manual' },
        { code: 'TMPM02', desc: 'Penyemprotan manual' },
        { code: 'TMPM03', desc: 'Pemangkasan manual' },
        { code: 'TMPM04', desc: 'Penyiangan manual' }
      ],
      panen: [
        { code: 'PNKT01', desc: 'Panen' },
        { code: 'PNKT02', desc: 'Transport' }
      ]
    };
    saveData();
    location.reload();
  }

  // Simpan user info dari modal
  function saveUserInfo() {
    const menteeName = $('#menteeName').value.trim();
    const mentorName = $('#mentorName').value.trim();
    if (!menteeName || !mentorName) { alert('Nama Mentee dan Mentor harus diisi!'); return; }
    state.userInfo = { menteeName, mentorName };
    saveData();
    $('#userInfoModal').style.display = 'none';
    updateUserInfoDisplay();
  }

  // Edit user info (prefill modal)
  function editUserInfo() {
    $('#menteeName').value = state.userInfo.menteeName || '';
    $('#mentorName').value = state.userInfo.mentorName || '';
    $('#userInfoModal').style.display = 'block';
  }

  /* ========================= 12) APP INIT ==================================== */
  document.addEventListener('DOMContentLoaded', () => {
    loadData();

    // Jika user info kosong, paksa isi modal
    if (!state.userInfo.menteeName || !state.userInfo.mentorName) {
      $('#userInfoModal').style.display = 'block';
    } else {
      updateUserInfoDisplay();
    }

    wireTabs();
    wireModals();
    setupEventListeners();
    initializeForms();

    // Input default baris tenaga kerja & bahan saat awal
    addPerawatanWorkerInput();
    addPerawatanMaterialInput();

    // Tampilkan history & dropdown ActTyp
    displayHistory();
    loadActTyps();
  });

})();
