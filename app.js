'use strict';

// ============================================================
// STORAGE
// ============================================================
const STORAGE_KEY = 'jne_ap_v2_data';
const STORAGE_META = 'jne_ap_v2_meta';

function saveToStorage() {
  try {
    const payload = {
      jne: SHEETS.jne.data,
      sla: SHEETS.sla.data,
      price: SHEETS.price.data
    };
    localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    const now = new Date();
    const meta = { savedAt: now.toISOString() };
    localStorage.setItem(STORAGE_META, JSON.stringify(meta));
    updateStorageBadge(now);
    document.getElementById('sb-saved').textContent = now.toLocaleTimeString('id-ID');
  } catch(e) {
    showToast('⚠ Gagal menyimpan ke browser storage', 'warn-toast');
  }
}

function loadFromStorage() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    const metaRaw = localStorage.getItem(STORAGE_META);
    if (!raw) return false;
    const payload = JSON.parse(raw);
    if (payload.jne && Array.isArray(payload.jne)) SHEETS.jne.data = payload.jne;
    if (payload.sla && Array.isArray(payload.sla)) SHEETS.sla.data = payload.sla;
    if (payload.price && Array.isArray(payload.price)) SHEETS.price.data = payload.price;
    if (metaRaw) {
      const meta = JSON.parse(metaRaw);
      const d = new Date(meta.savedAt);
      updateStorageBadge(d);
      document.getElementById('sb-saved').textContent = d.toLocaleTimeString('id-ID');
    }
    return true;
  } catch(e) { return false; }
}

function updateStorageBadge(date) {
  const b = document.getElementById('storage-badge');
  b.textContent = '● Tersimpan ' + date.toLocaleTimeString('id-ID');
  b.className = 'storage-badge saved';
}

// ============================================================
// DATE HELPERS
// ============================================================
function addDaysToDate(dateStr, days) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d)) return '';
  d.setDate(d.getDate() + days);
  return d.toISOString().split('T')[0];
}

function diffDays(dateStr1, dateStr2) {
  if (!dateStr1 || !dateStr2) return 0;
  const d1 = new Date(dateStr1);
  const d2 = new Date(dateStr2);
  if (isNaN(d1) || isNaN(d2)) return 0;
  return Math.round((d2 - d1) / (1000 * 60 * 60 * 24));
}

function addMonthToDate(dateStr, months) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d)) return '';
  d.setMonth(d.getMonth() + months);
  return d.toISOString().split('T')[0];
}

function formatDateDisplay(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  if (isNaN(d)) return dateStr;
  const months = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des'];
  return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

// ============================================================
// FLOW: SYNC GABUNGAN PI
// Derived from JNE where ver_ap === 'Verified'
// + SLA calculation from SLA reference table
// ============================================================
function syncGabungan() {
  const verifiedRows = SHEETS.jne.data.filter(r => r.ver_ap === 'Verified');

  SHEETS.gabungan.data = verifiedRows.map((jneRow, idx) => {
    const resi10 = (jneRow.no_resi || '').replace(/\s+/g,'').slice(-10);
    const slaRef = SHEETS.sla.data.find(s =>
      s.no_resi === jneRow.no_resi || s.resi10 === resi10
    );

    const slaMax = slaRef ? (parseInt(slaRef.sla_max) || 3) : 3;
    const tglDelivery = jneRow.tanggal || jneRow.tgl_pi || '';
    const eta = (slaRef && slaRef.est_pod) ? slaRef.est_pod : addDaysToDate(tglDelivery, slaMax);
    const tglPod = (slaRef && slaRef.final_pod) ? slaRef.final_pod : eta;

    const actual = diffDays(tglDelivery, tglPod);
    const hariTelat = Math.max(0, diffDays(eta, tglPod));
    const pctDenda = hariTelat > 5 ? 0.05 : hariTelat * 0.01;
    const biaya = parseFloat(jneRow.biaya_kirim) || 0;
    const amountDenda = Math.round(pctDenda * biaya);
    const hargaFinalBase = (parseFloat(jneRow.harga_final) > 0) ? parseFloat(jneRow.harga_final) : biaya;
    const hargaFinal = hargaFinalBase - amountDenda;

    return {
      no: idx + 1,
      no_resi: jneRow.no_resi || '',
      no_pi: jneRow.no_pi || '',
      tgl_pi: jneRow.tgl_pi || '',
      no_ap: jneRow.no_ap || '',
      vendor: jneRow.vendor || '',
      resi_vendor: jneRow.resi_vendor || '',
      tanggal: jneRow.tanggal || '',
      asal: jneRow.asal || '',
      kota_tujuan: jneRow.kota_tujuan || '',
      layanan: jneRow.layanan || '',
      berat: jneRow.berat || 0,
      koli: jneRow.koli || 0,
      biaya_kirim: biaya,
      remarks: jneRow.remarks || '',
      ver_sys: jneRow.ver_sys || '',
      ver_bast: jneRow.ver_bast || '',
      ver_ap: jneRow.ver_ap || '',
      pi: jneRow.pi || '',
      tgl_delivery: tglDelivery,
      eta: eta,
      sla: slaMax,
      tgl_pod: tglPod,
      actual: actual,
      hari_telat: hariTelat,
      pct_denda: pctDenda,
      amount_denda: amountDenda,
      resi10: resi10,
      harga_final: hargaFinal,
      diff: hargaFinal - biaya
    };
  });
}

// ============================================================
// FLOW: SYNC DAFTAR AP
// Derived from Gabungan PI, grouped by no_pi
// ============================================================
let daftarAPOverrides = {};

function syncDaftarAP() {
  const groups = {};
  SHEETS.gabungan.data.forEach(row => {
    const key = row.no_pi || 'UNKNOWN';
    if (!groups[key]) {
      groups[key] = {
        vendor: row.vendor || '',
        no_pi: row.no_pi || '',
        no_ap: row.no_ap || '',
        tgl_pi: row.tgl_pi || '',
        rows: []
      };
    }
    groups[key].rows.push(row);
  });

  SHEETS.daftarap.data = Object.values(groups).map((g, idx) => {
    const ov = daftarAPOverrides[g.no_pi] || {};
    const totalDpp = g.rows.reduce((s, r) => s + (r.harga_final || r.biaya_kirim || 0), 0);
    const ppn = Math.round(totalDpp * 0.11);
    const totalPpn = totalDpp + ppn;
    const pph23 = Math.round(totalDpp * 0.02);
    const netto = totalPpn - pph23;
    const jmlResi = g.rows.length;
    const dates = g.rows.map(r => r.tanggal).filter(Boolean).sort();
    const periodeStr = dates.length
      ? (dates[0] === dates[dates.length-1]
          ? formatDateDisplay(dates[0])
          : formatDateDisplay(dates[0]) + ' – ' + formatDateDisplay(dates[dates.length-1]))
      : formatDateDisplay(g.tgl_pi);
    const payDate = ov.tgl_pay || addMonthToDate(g.tgl_pi, 1);

    return {
      no: idx + 1,
      vendor: g.vendor,
      vendor_r: ov.vendor_r || 'JNE',
      no_pi: g.no_pi,
      no_ap: g.no_ap,
      tgl_pi: g.tgl_pi,
      no_inv: ov.no_inv || '',
      no_jha: ov.no_jha || '',
      tgl_inv: ov.tgl_inv || g.tgl_pi,
      tgl_recv: ov.tgl_recv || '',
      periode: periodeStr,
      jml_resi: jmlResi,
      deskripsi: `Jasa Pengiriman JNE – ${jmlResi} Resi – ${g.no_pi}`,
      dpp: totalDpp,
      ppn: ppn,
      total_ppn: totalPpn,
      pph23: pph23,
      netto: netto,
      tgl_pay: payDate,
      status: ov.status || 'Menunggu'
    };
  });
}

function syncFlow(showNotif) {
  const prevGabCount = SHEETS.gabungan.data.length;
  syncGabungan();
  syncDaftarAP();
  const newGabCount = SHEETS.gabungan.data.length;
  updateNavCounts();
  if (showNotif && newGabCount !== prevGabCount) {
    showToast(`⟳ ${newGabCount} data di Gabungan PI, ${SHEETS.daftarap.data.length} batch Daftar AP`, 'flow-toast');
  }
}

// ============================================================
// SHEET DEFINITIONS
// ============================================================
const SHEETS = {
  gabungan: {
    name: 'Gabungan PI', color: '#1a6b3c', frozenCols: 2, derived: true,
    columns: [
      { key:'no',           label:'No',               w:40,  type:'num',      editable:false },
      { key:'no_resi',      label:'NO RESI',           w:150, type:'text',     editable:false, frozen:true },
      { key:'no_pi',        label:'No. PI',            w:130, type:'text',     editable:false, frozen:true },
      { key:'tgl_pi',       label:'Tgl PI',            w:90,  type:'date',     editable:false },
      { key:'no_ap',        label:'No AP',             w:120, type:'text',     editable:false },
      { key:'vendor',       label:'VENDOR',            w:120, type:'text',     editable:false },
      { key:'tanggal',      label:'TANGGAL',           w:90,  type:'date',     editable:false },
      { key:'asal',         label:'ASAL',              w:100, type:'text',     editable:false },
      { key:'kota_tujuan',  label:'TUJUAN',            w:120, type:'text',     editable:false },
      { key:'layanan',      label:'LAYANAN',           w:70,  type:'text',     editable:false },
      { key:'berat',        label:'BERAT (KG)',         w:80,  type:'num',      editable:false },
      { key:'biaya_kirim',  label:'BIAYA KIRIM',        w:120, type:'currency', editable:false },
      { key:'ver_ap',       label:'Ver. AP',            w:110, type:'dropdown', editable:false },
      { key:'tgl_delivery', label:'TGL DELIVERY',       w:110, type:'date',     editable:false },
      { key:'eta',          label:'ETA',               w:90,  type:'date',     editable:false, formula:'tgl_delivery + SLA' },
      { key:'sla',          label:'SLA (hr)',           w:65,  type:'num',      editable:false, formula:'LOOKUP(SLA.sla_max)' },
      { key:'tgl_pod',      label:'Tgl POD',           w:90,  type:'date',     editable:false, formula:'LOOKUP(SLA.final_pod)' },
      { key:'actual',       label:'Actual (hr)',        w:80,  type:'num',      editable:false, formula:'DATEDIF(delivery, pod)' },
      { key:'hari_telat',   label:'Hari Telat',         w:80,  type:'num',      editable:false, formula:'MAX(0, pod - eta)' },
      { key:'pct_denda',    label:'% Denda',           w:75,  type:'pct',      editable:false, formula:'IF(telat>5, 5%, telat*1%)' },
      { key:'amount_denda', label:'Denda (Rp)',         w:110, type:'currency', editable:false, formula:'pct * biaya_kirim' },
      { key:'harga_final',  label:'Harga Final',        w:120, type:'currency', editable:false, formula:'harga_final - denda' },
      { key:'diff',         label:'Diff',              w:100, type:'currency', editable:false, formula:'harga_final - biaya_kirim' },
    ],
    data: []
  },
  jne: {
    name: 'JNE PI Check', color: '#1a4b8c', frozenCols: 1, derived: false,
    columns: [
      { key:'no',           label:'No',               w:40,  type:'num',      editable:false },
      { key:'no_resi',      label:'No Resi',           w:150, type:'text',     editable:true,  frozen:true },
      { key:'no_pi',        label:'No. Proforma Inv',  w:130, type:'text',     editable:true },
      { key:'tgl_pi',       label:'Tgl PI',            w:90,  type:'date',     editable:true },
      { key:'no_ap',        label:'No AP Approve',     w:120, type:'text',     editable:true },
      { key:'vendor',       label:'VENDOR',            w:120, type:'text',     editable:true },
      { key:'resi_vendor',  label:'RESI VENDOR',       w:130, type:'text',     editable:true },
      { key:'tanggal',      label:'TANGGAL',           w:90,  type:'date',     editable:true },
      { key:'asal',         label:'ASAL',              w:100, type:'text',     editable:true },
      { key:'kota_tujuan',  label:'KOTA TUJUAN',       w:130, type:'text',     editable:true },
      { key:'layanan',      label:'LAYANAN',           w:70,  type:'dropdown', editable:true,  options:['YES','OKE','REG','JTR','SS','SPS'] },
      { key:'berat',        label:'BERAT (KG)',         w:80,  type:'num',      editable:true },
      { key:'koli',         label:'KOLI',              w:55,  type:'num',      editable:true },
      { key:'biaya_kirim',  label:'BIAYA KIRIM (Rp.)', w:130, type:'currency', editable:true },
      { key:'remarks',      label:'Remarks',           w:130, type:'text',     editable:true },
      { key:'ver_sys',      label:'Ver. System (TMS)', w:130, type:'dropdown', editable:true, options:['DONE','Pending','Review'] },
      { key:'ver_bast',     label:'Ver. BAST Fisik',   w:120, type:'dropdown', editable:true, options:['DONE','Pending','Review'] },
      { key:'ver_ap',       label:'Verifikasi AP',     w:120, type:'dropdown', editable:true, options:['Verified','Pending','Review'], trigger:true },
      { key:'pi',           label:'PI',                w:100, type:'text',     editable:true },
      { key:'tujuan_final', label:'Tujuan Final',      w:130, type:'text',     editable:true },
      { key:'harga_pks',    label:'Harga PKS',         w:110, type:'currency', editable:true },
      { key:'harga_final',  label:'Harga Final',       w:120, type:'currency', editable:true },
      { key:'remarks2',     label:'Remarks 2',         w:130, type:'text',     editable:true },
      { key:'packing',      label:'Harga Packing?',    w:120, type:'dropdown', editable:true, options:['Ya','Tidak'] },
    ],
    data: [
      { no:1, no_resi:'JD0072345812', no_pi:'PI-2026-031', tgl_pi:'2026-01-15', no_ap:'AP-2026-001', vendor:'JNE Express', resi_vendor:'JNE-001', tanggal:'2026-01-15', asal:'Jakarta', kota_tujuan:'Surabaya', layanan:'YES', berat:2.5, koli:1, biaya_kirim:185000, remarks:'', ver_sys:'DONE', ver_bast:'DONE', ver_ap:'Verified', pi:'PI-2026-031', tujuan_final:'Surabaya', harga_pks:180000, harga_final:179450, remarks2:'', packing:'Tidak' },
      { no:2, no_resi:'JD0072345813', no_pi:'PI-2026-031', tgl_pi:'2026-01-15', no_ap:'AP-2026-001', vendor:'JNE Express', resi_vendor:'JNE-002', tanggal:'2026-01-15', asal:'Jakarta', kota_tujuan:'Medan', layanan:'OKE', berat:5.0, koli:2, biaya_kirim:320000, remarks:'', ver_sys:'DONE', ver_bast:'DONE', ver_ap:'Verified', pi:'PI-2026-031', tujuan_final:'Medan', harga_pks:315000, harga_final:320000, remarks2:'', packing:'Tidak' },
      { no:3, no_resi:'JD0072345814', no_pi:'PI-2026-032', tgl_pi:'2026-01-22', no_ap:'AP-2026-002', vendor:'JNE Express', resi_vendor:'JNE-003', tanggal:'2026-01-22', asal:'Jakarta', kota_tujuan:'Makassar', layanan:'REG', berat:1.2, koli:1, biaya_kirim:95000, remarks:'', ver_sys:'DONE', ver_bast:'DONE', ver_ap:'Verified', pi:'PI-2026-032', tujuan_final:'Makassar', harga_pks:90000, harga_final:90250, remarks2:'', packing:'Tidak' },
      { no:4, no_resi:'JD0072345815', no_pi:'PI-2026-033', tgl_pi:'2026-02-01', no_ap:'', vendor:'JNE Express', resi_vendor:'JNE-004', tanggal:'2026-02-01', asal:'Jakarta', kota_tujuan:'Bandung', layanan:'YES', berat:3.0, koli:1, biaya_kirim:115000, remarks:'', ver_sys:'DONE', ver_bast:'Pending', ver_ap:'Pending', pi:'', tujuan_final:'Bandung', harga_pks:110000, harga_final:0, remarks2:'', packing:'Tidak' },
      { no:5, no_resi:'JD0072345816', no_pi:'PI-2026-033', tgl_pi:'2026-02-01', no_ap:'', vendor:'JNE Express', resi_vendor:'JNE-005', tanggal:'2026-02-01', asal:'Jakarta', kota_tujuan:'Semarang', layanan:'REG', berat:7.5, koli:3, biaya_kirim:245000, remarks:'', ver_sys:'Pending', ver_bast:'Pending', ver_ap:'Pending', pi:'', tujuan_final:'Semarang', harga_pks:0, harga_final:0, remarks2:'', packing:'Tidak' },
      { no:6, no_resi:'JD0072345817', no_pi:'PI-2026-034', tgl_pi:'2026-02-10', no_ap:'', vendor:'JNE Express', resi_vendor:'JNE-006', tanggal:'2026-02-10', asal:'Surabaya', kota_tujuan:'Balikpapan', layanan:'JTR', berat:12.0, koli:2, biaya_kirim:480000, remarks:'', ver_sys:'Pending', ver_bast:'Pending', ver_ap:'Pending', pi:'', tujuan_final:'Balikpapan', harga_pks:0, harga_final:0, remarks2:'', packing:'Tidak' },
      { no:7, no_resi:'JD0072345818', no_pi:'PI-2026-034', tgl_pi:'2026-02-10', no_ap:'', vendor:'JNE Express', resi_vendor:'JNE-007', tanggal:'2026-02-10', asal:'Jakarta', kota_tujuan:'Pekanbaru', layanan:'OKE', berat:4.2, koli:2, biaya_kirim:175000, remarks:'', ver_sys:'Pending', ver_bast:'Pending', ver_ap:'Pending', pi:'', tujuan_final:'Pekanbaru', harga_pks:0, harga_final:0, remarks2:'', packing:'Tidak' },
    ]
  },
  daftarap: {
    name: 'Daftar AP', color: '#b45309', frozenCols: 1, derived: true,
    columns: [
      { key:'no',         label:'No',              w:40,  type:'num',      editable:false },
      { key:'vendor',     label:'Vendor',           w:120, type:'text',     editable:false, frozen:true },
      { key:'vendor_r',   label:'Vendor (Rename)',  w:120, type:'text',     editable:true },
      { key:'no_pi',      label:'Proforma Invoice', w:130, type:'text',     editable:false },
      { key:'no_ap',      label:'AP Approve',       w:120, type:'text',     editable:true },
      { key:'tgl_pi',     label:'PI Date',          w:90,  type:'date',     editable:false },
      { key:'no_inv',     label:'Invoice Number',   w:130, type:'text',     editable:true },
      { key:'no_jha',     label:'No JHA',           w:100, type:'text',     editable:true },
      { key:'tgl_inv',    label:'Invoice Date',     w:90,  type:'date',     editable:true },
      { key:'tgl_recv',   label:'Invoice Receive',  w:100, type:'date',     editable:true },
      { key:'periode',    label:'Periode',          w:170, type:'text',     editable:false, formula:'MIN/MAX tanggal pengiriman' },
      { key:'jml_resi',   label:'Jumlah Resi',      w:90,  type:'num',      editable:false, formula:'COUNT dari Gabungan PI' },
      { key:'deskripsi',  label:'Deskripsi',        w:240, type:'text',     editable:false, formula:'auto-generate' },
      { key:'dpp',        label:'DPP (Rp.)',         w:130, type:'currency', editable:false },
      { key:'ppn',        label:'PPN (11%)',         w:110, type:'currency', editable:false, formula:'DPP × 11%' },
      { key:'total_ppn',  label:'Total + PPN',       w:120, type:'currency', editable:false, formula:'DPP + PPN' },
      { key:'pph23',      label:'PPh 23 (2%)',       w:110, type:'currency', editable:false, formula:'DPP × 2%' },
      { key:'netto',      label:'Netto',             w:120, type:'currency', editable:false, formula:'Total+PPN − PPh23' },
      { key:'tgl_pay',    label:'Payment Date',      w:100, type:'date',     editable:true },
      { key:'status',     label:'Status',            w:110, type:'dropdown', editable:true,  options:['Menunggu','Lunas','Dibatalkan'] },
    ],
    data: []
  },
  sla: {
    name: 'SLA', color: '#6b3a1a', frozenCols: 0, derived: false,
    columns: [
      { key:'no_resi',    label:'No Resi',          w:150, type:'text', editable:true },
      { key:'resi10',     label:'Resi 10',           w:110, type:'text', editable:true },
      { key:'tanggal',    label:'Tanggal',           w:100, type:'date', editable:true },
      { key:'sla_max',    label:'SLA Max (Days)',     w:110, type:'num',  editable:true },
      { key:'est_pod',    label:'Est POD / ETA',      w:110, type:'date', editable:true },
      { key:'final_pod',  label:'Final Tanggal POD',  w:130, type:'date', editable:true },
      { key:'kota',       label:'Kabupaten/Kota',     w:160, type:'text', editable:true },
    ],
    data: [
      { no_resi:'JD0072345812', resi10:'JD72345812', tanggal:'2026-01-15', sla_max:3, est_pod:'2026-01-17', final_pod:'2026-01-20', kota:'Kab. Surabaya' },
      { no_resi:'JD0072345813', resi10:'JD72345813', tanggal:'2026-01-15', sla_max:5, est_pod:'2026-01-19', final_pod:'2026-01-19', kota:'Kota Medan' },
      { no_resi:'JD0072345814', resi10:'JD72345814', tanggal:'2026-01-22', sla_max:5, est_pod:'2026-01-27', final_pod:'2026-02-01', kota:'Kota Makassar' },
    ]
  },
  price: {
    name: 'PRICE', color: '#5a5a56', frozenCols: 0, derived: false,
    columns: [
      { key:'kota_asal',  label:'Kota Asal',      w:130, type:'text',     editable:true },
      { key:'kota_tuj',   label:'Kota Tujuan',    w:130, type:'text',     editable:true },
      { key:'layanan',    label:'Layanan',         w:80,  type:'dropdown', editable:true, options:['YES','OKE','REG','JTR','SS'] },
      { key:'berat_min',  label:'Berat Min (kg)',  w:110, type:'num',      editable:true },
      { key:'berat_max',  label:'Berat Max (kg)',  w:110, type:'num',      editable:true },
      { key:'harga_per_kg',label:'Harga (Rp/kg)',  w:120, type:'currency', editable:true },
    ],
    data: [
      { kota_asal:'Jakarta', kota_tuj:'Surabaya',   layanan:'YES', berat_min:0, berat_max:10, harga_per_kg:74000 },
      { kota_asal:'Jakarta', kota_tuj:'Medan',       layanan:'OKE', berat_min:0, berat_max:10, harga_per_kg:64000 },
      { kota_asal:'Jakarta', kota_tuj:'Makassar',    layanan:'REG', berat_min:0, berat_max:10, harga_per_kg:79167 },
      { kota_asal:'Jakarta', kota_tuj:'Bandung',     layanan:'YES', berat_min:0, berat_max:10, harga_per_kg:42000 },
      { kota_asal:'Jakarta', kota_tuj:'Semarang',    layanan:'REG', berat_min:0, berat_max:10, harga_per_kg:35000 },
      { kota_asal:'Jakarta', kota_tuj:'Pekanbaru',   layanan:'OKE', berat_min:0, berat_max:10, harga_per_kg:48000 },
      { kota_asal:'Surabaya',kota_tuj:'Balikpapan',  layanan:'JTR', berat_min:0, berat_max:50, harga_per_kg:40000 },
    ]
  }
};

// ============================================================
// STATE
// ============================================================
let currentSheet = 'daftarap';
let filteredData = [];
let selectedCell = null;
let selectedRowIdx = null;
let clipboard = null;
let panelRowIdx = null;
let freezeEnabled = true;
let columnWidths = {};
let ctxTarget = null;
let ddTarget = null;

// ============================================================
// INIT
// ============================================================
function init() {
  const loaded = loadFromStorage();
  if (!loaded) {
    document.getElementById('storage-badge').textContent = '○ Contoh data';
    document.getElementById('storage-badge').className = 'storage-badge';
  }
  syncFlow(false);
  renderSheet(currentSheet);
  showBanner();
  document.addEventListener('click', () => { hideDDPopup(); hideCtxMenu(); });
  document.addEventListener('keydown', handleKeyDown);
}

function showBanner() {
  const b = document.getElementById('flow-banner');
  b.classList.add('show');
}

function closeBanner(id) {
  document.getElementById(id).classList.remove('show');
}

function getSheet() { return SHEETS[currentSheet]; }

// ============================================================
// SWITCH SHEET
// ============================================================
function switchSheet(name) {
  currentSheet = name;
  selectedCell = null;
  selectedRowIdx = null;

  document.querySelectorAll('.nav-item').forEach(el => el.classList.remove('active'));
  const navEl = document.getElementById('nav-' + name);
  if(navEl) navEl.classList.add('active');

  const sheetNames = ['gabungan','jne','daftarap','sla','price'];
  document.querySelectorAll('.sheet-tab').forEach((el,i) => {
    el.classList.toggle('active', sheetNames[i] === name);
  });

  document.getElementById('sb-sheet').textContent = getSheet().name;
  closePanel();
  document.getElementById('tb-search').value = '';

  const isDerived = getSheet().derived;
  const lbanner = document.getElementById('locked-banner');
  lbanner.classList.toggle('show', isDerived);
  document.getElementById('btn-add').style.opacity = isDerived ? '0.4' : '1';
  document.getElementById('btn-add').style.pointerEvents = isDerived ? 'none' : 'auto';
  document.getElementById('btn-del').style.opacity = isDerived ? '0.4' : '1';
  document.getElementById('btn-del').style.pointerEvents = isDerived ? 'none' : 'auto';

  renderSheet(name);
}

// ============================================================
// RENDER
// ============================================================
function renderSheet(sheetKey) {
  const sheet = SHEETS[sheetKey];
  filteredData = [...sheet.data];
  renderHeader(sheet);
  renderBody(sheet, filteredData);
  updateStatusBar();
}

function renderHeader(sheet) {
  const tr = document.getElementById('col-header-row');
  tr.innerHTML = '';
  const corner = document.createElement('th');
  corner.className = 'rn';
  corner.style.position = 'sticky';
  corner.style.top = corner.style.zIndex = '';
  corner.style.zIndex = '15';
  corner.textContent = '#';
  tr.appendChild(corner);

  sheet.columns.forEach((col, ci) => {
    const th = document.createElement('th');
    th.dataset.ci = ci;
    const w = columnWidths[currentSheet + '_' + col.key] || col.w;
    th.style.width = th.style.minWidth = th.style.maxWidth = w + 'px';
    if(freezeEnabled && col.frozen) {
      th.style.position = 'sticky';
      let lo = 40;
      for(let j = 0; j < ci; j++) {
        if(sheet.columns[j].frozen) lo += (columnWidths[currentSheet+'_'+sheet.columns[j].key] || sheet.columns[j].w);
      }
      th.style.left = lo + 'px';
      th.style.zIndex = '12';
      th.style.background = '#e4e4e2';
      th.style.borderRight = '2px solid var(--frozen-border)';
    }
    if(col.formula) th.style.color = 'var(--blue)';
    th.innerHTML = col.label + (col.formula ? ' <span style="color:var(--blue);font-size:8px">ƒ</span>' : '') +
      '<div class="col-resize-handle" onmousedown="startResize(event,' + ci + ')"></div>';
    th.onclick = (e) => { if(!e.target.classList.contains('col-resize-handle')) selectColumn(ci); };
    tr.appendChild(th);
  });
}

function renderBody(sheet, data) {
  const tbody = document.getElementById('grid-body');
  tbody.innerHTML = '';
  data.forEach((row, ri) => {
    const tr = document.createElement('tr');
    tr.dataset.ri = ri;
    if(selectedRowIdx === ri) tr.classList.add('row-selected');

    const rn = document.createElement('td');
    rn.className = 'rn';
    rn.textContent = ri + 1;
    rn.onclick = () => selectRow(ri);
    tr.appendChild(rn);

    sheet.columns.forEach((col, ci) => {
      const td = document.createElement('td');
      td.className = 'cell ' + col.type;
      td.dataset.ri = ri;
      td.dataset.ci = ci;
      if(selectedCell && selectedCell.ri === ri && selectedCell.ci === ci) td.classList.add('selected');

      const w = columnWidths[currentSheet + '_' + col.key] || col.w;
      td.style.width = td.style.minWidth = td.style.maxWidth = w + 'px';

      if(freezeEnabled && col.frozen) {
        td.classList.add('frozen');
        let lo = 40;
        for(let j = 0; j < ci; j++) {
          if(sheet.columns[j].frozen) lo += (columnWidths[currentSheet+'_'+sheet.columns[j].key] || sheet.columns[j].w);
        }
        td.style.left = lo + 'px';
      }

      const val = row[col.key];

      if(col.type === 'dropdown' && col.editable && !sheet.derived) {
        td.innerHTML = '';
        const wrapper = document.createElement('div');
        wrapper.className = 'cell-dropdown';
        const pill = document.createElement('span');
        const pillClass = (val || '').toString().toLowerCase().replace(/\s+/g,'');
        pill.className = 'dd-pill ' + pillClass;
        pill.textContent = val || '—';
        wrapper.appendChild(pill);
        wrapper.onclick = (e) => { e.stopPropagation(); showDDPopup(e, ri, ci, col); };
        td.appendChild(wrapper);
      } else if(col.type === 'dropdown') {
        td.innerHTML = '';
        const pill = document.createElement('span');
        const pillClass = (val || '').toString().toLowerCase().replace(/\s+/g,'');
        pill.className = 'dd-pill ' + pillClass;
        pill.textContent = val || '—';
        td.appendChild(pill);
        td.classList.add('readonly');
      } else {
        td.textContent = formatCellValue(val, col.type);
        if(col.type === 'currency' && parseFloat(val) < 0) td.classList.add('neg');
        if(col.formula) td.classList.add('formula-col');
        if(sheet.derived || !col.editable) td.classList.add('readonly');
      }

      td.onclick = (e) => { e.stopPropagation(); selectCell(ri, ci); };
      td.ondblclick = () => {
        if(!sheet.derived && col.editable && col.type !== 'dropdown') startEdit(ri, ci, col);
      };
      td.oncontextmenu = (e) => { e.preventDefault(); showCtxMenu(e, ri, ci); };

      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

// ============================================================
// CELL FORMAT
// ============================================================
function formatCellValue(val, type) {
  if(val === null || val === undefined || val === '') return '';
  if(type === 'currency') {
    const n = parseFloat(val);
    if(isNaN(n)) return val;
    return (n < 0 ? '(' : '') + 'Rp ' + Math.abs(n).toLocaleString('id-ID') + (n < 0 ? ')' : '');
  }
  if(type === 'pct') {
    const n = parseFloat(val);
    if(isNaN(n)) return val;
    return (n * 100).toFixed(0) + '%';
  }
  if(type === 'num') {
    const n = parseFloat(val);
    if(isNaN(n)) return val;
    return n.toLocaleString('id-ID');
  }
  return val.toString();
}

// ============================================================
// CELL SELECTION & EDITING
// ============================================================
function selectCell(ri, ci) {
  selectedCell = { ri, ci };
  selectedRowIdx = ri;
  const sheet = getSheet();
  const col = sheet.columns[ci];
  const row = filteredData[ri];

  document.getElementById('cell-ref').textContent = String.fromCharCode(65 + ci) + (ri + 1);
  const val = row ? row[col.key] : '';
  document.getElementById('formula-input').value = col.formula ? col.formula : (val !== null && val !== undefined ? val : '');

  if(col.type === 'currency' || col.type === 'num') {
    const vals = filteredData.map(r => parseFloat(r[col.key])).filter(v => !isNaN(v));
    const sum = vals.reduce((a,b)=>a+b, 0);
    document.getElementById('sb-sum').textContent = formatCellValue(sum, col.type);
    document.getElementById('sb-avg').textContent = vals.length ? formatCellValue(sum/vals.length, col.type) : '—';
  } else {
    document.getElementById('sb-sum').textContent = '—';
    document.getElementById('sb-avg').textContent = '—';
  }

  renderBody(sheet, filteredData);
  if(document.getElementById('side-panel').classList.contains('open')) openPanelForRow(ri);
}

function selectRow(ri) {
  selectedRowIdx = ri;
  selectedCell = { ri, ci: 0 };
  renderBody(getSheet(), filteredData);
}

function selectColumn(ci) {
  if(filteredData.length > 0) selectCell(0, ci);
}

function startEdit(ri, ci, col) {
  if(getSheet().derived || !col.editable) return;
  const td = document.querySelector(`td[data-ri="${ri}"][data-ci="${ci}"]`);
  if(!td) return;
  td.classList.add('editing');
  td.innerHTML = '';
  const inp = document.createElement('input');
  inp.className = 'cell-editor';
  inp.type = (col.type === 'num' || col.type === 'currency') ? 'number' : (col.type === 'date' ? 'date' : 'text');
  const val = filteredData[ri][col.key];
  inp.value = val !== null && val !== undefined ? val : '';
  td.appendChild(inp);
  inp.focus();
  inp.select();
  inp.onblur = () => commitEdit(ri, ci, col, inp.value);
  inp.onkeydown = (e) => {
    if(e.key === 'Enter') { e.preventDefault(); inp.blur(); }
    if(e.key === 'Escape') { td.classList.remove('editing'); renderBody(getSheet(), filteredData); }
    e.stopPropagation();
  };
}

function commitEdit(ri, ci, col, newVal) {
  let parsed = newVal;
  if(col.type === 'num' || col.type === 'currency') parsed = parseFloat(newVal) || 0;
  const sheet = getSheet();
  if(sheet.derived) return;

  const oldVerAP = filteredData[ri] && filteredData[ri].ver_ap;
  filteredData[ri][col.key] = parsed;
  sheet.data[ri][col.key] = parsed;

  const isFlowTrigger = (currentSheet === 'jne' && col.key === 'ver_ap');
  const isDaftarAPEditable = (currentSheet === 'daftarap' && col.editable);

  if(isDaftarAPEditable) {
    const row = filteredData[ri];
    if(!daftarAPOverrides[row.no_pi]) daftarAPOverrides[row.no_pi] = {};
    daftarAPOverrides[row.no_pi][col.key] = parsed;
  }

  if(isFlowTrigger) {
    const newVerAP = parsed;
    syncFlow(true);
    if(newVerAP === 'Verified' && oldVerAP !== 'Verified') {
      showToast(`✓ Resi dipindah ke Gabungan PI & Daftar AP`, 'flow-toast');
    } else if(oldVerAP === 'Verified' && newVerAP !== 'Verified') {
      showToast(`↩ Resi ditarik dari Gabungan PI`, 'warn-toast');
    }
  } else if(currentSheet === 'jne') {
    syncFlow(false);
  }

  saveToStorage();
  renderBody(sheet, filteredData);
  updateStatusBar();
  updateNavCounts();
}

// ============================================================
// DROPDOWN
// ============================================================
function showDDPopup(e, ri, ci, col) {
  ddTarget = { ri, ci, col };
  const popup = document.getElementById('dd-popup');
  popup.innerHTML = '';
  col.options.forEach(opt => {
    const item = document.createElement('div');
    item.className = 'dd-popup-item' + (filteredData[ri][col.key] === opt ? ' active' : '');
    const pillClass = opt.toLowerCase().replace(/\s+/g,'');
    item.innerHTML = `<span class="dd-pill ${pillClass}">${opt}</span>`;
    item.onclick = () => {
      commitEdit(ri, ci, col, opt);
      hideDDPopup();
    };
    popup.appendChild(item);
  });
  popup.style.display = 'block';
  popup.style.top = Math.min(e.clientY + 4, window.innerHeight - popup.offsetHeight - 8) + 'px';
  popup.style.left = Math.min(e.clientX, window.innerWidth - 180) + 'px';
}

function hideDDPopup() { document.getElementById('dd-popup').style.display = 'none'; }

// ============================================================
// KEYBOARD
// ============================================================
function handleKeyDown(e) {
  if(!selectedCell) return;
  const { ri, ci } = selectedCell;
  const sheet = getSheet();

  if(e.key === 'Enter' || e.key === 'F2') {
    const col = sheet.columns[ci];
    if(!sheet.derived && col.editable && col.type !== 'dropdown') startEdit(ri, ci, col);
    return;
  }
  if(e.key === 'Delete' || e.key === 'Backspace') {
    if(!sheet.derived) {
      const col = sheet.columns[ci];
      if(col.editable && col.type !== 'dropdown') { commitEdit(ri, ci, col, ''); }
    }
    return;
  }
  const nr = e.key === 'ArrowDown' ? ri+1 : e.key === 'ArrowUp' ? ri-1 : ri;
  const nc = e.key === 'ArrowRight' || e.key === 'Tab' ? ci+1 : e.key === 'ArrowLeft' ? ci-1 : ci;
  if((nr !== ri || nc !== ci) && nr >= 0 && nr < filteredData.length && nc >= 0 && nc < sheet.columns.length) {
    if(e.key === 'Tab') e.preventDefault();
    selectCell(nr, nc);
  }
  if((e.ctrlKey || e.metaKey) && e.key === 'c') {
    const col = sheet.columns[ci];
    clipboard = { val: filteredData[ri][col.key], colType: col.type };
    showToast('Disalin');
  }
  if((e.ctrlKey || e.metaKey) && e.key === 'v') {
    if(clipboard) {
      const col = sheet.columns[ci];
      if(!sheet.derived && col.editable && col.type !== 'dropdown') commitEdit(ri, ci, col, clipboard.val);
    }
  }
  if((e.ctrlKey || e.metaKey) && e.key === 's') {
    e.preventDefault();
    saveToStorage();
    showToast('💾 Disimpan ke browser storage');
  }
}

// ============================================================
// CONTEXT MENU
// ============================================================
function showCtxMenu(e, ri, ci) {
  ctxTarget = { ri, ci };
  selectCell(ri, ci);
  const menu = document.getElementById('ctx-menu');
  menu.style.display = 'block';
  menu.style.top = Math.min(e.clientY, window.innerHeight - 220) + 'px';
  menu.style.left = Math.min(e.clientX, window.innerWidth - 200) + 'px';
}

function hideCtxMenu() { document.getElementById('ctx-menu').style.display = 'none'; }

function ctxCopy() {
  if(!ctxTarget) return;
  const col = getSheet().columns[ctxTarget.ci];
  clipboard = { val: filteredData[ctxTarget.ri][col.key], colType: col.type };
  showToast('Disalin'); hideCtxMenu();
}

function ctxPaste() {
  if(!ctxTarget || !clipboard) return;
  const col = getSheet().columns[ctxTarget.ci];
  if(!getSheet().derived && col.editable && col.type !== 'dropdown') commitEdit(ctxTarget.ri, ctxTarget.ci, col, clipboard.val);
  hideCtxMenu();
}

function ctxInsertAbove() { if(ctxTarget) insertRowAt(ctxTarget.ri); hideCtxMenu(); }
function ctxInsertBelow() { if(ctxTarget) insertRowAt(ctxTarget.ri + 1); hideCtxMenu(); }
function ctxDeleteRow()   { if(ctxTarget) deleteRowAt(ctxTarget.ri); hideCtxMenu(); }
function ctxOpenPanel()   { if(ctxTarget) { openPanelForRow(ctxTarget.ri); togglePanel(true); } hideCtxMenu(); }

// ============================================================
// ROW OPERATIONS
// ============================================================
function addRow() {
  const sheet = getSheet();
  if(sheet.derived) return;
  const newRow = {};
  sheet.columns.forEach(c => { newRow[c.key] = ''; });
  newRow.no = sheet.data.length + 1;
  if(currentSheet === 'jne') {
    newRow.ver_sys = 'Pending';
    newRow.ver_bast = 'Pending';
    newRow.ver_ap = 'Pending';
    newRow.vendor = 'JNE Express';
    newRow.packing = 'Tidak';
  }
  sheet.data.push(newRow);
  filteredData = [...sheet.data];
  renderBody(sheet, filteredData);
  updateStatusBar();
  updateNavCounts();
  document.getElementById('sheet-scroll').scrollTop = 999999;
  selectCell(filteredData.length - 1, 1);
  showToast('＋ Baris baru ditambahkan');
  if(currentSheet === 'jne') syncFlow(false);
  saveToStorage();
}

function insertRowAt(idx) {
  const sheet = getSheet();
  if(sheet.derived) return;
  const newRow = {};
  sheet.columns.forEach(c => { newRow[c.key] = ''; });
  if(currentSheet === 'jne') { newRow.ver_sys = 'Pending'; newRow.ver_bast = 'Pending'; newRow.ver_ap = 'Pending'; }
  sheet.data.splice(idx, 0, newRow);
  filteredData = [...sheet.data];
  filteredData.forEach((r, i) => r.no = i + 1);
  renderBody(sheet, filteredData);
  updateStatusBar();
  showToast('Baris disisipkan');
  saveToStorage();
}

function deleteRowAt(idx) {
  const sheet = getSheet();
  if(sheet.derived) return;
  if(!confirm('Hapus baris ini?')) return;
  sheet.data.splice(idx, 1);
  filteredData = [...sheet.data];
  filteredData.forEach((r, i) => r.no = i + 1);
  renderBody(sheet, filteredData);
  updateStatusBar();
  updateNavCounts();
  if(currentSheet === 'jne') syncFlow(false);
  showToast('Baris dihapus');
  saveToStorage();
}

function deleteSelectedRow() {
  if(selectedRowIdx !== null) deleteRowAt(selectedRowIdx);
}

// ============================================================
// FILTER
// ============================================================
function filterRows(q) {
  const sheet = getSheet();
  const query = q.toLowerCase();
  filteredData = !query ? [...sheet.data] :
    sheet.data.filter(row => Object.values(row).some(v => v !== null && v !== undefined && v.toString().toLowerCase().includes(query)));
  renderBody(sheet, filteredData);
  updateStatusBar();
}

function filterByStatus(val) {
  const sheet = getSheet();
  const key = (currentSheet === 'daftarap') ? 'status' : 'ver_ap';
  filteredData = !val ? [...sheet.data] : sheet.data.filter(r => r[key] === val);
  renderBody(sheet, filteredData);
  updateStatusBar();
}

// ============================================================
// SIDE PANEL
// ============================================================
function togglePanel(forceOpen) {
  const panel = document.getElementById('side-panel');
  const isOpen = panel.classList.contains('open');
  if(forceOpen || !isOpen) {
    panel.classList.add('open');
    document.getElementById('btn-panel').classList.add('active');
    if(selectedRowIdx !== null) openPanelForRow(selectedRowIdx);
    else if(filteredData.length > 0) openPanelForRow(0);
  } else {
    closePanel();
  }
}

function closePanel() {
  document.getElementById('side-panel').classList.remove('open');
  document.getElementById('btn-panel').classList.remove('active');
}

function openPanelForRow(ri) {
  const sheet = getSheet();
  const row = filteredData[ri];
  if(!row) return;
  panelRowIdx = ri;
  document.getElementById('sp-title').textContent = 'Baris ' + (ri+1) + ' — ' + (row.no_resi || row.no_pi || 'Detail');
  const body = document.getElementById('sp-body');

  const editableCols = sheet.columns.filter(c => c.editable && !c.formula);
  const readonlyCols = sheet.columns.filter(c => !c.editable || c.formula);

  let html = '';

  if(sheet.derived) {
    html += '<div style="background:var(--blue-light);border:1px solid #c8d8f0;border-radius:3px;padding:8px 10px;margin-bottom:12px;font-size:10px;color:var(--blue);">🔒 Sheet ini read-only — data diturunkan otomatis dari JNE PI Check.</div>';
  }

  if(editableCols.length && !sheet.derived) {
    html += '<div class="sp-section"><div class="sp-section-label">Data Input</div>';
    editableCols.forEach(col => {
      const val = row[col.key];
      html += `<div class="sp-field"><div class="sp-label">${col.label}</div>`;
      if(col.type === 'dropdown') {
        html += `<select class="sp-select" data-key="${col.key}" onchange="updatePanelField('${col.key}', this.value)">
          ${col.options.map(o => `<option ${o===val?'selected':''}>${o}</option>`).join('')}
        </select>`;
      } else {
        html += `<input class="sp-input" type="${col.type==='num'||col.type==='currency'?'number':col.type==='date'?'date':'text'}"
          value="${val !== null && val !== undefined ? val : ''}"
          data-key="${col.key}"
          onchange="updatePanelField('${col.key}', this.value)">`;
      }
      html += '</div>';
    });
    html += '</div>';
  }

  const showCols = sheet.derived ? sheet.columns : readonlyCols;
  if(showCols.length) {
    html += '<div class="sp-section"><div class="sp-section-label">' + (sheet.derived ? 'Data' : 'Kalkulasi Otomatis') + '</div>';
    showCols.forEach(col => {
      const val = row[col.key];
      let dv = formatCellValue(val, col.type) || '—';
      html += `<div class="sp-field">
        <div class="sp-label">${col.label} ${col.formula ? '<span style="color:var(--blue);font-size:9px">ƒ</span>' : ''}</div>
        ${col.formula
          ? `<div class="sp-formula">${col.formula}</div><div class="sp-calc"><div class="sp-calc-label">Hasil</div><div class="sp-calc-value">${dv}</div></div>`
          : `<div class="sp-value">${dv}</div>`
        }
      </div>`;
    });
    html += '</div>';
  }

  body.innerHTML = html;
  document.getElementById('sp-save-btn').style.display = sheet.derived ? 'none' : '';
}

function updatePanelField(key, val) {
  if(panelRowIdx === null) return;
  const col = getSheet().columns.find(c => c.key === key);
  if(!col) return;
  let parsed = val;
  if(col.type === 'num' || col.type === 'currency') parsed = parseFloat(val) || 0;
  filteredData[panelRowIdx][key] = parsed;
  SHEETS[currentSheet].data[panelRowIdx][key] = parsed;
}

function savePanelRow() {
  if(panelRowIdx === null) return;
  if(currentSheet === 'jne') {
    syncFlow(true);
  }
  if(currentSheet === 'daftarap') {
    const row = filteredData[panelRowIdx];
    if(!daftarAPOverrides[row.no_pi]) daftarAPOverrides[row.no_pi] = {};
    const editableKeys = getSheet().columns.filter(c => c.editable).map(c => c.key);
    editableKeys.forEach(k => { daftarAPOverrides[row.no_pi][k] = row[k]; });
  }
  saveToStorage();
  renderBody(getSheet(), filteredData);
  updateStatusBar();
  updateNavCounts();
  openPanelForRow(panelRowIdx);
  showToast('✓ Baris ' + (panelRowIdx+1) + ' disimpan');
}

// ============================================================
// FREEZE
// ============================================================
function freezeToggle() {
  freezeEnabled = !freezeEnabled;
  document.getElementById('btn-freeze').classList.toggle('active', freezeEnabled);
  renderSheet(currentSheet);
  showToast(freezeEnabled ? '❄ Kolom di-freeze' : 'Freeze dinonaktifkan');
}

// ============================================================
// COLUMN RESIZE
// ============================================================
function startResize(e, ci) {
  e.preventDefault(); e.stopPropagation();
  const sheet = getSheet();
  const col = sheet.columns[ci];
  const startX = e.clientX;
  const startW = columnWidths[currentSheet+'_'+col.key] || col.w;
  document.body.style.cursor = 'col-resize';

  const onMove = (ev) => {
    const newW = Math.max(40, startW + ev.clientX - startX);
    columnWidths[currentSheet+'_'+col.key] = newW;
    document.querySelectorAll(`[data-ci="${ci}"]`).forEach(el => {
      el.style.width = el.style.minWidth = el.style.maxWidth = newW + 'px';
    });
  };
  const onUp = () => {
    document.body.style.cursor = '';
    document.removeEventListener('mousemove', onMove);
    document.removeEventListener('mouseup', onUp);
  };
  document.addEventListener('mousemove', onMove);
  document.addEventListener('mouseup', onUp);
}

// ============================================================
// STATUS & NAV
// ============================================================
function updateStatusBar() {
  document.getElementById('sb-rows').textContent = filteredData.length;
}

function updateNavCounts() {
  const jneTotal = SHEETS.jne.data.length;
  const gabTotal = SHEETS.gabungan.data.length;
  const dapTotal = SHEETS.daftarap.data.length;
  const slaTotal = SHEETS.sla.data.length;
  const priceTotal = SHEETS.price.data.length;

  document.getElementById('nc-jne').textContent = jneTotal;
  document.getElementById('nc-gabungan').textContent = gabTotal;
  document.getElementById('nc-daftarap').textContent = dapTotal;
  document.getElementById('nc-sla').textContent = slaTotal;
  document.getElementById('nc-price').textContent = priceTotal;

  document.getElementById('fp-jne').textContent = jneTotal;
  document.getElementById('fp-gabungan').textContent = gabTotal;
  document.getElementById('fp-daftarap').textContent = dapTotal;
}

// ============================================================
// EXPORT CSV
// ============================================================
function exportCSV() {
  const sheet = getSheet();
  const headers = sheet.columns.map(c => c.label).join(',');
  const rows = filteredData.map(row =>
    sheet.columns.map(col => {
      const v = row[col.key];
      if(v === null || v === undefined) return '';
      const s = v.toString();
      return s.includes(',') || s.includes('"') ? `"${s.replace(/"/g,'""')}"` : s;
    }).join(',')
  );
  const csv = [headers, ...rows].join('\n');
  const blob = new Blob(['\uFEFF' + csv], {type:'text/csv;charset=utf-8;'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = sheet.name.replace(/\s+/g,'-') + '-' + new Date().toISOString().split('T')[0] + '.csv';
  a.click();
  showToast('↓ CSV diunduh: ' + a.download);
}

// ============================================================
// IMPORT CSV
// ============================================================
function openImportModal() {
  document.getElementById('import-modal').classList.add('show');
}

function closeModal(id) {
  document.getElementById(id).classList.remove('show');
}

function doImportCSV() {
  const text = document.getElementById('import-textarea').value.trim();
  if(!text) { showToast('⚠ Tidak ada data untuk diimport', 'warn-toast'); return; }

  const lines = text.split('\n').map(l => l.trim()).filter(l => l);
  if(lines.length < 2) { showToast('⚠ Minimal 2 baris (header + data)', 'warn-toast'); return; }

  const headers = lines[0].split(',').map(h => h.trim().toLowerCase().replace(/\s+/g,'_').replace(/[^a-z0-9_]/g,''));
  const jneCols = SHEETS.jne.columns.map(c => c.key);
  let imported = 0;

  for(let i = 1; i < lines.length; i++) {
    const vals = parseCSVLine(lines[i]);
    if(vals.length === 0) continue;
    const newRow = {};
    SHEETS.jne.columns.forEach(c => { newRow[c.key] = ''; });
    newRow.ver_sys = 'Pending';
    newRow.ver_bast = 'Pending';
    newRow.ver_ap = 'Pending';
    newRow.vendor = newRow.vendor || 'JNE Express';
    newRow.packing = 'Tidak';

    headers.forEach((h, idx) => {
      if(jneCols.includes(h) && vals[idx] !== undefined) {
        let v = vals[idx].trim();
        const col = SHEETS.jne.columns.find(c => c.key === h);
        if(col && (col.type === 'num' || col.type === 'currency')) v = parseFloat(v) || 0;
        newRow[h] = v;
      }
    });

    newRow.no = SHEETS.jne.data.length + 1;
    SHEETS.jne.data.push(newRow);
    imported++;
  }

  if(imported > 0) {
    syncFlow(true);
    saveToStorage();
    if(currentSheet === 'jne') renderSheet('jne');
    updateNavCounts();
    showToast(`✓ ${imported} baris berhasil diimport ke JNE PI Check`, 'flow-toast');
    closeModal('import-modal');
  } else {
    showToast('⚠ Tidak ada baris valid yang dapat diimport', 'warn-toast');
  }
}

function parseCSVLine(line) {
  const result = [];
  let cur = '';
  let inQuote = false;
  for(let i = 0; i < line.length; i++) {
    const ch = line[i];
    if(ch === '"') { inQuote = !inQuote; }
    else if(ch === ',' && !inQuote) { result.push(cur); cur = ''; }
    else { cur += ch; }
  }
  result.push(cur);
  return result;
}

// ============================================================
// RESET
// ============================================================
function confirmReset() {
  document.getElementById('reset-modal').classList.add('show');
}

function doReset() {
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(STORAGE_META);
  daftarAPOverrides = {};
  SHEETS.jne.data = [
    { no:1, no_resi:'JD0072345812', no_pi:'PI-2026-031', tgl_pi:'2026-01-15', no_ap:'AP-2026-001', vendor:'JNE Express', resi_vendor:'JNE-001', tanggal:'2026-01-15', asal:'Jakarta', kota_tujuan:'Surabaya', layanan:'YES', berat:2.5, koli:1, biaya_kirim:185000, remarks:'', ver_sys:'DONE', ver_bast:'DONE', ver_ap:'Verified', pi:'PI-2026-031', tujuan_final:'Surabaya', harga_pks:180000, harga_final:179450, remarks2:'', packing:'Tidak' },
    { no:2, no_resi:'JD0072345813', no_pi:'PI-2026-031', tgl_pi:'2026-01-15', no_ap:'AP-2026-001', vendor:'JNE Express', resi_vendor:'JNE-002', tanggal:'2026-01-15', asal:'Jakarta', kota_tujuan:'Medan', layanan:'OKE', berat:5.0, koli:2, biaya_kirim:320000, remarks:'', ver_sys:'DONE', ver_bast:'DONE', ver_ap:'Verified', pi:'PI-2026-031', tujuan_final:'Medan', harga_pks:315000, harga_final:320000, remarks2:'', packing:'Tidak' },
    { no:3, no_resi:'JD0072345814', no_pi:'PI-2026-032', tgl_pi:'2026-01-22', no_ap:'AP-2026-002', vendor:'JNE Express', resi_vendor:'JNE-003', tanggal:'2026-01-22', asal:'Jakarta', kota_tujuan:'Makassar', layanan:'REG', berat:1.2, koli:1, biaya_kirim:95000, remarks:'', ver_sys:'DONE', ver_bast:'DONE', ver_ap:'Verified', pi:'PI-2026-032', tujuan_final:'Makassar', harga_pks:90000, harga_final:90250, remarks2:'', packing:'Tidak' },
    { no:4, no_resi:'JD0072345815', no_pi:'PI-2026-033', tgl_pi:'2026-02-01', no_ap:'', vendor:'JNE Express', resi_vendor:'JNE-004', tanggal:'2026-02-01', asal:'Jakarta', kota_tujuan:'Bandung', layanan:'YES', berat:3.0, koli:1, biaya_kirim:115000, remarks:'', ver_sys:'DONE', ver_bast:'Pending', ver_ap:'Pending', pi:'', tujuan_final:'Bandung', harga_pks:110000, harga_final:0, remarks2:'', packing:'Tidak' },
    { no:5, no_resi:'JD0072345816', no_pi:'PI-2026-033', tgl_pi:'2026-02-01', no_ap:'', vendor:'JNE Express', resi_vendor:'JNE-005', tanggal:'2026-02-01', asal:'Jakarta', kota_tujuan:'Semarang', layanan:'REG', berat:7.5, koli:3, biaya_kirim:245000, remarks:'', ver_sys:'Pending', ver_bast:'Pending', ver_ap:'Pending', pi:'', tujuan_final:'Semarang', harga_pks:0, harga_final:0, remarks2:'', packing:'Tidak' },
    { no:6, no_resi:'JD0072345817', no_pi:'PI-2026-034', tgl_pi:'2026-02-10', no_ap:'', vendor:'JNE Express', resi_vendor:'JNE-006', tanggal:'2026-02-10', asal:'Surabaya', kota_tujuan:'Balikpapan', layanan:'JTR', berat:12.0, koli:2, biaya_kirim:480000, remarks:'', ver_sys:'Pending', ver_bast:'Pending', ver_ap:'Pending', pi:'', tujuan_final:'Balikpapan', harga_pks:0, harga_final:0, remarks2:'', packing:'Tidak' },
    { no:7, no_resi:'JD0072345818', no_pi:'PI-2026-034', tgl_pi:'2026-02-10', no_ap:'', vendor:'JNE Express', resi_vendor:'JNE-007', tanggal:'2026-02-10', asal:'Jakarta', kota_tujuan:'Pekanbaru', layanan:'OKE', berat:4.2, koli:2, biaya_kirim:175000, remarks:'', ver_sys:'Pending', ver_bast:'Pending', ver_ap:'Pending', pi:'', tujuan_final:'Pekanbaru', harga_pks:0, harga_final:0, remarks2:'', packing:'Tidak' },
  ];
  SHEETS.sla.data = [
    { no_resi:'JD0072345812', resi10:'JD72345812', tanggal:'2026-01-15', sla_max:3, est_pod:'2026-01-17', final_pod:'2026-01-20', kota:'Kab. Surabaya' },
    { no_resi:'JD0072345813', resi10:'JD72345813', tanggal:'2026-01-15', sla_max:5, est_pod:'2026-01-19', final_pod:'2026-01-19', kota:'Kota Medan' },
    { no_resi:'JD0072345814', resi10:'JD72345814', tanggal:'2026-01-22', sla_max:5, est_pod:'2026-01-27', final_pod:'2026-02-01', kota:'Kota Makassar' },
  ];
  closeModal('reset-modal');
  syncFlow(false);
  renderSheet(currentSheet);
  updateNavCounts();
  document.getElementById('storage-badge').textContent = '○ Contoh data';
  document.getElementById('storage-badge').className = 'storage-badge';
  document.getElementById('sb-saved').textContent = '—';
  showToast('↩ Data direset ke contoh awal');
}

// ============================================================
// TOAST
// ============================================================
let toastTimer;
function showToast(msg, cls) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast ' + (cls || '');
  t.classList.add('show');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 2800);
}

// ============================================================
// INIT CALL
// ============================================================
init();