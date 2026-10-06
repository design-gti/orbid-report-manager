/* ORBID Report Generator
 * Render report 5 halaman (replika desain Figma) dari data agregat.
 * Data masuk via file Excel (drag-drop / pilih file) atau data contoh bawaan.
 */

"use strict";

// ═══════════════════════════════════════════════════════════════
// DATA CONTOH — struktur ini juga jadi acuan format Excel
// ═══════════════════════════════════════════════════════════════
const SAMPLE = {
  info: {
    tim: "Marketing & Customer Growth",
    manager: "Head of Marketing",
    batch: "Batch Q3 2026",
    penerima: "Maya",
    label: "CONTOH · DATA ILUSTRASI",
    stdCommitment: 3.8,
    stdCompetency: 3.6,
    stdContribution: 3.6,
  },
  members: [
    { nama: "Rina Wijaya",   commitment: 3.4, competency: 3.8, contribution: 4.5, blindspot: 2 },
    { nama: "Andre Kusuma",  commitment: 3.3, competency: 4.2, contribution: 4.6, blindspot: 2 },
    { nama: "Budi Hartono",  commitment: 3.1, competency: 3.2, contribution: 3.0, blindspot: 3 },
    { nama: "Eko Santoso",   commitment: 3.6, competency: 3.1, contribution: 3.2, blindspot: 1 },
    { nama: "Nadia Rahma",   commitment: 3.9, competency: 3.4, contribution: 3.5, blindspot: 0 },
    { nama: "Dimas Pratama", commitment: 4.2, competency: 3.9, contribution: 3.8, blindspot: 1 },
    { nama: "Sari Lestari",  commitment: 4.0, competency: 4.1, contribution: 4.2, blindspot: 0 },
    { nama: "Maya Putri",    commitment: 4.4, competency: 3.6, contribution: 3.9, blindspot: 1 },
  ],
  commitmentItems: ["Kedisiplinan & Kehadiran", "Inisiatif & Proaktivitas", "Tanggung Jawab", "Komitmen Target", "Konsistensi Performa"],
  commitment: {
    "Rina Wijaya":   [3.5, 3.3, 3.4, 3.6, 3.2],
    "Andre Kusuma":  [3.4, 3.6, 3.1, 3.2, 3.2],
    "Budi Hartono":  [3.0, 3.2, 3.1, 3.0, 3.2],
    "Eko Santoso":   [3.7, 3.5, 3.6, 3.8, 3.4],
    "Nadia Rahma":   [4.0, 3.8, 3.9, 4.1, 3.7],
    "Dimas Pratama": [4.3, 4.1, 4.2, 4.4, 4.0],
    "Sari Lestari":  [4.1, 3.9, 4.0, 4.2, 3.8],
    "Maya Putri":    [4.5, 4.3, 4.4, 4.6, 4.2],
  },
  competencyAspects: ["Strategic Thinking", "Business Insight", "Collaboration", "Decision Making", "People Dev.", "Planning Org.", "Change Lead.", "Communication"],
  competencyStd: [3.6, 3.5, 3.6, 3.5, 3.5, 3.6, 3.5, 3.6],
  competency: {
    "Rina Wijaya":   [3.5, 3.4, 3.9, 3.8, 3.7, 4.0, 4.1, 4.2],
    "Andre Kusuma":  [4.0, 3.8, 4.3, 4.3, 4.3, 4.4, 4.5, 4.6],
    "Budi Hartono":  [2.9, 2.8, 3.3, 3.1, 3.3, 3.4, 3.5, 3.6],
    "Eko Santoso":   [2.8, 2.8, 3.3, 3.0, 3.1, 3.3, 3.4, 3.4],
    "Nadia Rahma":   [3.1, 3.0, 3.5, 3.5, 3.5, 3.6, 3.7, 3.8],
    "Dimas Pratama": [3.6, 3.5, 4.0, 4.0, 3.8, 4.1, 4.2, 4.3],
    "Sari Lestari":  [3.8, 3.6, 4.2, 4.1, 4.0, 4.3, 4.4, 4.6],
    "Maya Putri":    [3.3, 3.2, 3.8, 3.7, 3.5, 3.8, 3.9, 4.0],
  },
  raters: {
    "Rina Wijaya":   { self: 4.4, atasan: 3.6, rekan: 3.5, bawahan: 3.4 },
    "Andre Kusuma":  { self: 4.7, atasan: 4.0, rekan: 3.9, bawahan: 3.8 },
    "Budi Hartono":  { self: 4.1, atasan: 2.8, rekan: 2.9, bawahan: 3.0 },
    "Eko Santoso":   { self: 3.5, atasan: 3.0, rekan: 2.9, bawahan: 2.9 },
    "Nadia Rahma":   { self: 3.4, atasan: 3.5, rekan: 3.3, bawahan: 3.2 },
    "Dimas Pratama": { self: 4.3, atasan: 3.9, rekan: 3.8, bawahan: 3.7 },
    "Sari Lestari":  { self: 4.1, atasan: 4.1, rekan: 4.0, bawahan: 3.9 },
    "Maya Putri":    { self: 3.9, atasan: 3.7, rekan: 3.5, bawahan: 3.4 },
  },
  contributionAspects: ["Output & Kualitas", "Inisiatif Proyek", "Dampak pada Tim", "Kontrib. Lintas Tim"],
  contribution: {
    "Rina Wijaya":   [4.8, 4.6, 4.4, 4.3],
    "Andre Kusuma":  [4.8, 4.7, 4.5, 4.4],
    "Budi Hartono":  [3.2, 3.0, 2.9, 2.8],
    "Eko Santoso":   [3.5, 3.3, 3.1, 2.9],
    "Nadia Rahma":   [3.7, 3.5, 3.4, 3.4],
    "Dimas Pratama": [4.0, 3.9, 3.7, 3.6],
    "Sari Lestari":  [4.4, 4.3, 4.1, 4.0],
    "Maya Putri":    [4.1, 4.0, 3.8, 3.7],
  },
  star: {
    "Rina Wijaya":   { jumlah: 2, validasi: "Tervalidasi", ringkasan: "Konsistensi delivery di tengah tekanan deadline dan pengelolaan perubahan scope mendadak — dua kontribusi tervalidasi atasan" },
    "Andre Kusuma":  { jumlah: 2, validasi: "Tervalidasi", ringkasan: "Resolusi konflik prioritas lintas divisi dan koordinasi ulang delivery — bukti dampak nyata pada hasil tim dan kesinambungan kontrak" },
    "Budi Hartono":  { jumlah: 1, validasi: "Menunggu",    ringkasan: "Migrasi sistem lama tanpa henti operasi — downtime <15 menit dari target SLA. Menunggu validasi atasan untuk dilengkapi" },
    "Eko Santoso":   { jumlah: 1, validasi: "Menunggu",    ringkasan: "Penyesuaian scope mendadak di tengah sprint — delivery tepat waktu, klien memperbarui kontrak. Menunggu validasi atasan" },
    "Nadia Rahma":   { jumlah: 1, validasi: "Tervalidasi", ringkasan: "Proyek strategis dengan deadline sangat ketat — selesai H-2 dari tenggat dengan apresiasi langsung dari Direktur" },
    "Dimas Pratama": { jumlah: 2, validasi: "Tervalidasi", ringkasan: "Fasilitasi kesepakatan SLA antar dua divisi dan koordinasi perubahan scope secara proaktif — dampak dirasakan lintas tim" },
    "Sari Lestari":  { jumlah: 2, validasi: "Tervalidasi", ringkasan: "Migrasi sistem tanpa downtime berarti dan pengelolaan delivery di tengah perubahan scope — konsistensi kontribusi tertinggi di tim" },
    "Maya Putri":    { jumlah: 1, validasi: "Tervalidasi", ringkasan: "Penyesuaian scope mendadak dari klien — delivery tetap tepat waktu dan klien memperpanjang kontrak untuk kuartal berikutnya" },
  },
  // upward feedback: penilaian anggota terhadap manager — dist = jumlah jawaban skala 1..5
  upward: [
    { item: "My manager assigns stretch opportunities to help me develop new skills.", dist: [0, 0, 2, 3, 3] },
    { item: "My manager communicates clear goals for our team.", dist: [0, 0, 1, 2, 5] },
    { item: "My manager gives me actionable feedback on a regular basis.", dist: [0, 1, 2, 3, 2] },
    { item: "My manager consistently shows consideration for me as a person.", dist: [0, 0, 0, 2, 6] },
    { item: "My manager keeps the team focused on our priority results.", dist: [0, 0, 1, 4, 3] },
    { item: "My manager regularly shares relevant information from their manager.", dist: [0, 0, 3, 3, 2] },
    { item: "My manager has had a meaningful discussion with me about career development.", dist: [0, 0, 2, 4, 2] },
    { item: "My manager has the technical expertise required to effectively manage me.", dist: [0, 0, 0, 1, 7] },
    { item: "The actions of my manager show they value the perspective I bring to the team.", dist: [0, 0, 1, 3, 4] },
    { item: "My manager makes tough decisions effectively.", dist: [0, 1, 3, 2, 2] },
    { item: "My manager effectively collaborates across boundaries.", dist: [0, 0, 1, 4, 3] },
    { item: "I would recommend my manager to others.", dist: [0, 0, 0, 3, 5] },
  ],
  // uraian upward feedback (open feedback anggota ke manager, anonim)
  upwardNotes: [
    { title: "APA YANG SEBAIKNYA DILANJUTKAN OLEH MANAGER KAMU?", points: [
      "Tetap terbuka untuk berdiskusi dan mendengarkan masukan dari tim.",
      "Memberikan kepercayaan kepada tim untuk menjalankan tugas.",
    ]},
    { title: "APA YANG SEBAIKNYA DIUBAH OLEH MANAGER KAMU?", points: [
      "Menyampaikan perubahan prioritas atau keputusan strategis lebih awal.",
    ]},
    { title: "HAL LAIN YANG PERLU DIKETAHUI MANAGER TENTANG KONDISI KAMU DAN TIM.", points: [
      "Komunikasi dalam tim berjalan baik, koordinasi lintas fungsi masih bisa ditingkatkan.",
      "Pembagian task selalu diupayakan agar seimbang.",
    ]},
    { title: "DUKUNGAN YANG KAMU HARAPKAN DARI MANAGER.", points: [
      "Kejelasan informasi terkait perubahan kebijakan yang berdampak pada tim.",
    ]},
  ],
  // riwayat batch SEBELUMNYA (batch berjalan diambil otomatis dari data di atas)
  tren: {
    batches: ["Batch Q1 2026", "Batch Q2 2026"],
    members: {
      "Rina Wijaya":   { commitment: [3.9, 3.8], competency: [3.6, 3.7], contribution: [4.3, 4.4] },
      "Andre Kusuma":  { commitment: [3.8, 3.7], competency: [4.0, 4.1], contribution: [4.4, 4.5] },
      "Budi Hartono":  { commitment: [3.3, 3.2], competency: [3.3, 3.2], contribution: [3.1, 3.0] },
      "Eko Santoso":   { commitment: [3.6, 3.6], competency: [3.2, 3.1], contribution: [3.3, 3.2] },
      "Nadia Rahma":   { commitment: [3.6, 3.8], competency: [3.2, 3.3], contribution: [3.3, 3.4] },
      "Dimas Pratama": { commitment: [4.0, 4.1], competency: [3.8, 3.8], contribution: [3.7, 3.8] },
      "Sari Lestari":  { commitment: [3.9, 4.0], competency: [3.9, 4.0], contribution: [4.0, 4.1] },
      "Maya Putri":    { commitment: [4.2, 4.3], competency: [3.5, 3.6], contribution: [3.8, 3.8] },
    },
  },
  insights: {
    1: [
      "Commitment adalah satu-satunya dimensi di bawah standar (3.7 vs 3.8), dengan 4 dari 8 anggota di bawah standar. Competency (3.7) dan Contribution (3.8) berada di atas standar.",
      "Rina W. dan Andre K. berkontribusi tinggi namun komitmennya di bawah standar. Ini kelompok dengan risiko retensi tertinggi dan layak dijadwalkan one-on-one lebih dulu.",
      "Budi H. dan Eko S. berada di bawah standar pada ketiga dimensi dan membutuhkan rencana pengembangan terstruktur.",
    ],
    2: [
      "Strategic Thinking dan Business Insight lemah di level tim (rata-rata di bawah standar; 5 dan 6 dari 8 anggota di bawah standar). Pertimbangkan program tim, bukan coaching individual satu per satu.",
    ],
    3: [
      "3 anggota di bawah standar — Budi (3.1), Andre (3.3), Rina (3.4) — prioritaskan one-on-one untuk mengidentifikasi hambatan spesifik masing-masing.",
      "Pola rendah Budi konsisten di semua 5 item; Andre menunjukkan gap antara Inisiatif vs Konsistensi Performa — rencana aksi keduanya perlu dibedakan.",
      "Maya (4.4), Dimas (4.2), Sari (4.0) dapat dijadikan role model commitment — pertimbangkan sebagai buddy mentor untuk anggota yang sedang berkembang.",
    ],
    4: [
      "Business Insight (rata-rata 3.26) dan Strategic Thinking (3.38) adalah gap tim secara kolektif — butuh program pelatihan terstruktur, bukan hanya intervensi individual.",
      "Budi memiliki blind spot signifikan: Self 4.1 vs Atasan 2.8 — fasilitasi sesi kalibrasi persepsi sebelum rencana pengembangannya disusun.",
      "Communication (4.06) dan Change Leadership (3.96) adalah kekuatan tim — pertimbangkan sebagai fondasi untuk peran mentoring dan proyek lintas divisi.",
    ],
    5: [
      "Rina (4.5) dan Andre (4.6) masing-masing memiliki 2 bukti STAR tervalidasi — keduanya kandidat kuat untuk peran project lead atau mentor kontribusi lintas tim.",
      "Budi dan Eko memiliki bukti STAR yang belum tervalidasi — atasan perlu segera melakukan review agar penilaian kontribusi dapat diselesaikan secara adil.",
    ],
  },
};

// ═══════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════
const HEAT_STD = 3.5; // palet heatmap halaman 3–5 (konsisten dgn web app)

function heat5(v) {
  if (v >= 4.0)  return { bg: "#D1FAE5", fg: "#065F46" };
  if (v >= 3.75) return { bg: "#DBEAFE", fg: "#1E40AF" };
  if (v >= 3.5)  return { bg: "#EDE9FE", fg: "#5B21B6" };
  if (v >= 3.25) return { bg: "#FEF3C7", fg: "#92400E" };
  return { bg: "#FEE2E2", fg: "#991B1B" };
}
// palet halaman 2: relatif terhadap standar posisi
function heat2(v, std) {
  const d = v - std;
  if (d <= -0.4) return { bg: "#FECACA", fg: "#7F1D1D" };
  if (d < 0)     return { bg: "#FEF3C7", fg: "#78350F" };
  if (d < 0.5)   return { bg: "#DBEAFE", fg: "#1E3A8A" };
  return { bg: "#93C5FD", fg: "#1E3A8A" };
}

const STATUS = {
  risk:  { label: "Risiko retensi",  fokus: "One-on-one lebih dulu",   fg: "#DC2626", bg: "#FEE2E2" },
  watch: { label: "Perlu perhatian", fokus: "Rencana pengembangan",    fg: "#4B5563", bg: "#E5E7EB" },
  push:  { label: "Perlu didorong",  fokus: "Dorong inisiatif",        fg: "#D97706", bg: "#FEF3C7" },
  keep:  { label: "Pertahankan",     fokus: "Beri tantangan baru",     fg: "#15803D", bg: "#D1FAE5" },
};
const STATUS_ORDER = ["risk", "watch", "push", "keep"];

function statusOf(m, info) {
  const cOK = m.commitment >= info.stdCommitment;
  const bOK = m.contribution >= info.stdContribution;
  if (!cOK && bOK) return "risk";
  if (!cOK && !bOK) return "watch";
  if (cOK && !bOK) return "push";
  return "keep";
}

// ── TREN ──
// Seri nilai lintas batch (riwayat + batch berjalan). null = data tren tidak tersedia.
function trendSeries(data, key, nama) {
  const t = data.tren;
  if (!t || !t.batches || !t.batches.length) return null;
  if (nama) {
    const h = t.members[nama];
    if (!h || !h[key]) return null;
    const hist = t.batches.map((_, i) => h[key][i]);
    if (hist.some(v => v == null || isNaN(v))) return null;
    const cur = data.members.find(m => m.nama === nama);
    return [...hist, cur[key]];
  }
  const hist = t.batches.map((_, i) => {
    const vals = data.members.map(m => t.members[m.nama] && t.members[m.nama][key] ? t.members[m.nama][key][i] : null)
      .filter(v => v != null && !isNaN(v));
    return vals.length ? avg(vals) : null;
  });
  if (hist.some(v => v == null)) return null;
  return [...hist, avg(data.members.map(m => m[key]))];
}

// seri rata-rata 3C per anggota (untuk kolom tren Hal 2)
function trendSeries3C(data, nama) {
  const s = ["commitment", "competency", "contribution"].map(k => trendSeries(data, k, nama));
  if (s.some(x => !x)) return null;
  return s[0].map((_, i) => avg(s.map(x => x[i])));
}

const TREND_UP = "#16A34A", TREND_DOWN = "#DC2626", TREND_FLAT = "#6B7280";
function trendDelta(series) {
  const d = series[series.length - 1] - series[series.length - 2];
  if (d >= 0.05)  return { d, sym: "▲", color: TREND_UP,   txt: "+" + d.toFixed(1) };
  if (d <= -0.05) return { d, sym: "▼", color: TREND_DOWN, txt: "−" + Math.abs(d).toFixed(1) };
  return { d, sym: "▬", color: TREND_FLAT, txt: "±0.0" };
}

// sparkline berlabel: titik tiap batch + angka nilai di atasnya;
// batchLabels (opsional) menambah label batch di bawah titik
function sparkline(series, w = 56, h = 24, batchLabels) {
  const dd = trendDelta(series);
  let lo = Math.min(...series), hi = Math.max(...series);
  if (hi - lo < 0.4) { const mid = (hi + lo) / 2; lo = mid - 0.2; hi = mid + 0.2; }
  const TOP = 8, BOT = batchLabels ? 9 : 2.5;
  const X = (i) => 9 + (i / (series.length - 1)) * (w - 18);
  const Y = (v) => TOP + 1.5 + (1 - (v - lo) / (hi - lo)) * (h - TOP - BOT - 3);
  const pts = series.map((v, i) => `${X(i).toFixed(1)},${Y(v).toFixed(1)}`).join(" ");
  const li = series.length - 1;
  const dots = series.map((v, i) =>
    `<circle cx="${X(i).toFixed(1)}" cy="${Y(v).toFixed(1)}" r="${i === li ? 2 : 1.4}" fill="${i === li ? dd.color : "#9CA3AF"}"/>`).join("");
  const valLabels = series.map((v, i) =>
    `<text x="${X(i).toFixed(1)}" y="${(Y(v) - 3.5).toFixed(1)}" text-anchor="middle" font-size="6" font-weight="${i === li ? 700 : 600}" fill="${i === li ? dd.color : "#9CA3AF"}">${f1(v)}</text>`).join("");
  const bLabels = batchLabels ? batchLabels.map((b, i) =>
    `<text x="${X(i).toFixed(1)}" y="${h - 1.5}" text-anchor="middle" font-size="5.5" fill="#9CA3AF">${esc(b)}</text>`).join("") : "";
  return `<svg class="spark" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}"><polyline points="${pts}" fill="none" stroke="${dd.color}" stroke-width="1.2" stroke-linejoin="round" stroke-linecap="round" opacity="0.85"/>${dots}${valLabels}${bLabels}</svg>`;
}

// label batch ringkas: buang kata umum ("Batch Q1 2026" / "ORBID Q1" → "Q1 2026" / "Q1")
const shortBatch = (b) => String(b).replace(/^((batch|orbid)\s+)+/i, "");
function batchLabels(data) {
  if (!data.tren) return null;
  return [...data.tren.batches, data.info.batch].map(shortBatch);
}

const prevBatchShort = (data) =>
  data.tren && data.tren.batches.length
    ? shortBatch(data.tren.batches[data.tren.batches.length - 1])
    : "";

const AV_FG = ["#5B21B6", "#1E40AF", "#065F46", "#991B1B", "#92400E", "#1D4ED8", "#047857", "#7C3AED"];
const initials = (name) => name.split(" ").slice(0, 2).map(w => w[0]).join("").toUpperCase();
const firstName = (name) => {
  const p = name.split(" ");
  return p[0].length <= 2 && p[1] ? p[0] + " " + p[1] : p[0]; // "H. Naufalinda" → "H. Naufalinda"
};
const avg = (a) => a.reduce((x, y) => x + y, 0) / a.length;
const f1 = (v) => v.toFixed(1);
const f2 = (v) => v.toFixed(2);
const esc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));

// ═══════════════════════════════════════════════════════════════
// PAGE SHELL
// ═══════════════════════════════════════════════════════════════
function pageShell(num, title, data, bodyHtml) {
  const ribbon = data.info.label ? `<div class="ribbon">${esc(data.info.label)}</div>` : "";
  return `
  <div class="page">
    <div class="corner"></div>
    <div class="brand-title">Laporan ORBID</div>
    <div class="title-row"><img src="assets/title-accent.svg" alt=""><h1>${esc(title)}</h1></div>
    ${ribbon}
    <div class="body">${bodyHtml}</div>
    <img class="hex-pattern" src="assets/pattern-hex-bottom.svg" alt="">
    <div class="footer">
      <div class="provided">
        <span class="lbl">Provided By</span>
        <div class="row"><img src="assets/logo-talentlytica.svg" alt="talentlytica"><span class="divider"></span><span class="name">${esc(data.info.penerima)}</span></div>
      </div>
      <div class="pageno">${num}</div>
    </div>
  </div>`;
}

function insightBox(theme, title, bullets) {
  if (!bullets || !bullets.length) return "";
  return `
  <div class="insight-box ${theme}">
    <div class="accent"></div>
    <div class="ititle">${esc(title)}</div>
    <ul>${bullets.map(b => `<li>${esc(b)}</li>`).join("")}</ul>
  </div>`;
}

const heading = (text) => `<div class="heading"><div class="bar"></div><span>${esc(text)}</span></div>`;

// ═══════════════════════════════════════════════════════════════
// HALAMAN 1 — RINGKASAN TIM: PROFIL 3C
// ═══════════════════════════════════════════════════════════════
function renderPage1(data) {
  const { info, members } = data;
  const n = members.length;
  const dims = [
    { key: "commitment",   label: "COMMITMENT",   dot: "#16A34A", std: info.stdCommitment },
    { key: "competency",   label: "COMPETENCY",   dot: "#2563EB", std: info.stdCompetency },
    { key: "contribution", label: "CONTRIBUTION", dot: "#F59E0B", std: info.stdContribution },
  ];

  const prevShort = prevBatchShort(data);
  const cards = dims.map(d => {
    const vals = members.map(m => m[d.key]);
    const teamAvg = avg(vals);
    const ok = vals.filter(v => v >= d.std).length;
    const below = teamAvg < d.std;
    const badge = below
      ? `<span class="tag" style="background:#FEE2E2;color:#991B1B">Di bawah standar</span>`
      : `<span class="tag" style="background:#D1FAE5;color:#065F46">Di atas standar</span>`;
    const series = trendSeries(data, d.key);
    const trendRow = series ? (() => {
      const dd = trendDelta(series);
      return `<div class="trend-row">${sparkline(series, 112, 36, batchLabels(data))}<span class="tchip" style="color:${dd.color}">${dd.sym} ${dd.txt}<small>vs ${esc(prevShort)}</small></span></div>`;
    })() : "";
    return `
    <div class="card3c">
      <div class="top">
        <div class="dim"><span class="dot" style="background:${d.dot}"></span>${d.label}</div>
        ${badge}
      </div>
      <div class="big">${f1(teamAvg)}<small>rata-rata tim</small></div>
      <div class="track"><div class="seg-ok" style="width:${(ok / n) * 100}%"></div><div class="seg-no" style="width:${((n - ok) / n) * 100}%"></div></div>
      <div class="foot"><span class="ok">${ok} memenuhi standar</span><span class="no">${n - ok} di bawah</span></div>
      ${trendRow}
    </div>`;
  }).join("");

  // kuadran: x = competency, y = contribution, bubble size = commitment
  const CW = 394, CH = 228, PL = 30, PB = 22, PT = 8, PR = 8;
  const IW = CW - PL - PR, IH = CH - PT - PB;
  const MIN = 2.5, MAX = 5.0, R = MAX - MIN;
  const tx = (v) => PL + ((Math.min(MAX, Math.max(MIN, v)) - MIN) / R) * IW;
  const ty = (v) => PT + IH - ((Math.min(MAX, Math.max(MIN, v)) - MIN) / R) * IH;
  // X divider = stdCompetency, Y divider = stdContribution
  const sx = tx(info.stdCompetency), sy = ty(info.stdContribution);

  // warna bubble = commitment vs standar (hijau = memenuhi, merah = di bawah)
  const COMMIT_HI = "#16A34A", COMMIT_LO = "#DC2626";
  const BR = 11;
  const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

  // 1) koordinat identik dilebur jadi satu bubble multi-anggota
  const groupMap = {};
  members.forEach(m => {
    const key = m.competency + "|" + m.contribution;
    (groupMap[key] = groupMap[key] || []).push(m);
  });
  const nodes = Object.values(groupMap).map(g => {
    const x = tx(g[0].competency), y = ty(g[0].contribution);
    const r = BR + (g.length > 1 ? 2 : 0);
    return { x, y, ox: x, oy: y, r, members: g };
  });

  // 2) dodge: bubble yang bertabrakan saling didorong menjauh sedikit
  for (let iter = 0; iter < 40; iter++) {
    let moved = false;
    for (let i = 0; i < nodes.length; i++) for (let j = i + 1; j < nodes.length; j++) {
      const a = nodes[i], b = nodes[j];
      let dx = b.x - a.x, dy = b.y - a.y, d = Math.hypot(dx, dy);
      const min = a.r + b.r + 1.5;
      if (d < min) {
        if (d < 0.01) { dx = 1; dy = -0.5; d = Math.hypot(dx, dy); }
        const push = (min - d) / 2;
        a.x -= (dx / d) * push; a.y -= (dy / d) * push;
        b.x += (dx / d) * push; b.y += (dy / d) * push;
        moved = true;
      }
    }
    if (!moved) break;
  }
  nodes.forEach(n => { n.x = clamp(n.x, PL + n.r, PL + IW - n.r); n.y = clamp(n.y, PT + n.r, PT + IH - n.r); });

  const bubbles = nodes.map(n => {
    const ok = n.members.filter(m => m.commitment >= info.stdCommitment).length;
    let fillAttr;
    let defs = "";
    if (ok === n.members.length) fillAttr = COMMIT_HI;
    else if (ok === 0) fillAttr = COMMIT_LO;
    else {
      const gid = "mix" + n.ox.toFixed(0) + n.oy.toFixed(0);
      defs = `<linearGradient id="${gid}" x1="0" y1="0" x2="1" y2="0"><stop offset="50%" stop-color="${COMMIT_HI}"/><stop offset="50%" stop-color="${COMMIT_LO}"/></linearGradient>`;
      fillAttr = `url(#${gid})`;
    }
    const label = n.members.length <= 2
      ? n.members.map(m => initials(m.nama)).join("·")
      : `${n.members.length} org`;
    const fsize = n.members.length > 1 ? 5.5 : 7;
    // garis penunjuk ke titik asli bila geseran dodge cukup jauh
    const moved = Math.hypot(n.x - n.ox, n.y - n.oy);
    const leader = moved > 5
      ? `<line x1="${n.ox.toFixed(1)}" y1="${n.oy.toFixed(1)}" x2="${n.x.toFixed(1)}" y2="${n.y.toFixed(1)}" stroke="#9CA3AF" stroke-width="0.8"/><circle cx="${n.ox.toFixed(1)}" cy="${n.oy.toFixed(1)}" r="1.5" fill="#9CA3AF"/>`
      : "";
    return `${defs}${leader}<circle cx="${n.x.toFixed(1)}" cy="${n.y.toFixed(1)}" r="${n.r}" fill="${fillAttr}" opacity="0.92" stroke="#fff" stroke-width="1"/>
            <text x="${n.x.toFixed(1)}" y="${(n.y + 2.5).toFixed(1)}" text-anchor="middle" font-size="${fsize}" font-weight="700" fill="#fff">${label}</text>`;
  }).join("");

  const legend = `
  <div class="map-legend">
    <span class="li"><svg width="12" height="12"><circle cx="6" cy="6" r="5" fill="${COMMIT_HI}"/></svg> Commitment &ge; standar (${f1(info.stdCommitment)})</span>
    <span class="li"><svg width="12" height="12"><circle cx="6" cy="6" r="5" fill="${COMMIT_LO}"/></svg> Commitment &lt; standar</span>
  </div>`;

  const ticks = [];
  for (let v = MIN; v <= MAX + 0.001; v += 0.5) {
    ticks.push(`<text x="${tx(v)}" y="${CH - 6}" text-anchor="middle" font-size="7.5" fill="#9CA3AF">${v.toFixed(1)}</text>`);
  }

  const quad = `
  <svg width="${CW}" height="${CH}" viewBox="0 0 ${CW} ${CH}" xmlns="http://www.w3.org/2000/svg">
    <rect x="${PL}" y="${PT}" width="${sx - PL}" height="${sy - PT}" fill="#FEF2F2"/>
    <rect x="${sx}" y="${PT}" width="${PL + IW - sx}" height="${sy - PT}" fill="#F0FDF4"/>
    <rect x="${PL}" y="${sy}" width="${sx - PL}" height="${PT + IH - sy}" fill="#F3F4F6"/>
    <rect x="${sx}" y="${sy}" width="${PL + IW - sx}" height="${PT + IH - sy}" fill="#FFFBEB"/>
    <line x1="${sx}" y1="${PT}" x2="${sx}" y2="${PT + IH}" stroke="#9CA3AF" stroke-width="1" stroke-dasharray="4,3"/>
    <line x1="${PL}" y1="${sy}" x2="${PL + IW}" y2="${sy}" stroke="#9CA3AF" stroke-width="1" stroke-dasharray="4,3"/>
    <text x="${PL + 6}" y="${PT + 13}" font-size="7.5" font-weight="700" fill="#DC2626">RISIKO RETENSI</text>
    <text x="${PL + IW - 6}" y="${PT + 13}" text-anchor="end" font-size="7.5" font-weight="700" fill="#15803D">PERTAHANKAN</text>
    <text x="${PL + 6}" y="${PT + IH - 7}" font-size="7.5" font-weight="700" fill="#6B7280">PERLU PERHATIAN</text>
    <text x="${PL + IW - 6}" y="${PT + IH - 7}" text-anchor="end" font-size="7.5" font-weight="700" fill="#D97706">PERLU DIDORONG</text>
    ${bubbles}
    ${ticks.join("")}
  </svg>`;

  const vchip = (v, std) => {
    const c = heat2(v, std);
    return `<span class="vchip" style="background:${c.bg};color:${c.fg}">${f1(v)}</span>`;
  };
  const byStatus = STATUS_ORDER.map(key => {
    const list = members.filter(m => statusOf(m, info) === key);
    if (!list.length) return "";
    const st = STATUS[key];
    return `
    <div class="group">
      <span class="tag" style="background:${st.bg};color:${st.fg}">${st.label}</span>
      ${list.map(m => `
      <div class="member">
        <span class="mavatar" style="color:${AV_FG[members.indexOf(m) % 8]}">${initials(m.nama)}</span>
        <span class="nm">${esc(m.nama)}</span>
        <span class="chips">${vchip(m.commitment, info.stdCommitment)}${vchip(m.competency, info.stdCompetency)}${vchip(m.contribution, info.stdContribution)}</span>
      </div>`).join("")}
    </div>`;
  }).join("");

  const lengkap = `Laporan lengkap ${n} dari ${n}`;
  const body = `
  <div class="team-strip">
    <div class="avatar">${esc(info.tim[0] || "T")}</div>
    <div class="kv"><span class="k">Tim</span><span class="v">${esc(info.tim)}</span></div>
    <div class="kv"><span class="k">Manager</span><span class="v">${esc(info.manager)}</span></div>
    <div class="kv"><span class="k">Project / Batch</span><span class="v">${esc(info.batch)}</span></div>
    <div class="spacer"></div>
    <div class="quality"><div class="k">CAKUPAN DATA</div><div class="v">${n} anggota tim</div><div class="ok">${lengkap}</div></div>
  </div>
  ${heading("Rata-rata tim dibandingkan standar posisi")}
  <div class="cards3c">${cards}</div>
  ${heading("Peta tim: Contribution vs Competency")}
  <div class="map-row">
    <div class="map-card">
      <div class="mlabel">&uarr; Contribution</div>
      ${quad}
      <div class="mcaption">Competency &rarr; &nbsp;&nbsp; (garis pemisah = standar: Competency ${f1(info.stdCompetency)}, Contribution ${f1(info.stdContribution)})</div>
      ${legend}
    </div>
    <div class="sorotan">
      <h3>SOROTAN ANGGOTA TIM</h3>
      <div class="chead"><span class="cl">COMMIT<br>MENT</span><span class="cl">COMPE<br>TENCY</span><span class="cl">CONTRI<br>BUTION</span></div>
      ${byStatus}
      <div class="note">Warna sel mengikuti posisi nilai terhadap standar</div>
    </div>
  </div>
  ${insightBox("blue", "INSIGHT UTAMA TIM", data.insights[1])}`;

  return pageShell(1, "RINGKASAN TIM — PROFIL 3C", data, body);
}

// ═══════════════════════════════════════════════════════════════
// HALAMAN 2 — KOMPETENSI DAN ANGGOTA
// ═══════════════════════════════════════════════════════════════
function renderPage2(data) {
  const { info, members } = data;
  const n = members.length;
  const sorted = [...members].sort((a, b) =>
    STATUS_ORDER.indexOf(statusOf(a, info)) - STATUS_ORDER.indexOf(statusOf(b, info)));

  const hasTrend = !!trendSeries(data, "commitment");
  const rows = sorted.map(m => {
    const st = STATUS[statusOf(m, info)];
    const cells = [
      [m.commitment, info.stdCommitment],
      [m.competency, info.stdCompetency],
      [m.contribution, info.stdContribution],
    ].map(([v, std]) => {
      const c = heat2(v, std);
      return `<td class="c"><span class="cell" style="background:${c.bg};color:${c.fg}">${f1(v)}</span></td>`;
    }).join("");
    const bsColor = m.blindspot >= 2 ? "#DC2626" : m.blindspot === 1 ? "#374151" : "#9CA3AF";
    let trendCell = "";
    if (hasTrend) {
      const s3 = trendSeries3C(data, m.nama);
      trendCell = s3
        ? (() => { const dd = trendDelta(s3); return `<td class="c trc">${sparkline(s3, 50, 19)}<span style="color:${dd.color}">${dd.sym} ${dd.txt}</span></td>`; })()
        : `<td class="c trc" style="color:#9CA3AF">&ndash;</td>`;
    }
    return `
    <tr>
      <td class="name">${esc(m.nama)}</td>
      ${cells}
      ${trendCell}
      <td class="bs" style="color:${bsColor}">${m.blindspot}</td>
      <td><span class="tag" style="background:${st.bg};color:${st.fg}">${st.label}</span><span class="fokus">${st.fokus}</span></td>
    </tr>`;
  }).join("");

  const compRows = data.competencyAspects.map((a, i) => {
    const vals = members.map(m => data.competency[m.nama][i]);
    const tAvg = avg(vals);
    const std = data.competencyStd[i];
    const below = vals.filter(v => v < std).length;
    const d = tAvg - std;
    const isGap = d < 0;
    const isStrength = d >= 0.4;
    const barColor = isGap ? "#F08A6E" : "#2563EB";
    const barW = ((Math.min(5, tAvg) - 1) / 4) * 100;
    const stdPos = ((std - 1) / 4) * 100;
    const statusTag = isGap
      ? `<span class="tag" style="background:#FEE2E2;color:#991B1B">Gap tim (&minus;${Math.abs(d) < 0.05 ? Math.abs(d).toFixed(2) : Math.abs(d).toFixed(1)})</span>`
      : isStrength
        ? `<span class="tag" style="background:#D1FAE5;color:#065F46">Kekuatan tim (+${d.toFixed(1)})</span>`
        : `<span class="tag" style="background:#F3F4F6;color:#4B5563">Memenuhi (+${d.toFixed(1)})</span>`;
    return `
    <tr>
      <td class="name" style="font-size:10.5px">${esc(a)}</td>
      <td><div class="bar-wrap"><div class="bar-fill" style="width:${barW}%;background:${barColor}"></div><div class="bar-std" style="left:${stdPos}%"></div></div></td>
      <td class="c tval" style="color:${isGap ? "#DC2626" : "#111827"}">${f1(tAvg)}</td>
      <td class="c std">${f1(std)}</td>
      <td class="c below" style="color:${isGap ? "#DC2626" : "#6B7280"};${isGap ? "font-weight:700" : ""}">${below} dari ${n}</td>
      <td class="c">${statusTag}</td>
    </tr>`;
  }).join("");

  const body = `
  ${heading("Seluruh anggota tim (diurutkan dari yang paling perlu perhatian)")}
  <div>
    <table class="tbl">
      <thead><tr>
        <th>ANGGOTA</th><th class="c">COMMITMENT</th><th class="c">COMPETENCY</th><th class="c">CONTRIBUTION</th>${hasTrend ? '<th class="c">TREN 3C</th>' : ""}<th class="c">BLINDSPOT</th><th>KATEGORI DAN FOKUS</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <div class="tbl-note">Warna sel: merah muda = jauh di bawah standar, kuning = di bawah standar, biru muda = memenuhi, biru tua = &ge; 0.5 di atas standar.${hasTrend ? " Tren 3C = perubahan rata-rata 3 dimensi dibanding batch sebelumnya." : ""} Blindspot = jumlah aspek Competency di mana Self menilai &ge; 0.8 lebih tinggi dari orang lain.</div>
  </div>
  ${heading("Kompetensi tim dibandingkan standar")}
  <table class="tbl comp">
    <thead><tr>
      <th>ASPEK</th><th>RATA-RATA TIM (1&ndash;5)</th><th class="c">TIM</th><th class="c">STANDAR</th><th class="c">DI BAWAH</th><th class="c">STATUS</th>
    </tr></thead>
    <tbody>${compRows}</tbody>
  </table>
  ${insightBox("blue", "IMPLIKASI UNTUK TIM", data.insights[2])}`;

  return pageShell(2, "RINGKASAN TIM — KOMPETENSI DAN ANGGOTA", data, body);
}

// ═══════════════════════════════════════════════════════════════
// HALAMAN 3 — DETAIL KLUSTER COMMITMENT
// ═══════════════════════════════════════════════════════════════
function renderPage3(data) {
  const { members } = data;
  const items = data.commitmentItems;
  const teamAvg = avg(members.map(m => m.commitment));

  const hasTrend = !!trendSeries(data, "commitment");

  // kode singkat per item: pakai prefix "Q01." bila ada, selain itu I1..In
  const itemCode = (it, i) => {
    const m = String(it).match(/^(Q?\d+)\./i);
    return m ? m[1].toUpperCase() : "I" + (i + 1);
  };
  const itemText = (it) => String(it).replace(/^(Q?\d+)\.\s*/i, "");

  const trendTd = (nama) => {
    if (!hasTrend) return "";
    const s = trendSeries(data, "commitment", nama);
    return `<td class="spark-td">${s ? sparkline(s, 54, 24) : '<span style="color:#9CA3AF">&ndash;</span>'}</td>`;
  };
  const rows = members.map(m => {
    const vals = data.commitment[m.nama] || [];
    const agg = heat5(m.commitment);
    return `
    <tr>
      <td class="name">${esc(m.nama)}</td>
      <td><span class="hcell avgcell" style="background:${agg.bg};color:${agg.fg}">${f1(m.commitment)}</span></td>
      ${vals.map(v => { const c = heat5(v); return `<td><span class="hcell" style="background:${c.bg};color:${c.fg}">${f1(v)}</span></td>`; }).join("")}
      ${trendTd(m.nama)}
    </tr>`;
  }).join("");

  const avgCols = items.map((_, i) => {
    const a = avg(members.map(m => (data.commitment[m.nama] || [])[i] || 0));
    const c = heat5(a);
    return `<td><span class="hcell" style="background:${c.bg};color:${c.fg}">${f2(a)}</span></td>`;
  }).join("");
  const tc = heat5(teamAvg);

  // keterangan item (dua kolom)
  const defs = `
  <div class="item-defs">
    ${items.map((it, i) => `<div class="def"><b>${esc(itemCode(it, i))}</b><span>${esc(itemText(it))}</span></div>`).join("")}
  </div>`;

  const legend = `
  <div class="legend">
    ${[["&ge; 4.0 — Sangat Baik", "#D1FAE5", "#065F46"], ["3.75–3.99", "#DBEAFE", "#1E40AF"], ["3.5–3.74 — Standar", "#EDE9FE", "#5B21B6"], ["3.25–3.49", "#FEF3C7", "#92400E"], ["&lt; 3.25 — Perlu Intervensi", "#FEE2E2", "#991B1B"]]
      .map(([l, bg, fg]) => `<span class="it"><span class="sw" style="background:${bg};border-color:${fg}"></span>${l}</span>`).join("")}
  </div>`;

  const body = `
  <div class="cluster-strip" style="background:#F0FDF4;border-color:#86EFAC">
    <div class="accent" style="background:#16A34A"></div>
    <h2 style="color:#15803D">Commitment</h2>
    <p>Mengukur komitmen dan keterikatan anggota terhadap pekerjaan dan tim melalui ${items.length} item.</p>
    <div class="meta" style="color:#15803D">Rata-rata tim: ${f2(teamAvg)} &nbsp;&middot;&nbsp; Standar minimum: ${f1(HEAT_STD)}</div>
  </div>
  ${heading("Skor per item — seluruh anggota tim")}
  <div>
    <table class="heat">
      <thead><tr>
        <th class="l">Anggota</th><th style="width:52px">Agregat</th>
        ${items.map((it, i) => `<th>${esc(itemCode(it, i))}</th>`).join("")}
        ${hasTrend ? '<th style="width:58px">Tren</th>' : ""}
      </tr></thead>
      <tbody>
        ${rows}
        <tr class="avg">
          <td class="name">Rata-rata Tim</td>
          <td><span class="hcell avgcell" style="background:${tc.bg};color:${tc.fg}">${f2(teamAvg)}</span></td>
          ${avgCols}
          ${hasTrend ? `<td class="spark-td">${sparkline(trendSeries(data, "commitment"), 54, 24)}</td>` : ""}
        </tr>
      </tbody>
    </table>
    ${legend}
    ${hasTrend ? `<div class="tbl-note">Tren = nilai Commitment lintas batch: ${batchLabels(data).map(esc).join(" &rarr; ")} (batch berjalan).</div>` : ""}
  </div>
  ${heading("Keterangan item")}
  <div>
    ${defs}
  </div>
  ${insightBox("green", "CATATAN INTERPRETATIF", data.insights[3])}`;

  return pageShell(3, "DETAIL KLUSTER — COMMITMENT", data, body);
}

// ═══════════════════════════════════════════════════════════════
// HALAMAN 4 — DETAIL KLUSTER COMPETENCY
// ═══════════════════════════════════════════════════════════════
function renderPage4(data) {
  const { members } = data;
  const aspects = data.competencyAspects;
  const teamAvg = avg(members.map(m => avg(data.competency[m.nama])));

  const hasTrend = !!trendSeries(data, "competency");
  const trendTd = (nama) => {
    if (!hasTrend) return "";
    const s = trendSeries(data, "competency", nama);
    return `<td class="spark-td">${s ? sparkline(s, 54, 24) : '<span style="color:#9CA3AF">&ndash;</span>'}</td>`;
  };
  const rows = members.map(m => {
    const vals = data.competency[m.nama];
    const mAvg = avg(vals);
    const ac = heat5(mAvg);
    const bs = m.blindspot > 0 ? `<span class="bs-chip">BS${m.blindspot}</span>` : "";
    return `
    <tr>
      <td class="name">${esc(m.nama)}</td>
      <td><span class="hcell avgcell" style="background:${ac.bg};color:${ac.fg}">${f2(mAvg)}${bs}</span></td>
      ${vals.map(v => { const c = heat5(v); return `<td><span class="hcell" style="background:${c.bg};color:${c.fg}">${f2(v)}</span></td>`; }).join("")}
      ${trendTd(m.nama)}
    </tr>`;
  }).join("");

  const avgCols = aspects.map((_, i) => {
    const a = avg(members.map(m => data.competency[m.nama][i]));
    const c = heat5(a);
    return `<td><span class="hcell" style="background:${c.bg};color:${c.fg}">${f2(a)}</span></td>`;
  }).join("");
  const tc = heat5(teamAvg);

  const RCOLORS = { self: "#6366F1", atasan: "#2563EB", rekan: "#059669", bawahan: "#D97706" };
  const raterCards = members.map(m => {
    const r = data.raters[m.nama] || {};
    const rows2 = ["self", "atasan", "rekan", "bawahan"].map(k => {
      const v = r[k];
      if (v == null || isNaN(v) || v === 0) return ""; // penilai tidak tersedia (mis. tanpa bawahan)
      const w = Math.max(4, ((v - 2) / 3) * 100);
      const stdPos = ((HEAT_STD - 2) / 3) * 100;
      const vc = v >= HEAT_STD ? "#1D4ED8" : "#DC2626";
      return `
      <div class="rrow">
        <span class="rlbl">${k[0].toUpperCase() + k.slice(1)}</span>
        <span class="rbar"><span class="rfill" style="width:${w}%;background:${RCOLORS[k]}"></span><span class="rstd" style="left:${stdPos}%"></span></span>
        <span class="rval" style="color:${vc}">${f1(v)}</span>
      </div>`;
    }).join("");
    return `<div class="rater-card"><div class="rname">${esc(firstName(m.nama))}</div>${rows2}</div>`;
  }).join("");

  const body = `
  <div class="cluster-strip" style="background:#F0FDF4;border-color:#86EFAC">
    <div class="accent" style="background:#16A34A"></div>
    <h2 style="color:#15803D">Competency</h2>
    <p>Mengukur ${aspects.length} sub-kompetensi melalui penilaian 360&deg; (self, atasan, rekan, bawahan). Ditampilkan sebagai non-self average.</p>
    <div class="meta" style="color:#15803D">Rata-rata tim: ${f2(teamAvg)} &nbsp;&middot;&nbsp; Standar minimum: ${f1(HEAT_STD)} &nbsp;&middot;&nbsp; BS = jumlah item blindspot (self &ge; 0.8 lebih tinggi dari non-self)</div>
  </div>
  ${heading("Skor per sub-kompetensi (non-self avg) — seluruh anggota tim")}
  <table class="heat">
    <thead><tr>
      <th class="l">Anggota</th><th style="width:72px">Avg / BS</th>
      ${aspects.map(a => `<th>${esc(a).replace(/ (?=\S+$)/, "<br>")}</th>`).join("")}
      ${hasTrend ? '<th style="width:60px">Tren</th>' : ""}
    </tr></thead>
    <tbody>
      ${rows}
      <tr class="avg">
        <td class="name">Rata-rata Tim</td>
        <td><span class="hcell avgcell" style="background:${tc.bg};color:${tc.fg}">${f2(teamAvg)}</span></td>
        ${avgCols}
        ${hasTrend ? `<td class="spark-td">${sparkline(trendSeries(data, "competency"), 54, 24)}</td>` : ""}
      </tr>
    </tbody>
  </table>
  ${hasTrend ? `<div class="tbl-note" style="margin-top:-10px">Tren = nilai Competency (non-self avg) lintas batch: ${batchLabels(data).map(esc).join(" &rarr; ")} (batch berjalan).</div>` : ""}
  ${heading("Perbandingan perspektif penilai (avg semua sub-kompetensi)")}
  <div class="rater-grid">${raterCards}</div>
  ${insightBox("green", "CATATAN INTERPRETATIF", data.insights[4])}`;

  return pageShell(4, "DETAIL KLUSTER — COMPETENCY", data, body);
}

// ═══════════════════════════════════════════════════════════════
// HALAMAN 5 — DETAIL KLUSTER CONTRIBUTION
// ═══════════════════════════════════════════════════════════════
function renderPage5(data) {
  const { members } = data;
  const aspects = data.contributionAspects;
  const teamAvg = avg(members.map(m => m.contribution));

  const hasTrend = !!trendSeries(data, "contribution");
  // lebar sel aspek dinamis: 734 = 182 (nama) + n×(AW+4) + [48 tren] + 52 (avg)
  const AW = Math.floor((734 - 182 - 52 - (hasTrend ? 48 : 0)) / aspects.length) - 4;
  const labels = `
  <div class="collabels">
    <span class="sp" style="width:182px"></span>
    ${aspects.map(a => `<span class="cl" style="width:${AW}px;margin-right:4px">${esc(a).replace(/ (?=\S+$)/, "<br>")}</span>`).join("")}
    ${hasTrend ? '<span class="cl" style="width:44px;margin-right:4px">Tren</span>' : ""}
    <span class="cl avg" style="width:52px">Avg<br>Kontrib.</span>
  </div>
  <div class="hdr-div"></div>`;

  const blocks = members.map((m, mi) => {
    const vals = data.contribution[m.nama] || [];
    const s = data.star[m.nama] || { jumlah: 0, validasi: "-", ringkasan: "" };
    const ac = heat5(m.contribution);
    const valOk = s.validasi === "Tervalidasi";
    const chipBg = valOk ? "#D1FAE5" : "#FEF3C7";
    const chipFg = valOk ? "#065F46" : "#92400E";
    const countChip = s.jumlah > 1 ? `<span class="chip count">${s.jumlah} bukti</span>` : "";
    let sparkSlot = "";
    if (hasTrend) {
      const ser = trendSeries(data, "contribution", m.nama);
      sparkSlot = `<span class="spark-slot" style="width:44px;margin-right:4px">${ser ? sparkline(ser, 44, 26) : ""}</span>`;
    }
    return `
    <div class="mblock">
      <div class="mrow ${mi % 2 === 1 ? "alt" : ""}">
        <span class="avatar" style="color:${AV_FG[mi % 8]}">${initials(m.nama)}</span>
        <span class="mname">${esc(m.nama)}</span>
        ${vals.map(v => { const c = heat5(v); return `<span class="acell" style="width:${AW}px;background:${c.bg};color:${c.fg}">${f1(v)}</span>`; }).join("")}
        ${sparkSlot}
        <span class="acell avgcell" style="background:${ac.bg};color:${ac.fg}">${f1(m.contribution)}</span>
      </div>
      <div class="star-strip">
        <span class="chip" style="background:${chipBg};color:${chipFg}">${esc(s.validasi)}</span>
        ${countChip}
        <span class="sum">${esc(s.ringkasan)}</span>
      </div>
    </div>
    ${mi < members.length - 1 ? '<div class="sep"></div>' : ""}`;
  }).join("");

  const body = `
  <div class="cluster-strip" style="background:#FFFBEB;border-color:#FCD34D">
    <div class="accent" style="background:#D97706"></div>
    <h2 style="color:#B45309">Contribution</h2>
    <p>Mengukur dampak nyata terhadap hasil tim melalui ${aspects.length} aspek kontribusi, diperkuat bukti STAR (Situasi, Tugas, Aksi, Hasil) yang divalidasi atasan.</p>
    <div class="meta" style="color:#B45309">Rata-rata tim: ${f2(teamAvg)} &nbsp;&middot;&nbsp; Standar minimum: ${f1(HEAT_STD)}</div>
  </div>
  ${heading("Nilai per aspek kontribusi & ringkasan bukti STAR — seluruh anggota tim")}
  <div class="contrib-list">
    ${labels}
    ${blocks}
    ${hasTrend ? `<div class="tbl-note" style="margin-top:4px">Tren = nilai Contribution lintas batch: ${batchLabels(data).map(esc).join(" &rarr; ")} (batch berjalan).</div>` : ""}
  </div>
  ${insightBox("amber", "CATATAN INTERPRETATIF", data.insights[5])}`;

  return pageShell(5, "DETAIL KLUSTER — CONTRIBUTION", data, body);
}

// ═══════════════════════════════════════════════════════════════
// HALAMAN 6 — UPWARD FEEDBACK (penilaian anggota terhadap manager)
// ═══════════════════════════════════════════════════════════════
const UPWARD_COLORS = ["#F3F4F6", "#E5E7EB", "#93C5FD", "#1E40AF", "#2563EB"];
const UPWARD_LABELS = ["1 - Sangat tidak setuju", "2 - Tidak setuju", "3 - Netral", "4 - Setuju", "5 - Sangat Setuju"];

function renderPage6(data) {
  const rows = data.upward.map(u => {
    const n = u.dist.reduce((a, b) => a + b, 0) || 1;
    const avg = u.dist.reduce((s, c, i) => s + c * (i + 1), 0) / n;
    const maxC = Math.max(...u.dist, 1);
    const bars = u.dist.map((c, i) =>
      `<div class="ubar" style="height:${(3 + (c / maxC) * 30).toFixed(0)}px;background:${UPWARD_COLORS[i]}"></div>`).join("");
    return `
    <div class="urow">
      <div class="uitem">${esc(u.item)}</div>
      <div class="uavg"><span>${f1(avg)}</span></div>
      <div class="uhist">${bars}</div>
    </div>`;
  }).join("");

  const legend = `
  <div class="ulegend">
    ${UPWARD_LABELS.map((l, i) => `<span class="li"><span class="dot" style="background:${UPWARD_COLORS[i]}"></span>${esc(l)}</span>`).join("")}
  </div>`;

  const body = `
  <div class="upward-tbl">
    <div class="uhead"><span class="h-item">ITEM</span><span class="h-avg">AVERAGE</span><span class="h-hist">SEBARAN NILAI</span></div>
    ${rows}
    ${legend}
  </div>`;

  return pageShell(6, "UPWARD FEEDBACK — PENILAIAN ANGGOTA TERHADAP MANAGER", data, body);
}

// ═══════════════════════════════════════════════════════════════
// HALAMAN 7 — URAIAN UPWARD FEEDBACK (open feedback + tindak lanjut)
// ═══════════════════════════════════════════════════════════════
function renderPage7(data, num) {
  const boxes = data.upwardNotes.map(sec => `
  <div class="ufb-box">
    <div class="accent"></div>
    <div class="utitle">${esc(sec.title)}</div>
    ${sec.points.length
      ? `<ul>${sec.points.map(p => `<li>${esc(p)}</li>`).join("")}</ul>`
      : `<div class="empty">Tidak ada uraian dari anggota.</div>`}
  </div>`).join("");

  return pageShell(num, "URAIAN UPWARD FEEDBACK", data, boxes);
}

// ═══════════════════════════════════════════════════════════════
// RENDER
// ═══════════════════════════════════════════════════════════════
function renderReport(data) {
  let html = renderPage1(data) + renderPage2(data) + renderPage3(data) + renderPage4(data) + renderPage5(data);
  let num = 6;
  if (data.upward && data.upward.length) html += renderPage6(data), num++;
  if (data.upwardNotes && data.upwardNotes.length) html += renderPage7(data, num);
  document.getElementById("report").innerHTML = html;
}

// ═══════════════════════════════════════════════════════════════
// EXCEL — parse & template
// ═══════════════════════════════════════════════════════════════
function sheetRows(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return null;
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
}

function parseWorkbook(wb) {
  const err = (msg) => { throw new Error(msg); };

  // Info (key-value)
  const infoRows = sheetRows(wb, "Info") || err("Sheet 'Info' tidak ditemukan");
  const kv = {};
  infoRows.forEach(r => { if (r[0]) kv[String(r[0]).trim()] = r[1]; });
  const info = {
    tim: kv["Tim"] || "-",
    manager: kv["Manager"] || "-",
    batch: kv["Batch"] || "-",
    penerima: kv["Penerima"] || "-",
    label: kv["Label"] != null ? String(kv["Label"]).trim() : "",
    stdCommitment: Number(kv["Standar Commitment"]) || 3.5,
    stdCompetency: Number(kv["Standar Competency"]) || 3.5,
    stdContribution: Number(kv["Standar Contribution"]) || 3.5,
  };

  // Anggota
  const aRows = sheetRows(wb, "Anggota") || err("Sheet 'Anggota' tidak ditemukan");
  const members = aRows.slice(1).filter(r => r[0]).map(r => ({
    nama: String(r[0]).trim(),
    commitment: Number(r[1]),
    competency: Number(r[2]),
    contribution: Number(r[3]),
    blindspot: Number(r[4]) || 0,
  }));
  if (!members.length) err("Sheet 'Anggota' kosong");

  // Komitmen: header = [Nama, ...item names]
  const kRows = sheetRows(wb, "Komitmen") || err("Sheet 'Komitmen' tidak ditemukan");
  const commitmentItems = kRows[0].slice(1).filter(h => h !== "").map(String);
  const commitment = {};
  kRows.slice(1).filter(r => r[0]).forEach(r => {
    commitment[String(r[0]).trim()] = commitmentItems.map((_, i) => Number(r[i + 1]));
  });

  // Kompetensi: header = [Nama, ...aspects]; baris "STANDAR" = standar per aspek
  const cRows = sheetRows(wb, "Kompetensi") || err("Sheet 'Kompetensi' tidak ditemukan");
  const competencyAspects = cRows[0].slice(1).filter(h => h !== "").map(String);
  const competency = {};
  let competencyStd = competencyAspects.map(() => HEAT_STD);
  cRows.slice(1).filter(r => r[0]).forEach(r => {
    const key = String(r[0]).trim();
    const vals = competencyAspects.map((_, i) => Number(r[i + 1]));
    if (key.toUpperCase() === "STANDAR") competencyStd = vals;
    else competency[key] = vals;
  });

  // Penilai
  const pRows = sheetRows(wb, "Penilai") || err("Sheet 'Penilai' tidak ditemukan");
  const raters = {};
  pRows.slice(1).filter(r => r[0]).forEach(r => {
    raters[String(r[0]).trim()] = { self: Number(r[1]), atasan: Number(r[2]), rekan: Number(r[3]), bawahan: Number(r[4]) };
  });

  // Kontribusi
  const bRows = sheetRows(wb, "Kontribusi") || err("Sheet 'Kontribusi' tidak ditemukan");
  const contributionAspects = bRows[0].slice(1).filter(h => h !== "").map(String);
  const contribution = {};
  bRows.slice(1).filter(r => r[0]).forEach(r => {
    contribution[String(r[0]).trim()] = contributionAspects.map((_, i) => Number(r[i + 1]));
  });

  // STAR
  const sRows = sheetRows(wb, "STAR") || err("Sheet 'STAR' tidak ditemukan");
  const star = {};
  sRows.slice(1).filter(r => r[0]).forEach(r => {
    star[String(r[0]).trim()] = { jumlah: Number(r[1]) || 0, validasi: String(r[2] || "-"), ringkasan: String(r[3] || "") };
  });

  // Tren (opsional): riwayat batch SEBELUMNYA, long format [Batch, Nama, C, C, C]
  const tRows = sheetRows(wb, "Tren");
  let tren = null;
  if (tRows && tRows.length > 1) {
    const batches = [];
    const mT = {};
    tRows.slice(1).filter(r => r[0] !== "" && r[1] !== "").forEach(r => {
      const b = String(r[0]).trim(), nm = String(r[1]).trim();
      if (!batches.includes(b)) batches.push(b);
      const bi = batches.indexOf(b);
      if (!mT[nm]) mT[nm] = { commitment: [], competency: [], contribution: [] };
      const tnum = (v) => (v === "" || v == null) ? NaN : Number(v); // sel kosong = dimensi tanpa riwayat
      mT[nm].commitment[bi] = tnum(r[2]);
      mT[nm].competency[bi] = tnum(r[3]);
      mT[nm].contribution[bi] = tnum(r[4]);
    });
    if (batches.length) tren = { batches, members: mT };
  }

  // Upward (opsional): [Item, Average, N1..N5] — Average informasional, dihitung ulang dari dist
  const uRows = sheetRows(wb, "Upward");
  let upward = null;
  if (uRows && uRows.length > 1) {
    upward = uRows.slice(1).filter(r => r[0]).map(r => ({
      item: String(r[0]),
      dist: [2, 3, 4, 5, 6].map(i => Number(r[i]) || 0),
    })).filter(u => u.dist.some(c => c > 0));
    if (!upward.length) upward = null;
  }

  // Upward Uraian (opsional): [Bagian, Poin] long format, urutan bagian mengikuti kemunculan
  const unRows = sheetRows(wb, "Upward Uraian");
  let upwardNotes = null;
  if (unRows && unRows.length > 1) {
    const order = [];
    const map = {};
    unRows.slice(1).filter(r => r[0]).forEach(r => {
      const t = String(r[0]).trim();
      if (!map[t]) { map[t] = []; order.push(t); }
      const p = String(r[1] || "").trim();
      if (p) map[t].push(p);
    });
    upwardNotes = order.map(t => ({ title: t, points: map[t] }));
    if (!upwardNotes.length) upwardNotes = null;
  }

  // Insight
  const iRows = sheetRows(wb, "Insight") || [];
  const insights = { 1: [], 2: [], 3: [], 4: [], 5: [] };
  iRows.slice(1).filter(r => r[0] && r[1]).forEach(r => {
    const p = Number(r[0]);
    if (insights[p]) insights[p].push(String(r[1]));
  });

  // Validasi silang: anggota harus punya data di tiap sheet detail
  const missing = [];
  members.forEach(m => {
    if (!commitment[m.nama]) missing.push(`Komitmen: ${m.nama}`);
    if (!competency[m.nama]) missing.push(`Kompetensi: ${m.nama}`);
    if (!raters[m.nama]) missing.push(`Penilai: ${m.nama}`);
    if (!contribution[m.nama]) missing.push(`Kontribusi: ${m.nama}`);
    if (!star[m.nama]) missing.push(`STAR: ${m.nama}`);
  });
  if (missing.length) err("Data tidak lengkap — " + missing.slice(0, 4).join("; ") + (missing.length > 4 ? ` (+${missing.length - 4} lagi)` : ""));

  return { info, members, commitmentItems, commitment, competencyAspects, competencyStd, competency, raters, contributionAspects, contribution, star, tren, upward, upwardNotes, insights };
}

function buildTemplateWorkbook() {
  const wb = XLSX.utils.book_new();
  const d = SAMPLE;
  const aoa = (name, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);

  aoa("Info", [
    ["Tim", d.info.tim],
    ["Manager", d.info.manager],
    ["Batch", d.info.batch],
    ["Penerima", d.info.penerima],
    ["Label", d.info.label],
    ["Standar Commitment", d.info.stdCommitment],
    ["Standar Competency", d.info.stdCompetency],
    ["Standar Contribution", d.info.stdContribution],
  ]);
  aoa("Anggota", [
    ["Nama", "Commitment", "Competency", "Contribution", "Blindspot"],
    ...d.members.map(m => [m.nama, m.commitment, m.competency, m.contribution, m.blindspot]),
  ]);
  aoa("Komitmen", [
    ["Nama", ...d.commitmentItems],
    ...d.members.map(m => [m.nama, ...d.commitment[m.nama]]),
  ]);
  aoa("Kompetensi", [
    ["Nama", ...d.competencyAspects],
    ["STANDAR", ...d.competencyStd],
    ...d.members.map(m => [m.nama, ...d.competency[m.nama]]),
  ]);
  aoa("Penilai", [
    ["Nama", "Self", "Atasan", "Rekan", "Bawahan"],
    ...d.members.map(m => { const r = d.raters[m.nama]; return [m.nama, r.self, r.atasan, r.rekan, r.bawahan]; }),
  ]);
  aoa("Kontribusi", [
    ["Nama", ...d.contributionAspects],
    ...d.members.map(m => [m.nama, ...d.contribution[m.nama]]),
  ]);
  aoa("STAR", [
    ["Nama", "Jumlah Bukti", "Validasi", "Ringkasan"],
    ...d.members.map(m => { const s = d.star[m.nama]; return [m.nama, s.jumlah, s.validasi, s.ringkasan]; }),
  ]);
  aoa("Tren", [
    ["Batch", "Nama", "Commitment", "Competency", "Contribution"],
    ...d.tren.batches.flatMap((b, i) => d.members.map(m => {
      const h = d.tren.members[m.nama];
      return [b, m.nama, h.commitment[i], h.competency[i], h.contribution[i]];
    })),
  ]);
  aoa("Upward", [
    ["Item", "Average", "N1", "N2", "N3", "N4", "N5"],
    ...d.upward.map(u => {
      const n = u.dist.reduce((a, b) => a + b, 0) || 1;
      const avg = u.dist.reduce((s, c, i) => s + c * (i + 1), 0) / n;
      return [u.item, Math.round(avg * 100) / 100, ...u.dist];
    }),
  ]);
  aoa("Upward Uraian", [
    ["Bagian", "Poin"],
    ...d.upwardNotes.flatMap(sec => sec.points.map(p => [sec.title, p])),
  ]);
  aoa("Insight", [
    ["Halaman", "Teks"],
    ...Object.entries(d.insights).flatMap(([p, arr]) => arr.map(t => [Number(p), t])),
  ]);
  return wb;
}

// ═══════════════════════════════════════════════════════════════
// UI
// ═══════════════════════════════════════════════════════════════
function setStatus(msg, cls) {
  const el = document.getElementById("status");
  el.textContent = msg;
  el.className = cls || "";
}

function handleFile(file) {
  if (!file || !/\.(xlsx|xls)$/i.test(file.name)) {
    setStatus("File harus .xlsx", "err");
    return;
  }
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
      const data = parseWorkbook(wb);
      renderReport(data);
      setStatus(`✓ ${file.name} — ${data.members.length} anggota, tim "${data.info.tim}"`, "ok");
    } catch (ex) {
      setStatus("Gagal: " + ex.message, "err");
    }
  };
  reader.readAsArrayBuffer(file);
}

document.addEventListener("DOMContentLoaded", () => {
  renderReport(SAMPLE);

  document.getElementById("file-input").addEventListener("change", (e) => handleFile(e.target.files[0]));
  document.getElementById("btn-template").addEventListener("click", () => {
    XLSX.writeFile(buildTemplateWorkbook(), "orbid-report-template.xlsx");
  });
  document.getElementById("btn-print").addEventListener("click", () => window.print());

  let dragDepth = 0;
  document.body.addEventListener("dragenter", (e) => { e.preventDefault(); dragDepth++; document.body.classList.add("dragover"); });
  document.body.addEventListener("dragleave", (e) => { e.preventDefault(); if (--dragDepth <= 0) { dragDepth = 0; document.body.classList.remove("dragover"); } });
  document.body.addEventListener("dragover", (e) => e.preventDefault());
  document.body.addEventListener("drop", (e) => {
    e.preventDefault(); dragDepth = 0; document.body.classList.remove("dragover");
    handleFile(e.dataTransfer.files[0]);
  });
});
