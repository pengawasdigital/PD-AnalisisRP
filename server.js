import express from 'express';
import cors from 'cors';
import multer from 'multer';
import * as XLSX from 'xlsx';
import PDFDocument from 'pdfkit';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { createRequire } from 'module';
import mammoth from 'mammoth';

const require = createRequire(import.meta.url);
const pdfParseModule = require('pdf-parse');

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// File upload setup
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }
});

// Column mapping synonyms
const COLUMN_MAP = {
  no: ["no", "nomor", "#", "kode", "kode indikator", "no."],
  capaian: ["nama indikator", "nama_indikator", "indikator", "aspek", "kompetensi", "nama capaian"],
  skor_2025: ["skor rapor 2025", "skor 2025", "nilai 2025", "capaian 2025", "skor tahun ini", "skor berjalan", "skor satuan pendidikan 2025", "skor rapor", "skor", "nilai"],
  skor_2024: ["skor rapor 2024", "skor 2024", "nilai 2024", "capaian 2024", "skor tahun lalu", "skor satuan pendidikan 2024"],
  definisi: ["definisi capaian", "definisi indikator", "definisi", "pengertian", "deskripsi"],
  kategori: ["label capaian", "status capaian", "kategori capaian", "kategori", "kualifikasi", "predikat", "status", "label"],
  perubahan: ["perubahan skor", "perubahan", "selisih", "delta"]
};

// Mapping Kode Benahi & ARKAS
const KODE_BENAHI_MAP = {
  "A.1": {
    akar: "Kompetensi membaca dan analisis teks peserta didik masih rendah",
    benahi: [
      {
        kode_pbd: "PBD-LIT-01",
        kegiatan_pbd: "Peningkatan kompetensi guru dalam literasi membaca melalui Komunitas Belajar (Kombel)",
        kode_arkas: "03.02.01",
        kegiatan_arkas: "Pengembangan KKG/MGMP atau Gugus Sekolah"
      },
      {
        kode_pbd: "PBD-LIT-02",
        kegiatan_pbd: "Penyediaan bahan bacaan pengayaan untuk mendukung gerakan literasi sekolah",
        kode_arkas: "05.03.02",
        kegiatan_arkas: "Pengadaan Buku Teks Utama/Pendamping/Bacaan"
      }
    ]
  },
  "A.2": {
    akar: "Pemahaman konsep dasar matematika dan pemecahan masalah belum optimal",
    benahi: [
      {
        kode_pbd: "PBD-NUM-01",
        kegiatan_pbd: "Pelatihan pembelajaran numerasi berbasis media konkret dan kontekstual",
        kode_arkas: "03.02.04",
        kegiatan_arkas: "Peningkatan Kualitas Guru Mata Pelajaran/Kelas"
      }
    ]
  },
  "D.4": {
    akar: "Praktik pembelajaran interaktif dan diferensiasi belum berjalan konsisten",
    benahi: [
      {
        kode_pbd: "PBD-PBM-01",
        kegiatan_pbd: "Supervisi akademik dan diskusi peer-teaching antar guru",
        kode_arkas: "03.01.03",
        kegiatan_arkas: "Pelaksanaan Supervisi / Evaluasi Pembelajaran"
      }
    ]
  },
  "D.8": {
    akar: "Penerapan iklim keamanan, pencegahan perundungan, dan inklusivitas belum maksimal",
    benahi: [
      {
        kode_pbd: "PBD-IKL-01",
        kegiatan_pbd: "Sosialisasi dan pembentukan Tim Pencegahan dan Penanganan Kekerasan (TPPK)",
        kode_arkas: "06.07.01",
        kegiatan_arkas: "Penyelenggaraan Sekolah Sehat, Aman, Inklusif"
      }
    ]
  }
};

// ============ IN-MEMORY & FILE PERSISTENT HISTORY ============
const HISTORY_FILE = path.join(__dirname, 'history.json');
let history = [];
let nextHistoryId = 1;

try {
  if (fs.existsSync(HISTORY_FILE)) {
    const raw = fs.readFileSync(HISTORY_FILE, 'utf-8');
    history = JSON.parse(raw);
    if (history.length > 0) {
      nextHistoryId = Math.max(...history.map(h => h.id || 0)) + 1;
    }
  }
} catch (e) {
  console.warn('Could not load history.json:', e.message);
}

function saveHistory() {
  try {
    fs.writeFileSync(HISTORY_FILE, JSON.stringify(history, null, 2), 'utf-8');
  } catch (e) {
    console.error('Failed to save history:', e.message);
  }
}

// Seed with an initial sample report if history is empty
if (history.length === 0) {
  const sampleData = {
    judul: "Analisis Rapor Pendidikan UPT SD NEGERI 003 LUBUK SAKAT Tahun 2025",
    ringkasan: {
      total_indikator: 6,
      jumlah_per_kategori: { Capaian: 0, Baik: 3, Sedang: 2, Kurang: 1 },
      rata_rata_skor_2025: 68.42,
      indikator_naik: 4,
      indikator_turun: 2
    },
    kesimpulan: "Berdasarkan hasil analisis Rapor Pendidikan, dari 6 indikator utama yang dianalisis, sebanyak 3 indikator berada pada predikat Baik, 2 indikator pada predikat Sedang, 1 indikator pada predikat Kurang, dan 0 indikator pada predikat Capaian. Rata-rata skor capaian tahun 2025 adalah 68.42, dengan 4 indikator mengalami peningkatan skor dan 2 indikator mengalami penurunan dibandingkan tahun 2024. Capaian tertinggi terdapat pada indikator Iklim Keamanan Sekolah dengan skor 78.5, sedangkan capaian terendah pada indikator Kemampuan Literasi dengan skor 54.2. Terdapat 1 indikator yang masih berpredikat <span class=\"hl-kurang\">Kurang</span> sehingga menjadi fokus utama intervensi sekolah, yaitu: <span class=\"hl-kurang\">1) Kemampuan Literasi</span> (skor 54.2).",
    capaian: [],
    baik: [
      { no: "D.8", capaian: "Iklim Keamanan Sekolah", skor_2025: 78.5, definisi_capaian: "Kondisi lingkungan sekolah yang aman dan kondusif", perubahan_skor: 4.2, sub_indikator: [] },
      { no: "D.1", capaian: "Kualitas Pembelajaran", skor_2025: 72.1, definisi_capaian: "Manajemen kelas dan instruksi pembelajaran", perubahan_skor: 2.5, sub_indikator: [] },
      { no: "D.3", capaian: "Kepemimpinan Instruksional", skor_2025: 70.0, definisi_capaian: "Visi pembelajaran kepala sekolah", perubahan_skor: 1.8, sub_indikator: [] }
    ],
    sedang: [
      { no: "A.2", capaian: "Kemampuan Numerasi", skor_2025: 65.4, definisi_capaian: "Kompetensi peserta didik dalam pemecahan masalah matematika", perubahan_skor: -1.2, sub_indikator: [] },
      { no: "A.3", capaian: "Karakter", skor_2025: 66.3, definisi_capaian: "Penerapan Profil Pelajar Pancasila", perubahan_skor: 0.5, sub_indikator: [] }
    ],
    kurang: [
      { no: "A.1", capaian: "Kemampuan Literasi", skor_2025: 54.2, definisi_capaian: "Kompetensi membaca dan analisis teks peserta didik", perubahan_skor: -3.5, sub_indikator: [] }
    ]
  };

  history.push({
    id: nextHistoryId++,
    nama_sekolah: "UPT SD NEGERI 003 LUBUK SAKAT",
    tahun: "2025",
    filename: "Rapor_Pendidikan_SDN_003_Lubuk_Sakat_2025.xlsx",
    total_indikator: sampleData.ringkasan.total_indikator,
    rata_rata_skor: sampleData.ringkasan.rata_rata_skor_2025,
    indikator_naik: sampleData.ringkasan.indikator_naik,
    indikator_turun: sampleData.ringkasan.indikator_turun,
    jumlah_baik: sampleData.ringkasan.jumlah_per_kategori.Baik,
    jumlah_sedang: sampleData.ringkasan.jumlah_per_kategori.Sedang,
    jumlah_kurang: sampleData.ringkasan.jumlah_per_kategori.Kurang,
    data: sampleData,
    created_at: new Date().toISOString()
  });
  saveHistory();
}

// ============ HELPER FUNCTIONS ============
function cleanAngka(v) {
  if (v === null || v === undefined || v === '') return null;
  let s = String(v).trim().replace(/%/g, '').replace(/Rp/g, '');
  s = s.replace(/[^\d.,\-]/g, '');
  if (!s || s === '.' || s === '-' || s === ',') return null;
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function parsePerubahan(v) {
  if (v === null || v === undefined || v === '') return null;
  const s = String(v).trim().toLowerCase();
  if (s.includes('tidak berubah') || s.includes('tetap')) return 0.0;
  const angka = cleanAngka(s);
  if (angka === null) return null;
  if (s.includes('turun') || s.includes('menurun')) return -angka;
  return angka;
}

function ekstrakMeta(judul, namaFile) {
  const teks = judul || namaFile || '';
  const m = teks.match(/(20\d{2})/);
  const tahun = m ? m[1] : '';
  let bersih = teks.replace(/(20\d{2})/g, ' ');
  const generik = [
    "laporan", "rapor", "pendidikan", "tahun", "pbd", "unduhan",
    "satuan", "perencanaan", "berbasis", "data", "lembar", "dokumen"
  ];
  for (const g of generik) {
    bersih = bersih.replace(new RegExp('\\b' + g + '\\b', 'gi'), ' ');
  }
  bersih = bersih.replace(/\s+/g, ' ').replace(/^-+|-+$/g, '').trim();
  return {
    nama_sekolah: bersih || '-',
    tahun: tahun || '-'
  };
}

function normalizeColumns(rowObj) {
  const normalized = {};
  for (const [col, val] of Object.entries(rowObj)) {
    const cl = String(col).trim().toLowerCase();
    let matchedKey = col;
    for (const [baku, sinonim] of Object.entries(COLUMN_MAP)) {
      for (const s of sinonim) {
        if (cl === s || (s.length >= 5 && cl.includes(s))) {
          matchedKey = baku;
          break;
        }
      }
      if (matchedKey === baku) break;
    }
    normalized[matchedKey] = val;
  }
  return normalized;
}

// ============ FILE PARSING ============
function findBestSheet(wb) {
  const keywords = ['laporan rapor', 'laporan_rapor', 'rapor pendidikan', 'rapor_pendidikan', 'laporan', 'rapor', 'rekap', 'indikator', 'capaian', 'prioritas'];
  const skipWords = ['panduan', 'petunjuk', 'baca', 'info', 'keterangan'];

  // 1. Try keyword match excluding skipwords
  for (const kw of keywords) {
    for (const s of wb.SheetNames) {
      const sl = s.toLowerCase();
      if (sl.includes(kw) && !skipWords.some(w => sl.includes(w))) {
        return s;
      }
    }
  }

  // 2. Pick non-skip sheet with most rows
  let best = wb.SheetNames[0];
  let maxRows = -1;
  for (const s of wb.SheetNames) {
    const sl = s.toLowerCase();
    if (skipWords.some(w => sl.includes(w))) continue;
    const sheet = wb.Sheets[s];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (data.length > maxRows) {
      maxRows = data.length;
      best = s;
    }
  }
  return best;
}

function parseExcel(content, existingWb = null) {
  const wb = existingWb || XLSX.read(content, { type: 'buffer' });
  const target = findBestSheet(wb);
  const raw = XLSX.utils.sheet_to_json(wb.Sheets[target], { header: 1, defval: '' });

  let judul = '';
  for (let i = 0; i < Math.min(10, raw.length); i++) {
    const row = raw[i] || [];
    for (const val of row) {
      const strVal = String(val).toLowerCase();
      if (strVal.includes('rapor') && strVal.includes('pendidikan')) {
        if (/(20\d{2})/.test(strVal)) {
          judul = String(val).trim();
          break;
        }
      }
    }
    if (judul) break;
  }

  // Find header row by finding the row with the most keyword matches
  let bestHeaderIdx = null;
  let maxHeaderScore = 0;
  for (let i = 0; i < Math.min(25, raw.length); i++) {
    const row = raw[i] || [];
    let score = 0;
    const rowStr = row.map(c => String(c).trim().toLowerCase()).join(' ');
    if (rowStr.includes('no') || rowStr.includes('nomor') || rowStr.includes('kode')) score++;
    if (rowStr.includes('indikator') || rowStr.includes('aspek') || rowStr.includes('kompetensi')) score += 2;
    if (rowStr.includes('skor') || rowStr.includes('nilai')) score += 2;
    if (rowStr.includes('capaian') || rowStr.includes('kategori') || rowStr.includes('predikat') || rowStr.includes('label')) score++;
    if (rowStr.includes('definisi') || rowStr.includes('deskripsi')) score++;
    if (rowStr.includes('perubahan') || rowStr.includes('selisih') || rowStr.includes('delta')) score++;

    if (score > maxHeaderScore) {
      maxHeaderScore = score;
      bestHeaderIdx = i;
    }
  }

  let rows = [];
  if (bestHeaderIdx !== null && maxHeaderScore >= 2) {
    const headerRow = raw[bestHeaderIdx];
    const colMapping = {};

    // Analyze header cells
    for (let j = 0; j < headerRow.length; j++) {
      const headerText = String(headerRow[j] || '').trim().toLowerCase();
      if (!headerText) continue;

      let matched = null;
      for (const [baku, sinonim] of Object.entries(COLUMN_MAP)) {
        for (const s of sinonim) {
          if (headerText === s || (s.length >= 4 && headerText.includes(s))) {
            matched = baku;
            break;
          }
        }
        if (matched) break;
      }

      // Disambiguation for "capaian"
      if (matched === 'capaian' && colMapping.capaian !== undefined) {
        matched = 'kategori';
      } else if (headerText === 'capaian') {
        // Peek at data in this column: if it contains "baik", "sedang", "kurang", it's kategori
        let isKategoriValue = false;
        for (let k = bestHeaderIdx + 1; k < Math.min(bestHeaderIdx + 6, raw.length); k++) {
          const sampleVal = String((raw[k] || [])[j] || '').toLowerCase();
          if (sampleVal.includes('baik') || sampleVal.includes('sedang') || sampleVal.includes('kurang')) {
            isKategoriValue = true;
            break;
          }
        }
        if (isKategoriValue) matched = 'kategori';
      }

      if (matched && colMapping[matched] === undefined) {
        colMapping[matched] = j;
      }
    }

    // Fallbacks if some columns were not detected by name
    if (colMapping.no === undefined) {
      for (let j = 0; j < headerRow.length; j++) {
        if (j !== colMapping.capaian && j !== colMapping.skor_2025) {
          const sample = String((raw[bestHeaderIdx + 1] || [])[j] || '').trim();
          if (/^[A-Z]\.\d+(\.\d+)?$|^\d+$/.test(sample)) {
            colMapping.no = j;
            break;
          }
        }
      }
    }
    if (colMapping.skor_2025 === undefined) {
      for (let j = 0; j < headerRow.length; j++) {
        if (j !== colMapping.no && j !== colMapping.capaian) {
          const sample = parseFloat(String((raw[bestHeaderIdx + 1] || [])[j] || '').replace(',', '.'));
          if (!isNaN(sample) && sample >= 0 && sample <= 100) {
            colMapping.skor_2025 = j;
            break;
          }
        }
      }
    }

    for (let i = bestHeaderIdx + 1; i < raw.length; i++) {
      const r = raw[i];
      if (!r || r.every(c => c === '' || c === null || c === undefined)) continue;

      const rowObj = {
        no: colMapping.no !== undefined ? r[colMapping.no] : '',
        capaian: colMapping.capaian !== undefined ? r[colMapping.capaian] : '',
        kategori: colMapping.kategori !== undefined ? r[colMapping.kategori] : '',
        skor_2025: colMapping.skor_2025 !== undefined ? r[colMapping.skor_2025] : '',
        skor_2024: colMapping.skor_2024 !== undefined ? r[colMapping.skor_2024] : '',
        definisi: colMapping.definisi !== undefined ? r[colMapping.definisi] : '',
        perubahan: colMapping.perubahan !== undefined ? r[colMapping.perubahan] : ''
      };

      // Fallback: if capaian is empty, maybe column 1 has text
      if (!rowObj.capaian && r[1]) rowObj.capaian = r[1];
      if (!rowObj.no && r[0]) rowObj.no = r[0];

      if (rowObj.capaian && String(rowObj.capaian).trim().length > 1) {
        rows.push(rowObj);
      }
    }
  } else {
    // Fallback if no clean header: attempt positional read
    for (let i = 0; i < raw.length; i++) {
      const r = raw[i];
      if (!r || r.every(c => c === '' || c === null || c === undefined)) continue;
      const c1 = String(r[1] || '').trim();
      if (!c1 || c1.toLowerCase() === 'indikator' || c1.toLowerCase() === 'nama indikator') continue;
      rows.push({
        no: r[0],
        capaian: r[1],
        kategori: r[2],
        skor_2025: r[3],
        definisi: r[4],
        perubahan: r[5],
        skor_2024: r[6]
      });
    }
  }

  return { rows, judul };
}

function parseTextLines(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const rows = [];
  let judul = '';

  for (let i = 0; i < Math.min(10, lines.length); i++) {
    const l = lines[i];
    if (l.toLowerCase().includes('rapor') && l.toLowerCase().includes('pendidikan')) {
      judul = l;
      break;
    }
  }

  // Regex pattern for indicator line e.g. "A.1 Kemampuan Literasi 75.50 Baik"
  const indPattern = /^([A-Z]\.\d+(\.\d+)?|\d+)\s+([A-Za-z\s,()/-]+?)\s+(\d+[.,]?\d*)\s*(Baik|Sedang|Kurang|Tinggi|Menengah|Rendah)?/i;

  for (const line of lines) {
    const match = line.match(indPattern);
    if (match) {
      rows.push({
        no: match[1],
        capaian: match[3].trim(),
        skor_2025: match[4].replace(',', '.'),
        kategori: match[5] || '',
        skor_2024: '',
        perubahan: '',
        definisi: ''
      });
    }
  }

  return { rows, judul };
}

async function parseDocx(buffer) {
  const result = await mammoth.convertToHtml({ buffer });
  const html = result.value || '';
  if (html.includes('<table')) {
    const wb = XLSX.read(html, { type: 'string' });
    return parseExcel(null, wb);
  }
  return parseTextLines(result.value.replace(/<[^>]+>/g, '\n'));
}

async function parsePdf(buffer) {
  try {
    if (pdfParseModule && pdfParseModule.PDFParse) {
      const parser = new pdfParseModule.PDFParse({ data: buffer });
      const textResult = await parser.getText();
      await parser.destroy().catch(() => {});
      const extractedText = typeof textResult === 'string' ? textResult : (textResult?.text || '');
      return parseTextLines(extractedText);
    } else if (typeof pdfParseModule === 'function') {
      const data = await pdfParseModule(buffer);
      return parseTextLines(data.text || '');
    }
  } catch (err) {
    console.warn('PDF parsing warning:', err.message);
  }
  return { rows: [], judul: '' };
}

// Fallback plain text / generic parser
async function parseFileContent(buffer, ext) {
  if (ext === '.docx') {
    return await parseDocx(buffer);
  }
  if (ext === '.pdf') {
    return await parsePdf(buffer);
  }
  return parseExcel(buffer);
}

// ============ ANALISIS ENGINE ============
function runAnalisis(rawRows, judul = '') {
  const rows = [];
  for (const r of rawRows) {
    const rawCapaian = r.capaian !== undefined ? r.capaian : r['nama indikator'];
    if (!rawCapaian || String(rawCapaian).trim() === '') continue;

    const skor25 = cleanAngka(r.skor_2025);
    const skor24 = cleanAngka(r.skor_2024);
    let perubahan = parsePerubahan(r.perubahan);
    if (perubahan === null && skor25 !== null && skor24 !== null) {
      perubahan = Math.round((skor25 - skor24) * 100) / 100;
    }

    let kode = String(r.no || '').trim();
    if (!kode || kode.toLowerCase() === 'nan' || kode.toLowerCase() === 'none') {
      kode = '-';
    }

    let nama = String(rawCapaian).trim();
    let definisi = String(r.definisi || '').trim();
    if (nama.includes('\n')) {
      const bagian = nama.split('\n').map(b => b.trim()).filter(Boolean);
      nama = bagian[0];
      if (!definisi && bagian.length > 1) {
        definisi = bagian.slice(1).join(' ');
      }
    }

    let kateg = String(r.kategori || '').trim();
    if (!kateg || ['nan', 'none', '-'].includes(kateg.toLowerCase())) {
      kateg = '';
    }

    // Auto-infer kategori from score if missing
    if (!kateg && skor25 !== null) {
      if (skor25 >= 70) kateg = 'Baik';
      else if (skor25 >= 60) kateg = 'Sedang';
      else kateg = 'Kurang';
    }

    rows.push({
      no: kode,
      capaian: nama,
      skor_2025: skor25,
      definisi_capaian: definisi,
      perubahan_skor: perubahan,
      kategori: kateg
    });
  }

  if (rows.length === 0) {
    throw new Error('Tidak ada baris data yang bisa dianalisis dari file.');
  }

  function normKategori(k) {
    const kl = (k || '').toLowerCase();
    if (kl.includes('baik') || kl.includes('tinggi')) return 'Baik';
    if (kl.includes('sedang') || kl.includes('cukup') || kl.includes('menengah')) return 'Sedang';
    if (kl.includes('kurang') || kl.includes('rendah') || kl.includes('intervensi')) return 'Kurang';
    if (kl.includes('capaian')) return 'Capaian';
    return 'Capaian';
  }

  const grouped = { Capaian: [], Baik: [], Sedang: [], Kurang: [] };
  let parentAktif = null;
  const skorUtama = [];
  let naik = 0, turun = 0, jumlahUtama = 0;
  let terbaik = null, terendah = null;
  const kurangList = [];

  for (const r of rows) {
    let kateg = (r.kategori || '').toLowerCase().trim();
    if (!kateg && r.skor_2025 !== null) {
      if (r.skor_2025 >= 70) kateg = 'baik';
      else if (r.skor_2025 >= 60) kateg = 'sedang';
      else kateg = 'kurang';
    }

    // Sub-indicator detection: e.g. A.1.1, D.1.2 or lowercase sub-levels
    const isSubIndikator = /^[A-Z]\.\d+\.\d+/.test(r.no) || (parentAktif && r.no === '-' && !r.kategori);

    const item = {
      no: r.no,
      capaian: r.capaian,
      skor_2025: r.skor_2025,
      definisi_capaian: r.definisi_capaian,
      perubahan_skor: r.perubahan_skor,
      sub_indikator: []
    };

    if (isSubIndikator && parentAktif) {
      parentAktif.sub_indikator.push(item);
    } else {
      let g = normKategori(kateg);
      grouped[g].push(item);
      if (g === 'Kurang') kurangList.push([item.capaian, item.skor_2025]);
      parentAktif = item;
      jumlahUtama++;

      if (item.skor_2025 !== null) {
        skorUtama.push(item.skor_2025);
        if (terbaik === null || item.skor_2025 > terbaik[1]) terbaik = [item.capaian, item.skor_2025];
        if (terendah === null || item.skor_2025 < terendah[1]) terendah = [item.capaian, item.skor_2025];
      }
      if ((item.perubahan_skor || 0) > 0) naik++;
      else if ((item.perubahan_skor || 0) < 0) turun++;
    }
  }

  const jpk = {
    Capaian: grouped.Capaian.length,
    Baik: grouped.Baik.length,
    Sedang: grouped.Sedang.length,
    Kurang: grouped.Kurang.length
  };
  const totalDianalisis = jumlahUtama > 0 ? jumlahUtama : (jpk.Baik + jpk.Sedang + jpk.Kurang + jpk.Capaian);
  const rata2 = skorUtama.length > 0 ? Math.round((skorUtama.reduce((a, b) => a + b, 0) / skorUtama.length) * 100) / 100 : null;

  let kesimpulan = `Berdasarkan hasil analisis Rapor Pendidikan, dari ${totalDianalisis} indikator utama yang dianalisis, sebanyak ${jpk.Baik} indikator berada pada predikat Baik, ${jpk.Sedang} indikator pada predikat Sedang, ${jpk.Kurang} indikator pada predikat Kurang, dan ${jpk.Capaian} indikator pada predikat Capaian. Rata-rata skor capaian tahun 2025 adalah ${rata2 !== null ? rata2 : '-'}, dengan ${naik} indikator mengalami peningkatan skor dan ${turun} indikator mengalami penurunan dibandingkan tahun 2024.`;

  if (terbaik !== null || terendah !== null) {
    kesimpulan += ' Capaian tertinggi';
    if (terbaik !== null) kesimpulan += ` terdapat pada indikator ${terbaik[0]} dengan skor ${terbaik[1]}`;
    if (terbaik !== null && terendah !== null) kesimpulan += ',';
    if (terendah !== null) kesimpulan += ` sedangkan capaian terendah pada indikator ${terendah[0]} dengan skor ${terendah[1]}`;
    kesimpulan += '.';
  }

  if (jpk.Kurang > 0) {
    kesimpulan += ` Terdapat ${jpk.Kurang} indikator yang masih berpredikat <span class="hl-kurang">Kurang</span> sehingga menjadi fokus utama intervensi sekolah`;
    if (kurangList.length > 0) {
      const daftar = kurangList.map((item, i) => `<span class="hl-kurang">${i + 1}) ${item[0]}</span>` + (item[1] !== null ? ` (skor ${item[1]})` : '')).join('; ');
      kesimpulan += `, yaitu: ${daftar}.`;
    }
  } else if (jpk.Sedang > 0) {
    kesimpulan += ' Indikator yang masih berpredikat Sedang perlu diperkuat agar dapat mencapai predikat Baik.';
  } else {
    kesimpulan += ' Capaian mutu pendidikan secara keseluruhan sudah berjalan dengan baik.';
  }

  const ringkasan = {
    total_indikator: totalDianalisis,
    jumlah_per_kategori: jpk,
    rata_rata_skor_2025: rata2,
    indikator_naik: naik,
    indikator_turun: turun
  };

  return {
    judul,
    capaian: grouped.Capaian,
    baik: grouped.Baik,
    sedang: grouped.Sedang,
    kurang: grouped.Kurang,
    ringkasan,
    kesimpulan
  };
}

// Generate Rekomendasi
function generateRekomendasi(result) {
  const rekomendasi = [];
  const allKurang = [];
  for (const item of (result.kurang || [])) {
    allKurang.push(item);
    for (const sub of (item.sub_indikator || [])) {
      allKurang.push(sub);
    }
  }

  let idx = 1;
  for (const item of allKurang) {
    const kodeNo = String(item.no || '').trim().toUpperCase();
    const namaCapaian = String(item.capaian || '').toLowerCase();

    let matchedData = null;
    for (const [keyKode, valMap] of Object.entries(KODE_BENAHI_MAP)) {
      if (kodeNo.startsWith(keyKode) || namaCapaian.includes(keyKode.toLowerCase())) {
        matchedData = valMap;
        break;
      }
    }

    const skorFmt = item.skor_2025 !== null && item.skor_2025 !== undefined ? Number(item.skor_2025).toFixed(2) : '-';

    if (matchedData) {
      for (const b of matchedData.benahi) {
        rekomendasi.push({
          no: idx,
          kode_indikator: item.no,
          indikator: item.capaian,
          skor: skorFmt,
          identifikasi_masalah: `Capaian ${item.capaian} belum optimal`,
          refleksi_akar_masalah: matchedData.akar,
          kode_pbd: b.kode_pbd,
          kegiatan_pembenahan: b.kegiatan_pbd,
          kode_arkas: b.kode_arkas,
          kegiatan_arkas: b.kegiatan_arkas,
          estimasi_anggaran: 'Disesuaikan ARKAS'
        });
        idx++;
      }
    } else {
      rekomendasi.push({
        no: idx,
        kode_indikator: item.no,
        indikator: item.capaian,
        skor: skorFmt,
        identifikasi_masalah: `Capaian ${item.capaian} masih rendah`,
        refleksi_akar_masalah: 'Perlu analisis berbasis data bersama tim PBD',
        kode_pbd: 'PBD-GEN-01',
        kegiatan_pembenahan: 'Evaluasi dan penyusunan program pembenahan',
        kode_arkas: '03.01.01',
        kegiatan_arkas: 'Penyusunan RKT / RKAS',
        estimasi_anggaran: 'Disesuaikan'
      });
      idx++;
    }
  }

  return rekomendasi;
}

// ============ API ROUTES ============

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

// Analyze uploaded file
app.post('/api/analyze', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ detail: 'Pilih file dulu.' });
    }

    const originalName = req.file.originalname || '';
    const ext = path.extname(originalName).toLowerCase();
    const allowed = ['.xlsx', '.xls', '.docx', '.pdf', '.csv'];
    if (!allowed.includes(ext)) {
      return res.status(400).json({ detail: `Format tidak didukung: ${ext}. Gunakan .xlsx, .xls, .docx, atau .pdf.` });
    }

    const { rows, judul } = await parseFileContent(req.file.buffer, ext);
    if (!rows || rows.length === 0) {
      return res.status(422).json({ detail: 'Tidak ada data indikator rapor yang dapat diekstrak dari file ini. Pastikan file adalah dokumen Rapor Pendidikan yang valid.' });
    }

    const meta = ekstrakMeta(judul, originalName);
    if (req.body && req.body.nama_sekolah) {
      meta.nama_sekolah = String(req.body.nama_sekolah).trim();
    }
    const judulAkhir = `Analisis Rapor Pendidikan ${meta.nama_sekolah} Tahun ${meta.tahun}`;

    const hasil = runAnalisis(rows, judulAkhir);

    if (!hasil || !hasil.ringkasan) {
      return res.status(422).json({ detail: 'Gagal memproses ringkasan data dari file.' });
    }

    // Save to history
    const rData = hasil.ringkasan;
    const jpk = rData.jumlah_per_kategori || {};

    const newHistory = {
      id: nextHistoryId++,
      nama_sekolah: meta.nama_sekolah,
      tahun: meta.tahun,
      filename: originalName,
      total_indikator: rData.total_indikator || 0,
      rata_rata_skor: rData.rata_rata_skor_2025 || 0.0,
      indikator_naik: rData.indikator_naik || 0,
      indikator_turun: rData.indikator_turun || 0,
      jumlah_baik: jpk.Baik || 0,
      jumlah_sedang: jpk.Sedang || 0,
      jumlah_kurang: jpk.Kurang || 0,
      data: hasil,
      created_at: new Date().toISOString()
    };

    history.unshift(newHistory);
    saveHistory();

    return res.json(hasil);
  } catch (err) {
    console.error('Analysis error:', err);
    return res.status(422).json({ detail: `Analisis gagal: ${err.message}` });
  }
});

// Recommendations endpoint
app.post('/api/recommendations', (req, res) => {
  try {
    const data = req.body || {};
    const rekomendasi = generateRekomendasi(data);
    res.json({ rekomendasi });
  } catch (err) {
    console.error('Recommendations error:', err);
    res.status(500).json({ detail: `Gagal generate rekomendasi: ${err.message}` });
  }
});

// Compare multiple files
app.post('/api/compare', upload.array('files'), (req, res) => {
  try {
    const files = req.files || [];
    if (files.length < 2) {
      return res.status(400).json({ detail: 'Upload minimal 2 file.' });
    }

    const results = [];
    for (const f of files) {
      try {
        const { rows, judul } = parseExcel(f.buffer);
        const meta = ekstrakMeta(judul, f.originalname);
        const skors = rows.map(r => cleanAngka(r.skor_2025)).filter(s => s !== null);
        const avg = skors.length > 0 ? Math.round((skors.reduce((a, b) => a + b, 0) / skors.length) * 100) / 100 : 0;
        results.push({
          sekolah: meta.nama_sekolah,
          avg_score: avg
        });
      } catch (e) {
        console.warn('Skipping compare file:', f.originalname, e.message);
      }
    }

    res.json(results);
  } catch (err) {
    console.error('Compare error:', err);
    res.status(500).json({ detail: `Gagal membandingkan file: ${err.message}` });
  }
});

// History CRUD
app.get('/api/history', (req, res) => {
  const summary = history.map(({ data, ...rest }) => rest);
  res.json(summary);
});

app.get('/api/history/:hid', (req, res) => {
  const hid = parseInt(req.params.hid, 10);
  const found = history.find(h => h.id === hid);
  if (!found) {
    return res.status(404).json({ detail: 'Riwayat tidak ditemukan.' });
  }
  res.json(found);
});

app.delete('/api/history/:hid', (req, res) => {
  const hidParam = req.params.hid;
  const hidNum = parseInt(hidParam, 10);
  const idx = history.findIndex(h => h.id === hidNum || String(h.id) === String(hidParam));
  if (idx === -1) {
    return res.status(404).json({ detail: 'Riwayat tidak ditemukan.' });
  }
  history.splice(idx, 1);
  saveHistory();
  res.json({ status: 'ok', message: 'Riwayat berhasil dihapus.' });
});

// Export PDF (Full Report)
app.post('/api/export-pdf', (req, res) => {
  try {
    const data = req.body || {};
    const doc = new PDFDocument({ margin: 40 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=rapor_pendidikan.pdf');
    doc.pipe(res);

    doc.fontSize(16).text(data.judul || 'Analisis Rapor Pendidikan', { align: 'center' });
    doc.moveDown(1);

    const kesimpulan = (data.kesimpulan || '').replace(/<[^>]+>/g, '');
    doc.fontSize(11).text(kesimpulan, { align: 'justify' });
    doc.moveDown(1.5);

    for (const kelompok of ['baik', 'sedang', 'kurang']) {
      const items = data[kelompok] || [];
      if (items.length > 0) {
        doc.fontSize(12).fillColor('#333333').text(`KELOMPOK ${kelompok.toUpperCase()} (${items.length})`);
        doc.fontSize(10).fillColor('#000000');
        for (const item of items) {
          doc.text(`- ${item.no || ''} ${item.capaian || ''} (Skor: ${item.skor_2025 !== null && item.skor_2025 !== undefined ? item.skor_2025 : '-'})`);
        }
        doc.moveDown(0.8);
      }
    }

    doc.end();
  } catch (err) {
    console.error('Export PDF error:', err);
    res.status(500).json({ detail: `Gagal generate PDF: ${err.message}` });
  }
});

// Export Recommendations Excel (.xlsx)
app.post('/api/export-recommendations-excel', (req, res) => {
  try {
    const data = req.body || {};
    const rekomList = generateRekomendasi(data);

    const df = rekomList.map(r => ({
      "No": r.no,
      "Kode Indikator": r.kode_indikator,
      "Indikator / Capaian": r.indikator,
      "Skor 2025": r.skor,
      "Identifikasi Masalah": r.identifikasi_masalah,
      "Refleksi Akar Masalah": r.refleksi_akar_masalah,
      "Kode PBD": r.kode_pbd,
      "Program Pembenahan (PBD)": r.kegiatan_pembenahan,
      "Kode ARKAS": r.kode_arkas,
      "Kegiatan ARKAS": r.kegiatan_arkas,
      "Estimasi Anggaran": r.estimasi_anggaran
    }));

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(df);
    XLSX.utils.book_append_sheet(wb, ws, "Rekomendasi PBD-ARKAS");
    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Rekomendasi_PBD_ARKAS.xlsx');
    res.send(buffer);
  } catch (err) {
    console.error('Export Excel error:', err);
    res.status(500).json({ detail: `Gagal export Rekomendasi Excel: ${err.message}` });
  }
});

// Export Recommendations PDF
app.post('/api/export-recommendations-pdf', (req, res) => {
  try {
    const data = req.body || {};
    const rekomList = generateRekomendasi(data);

    const doc = new PDFDocument({ margin: 40 });
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'attachment; filename=Rekomendasi_PBD_ARKAS.pdf');
    doc.pipe(res);

    doc.fontSize(14).text("REKOMENDASI PROGRAM BENAHI (PBD) & ARKAS", { align: 'center' });
    doc.moveDown(1);

    if (rekomList.length === 0) {
      doc.fontSize(11).text("Tidak ada rekomendasi program yang diperlukan.");
    } else {
      for (const item of rekomList) {
        doc.fillColor('#b40000').fontSize(11).text(`[${item.kode_indikator}] ${item.indikator} (Skor: ${item.skor})`);
        doc.fillColor('#000000').fontSize(10);
        doc.text(`  - Akar Masalah: ${item.refleksi_akar_masalah}`);
        doc.text(`  - Program PBD (${item.kode_pbd}): ${item.kegiatan_pembenahan}`);
        doc.text(`  - Kegiatan ARKAS (${item.kode_arkas}): ${item.kegiatan_arkas}`);
        doc.moveDown(0.6);
      }
    }

    doc.end();
  } catch (err) {
    console.error('Export Recs PDF error:', err);
    res.status(500).json({ detail: `Gagal export Rekomendasi PDF: ${err.message}` });
  }
});

// ============ SERVE STATIC FRONTEND ============
const htmlDir = path.join(__dirname, 'html');
app.use(express.static(htmlDir));

app.get('*', (req, res) => {
  res.sendFile(path.join(htmlDir, 'index.html'));
});

// Start server on 0.0.0.0:3000
app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server running on http://0.0.0.0:${PORT}`);
});
