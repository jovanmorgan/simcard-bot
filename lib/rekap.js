const { google } = require("googleapis");
const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

// SEMUA FILE HASIL EXPORT DISIMPAN DI SINI, SUPAYA RAPI DAN TIDAK BERCAMPUR
// DENGAN FILE LAINNYA
const OUTPUT_DIR = path.join(__dirname, "..", "Hasil-Rekap");

// NAMA BULAN INGGRIS (dipakai untuk deteksi input) & INDONESIA (dipakai untuk penamaan file)
const MONTH_NAMES_EN = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];

const MONTH_NAMES_ID = [
  "Januari",
  "Februari",
  "Maret",
  "April",
  "Mei",
  "Juni",
  "Juli",
  "Agustus",
  "September",
  "Oktober",
  "November",
  "Desember",
];

// ALIAS/SINGKATAN TIAP BULAN (INGGRIS & INDONESIA), DIPAKAI SEBAGAI KAMUS
// SAAT MENCOCOKKAN TEKS YANG MENGANDUNG NAMA BULAN
const MONTH_ALIASES = [
  ["jan", "january", "januari"],
  ["feb", "february", "februari"],
  ["mar", "march", "maret"],
  ["apr", "april"],
  ["may", "mei"],
  ["jun", "june", "juni"],
  ["jul", "july", "juli"],
  ["aug", "agu", "ags", "august", "agustus"],
  ["sep", "sept", "september"],
  ["oct", "okt", "october", "oktober"],
  ["nov", "november"],
  ["dec", "des", "december", "desember"],
];

// JARAK EDIT (LEVENSHTEIN) ANTARA DUA STRING, DIPAKAI UNTUK TOLERANSI TYPO
// SAAT MENDETEKSI NAMA BULAN (MIS. "Jully", "Agustuz", "Novembar")
function levenshteinDistance(a, b) {
  a = a.toLowerCase();
  b = b.toLowerCase();
  const m = a.length;
  const n = b.length;
  const dp = Array.from({ length: m + 1 }, () => new Array(n + 1).fill(0));
  for (let i = 0; i <= m; i++) dp[i][0] = i;
  for (let j = 0; j <= n; j++) dp[0][j] = j;
  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      if (a[i - 1] === b[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
  }
  return dp[m][n];
}

/**
 * DETEKSI SATU KATA APAKAH ITU NAMA BULAN (INGGRIS/INDONESIA, LENGKAP/SINGKATAN),
 * TOLERAN TERHADAP TYPO. Return index bulan (0 = Januari .. 11 = Desember)
 * atau null kalau tidak ada kandidat yang cukup mirip.
 */
function detectMonthIndex(word) {
  if (!word) return null;
  const clean = word
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^a-z]/g, "");
  if (!clean) return null;

  let bestIndex = null;
  let bestDistance = Infinity;

  MONTH_ALIASES.forEach((aliases, idx) => {
    aliases.forEach((alias) => {
      // COCOK PERSIS ATAU SALAH SATU MERUPAKAN AWALAN DARI YANG LAIN -> LANGSUNG DIANGGAP COCOK
      if (
        clean === alias ||
        clean.startsWith(alias) ||
        alias.startsWith(clean)
      ) {
        if (bestDistance > 0) {
          bestIndex = idx;
          bestDistance = 0;
        }
        return;
      }
      const dist = levenshteinDistance(clean, alias);
      // MAKIN PANJANG KATANYA, MAKIN BANYAK TYPO YANG DITOLERIR (TAPI TETAP KETAT
      // SUPAYA KATA UMUM BAHASA INGGRIS/INDONESIA TIDAK KETEBAK JADI NAMA BULAN)
      const tolerance = alias.length <= 5 ? 1 : 2;
      if (dist <= tolerance && dist < bestDistance) {
        bestIndex = idx;
        bestDistance = dist;
      }
    });
  });

  return bestIndex;
}

// CARI NAMA BULAN (INGGRIS/INDONESIA, TERMASUK TYPO) DI DALAM SEBUAH STRING BEBAS
function findMonthInText(str) {
  if (!str) return null;
  const words = str
    .toString()
    .split(/[^a-zA-Z]+/)
    .filter(Boolean);
  for (const word of words) {
    if (word.length < 3) continue; // lewati sisa kata yang kependekan untuk dicocokkan
    const idx = detectMonthIndex(word);
    if (idx !== null) return idx;
  }
  return null;
}

// TERJEMAHKAN INDEX BULAN KE NAMA INDONESIA, DIPAKAI KHUSUS UNTUK PENAMAAN FILE DOWNLOAD.
// KALAU SUMBERNYA NAMA BULAN INGGRIS (ATAU TYPO), INDEX-NYA SUDAH DIDETEKSI LEBIH DULU
// LEWAT detectMonthIndex/findMonthInText, JADI DI SINI TINGGAL DITERJEMAHKAN SAJA.
function toIndonesianMonthName(monthIndex) {
  if (monthIndex === null || monthIndex === undefined) return "Unknown";
  if (monthIndex < 0 || monthIndex > 11) return "Unknown";
  return MONTH_NAMES_ID[monthIndex];
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

// AMBIL TANGGAL & JAM:MENIT:DETIK SAAT INI, UNTUK BIKIN NAMA FILE SELALU BEDA
function getTimestampParts() {
  const now = new Date();
  const dateStr = `${pad2(now.getDate())}-${pad2(now.getMonth() + 1)}-${now.getFullYear()}`;
  const timeStr = `${pad2(now.getHours())}-${pad2(now.getMinutes())}-${pad2(now.getSeconds())}`;
  return { dateStr, timeStr };
}

// PARSE NILAI KOLOM "Tanggal Pasang" MENJADI { month, year }
function parseTanggalPasang(value) {
  if (value === undefined || value === null) return null;
  const str = String(value).trim();
  if (!str) return null;

  let match = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (match) {
    let [, d, m, y] = match;
    m = parseInt(m, 10);
    y = parseInt(y, 10);
    if (y < 100) y += 2000;
    if (m >= 1 && m <= 12) return { month: m - 1, year: y };
  }

  match = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (match) {
    const [, y, m] = match;
    const mm = parseInt(m, 10);
    if (mm >= 1 && mm <= 12) return { month: mm - 1, year: parseInt(y, 10) };
  }

  // FORMAT DENGAN NAMA BULAN, MIS. "12 Juli 2024", "July 2024", ATAU TYPO SEPERTI
  // "Jully 2024" / "Agustuz 2024" -> TETAP TERDETEKSI SEBELUM JATUH KE DATE BAWAAN JS
  const yearMatch = str.match(/(\d{4})/);
  const monthIdxFromText = findMonthInText(str);
  if (monthIdxFromText !== null && yearMatch) {
    return { month: monthIdxFromText, year: parseInt(yearMatch[1], 10) };
  }

  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    return { month: d.getMonth(), year: d.getFullYear() };
  }

  return null;
}

// CARI INDEX KOLOM DI HEADER BERDASARKAN NAMA, TIDAK PEDULI BESAR/KECIL HURUF
function findColumnIndex(header, candidateNames) {
  const normalizedHeader = header.map((h) =>
    (h || "").toString().trim().toLowerCase(),
  );
  for (const candidate of candidateNames) {
    const idx = normalizedHeader.indexOf(candidate.toLowerCase());
    if (idx !== -1) return idx;
  }
  return -1;
}

// TERIMA LINK LENGKAP ATAU ID SAJA, LALU AMBIL ID-NYA SAJA
function extractSheetId(input) {
  const trimmed = String(input).trim();
  const urlMatch = trimmed.match(/\/d\/([a-zA-Z0-9-_]+)/);
  if (urlMatch) return urlMatch[1];
  return trimmed;
}

let _authClient = null;
function getAuth() {
  if (!_authClient) {
    _authClient = new google.auth.GoogleAuth({
      keyFile: path.join(__dirname, "..", "credentials.json"),
      scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
    });
  }
  return _authClient;
}

async function getSheetsClient() {
  const client = await getAuth().getClient();
  return google.sheets({ version: "v4", auth: client });
}

// AMBIL DAFTAR NAMA SHEET DARI SPREADSHEET, DIPAKAI UNTUK MENGISI DROPDOWN DI WEB
async function listSheets(sheetUrlOrId) {
  const SHEET_ID = extractSheetId(sheetUrlOrId);
  const sheets = await getSheetsClient();
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  return meta.data.sheets.map((s) => s.properties.title);
}

/**
 * Jalankan rekap CITY untuk range sheet [startSheetName..endSheetName] (inklusif).
 * Kalau startSheetName/endSheetName kosong -> pakai sheet paling awal/akhir.
 *
 * options:
 *  - sheetUrlOrId: string (wajib)
 *  - startSheetName: string|null
 *  - endSheetName: string|null
 *  - onLog: function(string) -> dipanggil untuk setiap baris log (opsional)
 *
 * return: { buffer, fileName, summary }
 */
// Helper delay/sleep untuk memberi jeda antar request
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

async function exportRekapPerCity({
  sheetUrlOrId,
  startSheetName,
  endSheetName,
  onLog,
}) {
  const log = onLog || (() => {});
  const SHEET_ID = extractSheetId(sheetUrlOrId);
  const sheets = await getSheetsClient();

  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const allSheets = meta.data.sheets;

  if (!allSheets || allSheets.length === 0) {
    throw new Error("Spreadsheet tidak punya sheet sama sekali.");
  }

  let startIndex = 0;
  let endIndex = allSheets.length - 1;

  if (startSheetName && startSheetName.trim() !== "") {
    const idx = allSheets.findIndex(
      (s) =>
        s.properties.title.trim().toLowerCase() ===
        startSheetName.trim().toLowerCase(),
    );
    if (idx !== -1) startIndex = idx;
    else
      log(`Sheet awal "${startSheetName}" tidak ditemukan, mulai dari awal.`);
  }

  if (endSheetName && endSheetName.trim() !== "") {
    const idx = allSheets.findIndex(
      (s) =>
        s.properties.title.trim().toLowerCase() ===
        endSheetName.trim().toLowerCase(),
    );
    if (idx !== -1) endIndex = idx;
    else
      log(
        `Sheet akhir "${endSheetName}" tidak ditemukan, sampai sheet paling akhir.`,
      );
  }

  if (endIndex < startIndex) {
    [startIndex, endIndex] = [endIndex, startIndex];
  }

  log(
    `Memproses sheet dari "${allSheets[startIndex].properties.title}" sampai "${allSheets[endIndex].properties.title}"`,
  );

  const cityData = {};
  const monthTally = {};
  const sheetsTanpaKolomCity = [];
  const sheetsTanpaKolomMsisdn = [];
  const barisTanpaMsisdnBySheet = {};
  const sheetSummaries = [];

  let lastProcessedIndex = -1;
  let hitRateLimit = false;

  // LOOP DENGAN PENANGANAN ERROR PER SHEET
  for (let sIndex = startIndex; sIndex <= endIndex; sIndex++) {
    const s = allSheets[sIndex];
    const sheetName = s.properties.title;

    log(`Membaca sheet (${sIndex + 1}/${allSheets.length}): ${sheetName}`);

    let rows = [];
    try {
      // Jeda 400ms setiap request agar tidak melampaui limit Google Sheets API (100 request / 100 detik)
      await sleep(400);

      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${sheetName}!A:Z`,
      });

      rows = res.data.values || [];
    } catch (err) {
      // CEK APAKAH ERROR RATE LIMIT (429 / QUOTA EXCEEDED) ATAU ERROR LAIN
      const isQuotaError =
        err.status === 429 ||
        (err.message && err.message.includes("Quota exceeded"));

      if (isQuotaError || err.code === 429) {
        log(
          `⚠️ TERKENA LIMIT GOOGLE API saat membaca sheet "${sheetName}". Proses membaca dihentikan sementara.`,
        );
      } else {
        log(
          `⚠️ Error tidak dikenal pada sheet "${sheetName}": ${err.message}. Dihentikan untuk amannya.`,
        );
      }

      hitRateLimit = true;
      // JANGAN SIMPAN SHEET INI! LANGSUNG BREAK AGAR PROGRESS SEMENTARA BISA DISIMPAN
      break;
    }

    // Jika berhasil baca data, proses seperti biasa
    lastProcessedIndex = sIndex;

    if (rows.length < 2) {
      log(`  -> Sheet kosong, dilewati.`);
      continue;
    }

    const header = rows[0];
    const cityCol = findColumnIndex(header, ["city", "kota"]);
    const msisdnCol = findColumnIndex(header, ["msisdn"]);
    const tanggalPasangCol = findColumnIndex(header, ["tanggal pasang"]);

    if (cityCol === -1) {
      log(`  -> Kolom "CITY" tidak ditemukan, dilewati.`);
      sheetsTanpaKolomCity.push(sheetName);
      continue;
    }

    if (msisdnCol === -1) {
      log(`  -> Kolom "MSISDN" tidak ditemukan, dilewati.`);
      sheetsTanpaKolomMsisdn.push(sheetName);
      continue;
    }

    const ringkasan = {};
    let lastTableRowIndex = 0;
    for (let i = rows.length - 1; i >= 1; i--) {
      const row = rows[i];
      const cityCheck = (row[cityCol] || "").toString().trim();
      const msisdnCheck = (row[msisdnCol] || "").toString().trim();
      if (cityCheck || msisdnCheck) {
        lastTableRowIndex = i;
        break;
      }
    }

    for (let i = 1; i <= lastTableRowIndex; i++) {
      const row = rows[i];
      const city = (row[cityCol] || "").toString().trim();
      const msisdn = (row[msisdnCol] || "").toString().trim();

      if (city && !msisdn) {
        if (!barisTanpaMsisdnBySheet[sheetName])
          barisTanpaMsisdnBySheet[sheetName] = [];
        barisTanpaMsisdnBySheet[sheetName].push(i + 1);
        continue;
      }

      if (!city) continue;

      if (!cityData[city]) cityData[city] = { columns: [], records: [] };
      const entry = cityData[city];

      header.forEach((col) => {
        if (col && !entry.columns.includes(col)) entry.columns.push(col);
      });

      const rec = {};
      header.forEach((col, idx) => {
        if (!col) return;
        rec[col] = row[idx] !== undefined ? row[idx] : "";
      });
      rec.__source_sheet = sheetName;

      entry.records.push(rec);
      ringkasan[city] = (ringkasan[city] || 0) + 1;

      if (tanggalPasangCol !== -1) {
        const parsed = parseTanggalPasang(row[tanggalPasangCol]);
        if (parsed) {
          const key = `${parsed.year}-${parsed.month}`;
          if (!monthTally[key]) {
            monthTally[key] = {
              month: parsed.month,
              year: parsed.year,
              count: 0,
            };
          }
          monthTally[key].count += 1;
        }
      }
    }

    log(`  -> Ringkasan: ${JSON.stringify(ringkasan)}`);
    sheetSummaries.push({ sheetName, ringkasan });
  }

  // Jika tidak ada satu pun sheet yang berhasil di-download sebelum terkena limit
  if (lastProcessedIndex === -1) {
    throw new Error(
      "Terkena limit Google API sejak sheet pertama. Silakan tunggu 1-2 menit lalu coba lagi.",
    );
  }

  log("Membuat file Excel dari sheet yang berhasil di-load...");

  const workbook = new ExcelJS.Workbook();
  const HEADER_COLOR = "FF1F4E78";
  const STRIPE_COLOR = "FFF2F6FA";
  const BORDER_COLOR = "FFD9D9D9";

  const cityWorksheetSummaries = [];

  for (const city of Object.keys(cityData)) {
    const { columns, records } = cityData[city];
    const header = [...columns, "Source Sheet"];
    const dataRows = records.map((rec) =>
      columns
        .map((c) => (rec[c] !== undefined ? rec[c] : ""))
        .concat(rec.__source_sheet),
    );

    const ws = workbook.addWorksheet(city.substring(0, 31), {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    ws.addRow(header);
    dataRows.forEach((r) => ws.addRow(r));

    const headerRow = ws.getRow(1);
    headerRow.height = 22;
    headerRow.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: HEADER_COLOR },
      };
      cell.alignment = { vertical: "middle", horizontal: "center" };
    });

    for (let r = 1; r <= dataRows.length + 1; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= header.length; c++) {
        const cell = row.getCell(c);
        cell.border = {
          top: { style: "thin", color: { argb: BORDER_COLOR } },
          left: { style: "thin", color: { argb: BORDER_COLOR } },
          bottom: { style: "thin", color: { argb: BORDER_COLOR } },
          right: { style: "thin", color: { argb: BORDER_COLOR } },
        };
      }
      if (r > 1 && r % 2 === 0) {
        row.eachCell((cell) => {
          cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: STRIPE_COLOR },
          };
        });
      }
    }

    header.forEach((h, idx) => {
      let maxLen = String(h).length;
      dataRows.forEach((r) => {
        const val =
          r[idx] !== undefined && r[idx] !== null ? String(r[idx]) : "";
        if (val.length > maxLen) maxLen = val.length;
      });
      ws.getColumn(idx + 1).width = Math.min(Math.max(maxLen + 2, 10), 45);
    });

    ws.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: header.length },
    };
    cityWorksheetSummaries.push({
      city,
      rows: dataRows.length,
      columns: header.length,
    });
  }

  let best = null;
  for (const key of Object.keys(monthTally)) {
    const cur = monthTally[key];
    if (!best || cur.count > best.count) best = cur;
  }
  const dominantMonthName = best
    ? toIndonesianMonthName(best.month)
    : "Unknown";
  const { dateStr, timeStr } = getTimestampParts();

  // Penanda jika file terpotong karena rate limit
  const statusSuffix = hitRateLimit ? "- PARTIAL" : "";
  const fileName = `Burn Kuota ${dominantMonthName} create-file ${dateStr} ${timeStr}${statusSuffix}.xlsx`;

  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  }
  const OUTPUT_FILE = path.join(OUTPUT_DIR, fileName);

  const buffer = await workbook.xlsx.writeBuffer();
  fs.writeFileSync(OUTPUT_FILE, buffer);

  const isCompleted = !hitRateLimit && lastProcessedIndex === endIndex;
  const nextSheetToProcess = isCompleted
    ? null
    : allSheets[lastProcessedIndex + 1]?.properties.title;

  log(`--------------------------------------------------`);
  log(`Proses Selesai Sementara.`);
  log(
    `Sheet Terakhir Berhasil: "${allSheets[lastProcessedIndex].properties.title}"`,
  );
  if (!isCompleted) {
    log(
      `⚠️ Terhenti karena limit. Sheet berikutnya untuk dilanjutkan: "${nextSheetToProcess}"`,
    );
  }
  log(`File tersimpan di: ${OUTPUT_FILE}`);

  const summary = {
    fileName,
    savedTo: OUTPUT_FILE,
    range: {
      startSheet: allSheets[startIndex].properties.title,
      lastSuccessfulSheet: allSheets[lastProcessedIndex].properties.title,
      targetEndSheet: allSheets[endIndex].properties.title,
    },
    isCompleted,
    nextSheetToProcess, // ACUAN UNTUK RESUME
    totalCity: Object.keys(cityData).length,
    dominantMonth: dominantMonthName,
    cityWorksheetSummaries,
    sheetsTanpaKolomCity,
    sheetsTanpaKolomMsisdn,
    barisTanpaMsisdnBySheet,
  };

  return { buffer, fileName, summary };
}

module.exports = {
  listSheets,
  exportRekapPerCity,
  extractSheetId,
};
