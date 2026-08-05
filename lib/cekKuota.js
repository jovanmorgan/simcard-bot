const { google } = require("googleapis");
const puppeteer = require("puppeteer");
const path = require("path");
const { extractSheetId } = require("./rekap");

// AUTH GOOGLE (PAKAI CREDENTIALS YANG SAMA DENGAN REKAP)
let _authClient = null;
function getAuth() {
  if (!_authClient) {
    _authClient = new google.auth.GoogleAuth({
      keyFile: path.join(__dirname, "..", "credentials.json"),
      scopes: ["https://www.googleapis.com/auth/spreadsheets"],
    });
  }
  return _authClient;
}

async function getSheetsClient() {
  const client = await getAuth().getClient();
  return google.sheets({ version: "v4", auth: client });
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// UBAH ANGKA KE HURUF KOLOM (A, B, C, ... Z, AA, AB, ...)
function columnLetter(col) {
  let temp = "";
  let letter = "";
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

// AMBIL ANGKA "Kuota Nasional" DARI ISI HALAMAN
async function getQuotaNasional(page) {
  let quota = "0 GB";

  const pageText = await page.evaluate(() => document.body.innerText);

  const nasionalRegex = /Kuota Nasional[\s\S]{0,100}?([0-9.,]+\s?(GB|MB|KB))/i;
  const nasionalMatch = pageText.match(nasionalRegex);

  if (nasionalMatch && nasionalMatch[1]) {
    quota = nasionalMatch[1].replace(/\s+/g, " ").trim();
  }

  return quota;
}

/**
 * Jalankan cek kuota untuk range sheet [startSheetName..endSheetName] (inklusif),
 * lalu tulis hasilnya ke kolom "Remaining Quota" di masing-masing sheet.
 * Kalau startSheetName/endSheetName kosong -> pakai sheet paling awal/akhir.
 *
 * options:
 *  - sheetUrlOrId: string (wajib, link lengkap ATAU ID spreadsheet)
 *  - startSheetName: string|null
 *  - endSheetName: string|null
 *  - onLog: function(string) -> dipanggil untuk setiap baris log (opsional)
 *
 * return: { results, hitRateLimit }
 */
async function cekKuotaPerCity({ sheetUrlOrId, startSheetName, endSheetName, onLog }) {
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
      (s) => s.properties.title.trim().toLowerCase() === startSheetName.trim().toLowerCase(),
    );
    if (idx !== -1) startIndex = idx;
    else log(`Sheet awal "${startSheetName}" tidak ditemukan, mulai dari awal.`);
  }

  if (endSheetName && endSheetName.trim() !== "") {
    const idx = allSheets.findIndex(
      (s) => s.properties.title.trim().toLowerCase() === endSheetName.trim().toLowerCase(),
    );
    if (idx !== -1) endIndex = idx;
    else log(`Sheet akhir "${endSheetName}" tidak ditemukan, sampai sheet paling akhir.`);
  }

  if (endIndex < startIndex) {
    [startIndex, endIndex] = [endIndex, startIndex];
  }

  log(`Memproses sheet dari "${allSheets[startIndex].properties.title}" sampai "${allSheets[endIndex].properties.title}"`);

  // BROWSER REALTIME, DIBUKA SEKALI DAN DIPAKAI UNTUK SEMUA SHEET
  const browser = await puppeteer.launch({
    headless: true,
    defaultViewport: null,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  const results = [];

  try {
    const page = await browser.newPage();

    for (let sIndex = startIndex; sIndex <= endIndex; sIndex++) {
      const s = allSheets[sIndex];
      const sheetName = s.properties.title;

      log(`Memproses sheet (${sIndex + 1}/${allSheets.length}): ${sheetName}`);

      const res = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: `${sheetName}!A:Z`,
      });

      const rows = res.data.values || [];

      if (rows.length < 2) {
        log("  -> Sheet kosong, dilewati.");
        results.push({ sheetName, status: "kosong", totalRow: 0 });
        continue;
      }

      const header = rows[0];
      const msisdnCol = header.indexOf("MSISDN");
      const linkCol = header.indexOf("Link Cek Kuota");
      const quotaCol = header.indexOf("Remaining Quota");

      if (msisdnCol === -1 || linkCol === -1 || quotaCol === -1) {
        log('  -> Kolom "MSISDN" / "Link Cek Kuota" / "Remaining Quota" tidak lengkap, dilewati.');
        results.push({ sheetName, status: "kolom-tidak-lengkap", totalRow: 0 });
        continue;
      }

      const output = [];

      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const msisdn = row[msisdnCol] || "";
        const url = row[linkCol] || "";

        if (!url) {
          output.push(["0 GB"]);
          log(`  ROW ${i + 1} | ${msisdn} | 0 GB (tanpa link)`);
          continue;
        }

        try {
          await page.goto(url, { waitUntil: "domcontentloaded", timeout: 60000 });

          // BERI WAKTU JAVASCRIPT WEBSITE
          await sleep(5000);

          let quota = await getQuotaNasional(page);

          if (quota === "0 GB") {
            log("  Kuota tidak ditemukan, reload halaman...");

            for (let retry = 1; retry <= 3; retry++) {
              await page.reload({ waitUntil: "domcontentloaded" });
              await sleep(5000);
              quota = await getQuotaNasional(page);

              if (quota !== "0 GB") {
                log(`  Kuota ditemukan pada retry ke-${retry}`);
                break;
              }
            }
          }

          log(`  ROW ${i + 1} | ${msisdn} | ${quota}`);
          output.push([quota]);
        } catch (err) {
          log(`  ROW ${i + 1} | ${msisdn} | 0 GB (error: ${err.message})`);
          output.push(["0 GB"]);
        }

        // DELAY ANTAR MSISDN
        await sleep(1000);
      }

      const colLetter = columnLetter(quotaCol + 1);

      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `${sheetName}!${colLetter}2:${colLetter}${output.length + 1}`,
        valueInputOption: "RAW",
        resource: { values: output },
      });

      log(`  -> SUCCESS UPDATE ${output.length} baris di sheet "${sheetName}"`);
      results.push({ sheetName, status: "sukses", totalRow: output.length });
    }
  } finally {
    // BROWSER SELALU DITUTUP SUPAYA TIDAK MENUMPUK DI SERVER
    await browser.close();
  }

  log("SEMUA SHEET SELESAI DIPROSES.");

  return { results };
}

module.exports = { cekKuotaPerCity };