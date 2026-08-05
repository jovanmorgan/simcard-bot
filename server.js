require("dotenv").config();

const express = require("express");
const path = require("path");
const { listSheets, exportRekapPerCity } = require("./lib/rekap");
const { cekKuotaPerCity } = require("./lib/cekKuota");

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

function resolveSheetUrl(bodyUrl) {
  if (bodyUrl && bodyUrl.trim() !== "") return bodyUrl.trim();
  if (process.env.SHEET_URL && process.env.SHEET_URL.trim() !== "") return process.env.SHEET_URL.trim();
  return null;
}

// KIRIM URL DEFAULT DARI .env (KALAU ADA) SUPAYA FORM DI WEB BISA TERISI OTOMATIS
app.get("/api/default-sheet-url", (req, res) => {
  res.json({ sheetUrl: process.env.SHEET_URL || "" });
});

// AMBIL DAFTAR NAMA SHEET, DIPAKAI UNTUK MENGISI DUA DROPDOWN (DARI - SAMPAI)
app.post("/api/list-sheets", async (req, res) => {
  try {
    const sheetUrl = resolveSheetUrl(req.body.sheetUrl);
    if (!sheetUrl) {
      return res.status(400).json({ error: "Link/ID spreadsheet wajib diisi." });
    }
    const sheetNames = await listSheets(sheetUrl);
    res.json({ sheets: sheetNames });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Gagal mengambil daftar sheet." });
  }
});

// JALANKAN REKAP UNTUK RANGE SHEET TERPILIH, LALU KIRIM FILE .xlsx SEBAGAI DOWNLOAD
app.post("/api/export", async (req, res) => {
  try {
    const sheetUrl = resolveSheetUrl(req.body.sheetUrl);
    const { startSheet, endSheet } = req.body;

    if (!sheetUrl) {
      return res.status(400).json({ error: "Link/ID spreadsheet wajib diisi." });
    }

    const logs = [];
    const result = await exportRekapPerCity({
      sheetUrlOrId: sheetUrl,
      startSheetName: startSheet || null,
      endSheetName: endSheet || null,
      onLog: (msg) => {
        console.log(msg);
        logs.push(msg);
      },
    });

    const summaryPayload = { ...result.summary, logs };
    const encodedSummary = Buffer.from(JSON.stringify(summaryPayload)).toString("base64");

    res.setHeader("X-Rekap-Summary", encodedSummary);
    res.setHeader("Access-Control-Expose-Headers", "X-Rekap-Summary");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${result.fileName}"`);
    res.send(result.buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Gagal membuat rekap." });
  }
});

// JALANKAN CEK KUOTA UNTUK RANGE SHEET TERPILIH, LALU UPDATE LANGSUNG KE GOOGLE SHEET
// (KOLOM "Remaining Quota" DI-ISI OTOMATIS BERDASARKAN "Link Cek Kuota")
app.post("/api/cek-kuota", async (req, res) => {
  try {
    const sheetUrl = resolveSheetUrl(req.body.sheetUrl);
    const { startSheet, endSheet } = req.body;

    if (!sheetUrl) {
      return res.status(400).json({ error: "Link/ID spreadsheet wajib diisi." });
    }

    const logs = [];
    const { results } = await cekKuotaPerCity({
      sheetUrlOrId: sheetUrl,
      startSheetName: startSheet || null,
      endSheetName: endSheet || null,
      onLog: (msg) => {
        console.log(msg);
        logs.push(msg);
      },
    });

    res.json({ results, logs });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Gagal cek kuota." });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server rekap-data jalan di http://localhost:${PORT}`);
});