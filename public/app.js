const sheetUrlInput = document.getElementById("sheetUrl");
const btnLoadSheets = document.getElementById("btnLoadSheets");
const stepRange = document.getElementById("step-range");
const startSheetSel = document.getElementById("startSheet");
const endSheetSel = document.getElementById("endSheet");
const tapePreview = document.getElementById("tapePreview");
const btnExport = document.getElementById("btnExport");
const stepLog = document.getElementById("step-log");
const summaryBox = document.getElementById("summaryBox");
const logBox = document.getElementById("logBox");
const statusEl = document.getElementById("status");

// ELEMEN UNTUK MODE CEK KUOTA
const modeTabs = document.getElementById("modeTabs");
const stepKuotaRange = document.getElementById("step-kuota-range");
const startSheetKuotaSel = document.getElementById("startSheetKuota");
const endSheetKuotaSel = document.getElementById("endSheetKuota");
const btnCekKuota = document.getElementById("btnCekKuota");
const stepKuotaLog = document.getElementById("step-kuota-log");
const kuotaSummaryBox = document.getElementById("kuotaSummaryBox");
const kuotaLogBox = document.getElementById("kuotaLogBox");

let sheetNames = [];
let currentMode = "rekap";

// PINDAH TAB: TAMPILKAN PANEL SESUAI MODE YANG DIPILIH (REKAP / KUOTA)
function setMode(mode) {
  currentMode = mode;

  document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.classList.toggle("active", btn.dataset.mode === mode);
  });

  document.querySelectorAll(".mode-panel").forEach((panel) => {
    const isThisMode = panel.dataset.mode === mode;
    // JANGAN PAKSA MUNCUL KALAU MEMANG BELUM ADA DAFTAR SHEET (SEBELUM "Muat daftar sheet")
    if (!isThisMode) {
      panel.hidden = true;
    } else if (sheetNames.length > 0) {
      panel.hidden = panel.id === "step-log" || panel.id === "step-kuota-log" ? panel.hidden : false;
    }
  });
}

modeTabs.addEventListener("click", (e) => {
  const btn = e.target.closest(".tab-btn");
  if (!btn) return;
  setMode(btn.dataset.mode);
});

function showStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle("error", isError);
  statusEl.classList.add("show");
  if (!isError) {
    clearTimeout(showStatus._t);
    showStatus._t = setTimeout(() => statusEl.classList.remove("show"), 3500);
  }
}

// ISI OTOMATIS DARI .env KALAU ADA, SUPAYA USER TIDAK PERLU KETIK ULANG
async function loadDefaultUrl() {
  try {
    const res = await fetch("/api/default-sheet-url");
    const data = await res.json();
    if (data.sheetUrl) sheetUrlInput.value = data.sheetUrl;
  } catch (e) {
    // diamkan, bukan hal fatal
  }
}
loadDefaultUrl();

function renderTapePreview() {
  tapePreview.innerHTML = "";
  const startIdx = startSheetSel.selectedIndex;
  const endIdx = endSheetSel.selectedIndex;
  const lo = Math.min(startIdx, endIdx);
  const hi = Math.max(startIdx, endIdx);

  sheetNames.forEach((name, idx) => {
    const seg = document.createElement("span");
    seg.className = "seg" + (idx >= lo && idx <= hi ? " active" : "");
    seg.textContent = name;
    tapePreview.appendChild(seg);
  });
}

function fillSelect(sel, names) {
  sel.innerHTML = "";
  names.forEach((name) => {
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
  });
}

btnLoadSheets.addEventListener("click", async () => {
  const sheetUrl = sheetUrlInput.value.trim();
  btnLoadSheets.disabled = true;
  btnLoadSheets.textContent = "Memuat...";
  showStatus("Mengambil daftar sheet...");

  try {
    const res = await fetch("/api/list-sheets", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheetUrl }),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Gagal mengambil daftar sheet.");

    sheetNames = data.sheets;

    // ISI DROPDOWN UNTUK MODE REKAP
    fillSelect(startSheetSel, sheetNames);
    fillSelect(endSheetSel, sheetNames);
    endSheetSel.selectedIndex = sheetNames.length - 1;

    // ISI DROPDOWN UNTUK MODE CEK KUOTA (PAKAI DAFTAR SHEET YANG SAMA)
    fillSelect(startSheetKuotaSel, sheetNames);
    fillSelect(endSheetKuotaSel, sheetNames);
    endSheetKuotaSel.selectedIndex = sheetNames.length - 1;

    modeTabs.hidden = false;
    setMode(currentMode);
    renderTapePreview();
    showStatus(`${sheetNames.length} sheet ditemukan.`);
  } catch (err) {
    showStatus(err.message, true);
  } finally {
    btnLoadSheets.disabled = false;
    btnLoadSheets.textContent = "Muat daftar sheet";
  }
});

startSheetSel.addEventListener("change", renderTapePreview);
endSheetSel.addEventListener("change", renderTapePreview);

btnExport.addEventListener("click", async () => {
  const sheetUrl = sheetUrlInput.value.trim();
  const startSheet = startSheetSel.value;
  const endSheet = endSheetSel.value;

  btnExport.disabled = true;
  btnExport.textContent = "Memproses...";
  showStatus("Membaca sheet & menyusun Excel, mohon tunggu...");

  try {
    const res = await fetch("/api/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheetUrl, startSheet, endSheet }),
    });

    if (!res.ok) {
      const data = await res.json();
      throw new Error(data.error || "Gagal membuat rekap.");
    }

    const encodedSummary = res.headers.get("X-Rekap-Summary");
    const blob = await res.blob();

    // TRIGGER DOWNLOAD FILE .xlsx DI BROWSER
    const disposition = res.headers.get("Content-Disposition") || "";
    const match = disposition.match(/filename="(.+)"/);
    const fileName = match ? match[1] : "rekap.xlsx";

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    if (encodedSummary) {
      const summary = JSON.parse(atob(encodedSummary));
      renderSummary(summary);
    }

    stepLog.hidden = false;
    showStatus("Selesai! File sudah ter-download.");
  } catch (err) {
    showStatus(err.message, true);
  } finally {
    btnExport.disabled = false;
    btnExport.textContent = "Export ke Excel";
  }
});

// MODE CEK KUOTA: JALANKAN LEWAT /api/cek-kuota, HASIL LANGSUNG TERTULIS KE SPREADSHEET
btnCekKuota.addEventListener("click", async () => {
  const sheetUrl = sheetUrlInput.value.trim();
  const startSheet = startSheetKuotaSel.value;
  const endSheet = endSheetKuotaSel.value;

  btnCekKuota.disabled = true;
  btnCekKuota.textContent = "Memproses...";
  showStatus("Membuka tiap link & membaca sisa kuota, mohon tunggu (bisa lama)...");

  try {
    const res = await fetch("/api/cek-kuota", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ sheetUrl, startSheet, endSheet }),
    });

    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Gagal cek kuota.");

    renderKuotaSummary(data.results, data.logs);
    stepKuotaLog.hidden = false;
    showStatus("Selesai! Kolom Remaining Quota sudah diperbarui di spreadsheet.");
  } catch (err) {
    showStatus(err.message, true);
  } finally {
    btnCekKuota.disabled = false;
    btnCekKuota.textContent = "Jalankan Cek Kuota";
  }
});

function renderKuotaSummary(results, logs) {
  const rows = (results || []).map((r) => {
    const statusLabel =
      r.status === "sukses"
        ? `${r.totalRow} baris ter-update`
        : r.status === "kosong"
        ? "sheet kosong, dilewati"
        : "kolom tidak lengkap, dilewati";
    return `<div class="row"><span>${r.sheetName}</span><span>${statusLabel}</span></div>`;
  });

  kuotaSummaryBox.innerHTML = rows.join("");
  kuotaLogBox.textContent = (logs || []).join("\n");
}

function renderSummary(summary) {
  const rows = [
    ["File", summary.fileName],
    ["Rentang sheet", `${summary.range.startSheet} → ${summary.range.endSheet}`],
    ["Total CITY terdeteksi", summary.totalCity],
    ["Bulan dominan", summary.dominantMonth],
    ["Tersimpan di server", summary.savedTo],
  ];

  let html = rows
    .map(([label, val]) => `<div class="row"><span>${label}</span><span>${val}</span></div>`)
    .join("");

  if (summary.sheetsTanpaKolomCity && summary.sheetsTanpaKolomCity.length) {
    html += `<p class="hint">Sheet tanpa kolom CITY (dilewati):</p><ul>${summary.sheetsTanpaKolomCity
      .map((s) => `<li>${s}</li>`)
      .join("")}</ul>`;
  }

  if (summary.sheetsTanpaKolomMsisdn && summary.sheetsTanpaKolomMsisdn.length) {
    html += `<p class="hint">Sheet tanpa kolom MSISDN (dilewati):</p><ul>${summary.sheetsTanpaKolomMsisdn
      .map((s) => `<li>${s}</li>`)
      .join("")}</ul>`;
  }

  const kosongEntries = Object.entries(summary.barisTanpaMsisdnBySheet || {});
  if (kosongEntries.length) {
    html += `<p class="hint">Baris dengan CITY ada tapi MSISDN kosong:</p><ul>${kosongEntries
      .map(([sheet, rowsArr]) => `<li>${sheet}: baris ke-${rowsArr.join(", ke-")}</li>`)
      .join("")}</ul>`;
  }

  summaryBox.innerHTML = html;
  logBox.textContent = (summary.logs || []).join("\n");
}