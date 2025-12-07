const express = require("express");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = 3000;

let holidayCache = [];
let lastDiag = {}; // expose loader diagnostics via /api/sensex/debug

/* ---------- Date helpers ---------- */
function formatDateIST(date) {
  const tzOffsetMs = date.getTimezoneOffset() * 60000;
  const local = new Date(date.getTime() - tzOffsetMs);
  return local.toISOString().slice(0, 10);
}
function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}
function isNonTradingDay(dateStr, holidays) {
  const d = new Date(dateStr);
  return isWeekend(d) || holidays.includes(dateStr);
}
function shiftToPreviousTradingDay(dateStr, holidays) {
  let d = new Date(dateStr);
  while (true) {
    d.setDate(d.getDate() - 1);
    const candidate = formatDateIST(d);
    if (!isNonTradingDay(candidate, holidays)) return candidate;
  }
}

/* ---------- Expiry generator ---------- */
function generateExpiries(startDate, monthsAhead = 3) {
  const expiries = [];
  const start = new Date(startDate);

  for (let m = 0; m < monthsAhead; m++) {
    const firstOfMonth = new Date(start.getFullYear(), start.getMonth() + m, 1);
    const month = firstOfMonth.getMonth();
    for (let d = new Date(firstOfMonth); d.getMonth() === month; d.setDate(d.getDate() + 1)) {
      if (d.getDay() === 4) expiries.push(formatDateIST(new Date(d))); // Thursday
    }
  }
  return Array.from(new Set(expiries)).sort();
}

/* ---------- Robust parsing ---------- */
function fromExcelSerial(n) {
  const base = new Date(Date.UTC(1899, 11, 30));
  return new Date(base.getTime() + n * 86400000);
}
function parseAnyDate(raw) {
  if (raw == null) return null;
  if (typeof raw === "number" && isFinite(raw)) {
    const d = fromExcelSerial(raw);
    return isNaN(d) ? null : d;
  }
  let s = String(raw).trim();
  if (!s) return null;
  s = s.replace(/,/g, " ").replace(/\s+/g, " ").trim();
  let d = new Date(s);
  if (!isNaN(d)) return d;
  const ddMmmYy = /^(\d{1,2})-(\w{3})-(\d{2,4})$/i;
  const m = s.match(ddMmmYy);
  if (m) {
    const day = parseInt(m[1], 10);
    const monStr = m[2].toUpperCase();
    const yr = parseInt(m[3], 10);
    const year = yr < 100 ? 2000 + yr : yr;
    const months = { JAN:0,FEB:1,MAR:2,APR:3,MAY:4,JUN:5,JUL:6,AUG:7,SEP:8,OCT:9,NOV:10,DEC:11 };
    if (months[monStr] != null) {
      d = new Date(year, months[monStr], day);
      if (!isNaN(d)) return d;
    }
  }
  const mw = s.match(/^([A-Za-z]+)\s+(\d{1,2})\s+(\d{2,4})$/);
  if (mw) {
    d = new Date(`${mw[1]} ${mw[2]} ${mw[3]}`);
    if (!isNaN(d)) return d;
  }
  return null;
}

/* ---------- File detection ---------- */
function findHolidayFileForYear(year) {
  const cwd = process.cwd();
  const candidates = fs.readdirSync(cwd)
    .filter(f => /^BSE_Holidays_\d{4}\.(xlsx|csv)$/i.test(f))
    .map(f => ({ f, y: parseInt((f.match(/\d{4}/) || [year])[0], 10) }))
    .filter(o => o.y === year);
  if (candidates.length === 0) {
    const expectedXlsx = path.join(cwd, `BSE_Holidays_${year}.xlsx`);
    const expectedCsv = path.join(cwd, `BSE_Holidays_${year}.csv`);
    return fs.existsSync(expectedXlsx) ? expectedXlsx :
           fs.existsSync(expectedCsv) ? expectedCsv : null;
  }
  return path.join(candidates[0].f);
}

/* ---------- Loaders with diagnostics ---------- */
function loadHolidaysFromXlsx(filePath) {
  const workbook = xlsx.readFile(filePath, { cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: true });

  const maxCols = Math.max(...rows.map(r => r.length));
  let bestCol = 0, bestCount = 0;
  for (let col = 0; col < maxCols; col++) {
    let count = 0;
    for (const row of rows) if (parseAnyDate(row[col])) count++;
    if (count > bestCount) { bestCount = count; bestCol = col; }
  }

  const holidays = [];
  const samples = [];
  rows.forEach((row, idx) => {
    const val = row[bestCol];
    const parsed = parseAnyDate(val);
    if (parsed) {
      const normalized = formatDateIST(parsed);
      holidays.push(normalized);
      if (samples.length < 8) samples.push({ row: idx + 1, raw: val, out: normalized });
    }
  });

  lastDiag = {
    source: "xlsx",
    filePath,
    sheetName,
    rowCount: rows.length,
    bestCol,
    bestCount,
    samples
  };
  return Array.from(new Set(holidays)).sort();
}

function loadHolidaysFromCsv(filePath) {
  const raw = fs.readFileSync(filePath, "utf8");
  const lines = raw.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const holidays = [];
  const samples = [];
  lines.forEach((line, idx) => {
    const parts = line.split(",").map(s => s.trim()).filter(Boolean);
    // Try each column in the line
    let normalized = null;
    for (const p of parts) {
      const d = parseAnyDate(p);
      if (d) { normalized = formatDateIST(d); break; }
    }
    if (normalized) {
      holidays.push(normalized);
      if (samples.length < 8) samples.push({ line: idx + 1, raw: line, out: normalized });
    }
  });

  lastDiag = {
    source: "csv",
    filePath,
    lineCount: lines.length,
    samples
  };
  return Array.from(new Set(holidays)).sort();
}

/* ---------- Hard fallback (your 2025 list) ---------- */
function fallbackHolidays(year) {
  const known2025 = [
    "February 26,2025","14-Mar-2025","March 31,2025","April 10,2025","April 14,2025",
    "April 18,2025","May 01,2025","August 15,2025","August 27,2025","October 02,2025",
    "October 21,2025","October 22,2025","November 05,2025","December 25,2025"
  ];
  const normalized = known2025
    .map(s => parseAnyDate(s))
    .filter(Boolean)
    .map(d => formatDateIST(d));
  lastDiag = { source: "fallback", count: normalized.length };
  return Array.from(new Set(normalized)).sort();
}

/* ---------- Refresh ---------- */
function refreshHolidayCache() {
  try {
    const currentYear = new Date().getFullYear();
    const filePath = findHolidayFileForYear(currentYear);

    let loaded = [];
    if (filePath && /\.xlsx$/i.test(filePath)) {
      loaded = loadHolidaysFromXlsx(filePath);
      console.log("Loaded holidays from XLSX:", loaded.length, "entries");
    } else if (filePath && /\.csv$/i.test(filePath)) {
      loaded = loadHolidaysFromCsv(filePath);
      console.log("Loaded holidays from CSV:", loaded.length, "entries");
    } else {
      console.warn("No holiday file found for", currentYear, "→ using fallback.");
      loaded = fallbackHolidays(currentYear);
    }

    if (!loaded || loaded.length === 0) {
      console.warn("Holiday list empty after load → using fallback.");
      loaded = fallbackHolidays(currentYear);
    }

    holidayCache = loaded;
  } catch (err) {
    console.error("Failed to refresh holiday cache:", err);
    holidayCache = fallbackHolidays(new Date().getFullYear());
  }
}

// Startup load
refreshHolidayCache();

/* ---------- Routes ---------- */
app.get("/api/sensex/holidays", (req, res) => {
  res.json({ holidays: holidayCache });
});

app.get("/api/sensex/debug", (req, res) => {
  res.json({ diagnostics: lastDiag });
});

app.get("/api/sensex/reload-holidays", (req, res) => {
  refreshHolidayCache();
  res.json({ status: "reloaded", holidays: holidayCache, diagnostics: lastDiag });
});

app.get("/api/sensex/expiries", (req, res) => {
  try {
    const monthsAhead = Number(req.query.monthsAhead || 3);
    const startDate = req.query.startDate || formatDateIST(new Date());
    const expiries = generateExpiries(startDate, monthsAhead);
    const adjusted = expiries.map(dateStr =>
      isNonTradingDay(dateStr, holidayCache)
        ? shiftToPreviousTradingDay(dateStr, holidayCache)
        : dateStr
    );
    res.json({ expiries: adjusted, holidays: holidayCache });
  } catch (err) {
    console.error("Error generating Sensex expiries:", err);
    res.status(500).json({ error: "Failed to generate Sensex expiries" });
  }
});

app.listen(PORT, () => {
  console.log(`Sensex test server running at http://localhost:${PORT}`);
});