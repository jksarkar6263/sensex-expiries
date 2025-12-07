const express = require("express");
const fs = require("fs");
const path = require("path");
const xlsx = require("xlsx");

const app = express();
const PORT = process.env.PORT || 3000;

let holidayCache = [];
let lastDiag = {};

/* ---------- Helpers ---------- */
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

/* ---------- Holiday loader ---------- */
function loadHolidayFile(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 });

  const holidays = [];
  rows.forEach(row => {
    const cell = row[0]; // only Column A
    if (!cell) return;

    let d;
    if (cell instanceof Date) {
      d = cell;
    } else if (typeof cell === "string") {
      d = new Date(cell.replace(/-/g, " "));
    } else if (typeof cell === "number") {
      const parsed = xlsx.SSF.parse_date_code(cell);
      if (parsed) {
        d = new Date(parsed.y, parsed.m - 1, parsed.d);
      }
    }

    if (d && !isNaN(d)) {
      holidays.push(d.toISOString().slice(0, 10));
    }
  });

  lastDiag = { source: "xlsx", filePath, count: holidays.length };
  return Array.from(new Set(holidays)).sort();
}

/* ---------- Fallback ---------- */
function fallbackHolidays(year) {
  if (year === 2025) {
    return [
      "2025-02-26","2025-03-14","2025-03-31","2025-04-10","2025-04-14","2025-04-18",
      "2025-05-01","2025-08-15","2025-08-27","2025-10-02","2025-10-21","2025-10-22",
      "2025-11-05","2025-12-25"
    ];
  }
  return [];
}

/* ---------- Refresh holiday cache ---------- */
function refreshHolidayCache() {
  const currentYear = new Date().getFullYear();
  const localFile = path.join(__dirname, `BSE_Holidays_${currentYear}.xlsx`);
  if (fs.existsSync(localFile)) {
    holidayCache = loadHolidayFile(localFile);
    console.log(`Holiday cache loaded from local file (${currentYear}):`, holidayCache.length, "entries");
  } else {
    holidayCache = fallbackHolidays(currentYear);
    lastDiag = { source: "fallback", year: currentYear, count: holidayCache.length };
    console.warn("No local holiday file found, using fallback");
  }
}

/* ---------- Routes ---------- */
app.get("/", (req, res) => {
  res.send("Sensex Expiry API is running. Try /api/sensex/expiries or /api/sensex/holidays");
});

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

/* ---------- Startup ---------- */
refreshHolidayCache();

app.listen(PORT, () => {
  console.log(`Sensex expiry server running at http://localhost:${PORT}`);
});
