// ====== CONFIG ======
const EXCEL_FILE = "Database.xlsx";      // <-- your real filename in repo
const SHEET_NAME = "14";

// Your requested solid colors (no transparency)
const AREA_SOLID_COLORS = {
  "MANAGEMENT": "#4169E1",   // Royal Blue
  "MARKETING":  "#CC0000",   // Red
  "SUPPLYCHAIN":"#228B22",   // Forest Green
  "ISQS":       "#00A3E0",   // Cyber Blue
  "FINANCE":    "#003366",   // Deep Navy Blue
  "ECBE":       "#FFBF00",   // Amber
  "ACCOUNTING": "#800020"    // Burgundy
};

const FALLBACK_COLOR = "#475467";
const charts = {};

function cleanStr(v){ return (v ?? "").toString().trim(); }
function yes(v){ return cleanStr(v).toLowerCase() === "yes"; }

// Map raw Area text -> category key (controls color)
function areaCategoryKey(area){
  const a = cleanStr(area).toLowerCase();

  // --- Highest priority: combined label should be Supply Chain (Forest Green) ---
  // Handles: "Marketing & Supply Chain Management", "Marketing and Supply Chain", etc.
  if (
    (a.includes("marketing") && a.includes("supply chain")) ||
    a.includes("marketing & supply chain") ||
    a.includes("marketing and supply chain")
  ) {
    return "SUPPLYCHAIN";
  }

  // Management
  if (a === "mgt" || a.includes("management")) return "MANAGEMENT";

  // Finance
  if (a === "fin" || a.includes("finance")) return "FINANCE";

  // Accounting
  if (a === "acct" || a.includes("account")) return "ACCOUNTING";

  // Energy commerce
  if (a.includes("energy") || a.includes("commerce") || a.includes("business economics") || a === "ecbe") return "ECBE";

  // ISQS
  if (a.includes("isqs") || a.includes("information") || a.includes("quantitative")) return "ISQS";

  // Supply Chain (includes SCM)
  if (
    a.includes("supply chain") ||
    a.includes("supply-chain") ||
    a.includes("supplychain") ||
    a.includes("scm")
  ) return "SUPPLYCHAIN";

  // Marketing (only marketing without supply chain)
  if (a.includes("marketing")) return "MARKETING";

  return "";
}

// Wrap long axis labels to stop overlap
function wrapLabel(label, maxLineLen = 14){
  const s = cleanStr(label);
  if (s.length <= maxLineLen) return s;

  const words = s.split(/\s+/);
  const lines = [];
  let line = "";

  for (const w of words){
    const test = line ? (line + " " + w) : w;
    if (test.length > maxLineLen){
      if (line) lines.push(line);
      line = w;
    } else {
      line = test;
    }
  }
  if (line) lines.push(line);
  return lines;
}

function destroy(key){
  if (charts[key]) { charts[key].destroy(); delete charts[key]; }
}

function buildBar(canvasId, labels, values, key, labelToColor){
  const el = document.getElementById(canvasId);
  if (!el) return;

  destroy(key);

  const bg = labels.map(l => labelToColor[l] || FALLBACK_COLOR);

  charts[key] = new Chart(el, {
    type: "bar",
    data: {
      labels: labels.map(l => wrapLabel(l, 14)),
      datasets: [{
        data: values,
        backgroundColor: bg,   // SOLID fill
        borderWidth: 0,
        borderRadius: 10,
        barThickness: 44,
        maxBarThickness: 54
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 650, easing: "easeOutQuart" },
      plugins: {
        legend: { display: false },
        tooltip: {
          enabled: true,
          callbacks: {
            title: (items) => labels[items[0].dataIndex]
          }
        }
      },
      layout: { padding: { left: 8, right: 8, top: 8, bottom: 0 } },
      scales: {
        x: {
          grid: { display: false },
          ticks: {
            autoSkip: false,
            maxRotation: 0,
            minRotation: 0,
            font: { size: 11, weight: "800" },
            color: "rgba(16,24,40,0.70)"
          }
        },
        y: {
          beginAtZero: true,
          ticks: {
            precision: 0,
            font: { weight: "800" },
            color: "rgba(16,24,40,0.60)"
          },
          grid: { color: "rgba(16,24,40,0.08)" }
        }
      }
    }
  });
}

async function loadExcelAndRender(){
  const status = document.getElementById("rrpStatus");
  const pill = document.getElementById("lastUpdated");

  const fromEl = document.getElementById("dateFrom");
  const toEl   = document.getElementById("dateTo");

  status.textContent = "Loading…";

  // --- helpers for date parsing/filtering ---
  function startOfDay(d){
    const x = new Date(d);
    x.setHours(0,0,0,0);
    return x;
  }
  function endOfDay(d){
    const x = new Date(d);
    x.setHours(23,59,59,999);
    return x;
  }

  // Converts Excel serial date (days since 1899-12-30) to JS Date
  function excelSerialToDate(n){
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const ms = n * 24 * 60 * 60 * 1000;
    return new Date(epoch.getTime() + ms);
  }

  // Robust pubdate parsing
  function parsePubDate(v){
    if (v === null || v === undefined || v === "") return null;

    // If SheetJS gave a number, it's likely an Excel serial date
    if (typeof v === "number" && isFinite(v)){
      const d = excelSerialToDate(v);
      return isNaN(d.getTime()) ? null : d;
    }

    const s = cleanStr(v);

    // Try ISO first (YYYY-MM-DD or YYYY/MM/DD)
    const iso = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (iso){
      const y = Number(iso[1]), m = Number(iso[2]) - 1, day = Number(iso[3]);
      const d = new Date(y, m, day);
      return isNaN(d.getTime()) ? null : d;
    }

    // Try MM/DD/YYYY or M/D/YYYY
    const us = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (us){
      const m = Number(us[1]) - 1, day = Number(us[2]), y = Number(us[3]);
      const d = new Date(y, m, day);
      return isNaN(d.getTime()) ? null : d;
    }

    // Fallback: Date.parse
    const t = Date.parse(s);
    if (!Number.isNaN(t)) return new Date(t);

    return null;
  }

  function getSelectedRange(){
    const fromVal = fromEl?.value ? new Date(fromEl.value) : null;
    const toVal   = toEl?.value ? new Date(toEl.value) : null;

    const from = fromVal ? startOfDay(fromVal) : null;
    const to   = toVal ? endOfDay(toVal) : null;

    return { from, to };
  }

  try{
    const url = `${EXCEL_FILE}?v=${Date.now()}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error("Unavailable");

    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[SHEET_NAME];
    if (!sheet) throw new Error("Unavailable");

    const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    // --- Date filter ---
    const { from, to } = getSelectedRange();

    const filtered = rows.filter(r => {
      // If no filter set, keep everything
      if (!from && !to) return true;

      const d = parsePubDate(r["pubdate"]);
      if (!d) return false;

      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    });

    // Aggregate by displayed Area label
    const byArea = {};
    for (const r of filtered){
      const areaLabel = cleanStr(r["Area"]);
      if (!areaLabel) continue;

      if (!byArea[areaLabel]) byArea[areaLabel] = { fiveStar:0, five:0, utd:0, ft:0 };

      const ranking = cleanStr(r["Ranking - Updated"]);
      if (ranking === "5*") byArea[areaLabel].fiveStar++;
      if (ranking === "5")  byArea[areaLabel].five++;

      if (yes(r["UTD Journal"])) byArea[areaLabel].utd++;
      if (yes(r["Financial Times Journal"])) byArea[areaLabel].ft++;
    }

    const areas = Object.keys(byArea).sort((a,b) => {
      const ta = byArea[a].fiveStar + byArea[a].five + byArea[a].utd + byArea[a].ft;
      const tb = byArea[b].fiveStar + byArea[b].five + byArea[b].utd + byArea[b].ft;
      return tb - ta;
    });

    // Map each displayed area label -> your requested solid color
    const labelToColor = {};
    for (const label of areas){
      const k = areaCategoryKey(label);
      labelToColor[label] = AREA_SOLID_COLORS[k] || FALLBACK_COLOR;
    }

    buildBar("chart5star", areas, areas.map(a => byArea[a].fiveStar), "fiveStar", labelToColor);
    buildBar("chart5",     areas, areas.map(a => byArea[a].five),     "five",     labelToColor);
    buildBar("chartUTD",   areas, areas.map(a => byArea[a].utd),      "utd",      labelToColor);
    buildBar("chartFT",    areas, areas.map(a => byArea[a].ft),       "ft",       labelToColor);

    const now = new Date();
    if (pill) pill.textContent = `Last updated: ${now.toLocaleString()}`;

    // Status message
    if (from || to){
      const fromTxt = from ? from.toLocaleDateString() : "—";
      const toTxt   = to ? to.toLocaleDateString() : "—";
      status.textContent = `Showing ${filtered.length} rows (pubdate ${fromTxt} → ${toTxt}).`;
    } else {
      status.textContent = `Showing all rows (${rows.length}).`;
    }

  }catch(e){
    console.error(e);
    status.textContent = "Charts are temporarily unavailable. Please try again.";
    if (pill) pill.textContent = "Last updated: —";
  }
}

// Buttons
document.getElementById("applyDate")?.addEventListener("click", loadExcelAndRender);

document.getElementById("clearDate")?.addEventListener("click", () => {
  const fromEl = document.getElementById("dateFrom");
  const toEl   = document.getElementById("dateTo");
  if (fromEl) fromEl.value = "";
  if (toEl) toEl.value = "";
  loadExcelAndRender();
});

document.getElementById("refreshBtn")?.addEventListener("click", loadExcelAndRender);

const refreshTopBtn = document.getElementById("refreshTop");
if (refreshTopBtn) {
  refreshTopBtn.addEventListener("click", loadExcelAndRender);
}

document.addEventListener("DOMContentLoaded", loadExcelAndRender);
