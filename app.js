/* SuperTools Excel Add-in - app.js */

// =============================================
// INIT
// =============================================
let selectedBgColor = "#4472C4";
let selectedFgColor = "#FFFFFF";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // Siap digunakan
    console.log("SuperTools Add-in loaded!");
  }
  // Setup dynamic listeners
  document.getElementById("split-delim").addEventListener("change", function () {
    const wrap = document.getElementById("split-custom-wrap");
    wrap.classList.toggle("hidden", this.value !== "custom");
  });
});

// =============================================
// NAVIGATION
// =============================================
function switchTab(name) {
  document.querySelectorAll(".content").forEach(el => el.classList.remove("active"));
  document.querySelectorAll(".nav-tab").forEach(el => el.classList.remove("active"));
  document.getElementById("tab-" + name).classList.add("active");
  event.target.classList.add("active");
}

// =============================================
// TOAST
// =============================================
function showToast(msg, color) {
  const t = document.getElementById("toast");
  t.textContent = msg;
  t.style.borderColor = color || "var(--accent)";
  t.style.color = color || "var(--accent)";
  t.classList.add("show");
  setTimeout(() => t.classList.remove("show"), 2800);
}

function showResult(id, html, show = true) {
  const el = document.getElementById(id);
  el.innerHTML = html;
  if (show) el.classList.add("show");
  else el.classList.remove("show");
}

// =============================================
// 1. STATISTIK
// =============================================
async function calcStats() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "address"]);
      await context.sync();

      const values = range.values.flat().filter(v => typeof v === "number" && !isNaN(v));

      if (values.length === 0) {
        showToast("⚠ Tidak ada data angka!", "var(--accent3)");
        return;
      }

      const n = values.length;
      const sorted = [...values].sort((a, b) => a - b);
      const sum = values.reduce((a, b) => a + b, 0);
      const mean = sum / n;
      const variance = values.reduce((acc, v) => acc + Math.pow(v - mean, 2), 0) / n;
      const stdDev = Math.sqrt(variance);
      const median = n % 2 === 0
        ? (sorted[n / 2 - 1] + sorted[n / 2]) / 2
        : sorted[Math.floor(n / 2)];

      // Mode
      const freq = {};
      values.forEach(v => freq[v] = (freq[v] || 0) + 1);
      const maxFreq = Math.max(...Object.values(freq));
      const mode = Object.keys(freq).filter(k => freq[k] === maxFreq).join(", ");

      const stats = [
        { label: "Jumlah Data", value: n, cls: "purple" },
        { label: "Total (SUM)", value: fmt(sum), cls: "green" },
        { label: "Rata-rata", value: fmt(mean), cls: "green" },
        { label: "Median", value: fmt(median), cls: "purple" },
        { label: "Modus", value: mode, cls: "yellow" },
        { label: "Std. Deviasi", value: fmt(stdDev), cls: "yellow" },
        { label: "Minimum", value: fmt(sorted[0]), cls: "red" },
        { label: "Maksimum", value: fmt(sorted[n - 1]), cls: "green" },
      ];

      const html = stats.map(s => `
        <div class="stat-item">
          <div class="stat-label">${s.label}</div>
          <div class="stat-value ${s.cls}">${s.value}</div>
        </div>
      `).join("");

      const container = document.getElementById("stats-result");
      container.innerHTML = html;
      container.style.display = "grid";

      showToast(`✅ Statistik dari ${n} angka dihitung!`);
    });
  } catch (e) {
    showToast("⚠ Pilih range data dulu!", "var(--accent3)");
  }
}

function fmt(n) {
  return typeof n === "number"
    ? (Number.isInteger(n) ? n.toLocaleString("id-ID") : n.toLocaleString("id-ID", { maximumFractionDigits: 2 }))
    : n;
}

// =============================================
// 2. DUPLIKAT & KOSONG
// =============================================
async function highlightDuplicates() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount", "address"]);
      await context.sync();

      const flat = range.values.flat();
      const seen = {};
      const dupes = new Set();
      flat.forEach(v => {
        const key = String(v).trim();
        if (key === "") return;
        seen[key] = (seen[key] || 0) + 1;
        if (seen[key] > 1) dupes.add(key);
      });

      let count = 0;
      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const val = String(range.values[r][c]).trim();
          if (dupes.has(val)) {
            const cell = range.getCell(r, c);
            cell.format.fill.color = "#FF6B6B";
            cell.format.font.color = "#FFFFFF";
            count++;
          }
        }
      }

      await context.sync();
      showResult("dup-result", `<span style="color:var(--accent3)">🔴 ${count} sel duplikat ditemukan & dihighlight</span>`, true);
      showToast(`🔍 ${count} duplikat ditemukan!`, "var(--accent3)");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

async function highlightBlanks() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      let count = 0;
      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          if (range.values[r][c] === "" || range.values[r][c] === null) {
            const cell = range.getCell(r, c);
            cell.format.fill.color = "#FFD166";
            count++;
          }
        }
      }

      await context.sync();
      showResult("dup-result", `<span style="color:var(--accent4)">🟡 ${count} sel kosong ditemukan & dihighlight</span>`, true);
      showToast(`🟡 ${count} sel kosong ditemukan!`, "var(--accent4)");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

async function removeDuplicates() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
      await context.sync();

      const rows = range.values;
      const seen = new Set();
      const toDelete = [];

      rows.forEach((row, idx) => {
        const key = JSON.stringify(row);
        if (seen.has(key)) {
          toDelete.push(idx);
        } else {
          seen.add(key);
        }
      });

      // Hapus dari bawah ke atas
      for (let i = toDelete.length - 1; i >= 0; i--) {
        const rowIdx = range.rowIndex + toDelete[i];
        context.workbook.worksheets.getActiveWorksheet()
          .getRange(`${rowIdx + 1}:${rowIdx + 1}`)
          .delete(Excel.DeleteShiftDirection.up);
      }

      await context.sync();
      showResult("dup-result", `<span style="color:var(--accent)">✅ ${toDelete.length} baris duplikat dihapus</span>`, true);
      showToast(`🗑 ${toDelete.length} duplikat dihapus!`);
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

async function clearHighlights() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.clear();
      range.format.font.color = "#000000";
      await context.sync();
      showToast("✖ Highlight dihapus!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

// =============================================
// 3. KONVERSI FORMAT
// =============================================
async function convertFormat() {
  const type = document.getElementById("conv-type").value;
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const newValues = range.values.map(row =>
        row.map(cell => {
          if (cell === null || cell === "") return cell;
          const s = String(cell);
          const n = Number(cell);

          switch (type) {
            case "upper": return s.toUpperCase();
            case "lower": return s.toLowerCase();
            case "proper": return s.replace(/\w\S*/g, w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
            case "trim": return s.replace(/\s+/g, " ").trim();
            case "rupiah":
              if (!isNaN(n)) return "Rp " + n.toLocaleString("id-ID");
              return cell;
            case "persen":
              if (!isNaN(n)) return (n * 100).toFixed(1) + "%";
              return cell;
            case "ribuan":
              if (!isNaN(n)) return n.toLocaleString("id-ID");
              return cell;
            case "round2":
              if (!isNaN(n)) return Math.round(n * 100) / 100;
              return cell;
            case "date-id":
              if (cell instanceof Date || !isNaN(Date.parse(s))) {
                const d = new Date(s);
                return `${String(d.getDate()).padStart(2, "0")}/${String(d.getMonth() + 1).padStart(2, "0")}/${d.getFullYear()}`;
              }
              return cell;
            case "date-long":
              if (!isNaN(Date.parse(s))) {
                return new Date(s).toLocaleDateString("id-ID", { weekday: "long", year: "numeric", month: "long", day: "numeric" });
              }
              return cell;
            case "date-iso":
              if (!isNaN(Date.parse(s))) {
                return new Date(s).toISOString().split("T")[0];
              }
              return cell;
            default: return cell;
          }
        })
      );

      range.values = newValues;
      await context.sync();
      showToast("✅ Konversi berhasil diterapkan!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

// =============================================
// 4. TEMPLATE STYLE
// =============================================
const STYLES = {
  professional: {
    header: { bg: "#1F3864", fg: "#FFFFFF", bold: true },
    odd: { bg: "#F2F2F2", fg: "#000000" },
    even: { bg: "#FFFFFF", fg: "#000000" },
    border: "#BDD7EE"
  },
  dark: {
    header: { bg: "#0D0D0F", fg: "#00E5A0", bold: true },
    odd: { bg: "#18181C", fg: "#F0F0F5" },
    even: { bg: "#222228", fg: "#F0F0F5" },
    border: "#2E2E38"
  },
  colorful: {
    header: { bg: "#7C6CFC", fg: "#FFFFFF", bold: true },
    odd: { bg: "#FFF3CD", fg: "#000000" },
    even: { bg: "#D1ECF1", fg: "#000000" },
    border: "#7C6CFC"
  },
  minimal: {
    header: { bg: "#FFFFFF", fg: "#000000", bold: true },
    odd: { bg: "#FFFFFF", fg: "#000000" },
    even: { bg: "#F9F9F9", fg: "#000000" },
    border: "#E0E0E0"
  },
  ocean: {
    header: { bg: "#006994", fg: "#FFFFFF", bold: true },
    odd: { bg: "#E0F4FF", fg: "#003049" },
    even: { bg: "#FFFFFF", fg: "#003049" },
    border: "#4AACDB"
  },
  sunset: {
    header: { bg: "#FF6B6B", fg: "#FFFFFF", bold: true },
    odd: { bg: "#FFF0EB", fg: "#333333" },
    even: { bg: "#FFFFFF", fg: "#333333" },
    border: "#FF8585"
  }
};

async function applyStyle(styleName) {
  const s = STYLES[styleName];
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["rowCount", "columnCount"]);
      await context.sync();

      // Header row
      const headerRow = range.getRow(0);
      headerRow.format.fill.color = s.header.bg;
      headerRow.format.font.color = s.header.fg;
      if (s.header.bold) headerRow.format.font.bold = true;

      // Data rows
      for (let r = 1; r < range.rowCount; r++) {
        const row = range.getRow(r);
        const style = r % 2 === 1 ? s.odd : s.even;
        row.format.fill.color = style.bg;
        row.format.font.color = style.fg;
      }

      // Border
      range.format.borders.getItem("InsideHorizontal").style = "Thin";
      range.format.borders.getItem("InsideHorizontal").color = s.border;
      range.format.borders.getItem("EdgeBottom").style = "Thin";
      range.format.borders.getItem("EdgeBottom").color = s.border;
      range.format.borders.getItem("EdgeTop").style = "Medium";
      range.format.borders.getItem("EdgeTop").color = s.header.bg;

      await context.sync();
      showToast(`✨ Style "${styleName}" diterapkan!`);
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

// =============================================
// 5. HEATMAP
// =============================================
async function applyHeatmap() {
  const scheme = document.getElementById("heatmap-scheme").value;
  const skipHeader = document.getElementById("heatmap-header").checked;

  const colorSchemes = {
    "red-green": { low: [255, 107, 107], high: [0, 229, 160] },
    "blue-red": { low: [100, 149, 237], high: [255, 107, 107] },
    "yellow-red": { low: [255, 209, 102], high: [255, 50, 50] },
    "white-blue": { low: [255, 255, 255], high: [0, 80, 180] }
  };

  const cs = colorSchemes[scheme];

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const startRow = skipHeader ? 1 : 0;
      const numbers = [];
      for (let r = startRow; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const v = Number(range.values[r][c]);
          if (!isNaN(v)) numbers.push(v);
        }
      }

      if (numbers.length === 0) {
        showToast("⚠ Tidak ada angka untuk heatmap!", "var(--accent3)");
        return;
      }

      const min = Math.min(...numbers);
      const max = Math.max(...numbers);

      for (let r = startRow; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const v = Number(range.values[r][c]);
          if (!isNaN(v) && max !== min) {
            const t = (v - min) / (max - min);
            const R = Math.round(cs.low[0] + t * (cs.high[0] - cs.low[0]));
            const G = Math.round(cs.low[1] + t * (cs.high[1] - cs.low[1]));
            const B = Math.round(cs.low[2] + t * (cs.high[2] - cs.low[2]));
            const hex = "#" + [R, G, B].map(x => x.toString(16).padStart(2, "0")).join("");
            const cell = range.getCell(r, c);
            cell.format.fill.color = hex;
            // Teks putih jika gelap
            const brightness = 0.299 * R + 0.587 * G + 0.114 * B;
            cell.format.font.color = brightness < 128 ? "#FFFFFF" : "#000000";
          }
        }
      }

      await context.sync();
      showToast("🌡 Heatmap berhasil diterapkan!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

// =============================================
// 6. CUSTOM COLOR
// =============================================
function selectSwatch(el, type, color) {
  const group = type === "bg" ? "bg-swatches" : "fg-swatches";
  document.querySelectorAll(`#${group} .swatch`).forEach(s => s.classList.remove("selected"));
  el.classList.add("selected");
  if (type === "bg") selectedBgColor = color;
  else selectedFgColor = color;
}

async function applyCustomColor() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = selectedBgColor;
      range.format.font.color = selectedFgColor;
      await context.sync();
      showToast("🎨 Warna diterapkan!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

async function clearFormat() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.clear();
      range.format.font.bold = false;
      range.format.font.italic = false;
      range.format.font.color = "#000000";
      await context.sync();
      showToast("✖ Format direset!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

// =============================================
// 7. CHART BUILDER
// =============================================
const CHART_TYPE_MAP = {
  "ColumnClustered": Excel.ChartType.columnClustered,
  "BarClustered": Excel.ChartType.barClustered,
  "Line": Excel.ChartType.line,
  "LineFilled": Excel.ChartType.areaStacked,
  "Pie": Excel.ChartType.pie,
  "Doughnut": Excel.ChartType.doughnut,
  "XYScatter": Excel.ChartType.xyscatter,
  "Radar": Excel.ChartType.radar
};

async function createChart() {
  const typeKey = document.getElementById("chart-type").value;
  const title = document.getElementById("chart-title").value || "Chart";
  const showLegend = document.getElementById("chart-legend").checked;
  const showDataLabel = document.getElementById("chart-datalabel").checked;

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();

      const chartType = CHART_TYPE_MAP[typeKey] || Excel.ChartType.columnClustered;
      const chart = sheet.charts.add(chartType, range, Excel.ChartSeriesBy.auto);

      chart.title.text = title;
      chart.legend.visible = showLegend;
      chart.legend.position = Excel.ChartLegendPosition.bottom;

      if (showDataLabel) {
        chart.dataLabels.showValue = true;
      }

      // Ukuran default
      chart.width = 480;
      chart.height = 300;

      await context.sync();
      showToast(`📈 Chart "${title}" berhasil dibuat!`);
    });
  } catch (e) {
    showToast("⚠ Pilih data range dulu!", "var(--accent3)");
  }
}

async function recommendChart() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount"]);
      await context.sync();

      const rows = range.rowCount;
      const cols = range.columnCount;
      const values = range.values;

      // Hitung berapa banyak angka
      const nums = values.flat().filter(v => typeof v === "number" && !isNaN(v)).length;
      const total = rows * cols;
      const numRatio = nums / total;

      let recommendation, reason;

      if (cols === 2 && numRatio > 0.4) {
        recommendation = "📊 Bar Chart / Kolom";
        reason = "Data 2 kolom cocok untuk perbandingan kategori";
      } else if (cols >= 3 && rows <= 10) {
        recommendation = "📈 Line Chart";
        reason = "Beberapa seri data dengan sedikit titik, cocok untuk tren";
      } else if (cols === 2 && rows <= 8 && numRatio > 0.5) {
        recommendation = "🥧 Pie Chart";
        reason = "Data proporsi dengan sedikit kategori";
      } else if (numRatio < 0.3) {
        recommendation = "📋 Tabel (tidak perlu chart)";
        reason = "Data kebanyakan teks, lebih baik sebagai tabel";
      } else if (rows > 20) {
        recommendation = "📈 Line Chart atau Area Chart";
        reason = "Data banyak titik, cocok untuk visualisasi tren";
      } else {
        recommendation = "📊 Column Chart";
        reason = "Pilihan umum untuk perbandingan data";
      }

      showResult("recommend-result",
        `<span style="color:var(--accent)">${recommendation}</span>\n\n` +
        `Alasan: ${reason}\n\n` +
        `📐 Dimensi data: ${rows} baris × ${cols} kolom\n` +
        `🔢 Proporsi angka: ${Math.round(numRatio * 100)}%`,
        true
      );
    });
  } catch (e) {
    showToast("⚠ Pilih data range dulu!", "var(--accent3)");
  }
}

// =============================================
// 8. FIND & REPLACE
// =============================================
async function findReplace() {
  const findText = document.getElementById("find-text").value;
  const replaceText = document.getElementById("replace-text").value;
  const caseSensitive = document.getElementById("fr-case").checked;
  const wholeWord = document.getElementById("fr-whole").checked;

  if (!findText) {
    showToast("⚠ Masukkan teks yang dicari!", "var(--accent3)");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load("values");
      await context.sync();

      let count = 0;
      const newValues = usedRange.values.map(row =>
        row.map(cell => {
          if (typeof cell !== "string" && typeof cell !== "number") return cell;
          const s = String(cell);
          const flags = caseSensitive ? "g" : "gi";
          const pattern = wholeWord ? `\\b${escapeRegex(findText)}\\b` : escapeRegex(findText);
          const re = new RegExp(pattern, flags);
          if (re.test(s)) {
            count++;
            return s.replace(re, replaceText);
          }
          return cell;
        })
      );

      usedRange.values = newValues;
      await context.sync();
      showResult("fr-result", `<span style="color:var(--accent)">✅ ${count} penggantian dilakukan</span>`, true);
      showToast(`✅ ${count} teks diganti!`);
    });
  } catch (e) {
    showToast("⚠ Error!", "var(--accent3)");
  }
}

async function findOnly() {
  const findText = document.getElementById("find-text").value;
  if (!findText) {
    showToast("⚠ Masukkan teks yang dicari!", "var(--accent3)");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["values", "address"]);
      await context.sync();

      const caseSensitive = document.getElementById("fr-case").checked;
      const found = [];

      usedRange.values.forEach((row, ri) => {
        row.forEach((cell, ci) => {
          const s = String(cell);
          const match = caseSensitive ? s.includes(findText) : s.toLowerCase().includes(findText.toLowerCase());
          if (match) found.push({ row: ri + 1, col: ci + 1, value: s });
        });
      });

      if (found.length === 0) {
        showResult("fr-result", `<span style="color:var(--accent3)">❌ Tidak ditemukan</span>`, true);
      } else {
        const preview = found.slice(0, 8).map(f => `Baris ${f.row}, Kol ${f.col}: "${f.value.substring(0, 30)}"`).join("\n");
        showResult("fr-result", `<span style="color:var(--accent)">🔍 ${found.length} sel ditemukan</span>\n\n${preview}${found.length > 8 ? `\n...dan ${found.length - 8} lainnya` : ""}`, true);
      }
    });
  } catch (e) {
    showToast("⚠ Error!", "var(--accent3)");
  }
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// =============================================
// 9. SPLIT KOLOM
// =============================================
async function splitColumn() {
  let delim = document.getElementById("split-delim").value;
  if (delim === "custom") {
    delim = document.getElementById("split-custom").value;
    if (!delim) {
      showToast("⚠ Masukkan pemisah custom!", "var(--accent3)");
      return;
    }
  } else if (delim === "\\t") {
    delim = "\t";
  }

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnIndex", "rowIndex"]);
      await context.sync();

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      let maxParts = 1;

      // Cari max parts
      range.values.forEach(row => {
        row.forEach(cell => {
          const parts = String(cell).split(delim).length;
          if (parts > maxParts) maxParts = parts;
        });
      });

      // Tulis ke kolom-kolom berikutnya
      for (let r = 0; r < range.rowCount; r++) {
        const val = String(range.values[r][0]);
        const parts = val.split(delim);
        for (let p = 0; p < maxParts; p++) {
          const col = range.columnIndex + 1 + p;
          const cell = sheet.getCell(range.rowIndex + r, col);
          cell.values = [[parts[p] || ""]];
        }
      }

      await context.sync();
      showToast(`✂️ Split menjadi ${maxParts} kolom!`);
    });
  } catch (e) {
    showToast("⚠ Pilih satu kolom dulu!", "var(--accent3)");
  }
}

// =============================================
// 10. FORMULA PINTAR
// =============================================
async function insertFormula(type) {
  try {
    await Excel.run(async (context) => {
      const cell = context.workbook.getSelectedRange();
      cell.load("address");
      await context.sync();

      // Deteksi kolom di atas
      const addr = cell.address.split("!").pop();
      const col = addr.replace(/[0-9]/g, "");
      const row = parseInt(addr.replace(/[A-Z]/g, ""));
      const above = `${col}1:${col}${row - 1}`;

      const formulas = {
        sum: `=SUM(${above})`,
        avg: `=AVERAGE(${above})`,
        count: `=COUNT(${above})`,
        max: `=MAX(${above})`,
        min: `=MIN(${above})`,
        counta: `=COUNTA(${above})`,
        iferror: `=IFERROR(,0)`,
        vlookup: `=VLOOKUP(,"",1,FALSE)`
      };

      cell.formulas = [[formulas[type] || ""]];
      await context.sync();
      showToast(`🧮 Formula ${type.toUpperCase()} dimasukkan!`);
    });
  } catch (e) {
    showToast("⚠ Pilih sel tujuan dulu!", "var(--accent3)");
  }
}

// =============================================
// 11. SORT
// =============================================
async function sortRange(dir) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["columnCount"]);
      await context.sync();

      range.sort.apply([{
        key: 0,
        ascending: dir === "asc"
      }]);

      await context.sync();
      showToast(`↕ Data diurutkan ${dir === "asc" ? "A→Z" : "Z→A"}!`);
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

async function autoFilter() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.autoFilter.apply(range);
      await context.sync();
      showToast("🔽 Auto Filter diterapkan!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

async function clearFilter() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.autoFilter.remove();
      await context.sync();
      showToast("✖ Filter dihapus!");
    });
  } catch (e) {
    showToast("⚠ Tidak ada filter aktif!", "var(--accent3)");
  }
}

// =============================================
// 12. EXPORT CSV
// =============================================
async function exportCSV() {
  let delim = document.getElementById("csv-delim").value;
  if (delim === "\\t") delim = "\t";
  const header = document.getElementById("csv-header").checked;

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      let rows = range.values;
      if (!header && rows.length > 0) rows = rows.slice(1);

      const csv = rows.map(row =>
        row.map(cell => {
          const s = String(cell === null ? "" : cell);
          return s.includes(delim) || s.includes('"') || s.includes("\n")
            ? `"${s.replace(/"/g, '""')}"`
            : s;
        }).join(delim)
      ).join("\n");

      downloadFile(csv, "export.csv", "text/csv;charset=utf-8;");
      showToast("⬇ CSV berhasil didownload!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

// =============================================
// 13. EXPORT JSON
// =============================================
async function exportJSON() {
  const pretty = document.getElementById("json-pretty").checked;
  const useHeader = document.getElementById("json-header").checked;

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const rows = range.values;
      let result;

      if (useHeader && rows.length > 1) {
        const keys = rows[0].map(k => String(k));
        result = rows.slice(1).map(row => {
          const obj = {};
          keys.forEach((k, i) => obj[k] = row[i]);
          return obj;
        });
      } else {
        result = rows;
      }

      const json = pretty ? JSON.stringify(result, null, 2) : JSON.stringify(result);
      downloadFile(json, "export.json", "application/json");
      showToast("⬇ JSON berhasil didownload!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu!", "var(--accent3)");
  }
}

// =============================================
// 14. COPY TO CLIPBOARD
// =============================================
async function copyToClipboard() {
  const format = document.getElementById("copy-format").value;

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const rows = range.values;
      let output = "";

      if (format === "tab") {
        output = rows.map(r => r.join("\t")).join("\n");
      } else if (format === "csv") {
        output = rows.map(r => r.join(",")).join("\n");
      } else if (format === "markdown") {
        const header = rows[0].map(c => String(c));
        const sep = header.map(() => "---");
        const body = rows.slice(1);
        output = [
          "| " + header.join(" | ") + " |",
          "| " + sep.join(" | ") + " |",
          ...body.map(r => "| " + r.join(" | ") + " |")
        ].join("\n");
      } else if (format === "html") {
        const header = `<tr>${rows[0].map(c => `<th>${c}</th>`).join("")}</tr>`;
        const body = rows.slice(1).map(r => `<tr>${r.map(c => `<td>${c}</td>`).join("")}</tr>`).join("\n");
        output = `<table>\n<thead>${header}</thead>\n<tbody>${body}</tbody>\n</table>`;
      }

      await navigator.clipboard.writeText(output);
      showResult("copy-result", `<span style="color:var(--accent)">✅ ${rows.length} baris disalin ke clipboard (format: ${format})</span>`, true);
      showToast("📋 Disalin ke clipboard!");
    });
  } catch (e) {
    showToast("⚠ Pilih range dulu atau izinkan clipboard!", "var(--accent3)");
  }
}

// =============================================
// HELPER: DOWNLOAD FILE
// =============================================
function downloadFile(content, filename, mimeType) {
  const blob = new Blob(["\uFEFF" + content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
