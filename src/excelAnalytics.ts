import * as XLSX from "xlsx";
import {
  computeColumnStatsAndOutliers,
  computePearsonCorrelation,
  forecastSeries,
  groupTimeSeries,
  ForecastResult,
  OutlierDetail
} from "./utils/mathEngine";

export type DataRow = Record<string, any>;

export type ColumnType = "numeric" | "date" | "boolean" | "text" | "mixed" | "empty";

export type ColumnProfile = {
  name: string;
  type: ColumnType;
  count: number;
  emptyCount: number;
  nonEmptyCount: number;
  uniqueCount: number;
  topValues: Array<{ value: any; count: number }>;
  sampleValues: any[];
  sum?: number;
  avg?: number;
  median?: number;
  stddev?: number;
  min?: number;
  max?: number;

  // Phase 2 Advanced Statistical Properties
  q1?: number;
  q3?: number;
  iqr?: number;
  lowerBound?: number;
  upperBound?: number;
  outliers?: number[];
  outliersList?: OutlierDetail[];
};

export type DataProfile = {
  totalRows: number;
  totalColumns: number;
  duplicateRows: number;
  headers: string[];
  sheetName?: string;
  headerRow?: number;
  columns: ColumnProfile[];
  topDuplicateRows: Array<{ row: string; count: number }>;

  // Phase 2 Analytical Heatmap & Predictions
  correlations?: Record<string, Record<string, number>>;
  forecasts?: Record<string, ForecastResult>;
};

function normalizeHeader(header: any) {
  const label = String(header ?? "").trim();
  if (!label) return "column";
  return label.replace(/\s+/g, " ").trim();
}

function candidateHeaderScore(value: any) {
  const text = String(value ?? "").trim();
  if (!text) return 0;
  let score = 0;
  if (/order\s*id|order\s*no|order\s*number|sku|product|qty|quantity|amount|date|name|id/i.test(text)) score += 5;
  if (/\d/.test(text)) score -= 2;
  if (text.length < 2) score -= 2;
  return score;
}

function detectHeaderRow(raw: any[][]) {
  let bestScore = -Infinity;
  let bestRow = 0;
  for (let i = 0; i < Math.min(raw.length, 6); i++) {
    const row = raw[i];
    if (!Array.isArray(row)) continue;
    const score = row.reduce((sum: number, cell: any) => sum + candidateHeaderScore(cell), 0);
    if (score > bestScore) {
      bestScore = score;
      bestRow = i;
    }
  }
  return bestRow;
}

function countNonEmpty(row: any[]) {
  return Array.isArray(row) ? row.reduce((count, cell) => count + (String(cell ?? "").trim() ? 1 : 0), 0) : 0;
}

export function detectHeaderRowForSheet(raw: any[][], maxRows = 12) {
  let bestScore = -Infinity;
  let bestRow = 0;
  for (let i = 0; i < Math.min(raw.length, maxRows); i++) {
    const row = raw[i];
    if (!Array.isArray(row)) continue;
    const nonEmpty = countNonEmpty(row);
    const headerScore = row.reduce((sum: number, cell: any) => sum + candidateHeaderScore(cell), 0);
    const typeHints = row.reduce((sum: number, cell: any) => {
      const text = String(cell ?? "").trim();
      if (!text) return sum;
      if (/^[0-9,.₹%\-]+$/.test(text)) return sum - 0.5;
      return sum + 0.25;
    }, 0);
    const score = headerScore + nonEmpty * 0.3 + typeHints;
    if (score > bestScore) {
      bestScore = score;
      bestRow = i;
    }
  }
  return bestRow;
}

export function normalizeHeaders(rawHeaders: any[]) {
  const headers = Array.isArray(rawHeaders) ? rawHeaders.map(normalizeHeader) : [];
  const uniqueHeaders: string[] = [];
  headers.forEach((name, idx) => {
    let candidate = name || `column_${idx + 1}`;
    let suffix = 1;
    while (uniqueHeaders.includes(candidate)) {
      candidate = `${name}_${suffix}`;
      suffix += 1;
    }
    uniqueHeaders.push(candidate);
  });
  return uniqueHeaders;
}

export function buildTable(raw: any[][], headerRow: number) {
  const rawHeaders = Array.isArray(raw[headerRow]) ? raw[headerRow] : [];
  const headers = normalizeHeaders(rawHeaders);
  const rows = raw.slice(headerRow + 1).filter((r: any) => Array.isArray(r) && r.some((c: any) => String(c ?? "").trim() !== ""));
  const data = rows.map((row: any[]) => {
    const obj: DataRow = {};
    headers.forEach((header, idx) => {
      obj[header] = row[idx] ?? "";
    });
    return obj;
  });
  return { headers, data };
}

export function chooseBestSheet(wb: XLSX.WorkBook) {
  let best: { sheetName: string; headerRow: number; headers: string[]; data: DataRow[]; score: number } | null = null;
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) as any[][];
    if (!raw.length) continue;
    const headerRow = detectHeaderRowForSheet(raw, 12);
    const { headers, data } = buildTable(raw, headerRow);
    if (!data.length || headers.length < 2) continue;
    const score = data.length * 1.2 + headers.length * 2 + countNonEmpty(raw[headerRow]) * 0.8;
    if (!best || score > best.score) {
      best = { sheetName, headerRow, headers, data, score };
    }
  }
  return best;
}

function isDateValue(value: any) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return true;
  const str = String(value ?? "").trim();
  if (!str) return false;
  const parsed = Date.parse(str);
  return !Number.isNaN(parsed);
}

export function parseWorkbook(wb: XLSX.WorkBook) {
  const candidate = chooseBestSheet(wb);
  if (candidate) {
    return { data: candidate.data, originalRows: candidate.data.length, headerRow: candidate.headerRow, sheetName: candidate.sheetName, headers: candidate.headers };
  }
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) as any[][];
  const headerRow = detectHeaderRowForSheet(raw, 12);
  const table = buildTable(raw, headerRow);
  return { data: table.data, originalRows: table.data.length, headerRow, sheetName, headers: table.headers };
}

export function parseExcel(file: File) {
  return new Promise<{ data: DataRow[]; originalRows: number; headerRow: number; sheetName: string; headers: string[] }>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target!.result, { type: "array" });
        resolve(parseWorkbook(wb));
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

export function parseExcelBuffer(buffer: any) {
  const isNodeBuffer = typeof Buffer !== "undefined" && Buffer.isBuffer && Buffer.isBuffer(buffer);
  const wb = XLSX.read(buffer, { type: isNodeBuffer ? "buffer" : "array" });
  return parseWorkbook(wb);
}

function getType(values: any[]) {
  const nonEmpty = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
  if (!nonEmpty.length) return "empty" as ColumnType;
  const numeric = nonEmpty.every((v) => !Number.isNaN(Number(v)));
  if (numeric) return "numeric" as ColumnType;
  const booleanish = nonEmpty.every((v) => /^(true|false|yes|no|0|1)$/i.test(String(v).trim()));
  if (booleanish) return "boolean" as ColumnType;
  const dateish = nonEmpty.every((v) => isDateValue(v));
  if (dateish) return "date" as ColumnType;
  const mixed = new Set(nonEmpty.map((v) => typeof v)).size > 1;
  return mixed ? "mixed" : "text";
}

function topValues(values: any[], limit = 8) {
  const counts = values.reduce((acc: Record<string, number>, value: any) => {
    const key = String(value ?? "").trim();
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);
  const result: { value: string; count: number }[] = (Object.entries(counts) as [string, number][]).map(([value, count]) => ({ value, count }));
  return result
    .sort((a, b) => b.count - a.count)
    .slice(0, limit);
}

export function analyzeData(data: DataRow[], headers: string[]): DataProfile {
  const totalRows = data.length;
  const columnStats: ColumnProfile[] = headers.map((header) => {
    const values = data.map((row) => row[header]);
    const clean = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== "");
    const uniqueCount = new Set(values.map((v) => String(v ?? "").trim())).size;
    const type = getType(values);
    const top = topValues(clean);
    const numericValues = clean
      .map((v) => Number(String(v ?? "").replace(/[,₹\s%]/g, "")))
      .filter((v) => !Number.isNaN(v));
    const sum = numericValues.length ? numericValues.reduce((a, b) => a + b, 0) : undefined;
    const min = numericValues.length ? Math.min(...numericValues) : undefined;
    const max = numericValues.length ? Math.max(...numericValues) : undefined;
    const avg = numericValues.length ? sum! / numericValues.length : undefined;
    const median = numericValues.length
      ? (() => {
          const sorted = [...numericValues].sort((a, b) => a - b);
          const mid = Math.floor(sorted.length / 2);
          return sorted.length % 2 === 1 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
        })()
      : undefined;
    const stddev = numericValues.length
      ? Math.sqrt(numericValues.reduce((acc, val) => acc + Math.pow(val - avg!, 2), 0) / numericValues.length)
      : undefined;

    // Advanced Phase 2: Compute IQR and Z-Score Outliers
    let q1: number | undefined;
    let q3: number | undefined;
    let iqr: number | undefined;
    let lowerBound: number | undefined;
    let upperBound: number | undefined;
    let outliers: number[] | undefined;
    let outliersList: OutlierDetail[] | undefined;

    if (type === "numeric" && numericValues.length > 0) {
      const stats = computeColumnStatsAndOutliers(numericValues, values);
      if (stats) {
        q1 = stats.q1;
        q3 = stats.q3;
        iqr = stats.iqr;
        lowerBound = stats.lowerBound;
        upperBound = stats.upperBound;
        outliers = stats.outliers;
        outliersList = stats.outliersList;
      }
    }

    return {
      name: header,
      type,
      count: values.length,
      emptyCount: values.filter((v) => v === null || v === undefined || String(v).trim() === "").length,
      nonEmptyCount: clean.length,
      uniqueCount,
      topValues: top,
      sampleValues: values.slice(0, 5),
      sum,
      avg,
      median,
      stddev,
      min,
      max,
      q1,
      q3,
      iqr,
      lowerBound,
      upperBound,
      outliers,
      outliersList
    };
  });

  const rowStrings = data.map((row) => JSON.stringify(row));
  const duplicates: Record<string, number> = {};
  rowStrings.forEach((row) => { duplicates[row] = (duplicates[row] || 0) + 1; });
  const topDuplicateRows = Object.entries(duplicates)
    .filter(([, count]) => count > 1)
    .map(([row, count]) => ({ row, count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);
  const duplicateRows = topDuplicateRows.reduce((sum, item) => sum + item.count - 1, 0);

  // Advanced Phase 2: Compute Pearson Correlation Matrix
  const correlations: Record<string, Record<string, number>> = {};
  const numericColumnNames = columnStats
    .filter((c) => c.type === "numeric" && c.nonEmptyCount > 1)
    .map((c) => c.name);

  numericColumnNames.forEach((col1) => {
    correlations[col1] = {};
    numericColumnNames.forEach((col2) => {
      if (col1 === col2) {
        correlations[col1][col2] = 1.0;
      } else {
        const x: number[] = [];
        const y: number[] = [];
        data.forEach((row) => {
          const val1 = Number(String(row[col1] ?? "").replace(/[,₹\s%]/g, ""));
          const val2 = Number(String(row[col2] ?? "").replace(/[,₹\s%]/g, ""));
          if (!Number.isNaN(val1) && !Number.isNaN(val2) && row[col1] !== null && row[col2] !== null) {
            x.push(val1);
            y.push(val2);
          }
        });
        if (x.length > 1) {
          correlations[col1][col2] = computePearsonCorrelation(x, y);
        } else {
          correlations[col1][col2] = 0;
        }
      }
    });
  });

  // Advanced Phase 2: Compute Holt-Winters Forecasting
  const forecasts: Record<string, ForecastResult> = {};
  const findHeaderByPatterns = (hdrs: string[], patterns: RegExp[]) => {
    const norm = (s: string) => s.toLowerCase();
    for (const p of patterns) {
      const found = hdrs.find((h) => p.test(norm(h)));
      if (found) return found;
    }
    return null;
  };
  const dateHeader = findHeaderByPatterns(headers, [/date|created|ordered|timestamp|time|ship/]);

  if (dateHeader && numericColumnNames.length > 0) {
    const dates = data.map((row) => row[dateHeader]);
    numericColumnNames.forEach((numCol) => {
      const vals = data.map((row) => {
        const raw = row[numCol];
        if (raw === null || raw === undefined) return 0;
        const n = Number(String(raw).replace(/[,₹\s%]/g, ""));
        return Number.isFinite(n) ? n : 0;
      });

      const grouped = groupTimeSeries(dates, vals);
      if (grouped.values.length >= 3) {
        const forecastSteps = 3;
        const result = forecastSeries(grouped.values, forecastSteps);

        const lastLabel = grouped.labels[grouped.labels.length - 1];
        const futureLabels: string[] = [];

        for (let i = 1; i <= forecastSteps; i++) {
          if (lastLabel && lastLabel.includes("-W")) {
            const [yearStr, weekStr] = lastLabel.split("-W");
            const w = parseInt(weekStr) + i;
            futureLabels.push(`${yearStr}-W${String(w).padStart(2, "0")}`);
          } else if (lastLabel && lastLabel.includes("-")) {
            const parts = lastLabel.split("-");
            if (parts.length === 2) {
              const yr = parseInt(parts[0]);
              let mo = parseInt(parts[1]) + i;
              const yAdd = Math.floor((mo - 1) / 12);
              mo = ((mo - 1) % 12) + 1;
              futureLabels.push(`${yr + yAdd}-${String(mo).padStart(2, "0")}`);
            } else {
              const dObj = new Date(lastLabel);
              dObj.setDate(dObj.getDate() + i);
              futureLabels.push(dObj.toISOString().split("T")[0]);
            }
          } else {
            futureLabels.push(`Projection +${i}`);
          }
        }

        forecasts[numCol] = {
          labels: [...grouped.labels, ...futureLabels],
          historicalValues: grouped.values,
          forecastValues: [...grouped.values, ...result.forecast],
          confidenceLower: [...grouped.values, ...result.lower],
          confidenceUpper: [...grouped.values, ...result.upper],
          model: result.model
        };
      }
    });
  }

  return {
    totalRows,
    totalColumns: headers.length,
    duplicateRows,
    headers,
    columns: columnStats,
    topDuplicateRows,
    correlations,
    forecasts
  };
}

export function buildAnalyticsWorkbook(fileName: string, data: DataRow[], profile: DataProfile) {
  const wb = XLSX.utils.book_new();

  const numVal = (v: any): number => {
    if (v == null) return 0;
    const n = typeof v === "number" ? v : parseFloat(String(v).replace(/[,₹\s%]/g, ""));
    return Number.isFinite(n) ? n : 0;
  };

  const cellCur = (v: any) => ({ v: numVal(v), t: "n", z: '"₹"#,##0.00' });
  const cellPct = (v: number) => ({ v: isNaN(v) ? 0 : v, t: "n", z: '0.0%' });
  const cellNum = (v: any) => ({ v: Math.round(numVal(v)), t: "n", z: '#,##0' });

  const isCurrencyCol = (name: string) => /revenue|amount|total|price|sale/i.test(name);
  
  const cellVal = (v: any, name: string) => {
    if (v == null || v === "") return "";
    const n = numVal(v);
    return isCurrencyCol(name) ? { v: n, t: "n", z: '"₹"#,##0.00' } : { v: n, t: "n", z: '#,##0.00' };
  };

  const formatCurrency = (v: any) => {
    if (v == null || v === "") return "";
    const n = numVal(v);
    return `₹${n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  };

  function findHeaderByPatterns(headers: string[], patterns: RegExp[]) {
    const norm = (s: string) => s.toLowerCase();
    for (const p of patterns) {
      const found = headers.find((h) => p.test(norm(h)));
      if (found) return found;
    }
    return null;
  }

  function sumColumn(colName: string | null) {
    if (!colName) return 0;
    return data.reduce((acc, row) => {
      const raw = row[colName];
      const n = Number(String(raw ?? "").replace(/[,₹\s%]/g, ""));
      return acc + (Number.isFinite(n) ? n : 0);
    }, 0);
  }

  function getMonthlyTrends(dateCol: string | null, revCol: string | null) {
    if (!dateCol || !revCol) return [];
    const map: Record<string, number> = {};
    for (const row of data) {
      const dval = row[dateCol];
      const rev = Number(String(row[revCol] ?? "").replace(/[,₹\s%]/g, ""));
      const dt = new Date(dval);
      if (isNaN(dt.getTime())) continue;
      const key = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, "0")}`;
      map[key] = (map[key] || 0) + (Number.isFinite(rev) ? rev : 0);
    }
    return Object.entries(map).sort(([a], [b]) => a.localeCompare(b)).map(([k, v]) => [k, v]);
  }

  function buildDashboard() {
    const headers = profile.headers || [];
    const productCol = findHeaderByPatterns(headers, [/product|item|title|name|sku/]);
    const revenueCol = findHeaderByPatterns(headers, [/revenue|amount|order total|order_total|total|price|sale/]) || (profile.columns.find((c) => c.type === "numeric" && (c.sum || 0) > 0)?.name ?? null);
    const dateCol = findHeaderByPatterns(headers, [/date|created|ordered|order date|ship/]);

    const totalRevenue = sumColumn(revenueCol);
    const avgRowRevenue = data.length ? totalRevenue / data.length : 0;
    const sampleTopProducts = [] as Array<[string, number]>;
    if (productCol && revenueCol) {
      const prodMap: Record<string, number> = {};
      for (const row of data) {
        const p = String(row[productCol] ?? "").trim() || "(unknown)";
        const r = Number(String(row[revenueCol] ?? "").replace(/[,₹\s%]/g, ""));
        prodMap[p] = (prodMap[p] || 0) + (Number.isFinite(r) ? r : 0);
      }
      sampleTopProducts.push(...Object.entries(prodMap).sort((a, b) => b[1] - a[1]).slice(0, 8));
    }

    const monthly = getMonthlyTrends(dateCol, revenueCol);
    const rows: any[] = [];

    rows.push([`📊 ${fileName} — Deep Dive Dashboard`, null, null, null]);
    rows.push([`Source sheet: ${profile.sheetName || "Auto-detected"}`, null, null, null]);
    rows.push([`Header row: ${profile.headerRow != null ? profile.headerRow + 1 : "Unknown"}`, null, null, null]);
    rows.push([`Generated: ${new Date().toLocaleString()}`, null, null, null]);
    rows.push([]);

    rows.push(["KEY METRICS", null, "TOP PRODUCTS", null]);
    rows.push(["Metric", "Value", "Product", "Revenue"]);
    rows.push(["Total Rows", cellNum(profile.totalRows), ...(sampleTopProducts[0] ? [sampleTopProducts[0][0], cellCur(sampleTopProducts[0][1])] : ["", ""])]);
    rows.push(["Total Columns", cellNum(profile.totalColumns), ...(sampleTopProducts[1] ? [sampleTopProducts[1][0], cellCur(sampleTopProducts[1][1])] : ["", ""])]);
    rows.push(["Duplicate Rows", cellNum(profile.duplicateRows), ...(sampleTopProducts[2] ? [sampleTopProducts[2][0], cellCur(sampleTopProducts[2][1])] : ["", ""])]);
    rows.push(["Total Revenue", cellCur(totalRevenue), ...(sampleTopProducts[3] ? [sampleTopProducts[3][0], cellCur(sampleTopProducts[3][1])] : ["", ""])]);
    rows.push(["Avg per Row", cellCur(avgRowRevenue), ...(sampleTopProducts[4] ? [sampleTopProducts[4][0], cellCur(sampleTopProducts[4][1])] : ["", ""])]);
    rows.push([]);

    rows.push(["MONTHLY REVENUE TRENDS", null, "", null]);
    rows.push(["Month", "Revenue", "", ""]);
    monthly.forEach(([month, revenue]) => rows.push([month, cellCur(revenue), "", ""]));
    rows.push([]);

    rows.push(["Notes", `Detected product column: ${productCol || "none"}`, null, null]);
    rows.push(["", `Detected revenue column: ${revenueCol || "none"}`, null, null]);
    rows.push(["", `Detected date column: ${dateCol || "none"}`, null, null]);

    return { rows, monthly, sampleTopProducts };
  }

  function buildTrendSheet(monthly: Array<[string, number]>, topProducts: any): any[][] {
    const rows: any[][] = [];
    rows.push(["MONTHLY REVENUE TRENDS"]);
    rows.push(["Month", "Revenue"]);
    monthly.forEach(([month, revenue]) => rows.push([month, cellCur(revenue)]));
    rows.push([]);
    rows.push(["TOP PRODUCTS BY REVENUE"]);
    rows.push(["Product", "Revenue"]);
    topProducts.forEach(([product, revenue]) => rows.push([product, cellCur(revenue)]));
    return rows;
  }

  const rawSummary = [
    ["Deep Dive Analytics Report"],
    [`Source file: ${fileName}`],
    [`Generated: ${new Date().toLocaleString()}`],
    [`Detected sheet: ${profile.sheetName || "Unknown"}`],
    [`Header row: ${profile.headerRow != null ? profile.headerRow + 1 : "Unknown"}`],
    [],
    ["Total Rows", cellNum(profile.totalRows)],
    ["Total Columns", cellNum(profile.totalColumns)],
    ["Duplicate Rows", cellNum(profile.duplicateRows)],
    ["Columns with empty values", cellNum(profile.columns.filter((col) => col.emptyCount > 0).length)],
  ];
  // Dashboard sheet similar to samples
  try {
    const { rows: dashboardRows, monthly, sampleTopProducts } = buildDashboard();
    const dashSheet = XLSX.utils.aoa_to_sheet(dashboardRows);
    dashSheet["!cols"] = [{ wch: 34 }, { wch: 16 }, { wch: 26 }, { wch: 18 }];
    dashSheet["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 3 } },
      { s: { r: 2, c: 0 }, e: { r: 2, c: 3 } },
      { s: { r: 3, c: 0 }, e: { r: 3, c: 3 } },
      { s: { r: 5, c: 0 }, e: { r: 5, c: 1 } },
      { s: { r: 5, c: 2 }, e: { r: 5, c: 3 } },
      { s: { r: 10, c: 0 }, e: { r: 10, c: 1 } },
    ];

    const titleCell = dashSheet["A1"];
    if (titleCell) titleCell.s = { font: { bold: true, sz: 16, color: { rgb: "FFFFFFFF" } }, fill: { fgColor: { rgb: "FF1F2937" } } };
    ["A6", "C6", "A11"].forEach((address) => {
      const cell = dashSheet[address];
      if (cell) cell.s = { font: { bold: true, color: { rgb: "FFFFFFFF" } }, fill: { fgColor: { rgb: "FF2563EB" } } };
    });

    XLSX.utils.book_append_sheet(wb, dashSheet, "Dashboard");

    if (monthly.length) {
      const trendAoA: any[][] = [
        ["MONTHLY REVENUE TRENDS"],
        ["Month", "Revenue"],
        ...monthly.map(([month, revenue]) => [month, cellCur(revenue)]),
        [],
        ["TOP PRODUCTS BY REVENUE"],
        ["Product", "Revenue"],
        ...sampleTopProducts.slice(0, 8).map(([product, revenue]) => [product, cellCur(revenue)]),
      ];
      const trendSheet = XLSX.utils.aoa_to_sheet(trendAoA);
      trendSheet["!cols"] = [{ wch: 24 }, { wch: 18 }];
      trendSheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }, { s: { r: monthly.length + 2, c: 0 }, e: { r: monthly.length + 2, c: 1 } }];
      const title = trendSheet["A1"];
      if (title) title.s = { font: { bold: true, sz: 14, color: { rgb: "FF111827" } } };
      XLSX.utils.book_append_sheet(wb, trendSheet, "Trends");
    }
  } catch (err) {
    // fallback: continue
  }

  const summarySheet = XLSX.utils.aoa_to_sheet(rawSummary);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");

  const columnHeader = [
    "Column",
    "Type",
    "Non-empty",
    "Empty",
    "Unique",
    "Sum",
    "Avg",
    "Median",
    "Std Dev",
    "Min",
    "Max",
    "Top values",
    "Sample values",
  ];
  const columnRows = profile.columns.map((col) => [
    col.name,
    col.type,
    cellNum(col.nonEmptyCount),
    cellNum(col.emptyCount),
    cellNum(col.uniqueCount),
    col.sum != null ? cellVal(col.sum, col.name) : "",
    col.avg != null ? cellVal(col.avg, col.name) : "",
    col.median != null ? cellVal(col.median, col.name) : "",
    col.stddev != null ? cellNum(col.stddev) : "",
    col.min != null ? cellVal(col.min, col.name) : "",
    col.max != null ? cellVal(col.max, col.name) : "",
    col.topValues.map((item) => `${item.value} (${item.count})`).join("; "),
    col.sampleValues.map((value) => String(value)).join("; "),
  ]);
  const columnSheet = XLSX.utils.aoa_to_sheet([columnHeader, ...columnRows]);
  XLSX.utils.book_append_sheet(wb, columnSheet, "Column Profile");

  const qualityRows = [
    ["Column", "Type", "Non-empty", "Empty", "Empty %", "Unique", "Unique %", "Sample values"],
    ...profile.columns.map((col) => [
      col.name,
      col.type,
      cellNum(col.nonEmptyCount),
      cellNum(col.emptyCount),
      col.count ? cellPct(col.emptyCount / col.count) : "",
      cellNum(col.uniqueCount),
      col.count ? cellPct(col.uniqueCount / col.count) : "",
      col.sampleValues.map((value) => String(value)).join("; "),
    ]),
  ];
  const qualitySheet = XLSX.utils.aoa_to_sheet(qualityRows);
  XLSX.utils.book_append_sheet(wb, qualitySheet, "Data Quality");

  const topValuesRows = [
    ["Column", "Value", "Count"],
    ...profile.columns.flatMap((col) => col.topValues.slice(0, 8).map((item) => [col.name, item.value, cellNum(item.count)])),
  ];
  const topValuesSheet = XLSX.utils.aoa_to_sheet(topValuesRows);
  XLSX.utils.book_append_sheet(wb, topValuesSheet, "Top Values");

  const numericColumns = profile.columns.filter((col) => col.type === "numeric");
  if (numericColumns.length) {
    const numericSheet = XLSX.utils.aoa_to_sheet([
      ["Column", "Min", "Max", "Sum", "Avg", "Median", "Std Dev", "Unique"],
      ...numericColumns.map((col) => [
        col.name,
        col.min != null ? cellVal(col.min, col.name) : "",
        col.max != null ? cellVal(col.max, col.name) : "",
        col.sum != null ? cellVal(col.sum, col.name) : "",
        col.avg != null ? cellVal(col.avg, col.name) : "",
        col.median != null ? cellVal(col.median, col.name) : "",
        col.stddev != null ? cellNum(col.stddev) : "",
        cellNum(col.uniqueCount),
      ]),
    ]);
    XLSX.utils.book_append_sheet(wb, numericSheet, "Numeric Summary");
  }

  const sampleLimit = Math.min(20, data.length);
  if (sampleLimit > 0) {
    const sampleRows = [profile.headers, ...data.slice(0, sampleLimit).map((row) => profile.headers.map((header) => row[header] ?? ""))];
    const sampleSheet = XLSX.utils.aoa_to_sheet(sampleRows);
    XLSX.utils.book_append_sheet(wb, sampleSheet, "Sample Data");
  }

  if (profile.topDuplicateRows.length) {
    const dupRows = [
      ["Duplicate Row JSON", "Duplicate Count"],
      ...profile.topDuplicateRows.map((item) => [item.row, cellNum(item.count)]),
    ];
    const dupSheet = XLSX.utils.aoa_to_sheet(dupRows);
    XLSX.utils.book_append_sheet(wb, dupSheet, "Duplicate Rows");
  }

  return wb;
}

export function buildLogisticsWorkbook(fileName: string, data: any[], mergedCount: number, analytics: any): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  const { total, totalRev, totalQty, totalCOD, totalFreight, delivered, rto, deliveryRate, rtoRate, statusCounts, pickupCounts, qtyCounts, courierCounts, zoneCounts, stateCounts, ndrCounts, payCounts } = analytics;

  function addSheet(name: string, aoa: any[][], colWidths?: number[]) {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    if (colWidths) ws["!cols"] = colWidths.map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  }

  const numVal = (v: any): number => {
    const n = parseFloat(v);
    return Number.isFinite(n) ? n : 0;
  };

  const sorted = (obj: any, key = "orders") => Object.entries(obj).map(([k, v]: [string, any]) => ({ key: k, ...v })).sort((a, b) => b[key] - a[key]);

  const cellCur = (v: any) => ({ v: numVal(v), t: "n", z: '"₹"#,##0.00' });
  const cellPct = (v: number) => ({ v: isNaN(v) ? 0 : v, t: "n", z: '0.0%' });
  const cellNum = (v: any) => ({ v: Math.round(numVal(v)), t: "n", z: '#,##0' });
  const cellDec = (v: any) => ({ v: numVal(v), t: "n", z: '#,##0.00' });

  const cod = payCounts["cod"] || {};
  const pre = payCounts["prepaid"] || {};

  addSheet("📊 Executive Summary", [
    [`${fileName} — Shiprocket Analytics Report`],
    [`Generated: ${new Date().toLocaleString("en-IN")} | Original Rows: ${data.length + mergedCount} → Merged: ${total} unique orders | ${mergedCount} duplicates resolved`],
    [],
    ["── KEY PERFORMANCE INDICATORS ──"],
    ["Metric", "Value", "", "Metric", "Value"],
    ["Total Orders", cellNum(total), "", "Delivery Rate", cellPct(deliveryRate)],
    ["Total Qty Shipped", cellNum(totalQty), "", "RTO Rate", cellPct(rtoRate)],
    ["Total Revenue", cellCur(totalRev), "", "COD Orders", cellNum(cod.orders || 0)],
    ["Delivered Orders", cellNum(delivered), "", "Prepaid Orders", cellNum(pre.orders || 0)],
    ["RTO Delivered", cellNum(rto), "", "COD Collected", cellCur(totalCOD)],
    ["Total Freight", cellCur(totalFreight), "", "Avg Order Value", cellCur(totalRev / total)],
    [],
    ["── STATUS WISE SUMMARY ──"],
    ["Status", "Orders", "Total Qty", "Revenue (₹)", "Avg Value (₹)", "COD (₹)", "% of Orders", "% of Revenue"],
    ...sorted(statusCounts).map((r: any) => [r.key, cellNum(r.orders), cellNum(r.qty), cellCur(r.revenue), cellCur(r.revenue / r.orders), cellCur(r.cod), cellPct(r.orders / total), cellPct(r.revenue / totalRev)]),
    [],
    ["── PAYMENT METHOD BREAKDOWN ──"],
    ["Payment Method", "Orders", "Revenue (₹)", "COD Collected (₹)", "% of Orders"],
    ...sorted(payCounts).map((r: any) => [r.key, cellNum(r.orders), cellCur(r.revenue), cellCur(r.cod), cellPct(r.orders / total)]),
  ], [30, 18, 4, 28, 18]);

  addSheet("📦 Status Wise", [
    ["STATUS-WISE ORDER ANALYTICS"],
    [],
    ["Status", "Orders", "Total Qty", "Revenue (₹)", "Avg Value (₹)", "Freight (₹)", "COD Orders", "Prepaid Orders", "COD Collected (₹)", "Delivery %", "RTO %"],
    ...sorted(statusCounts).map((r: any) => [r.key, cellNum(r.orders), cellNum(r.qty), cellCur(r.revenue), cellCur(r.revenue / r.orders), cellCur(r.freight), cellNum(r.codOrders), cellNum(r.prepaid), cellCur(r.cod), cellPct(r.delivered / r.orders), cellPct(r.rto / r.orders)]),
  ], [26, 10, 10, 18, 18, 14, 14, 16, 20, 12, 10]);

  addSheet("🏢 Pickup Address Wise", [
    ["PICKUP ADDRESS WISE ANALYTICS"],
    [],
    ["Pickup Address", "Orders", "Total Qty", "Revenue (₹)", "Delivered", "RTO", "Canceled", "Delivery %", "RTO %", "COD Collected (₹)"],
    ...sorted(pickupCounts).map((r: any) => [r.key, cellNum(r.orders), cellNum(r.qty), cellCur(r.revenue), cellNum(r.delivered), cellNum(r.rto), cellNum(r.canceled), cellPct(r.delivered / r.orders), cellPct(r.rto / r.orders), cellCur(r.cod)]),
  ], [34, 10, 10, 18, 11, 10, 10, 12, 10, 20]);

  const qtyRows = Object.entries(qtyCounts).sort((a, b) => parseFloat(a[0]) - parseFloat(b[0])).map(([k, v]: [string, any]) => [parseFloat(k), cellNum(v.orders), cellCur(v.revenue), cellNum(v.delivered), cellNum(v.rto), cellPct(v.delivered / v.orders), cellPct(v.rto / v.orders), cellCur(v.revenue / v.orders), cellCur(v.cod)]);
  addSheet("📦 Qty Wise", [
    ["QUANTITY-WISE ORDER ANALYTICS"],
    [],
    ["Qty / Order", "Orders", "Revenue (₹)", "Delivered", "RTO", "Delivery %", "RTO %", "Avg Revenue (₹)", "COD Collected (₹)"],
    ...qtyRows,
  ], [14, 10, 18, 11, 10, 12, 10, 18, 18]);

  const courierRows = sorted(courierCounts).map((r: any) => [r.key, cellNum(r.orders), cellNum(r.delivered), cellNum(r.rto), cellNum(r.canceled), cellPct(r.delivered / r.orders), cellPct(r.rto / r.orders), cellCur(r.revenue), cellCur(r.revenue / r.orders), cellCur(r.freight), cellCur(r.cod)]);
  const zoneMap: Record<string, string> = { z_a: "Zone A (Local)", z_b: "Zone B", z_c: "Zone C", z_d: "Zone D", z_e: "Zone E (Far)" };
  const zoneRows = sorted(zoneCounts).map((r: any) => [zoneMap[r.key] || r.key, cellNum(r.orders), cellNum(r.delivered), cellNum(r.rto), cellPct(r.delivered / r.orders), cellPct(r.rto / r.orders), cellCur(r.revenue), cellCur(r.cod), cellCur(r.freight)]);
  const stateRows = sorted(stateCounts).map((r: any) => [r.key, cellNum(r.orders), cellNum(r.delivered), cellNum(r.rto), cellNum(r.canceled), cellPct(r.delivered / r.orders), cellPct(r.rto / r.orders), cellCur(r.revenue), cellCur(r.cod)]);
  const ndrTotal: any = Object.values(ndrCounts).reduce((acc: number, val: any) => acc + Number(val), 0) || 1;
  const ndrRows = Object.entries(ndrCounts).sort((a: any, b: any) => b[1] - a[1]).map(([k, v]: [string, any]) => [k, cellNum(v), cellPct(Number(v) / ndrTotal)]);



  addSheet("🔍 Deep Analytics", [
    ["DEEP ANALYTICS — COURIER / ZONE / STATE / NDR"],
    [],
    ["── COURIER COMPANY PERFORMANCE ──"],
    ["Courier", "Orders", "Delivered", "RTO", "Canceled", "Delivery %", "RTO %", "Revenue (₹)", "Avg Value (₹)", "Freight (₹)", "COD (₹)"],
    ...courierRows,
    [],
    ["── ZONE-WISE ANALYSIS ──"],
    ["Zone", "Orders", "Delivered", "RTO", "Delivery %", "RTO %", "Revenue (₹)", "COD (₹)", "Freight (₹)"],
    ...zoneRows,
    [],
    ["── STATE-WISE ANALYSIS ──"],
    ["State", "Orders", "Delivered", "RTO", "Canceled", "Delivery %", "RTO %", "Revenue (₹)", "COD (₹)"],
    ...stateRows,
    [],
    ["── NDR (NON-DELIVERY REPORT) REASONS ──"],
    ["NDR Reason", "Count", "% Share"],
    ...ndrRows,
  ], [32, 10, 11, 10, 10, 12, 10, 18, 16, 14, 14]);

  const highRTOStates = sorted(stateCounts).filter((r: any) => r.orders >= 15).sort((a: any, b: any) => (b.rto / b.orders) - (a.rto / a.orders));
  const highRTOCouriers = sorted(courierCounts).filter((r: any) => r.orders >= 10).sort((a: any, b: any) => (b.rto / b.orders) - (a.rto / a.orders));

  addSheet("⚠️ RTO Risk", [
    ["RTO RISK & PERFORMANCE ANALYSIS"],
    [],
    ["── HIGH RTO STATES (min 15 orders) ──"],
    ["State", "Orders", "RTO", "Delivered", "RTO %", "Revenue (₹)"],
    ...highRTOStates.map((r: any) => [r.key, cellNum(r.orders), cellNum(r.rto), cellNum(r.delivered), cellPct(r.rto / r.orders), cellCur(r.revenue)]),
    [],
    ["── HIGH RTO COURIERS (min 10 orders) ──"],
    ["Courier", "Orders", "RTO", "Delivered", "RTO %"],
    ...highRTOCouriers.map((r: any) => [r.key, cellNum(r.orders), cellNum(r.rto), cellNum(r.delivered), cellPct(r.rto / r.orders)]),
  ], [26, 10, 10, 11, 10, 18]);

  const exportCols = ["Order ID", "Is Multi-SKU", "SKU Count", "Status", "Pickup Address Name", "Courier Company", "AWB Code", "Zone", "Payment Method", "Product Quantity", "Order Total", "COD Payble Amount", "Freight Total Amount", "Product Name", "Address State", "Address City", "Customer Name", "Shiprocket Created At"];
  const presentCols = exportCols.filter((c) => data[0] && c in data[0]);

  const formattedDataRows = data.map((r) => presentCols.map((c) => {
    const val = r[c];
    if (val == null || val === "") return "";
    if (c === "Order Total" || c === "COD Payble Amount" || c === "Freight Total Amount") {
      return cellCur(val);
    }
    if (c === "Product Quantity" || c === "SKU Count") {
      return cellNum(val);
    }
    return val;
  }));

  // 3. Freight & Weight Billing Audit Sheet
  const auditData = auditLogisticsData(data);
  const auditRows = auditData.map((r) => {
    const weightAnomalyText = r.isWeightAnomaly ? "⚠️ OVERBILLED (Ratio > 1.5x)" : "✅ Normal";
    const freightLeakText = r.isFreightLeak ? "⚠️ LEAKAGE (Freight > 30%)" : "✅ Normal";
    let recommendation = "✅ Passed Audit";
    if (r.isWeightAnomaly && r.isFreightLeak) {
      recommendation = "🚨 CRITICAL: Dispute weight with courier AND review shipping zones!";
    } else if (r.isWeightAnomaly) {
      recommendation = "⚠️ DISPUTE WEIGHT: Charged weight significantly exceeds physical weight!";
    } else if (r.isFreightLeak) {
      recommendation = "⚠️ REVIEW PRICING: High shipping fee relative to order value!";
    }
    
    return [
      r.orderId,
      r.courier,
      cellDec(r.chargedWeight),
      cellDec(r.physicalWeight),
      cellPct(r.weightRatio),
      weightAnomalyText,
      cellCur(r.freightCost),
      cellCur(r.orderRevenue),
      cellPct(r.freightRatio),
      freightLeakText,
      r.rtoStatus,
      r.state,
      recommendation
    ];
  });

  addSheet("🚨 Billing Audit", [
    ["🚨 Freight Cost & Courier Weight Overcharge Audit Ledger"],
    ["Automated logistics auditor mapping weight discrepancies (Charged > 1.5x Physical) and freight leakage (Freight > 30% of revenue)."],
    [],
    [
      "Order ID",
      "Courier Company",
      "Charged Weight (kg)",
      "Physical Weight (kg)",
      "Weight Ratio",
      "Weight Audit",
      "Freight Cost (₹)",
      "Order Revenue (₹)",
      "Freight % of Revenue",
      "Freight Audit",
      "RTO Status",
      "State",
      "Operational Action Recommendation"
    ],
    ...auditRows
  ], [16, 20, 18, 18, 14, 24, 18, 18, 18, 24, 16, 16, 48]);

  addSheet("📋 Merged Clean Data", [
    presentCols,
    ...formattedDataRows,
  ], presentCols.map((c) => c === "Product Name" ? 65 : c === "Order ID" ? 12 : 18));

  // 4. Courier Claims Ledger Sheet
  const claimsRows = auditData.filter(r => r.isWeightAnomaly).map((r) => {
    const claimLetter = generateDisputeLetterText(r);
    return [
      r.orderId,
      r.courier,
      cellDec(r.chargedWeight),
      cellDec(r.physicalWeight),
      cellPct(r.weightRatio),
      cellCur(r.freightCost),
      claimLetter
    ];
  });
  addSheet("Courier Claims Ledger", [
    ["📋 Automated Courier Billing Claims & Weight Dispute Letters"],
    ["Pre-drafted dispute claim letters ready to copy-paste for all orders flagged with weight anomalies."],
    [],
    ["Order ID", "Courier Company", "Charged Weight (kg)", "Physical Weight (kg)", "Weight Discrepancy", "Freight Cost Charged", "Legally Assertive Claim Dispute Letter"],
    ...claimsRows
  ], [16, 20, 20, 20, 18, 20, 120]);

  return wb;
}

export type MarketBasketRule = {
  itemA: string;
  itemB: string;
  support: number;
  confidence: number;
  lift: number;
  coPurchaseCount: number;
};

export type CustomerRfmProfile = {
  key?: string;
  customerId: string;
  customerName: string;
  email: string;
  phone: string;
  recencyDays: number;
  frequency: number;
  monetary: number;
  rScore: number;
  fScore: number;
  mScore: number;
  cohort: "Champions" | "Loyal Shoppers" | "Recent Starters" | "At Risk" | "Lost" | "Unknown";
};

export type LogisticsLeakageAnomaly = {
  orderId: string;
  courier: string;
  chargedWeight: number;
  physicalWeight: number;
  weightRatio: number;
  freightCost: number;
  orderRevenue: number;
  freightRatio: number;
  isFreightLeak: boolean;
  isWeightAnomaly: boolean;
  rtoStatus: string;
  state: string;
};

export type CohortRetentionData = {
  cohortMonth: string;
  totalCustomers: number;
  months: number[];
  rates: number[];
};

export type OrderRtoRisk = {
  orderId: string;
  customerName: string;
  paymentMethod: string;
  state: string;
  riskScore: number;
  riskLevel: "High" | "Medium" | "Low";
};

export type ShopifyAnalyticsSummary = {
  totalRows: number;
  totalOrders: number;
  totalCustomers: number;
  totalRevenue: number;
  totalUnits: number;
  productCount: number;
  topProduct: string;
  topCity: string;
  segmentCounts: Record<string, number>;
  statusCounts: Record<string, number>;
  aprioriRules?: MarketBasketRule[];
  rfmMatrix?: CustomerRfmProfile[];
  cohortRetention?: CohortRetentionData[];
};

type ShopifyOrderLine = {
  orderNo: string;
  date: Date | null;
  customerName: string;
  email: string;
  phone: string;
  financialStatus: string;
  fulfillmentStatus: string;
  orderStatus: string;
  total: number;
  paymentMethod: string;
  discount: number;
  refunded: number;
  city: string;
  state: string;
  country: string;
  risk: string;
  product: string;
  qty: number;
  unitPrice: number;
  lineRevenue: number;
  discountCode: string;
  emailMarketing: string;
  smsMarketing: string;
  customerId: string;
  address: string;
};

function toNumber(value: any): number {
  if (typeof value === "number") return Number.isFinite(value) ? value : 0;
  const n = parseFloat(String(value ?? "").replace(/[^0-9.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

function toText(value: any): string {
  return String(value ?? "").trim();
}

function parseShopifyDate(value: any): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const raw = toText(value);
  if (!raw) return null;
  const normalized = raw.replace(/ ([+-]\d{2})(\d{2})$/, " $1:$2");
  const parsed = new Date(normalized);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function monthKey(date: Date | null): string {
  if (!date) return "Unknown";
  return date.toLocaleString("en-US", { month: "short", year: "numeric" });
}

function cleanProductName(value: string): string {
  return value
    .replace(/^CureForever\s*/i, "")
    .replace(/\s+/g, " ")
    .trim() || "Unknown Product";
}

function deriveOrderStatus(financial: string, fulfillment: string, payment: string, cancelled: string, refunded: number): string {
  const fin = financial.toLowerCase();
  const fulf = fulfillment.toLowerCase();
  const pay = payment.toLowerCase();
  if (cancelled || fin.includes("void")) return "Cancelled";
  if (refunded > 0 || fin.includes("refund")) return "Refunded";
  if (fin.includes("partial")) return "Partially Paid";
  if (fulf.includes("fulfilled") && !fulf.includes("unfulfilled")) {
    if (pay.includes("cod") || fin.includes("pending")) return "Fulfilled - COD Pending";
    return "Delivered";
  }
  if (pay.includes("cod") || fin.includes("pending")) return "COD - Awaiting Delivery";
  if (fin.includes("paid")) return "Paid - Pending Fulfillment";
  return financial || fulfillment || "Unknown";
}

function segmentCustomer(totalSpent: number, totalOrders: number): string {
  if (totalSpent >= 15000 || totalOrders >= 10) return "VIP";
  if (totalSpent >= 3000 || totalOrders >= 3) return "High Value";
  if (totalOrders >= 2) return "Repeat Buyer";
  if (totalSpent > 0) return "First-Time Buyer";
  return "No Spend";
}

function mapShopifyRows(data: DataRow[]): ShopifyOrderLine[] {
  return data.map((row) => {
    const financial = toText(row["Financial Status"]);
    const fulfillment = toText(row["Fulfillment Status"] || row["Lineitem fulfillment status"]);
    const payment = toText(row["Payment Method"]);
    const refunded = toNumber(row["Refunded Amount"]);
    const qty = toNumber(row["Lineitem quantity"]) || 1;
    const unitPrice = toNumber(row["Lineitem price"]);
    const product = cleanProductName(toText(row["Lineitem name"]));
    const orderNo = toText(row["Name"] || row["Order #"]);
    const email = toText(row["Email"]).toLowerCase();
    const phone = toText(row["Phone"] || row["Billing Phone"] || row["Shipping Phone"]);
    const customerName = toText(row["Shipping Name"] || row["Billing Name"]);
    const cancelled = toText(row["Cancelled at"]);
    return {
      orderNo,
      date: parseShopifyDate(row["Created at"]),
      customerName,
      email,
      phone,
      financialStatus: financial,
      fulfillmentStatus: fulfillment,
      orderStatus: deriveOrderStatus(financial, fulfillment, payment, cancelled, refunded),
      total: toNumber(row["Total"]),
      paymentMethod: payment || "Unknown",
      discount: toNumber(row["Discount Amount"]) + toNumber(row["Lineitem discount"]),
      refunded,
      city: toText(row["Shipping City"] || row["Billing City"]) || "Unknown",
      state: toText(row["Shipping Province"] || row["Billing Province"] || row["Shipping Province Name"] || row["Billing Province Name"]) || "Unknown",
      country: toText(row["Shipping Country"] || row["Billing Country"]) || "Unknown",
      risk: toText(row["Risk Level"]) || "Low",
      product,
      qty,
      unitPrice,
      lineRevenue: qty * unitPrice,
      discountCode: toText(row["Discount Code"]),
      emailMarketing: toText(row["Accepts Marketing"]).toLowerCase() === "yes" ? "yes" : "no",
      smsMarketing: phone ? "yes" : "no",
      customerId: toText(row["Id"]),
      address: toText(row["Shipping Address1"] || row["Billing Address1"] || row["Shipping Street"] || row["Billing Street"]),
    };
  }).filter((line) => line.orderNo || line.email || line.product !== "Unknown Product");
}

function uniqueOrders(lines: ShopifyOrderLine[]): ShopifyOrderLine[] {
  const byOrder = new Map<string, ShopifyOrderLine>();
  for (const line of lines) {
    const key = line.orderNo || `${line.email}-${line.date?.toISOString() || ""}`;
    if (!byOrder.has(key)) byOrder.set(key, line);
  }
  return [...byOrder.values()];
}

function incrementMetric(map: Record<string, any>, key: string, patch: Record<string, number>) {
  const safeKey = key || "Unknown";
  if (!map[safeKey]) map[safeKey] = {};
  for (const [field, value] of Object.entries(patch)) {
    map[safeKey][field] = (map[safeKey][field] || 0) + value;
  }
}

function sortedEntries<T = any>(obj: Record<string, T>, metric?: string): Array<[string, any]> {
  return Object.entries(obj).sort((a: any, b: any) => {
    const av = metric ? a[1][metric] || 0 : a[1] || 0;
    const bv = metric ? b[1][metric] || 0 : b[1] || 0;
    return bv - av;
  });
}

export function mineAssociationRules(lines: ShopifyOrderLine[], totalOrders: number): MarketBasketRule[] {
  if (totalOrders === 0) return [];
  const orderItems = new Map<string, Set<string>>();
  const itemCounts: Record<string, number> = {};

  for (const line of lines) {
    const key = line.orderNo || `${line.email}-${line.date?.toISOString() || ""}`;
    if (!key) continue;
    if (!orderItems.has(key)) orderItems.set(key, new Set());
    const prod = line.product;
    orderItems.get(key)!.add(prod);
    itemCounts[prod] = (itemCounts[prod] || 0) + line.qty;
  }

  const pairCounts: Record<string, number> = {};
  orderItems.forEach((items) => {
    const arr = Array.from(items);
    for (let i = 0; i < arr.length; i++) {
      for (let j = i + 1; j < arr.length; j++) {
        const itemA = arr[i];
        const itemB = arr[j];
        const pairKey = itemA < itemB ? `${itemA}|||${itemB}` : `${itemB}|||${itemA}`;
        pairCounts[pairKey] = (pairCounts[pairKey] || 0) + 1;
      }
    }
  });

  const rules: MarketBasketRule[] = [];
  for (const [pairKey, count] of Object.entries(pairCounts)) {
    if (count < 2) continue;
    const [itemA, itemB] = pairKey.split("|||");
    const countA = itemCounts[itemA] || 1;
    const countB = itemCounts[itemB] || 1;

    const support = count / totalOrders;
    const confidenceA = count / countA;
    const lift = support / ((countA / totalOrders) * (countB / totalOrders));

    rules.push({
      itemA,
      itemB,
      support,
      confidence: confidenceA,
      lift,
      coPurchaseCount: count
    });
  }

  return rules
    .filter(r => r.lift > 1.0)
    .sort((a, b) => b.lift - a.lift)
    .slice(0, 10);
}

export function computeRfmSegmentation(lines: ShopifyOrderLine[]): CustomerRfmProfile[] {
  const customerMap: Record<string, {
    name: string;
    email: string;
    phone: string;
    spent: number;
    orders: number;
    lastOrderDate: Date;
    customerId: string;
  }> = {};

  let maxDateTime = -Infinity;

  for (const line of lines) {
    const key = line.email || line.phone || line.customerName || line.orderNo;
    if (!key) continue;

    const lineDate = line.date || new Date();
    if (lineDate.getTime() > maxDateTime) {
      maxDateTime = lineDate.getTime();
    }

    if (!customerMap[key]) {
      customerMap[key] = {
        name: line.customerName || "Customer",
        email: line.email || "",
        phone: line.phone || "",
        spent: 0,
        orders: 0,
        lastOrderDate: lineDate,
        customerId: line.customerId || "",
      };
    }

    customerMap[key].orders += 1;
    customerMap[key].spent += line.lineRevenue;
    if (lineDate.getTime() > customerMap[key].lastOrderDate.getTime()) {
      customerMap[key].lastOrderDate = lineDate;
    }
  }

  const maxDate = maxDateTime === -Infinity ? new Date() : new Date(maxDateTime);
  const profiles = Object.entries(customerMap).map(([key, c]) => {
    const diffTime = Math.max(0, maxDate.getTime() - c.lastOrderDate.getTime());
    const recencyDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
    return {
      key,
      customerId: c.customerId,
      customerName: c.name,
      email: c.email,
      phone: c.phone,
      recencyDays,
      frequency: c.orders,
      monetary: c.spent,
      rScore: 1,
      fScore: 1,
      mScore: 1,
      cohort: "Unknown" as CustomerRfmProfile["cohort"]
    };
  });

  if (profiles.length === 0) return [];

  const rSorted = [...profiles].sort((a, b) => a.recencyDays - b.recencyDays);
  const fSorted = [...profiles].sort((a, b) => b.frequency - a.frequency);
  const mSorted = [...profiles].sort((a, b) => b.monetary - a.monetary);

  const N = profiles.length;
  const getTertileScore = (rank: number, total: number) => {
    if (rank < total / 3) return 3;
    if (rank < (2 * total) / 3) return 2;
    return 1;
  };

  profiles.forEach(p => {
    const rRank = rSorted.findIndex(x => x.key === p.key);
    const fRank = fSorted.findIndex(x => x.key === p.key);
    const mRank = mSorted.findIndex(x => x.key === p.key);

    p.rScore = getTertileScore(rRank, N);
    p.fScore = getTertileScore(fRank, N);
    p.mScore = getTertileScore(mRank, N);

    const rfm = `${p.rScore}${p.fScore}${p.mScore}`;
    if (p.rScore === 3 && p.fScore === 3 && p.mScore === 3) {
      p.cohort = "Champions";
    } else if (p.rScore >= 2 && p.fScore >= 2 && p.mScore >= 2) {
      p.cohort = "Loyal Shoppers";
    } else if (p.rScore === 3 && p.fScore === 1) {
      p.cohort = "Recent Starters";
    } else if (p.rScore === 1 && (p.fScore >= 2 || p.mScore >= 2)) {
      p.cohort = "At Risk";
    } else {
      p.cohort = "Lost";
    }
  });

  return profiles;
}

export function auditLogisticsData(data: DataRow[]): LogisticsLeakageAnomaly[] {
  return data.map((r) => {
    const keys = Object.keys(r || {});
    const normalize = (key: string) => key.replace(/\s+/g, "").toLowerCase();
    const findKey = (pred: (k: string) => boolean) => keys.find((k) => !!k && pred(normalize(k))) as string | undefined;

    const idField = findKey((k) => k.includes("orderid") || (k.includes("order") && k.includes("id"))) || "Order ID";
    const courierField = findKey((k) => k.includes("courier.*company") || k.includes("courier") || k.includes("logisticspartner")) || "Courier Company";
    
    const chargedWeightField = findKey((k) => k.includes("chargedweight") || k.includes("billedweight") || k.includes("weightcharged")) || "Charged Weight";
    const physicalWeightField = findKey((k) => k.includes("physicalweight") || k.includes("actualweight") || k.includes("weightphysical")) || "Physical Weight";
    
    const freightField = findKey((k) => k.includes("freight.*total") || k.includes("shippingcharge") || k.includes("freight") || k.includes("couriercharge")) || "Freight Total Amount";
    const revenueField = findKey((k) => k.includes("order.*total") || k.includes("ordertotal") || k.includes("totalamount") || k.includes("invoiceamount") || k.includes("revenue")) || "Order Total";
    const stateField = findKey((k) => k.includes("address.*state") || k.includes("state") || k.includes("province")) || "State";
    const statusField = findKey((k) => k.includes("status") || k.includes("shipmentstatus") || k.includes("orderstatus")) || "Status";

    const orderId = String(r[idField] || "").trim();
    const courier = String(r[courierField] || "Unknown");
    const chargedWeight = toNumber(r[chargedWeightField]);
    const physicalWeight = toNumber(r[physicalWeightField]);
    const freightCost = toNumber(r[freightField]);
    const orderRevenue = toNumber(r[revenueField]);
    const state = String(r[stateField] || "Unknown");
    const rtoStatus = String(r[statusField] || "UNKNOWN");

    const weightRatio = physicalWeight > 0 ? chargedWeight / physicalWeight : 0;
    const freightRatio = orderRevenue > 0 ? freightCost / orderRevenue : 0;

    const isWeightAnomaly = physicalWeight > 0 && chargedWeight > 1.5 * physicalWeight;
    const isFreightLeak = orderRevenue > 0 && freightRatio > 0.30;

    return {
      orderId,
      courier,
      chargedWeight,
      physicalWeight,
      weightRatio,
      freightCost,
      orderRevenue,
      freightRatio,
      isFreightLeak,
      isWeightAnomaly,
      rtoStatus,
      state
    };
  }).filter(item => item.orderId);
}

export function calculateCohortRetention(lines: ShopifyOrderLine[]): CohortRetentionData[] {
  const customerFirstOrder: Record<string, Date> = {};
  const customerKeyToCohort: Record<string, string> = {};

  for (const line of lines) {
    const key = line.email || line.phone || line.customerName || line.orderNo;
    if (!key) continue;
    
    const lineDate = line.date || new Date();
    if (!customerFirstOrder[key] || lineDate.getTime() < customerFirstOrder[key].getTime()) {
      customerFirstOrder[key] = lineDate;
    }
  }

  // Assign cohort month (YYYY-MM)
  const cohortCustomers: Record<string, Set<string>> = {}; // cohortMonth -> Set of customerKeys
  Object.entries(customerFirstOrder).forEach(([key, firstDate]) => {
    const monthStr = `${firstDate.getFullYear()}-${String(firstDate.getMonth() + 1).padStart(2, "0")}`;
    customerKeyToCohort[key] = monthStr;
    if (!cohortCustomers[monthStr]) cohortCustomers[monthStr] = new Set();
    cohortCustomers[monthStr].add(key);
  });

  // cohortMonth -> Map of monthIndex -> Set of customerKeys who bought in that month
  const cohortActivity: Record<string, Record<number, Set<string>>> = {};

  for (const line of lines) {
    const key = line.email || line.phone || line.customerName || line.orderNo;
    if (!key) continue;

    const lineDate = line.date || new Date();
    const firstDate = customerFirstOrder[key];
    const cohortMonth = customerKeyToCohort[key];
    if (!firstDate || !cohortMonth) continue;

    const diffMonths = (lineDate.getFullYear() - firstDate.getFullYear()) * 12 + (lineDate.getMonth() - firstDate.getMonth());
    if (diffMonths < 0 || diffMonths > 12) continue;

    if (!cohortActivity[cohortMonth]) cohortActivity[cohortMonth] = {};
    if (!cohortActivity[cohortMonth][diffMonths]) cohortActivity[cohortMonth][diffMonths] = new Set();
    cohortActivity[cohortMonth][diffMonths].add(key);
  }

  const cohorts = Object.keys(cohortCustomers).sort();
  return cohorts.map((cohortMonth) => {
    const totalCustomers = cohortCustomers[cohortMonth].size;
    const months: number[] = [];
    const rates: number[] = [];

    for (let m = 0; m <= 5; m++) {
      const activeCount = cohortActivity[cohortMonth]?.[m]?.size || 0;
      months.push(activeCount);
      rates.push(totalCustomers > 0 ? (activeCount / totalCustomers) * 100 : 0);
    }

    return {
      cohortMonth,
      totalCustomers,
      months,
      rates
    };
  });
}

export function calculatePredictiveRtoRisk(data: DataRow[]): OrderRtoRisk[] {
  return data.map((r) => {
    const keys = Object.keys(r || {});
    const normalize = (key: string) => key.replace(/\s+/g, "").toLowerCase();
    const findKey = (pred: (k: string) => boolean) => keys.find((k) => !!k && pred(normalize(k))) as string | undefined;

    const idField = findKey((k) => k.includes("orderid") || (k.includes("order") && k.includes("id"))) || "Order ID";
    const nameField = findKey((k) => k.includes("customername") || k.includes("customer") || k.includes("consigneename")) || "Customer Name";
    const paymentField = findKey((k) => k.includes("paymentmethod") || k.includes("paymentmode") || k.includes("payment")) || "Payment Method";
    const stateField = findKey((k) => k.includes("address.*state") || k.includes("state") || k.includes("province")) || "State";

    const orderId = String(r[idField] || "").trim();
    const customerName = String(r[nameField] || "Customer");
    const paymentMethod = String(r[paymentField] || "Prepaid");
    const state = String(r[stateField] || "Unknown");

    let riskScore = 15;
    
    const payNorm = paymentMethod.toUpperCase();
    if (payNorm.includes("COD") || payNorm.includes("CASH")) {
      riskScore += 35;
    }

    const normalizedState = state.toLowerCase();
    if (
      normalizedState.includes("bihar") || 
      normalizedState.includes("uttar pradesh") || 
      normalizedState.includes("up") ||
      normalizedState.includes("west bengal") || 
      normalizedState.includes("bengal") ||
      normalizedState.includes("jharkhand") || 
      normalizedState.includes("assam") ||
      normalizedState.includes("northeast")
    ) {
      riskScore += 25;
    }

    const qtyField = findKey((k) => k.includes("productquantity") || k.includes("quantity") || k.includes("qty")) || "Product Quantity";
    const qty = toNumber(r[qtyField]);
    if (qty > 3) {
      riskScore += 15;
    }

    riskScore = Math.min(100, riskScore);

    let riskLevel: "High" | "Medium" | "Low" = "Low";
    if (riskScore >= 70) {
      riskLevel = "High";
    } else if (riskScore >= 40) {
      riskLevel = "Medium";
    }

    return {
      orderId,
      customerName,
      paymentMethod,
      state,
      riskScore,
      riskLevel
    };
  }).filter(item => item.orderId);
}

export function generateDisputeLetterText(r: LogisticsLeakageAnomaly): string {
  return `CLAIM REF: WT-DISPUTE-${r.orderId}
DATE: ${new Date().toLocaleDateString("en-IN")}
TO: Shiprocket Operations & Courier Billing Team
SUBJECT: Billing Dispute - Incorrect Weight Charged on Order ID #${r.orderId}

Dear Billing Team,

We are formalizing an immediate billing dispute regarding shipping charges for Order ID #${r.orderId}, shipped via ${r.courier || "our courier partner"}.

The audit of our warehouse logistics ledger highlights a significant weight discrepancy:
- Physical Weight (Actual weight of packed box): ${r.physicalWeight.toFixed(2)} kg
- Charged Weight (Billed by Courier): ${r.chargedWeight.toFixed(2)} kg
- Discrepancy Overcharge Ratio: ${r.weightRatio.toFixed(2)}x
- Total Freight Cost Charged: INR ${r.freightCost.toFixed(2)}

Under the standard courier service level agreement (SLA), billing must represent actual physical weight or volumetric weight, whichever is higher. Billed weight exceeds the physical weight by ${(r.chargedWeight - r.physicalWeight).toFixed(2)} kg, representing an audit failure.

We request an immediate credit note adjustment for the excess freight fee billed and a review of the courier dimensions scan for this shipment.

We have attached the package dimensions logs from our warehouse catalog for your quick resolution.

Sincerely,
Logistics Audit Manager, SheetCodeCrest Merchant Network`;
}

export function analyzeShopifyData(data: DataRow[]): ShopifyAnalyticsSummary {
  const lines = mapShopifyRows(data);
  const orders = uniqueOrders(lines);
  const customers = new Set(orders.map((o) => o.email || o.phone || o.customerName).filter(Boolean));
  const products: Record<string, number> = {};
  const cities: Record<string, number> = {};
  const statusCounts: Record<string, number> = {};
  const customerTotals: Record<string, { orders: number; spent: number }> = {};

  for (const line of lines) products[line.product] = (products[line.product] || 0) + line.qty;
  for (const order of orders) {
    cities[order.city] = (cities[order.city] || 0) + 1;
    statusCounts[order.orderStatus] = (statusCounts[order.orderStatus] || 0) + 1;
    const key = order.email || order.phone || order.customerName || order.orderNo;
    if (!customerTotals[key]) customerTotals[key] = { orders: 0, spent: 0 };
    customerTotals[key].orders += 1;
    customerTotals[key].spent += order.total;
  }

  const segmentCounts: Record<string, number> = {};
  Object.values(customerTotals).forEach((c) => {
    const segment = segmentCustomer(c.spent, c.orders);
    segmentCounts[segment] = (segmentCounts[segment] || 0) + 1;
  });

  const aprioriRules = mineAssociationRules(lines, orders.length);
  const rfmMatrix = computeRfmSegmentation(lines);
  const cohortRetention = calculateCohortRetention(lines);

  return {
    totalRows: data.length,
    totalOrders: orders.length,
    totalCustomers: customers.size,
    totalRevenue: orders.reduce((sum, order) => sum + order.total, 0),
    totalUnits: lines.reduce((sum, line) => sum + line.qty, 0),
    productCount: Object.keys(products).length,
    topProduct: sortedEntries(products)[0]?.[0] || "N/A",
    topCity: sortedEntries(cities)[0]?.[0] || "N/A",
    segmentCounts,
    statusCounts,
    aprioriRules,
    rfmMatrix,
    cohortRetention,
  };
}

export function buildShopifyAnalyticsWorkbook(fileName: string, data: DataRow[]): XLSX.WorkBook {
  const wb = XLSX.utils.book_new();
  const lines = mapShopifyRows(data);
  const orders = uniqueOrders(lines);
  const summary = analyzeShopifyData(data);
  const currency = (v: any) => ({ v: toNumber(v), t: "n", z: '"INR " #,##0.00' });
  const number = (v: any) => ({ v: Math.round(toNumber(v)), t: "n", z: "#,##0" });
  const pct = (v: number) => ({ v: Number.isFinite(v) ? v : 0, t: "n", z: "0.0%" });

  function addSheet(name: string, rows: any[][], widths?: number[]) {
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!freeze"] = { xSplit: 0, ySplit: rows[1] && rows[1].some(Boolean) ? 2 : 1 };
    if (widths) ws["!cols"] = widths.map((wch) => ({ wch }));
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  }

  const productMap: Record<string, any> = {};
  const statusProductMap: Record<string, any> = {};
  const monthMap: Record<string, any> = {};
  const cityMap: Record<string, any> = {};
  const stateMap: Record<string, any> = {};
  const discountMap: Record<string, any> = {};
  const codMap: Record<string, any> = {};
  const customerMap: Record<string, any> = {};

  for (const line of lines) {
    incrementMetric(productMap, line.product, {
      units: line.qty,
      revenue: line.lineRevenue,
      fulfilled: /fulfilled/i.test(line.fulfillmentStatus) && !/unfulfilled/i.test(line.fulfillmentStatus) ? line.qty : 0,
      pending: /unfulfilled|pending/i.test(line.fulfillmentStatus) ? line.qty : 0,
      restocked: /refund|void|cancel/i.test(line.financialStatus) ? line.qty : 0,
    });
    if (!statusProductMap[line.product]) statusProductMap[line.product] = {};
    statusProductMap[line.product][line.orderStatus] = (statusProductMap[line.product][line.orderStatus] || 0) + line.qty;
  }

  for (const order of orders) {
    const month = monthKey(order.date);
    incrementMetric(monthMap, month, {
      orders: 1,
      revenue: order.total,
      delivered: order.orderStatus === "Delivered" ? 1 : 0,
      cancelled: order.orderStatus === "Cancelled" ? 1 : 0,
      cod: /cod|cash/i.test(order.paymentMethod) ? 1 : 0,
      discounts: order.discount,
    });
    incrementMetric(cityMap, order.city, {
      orders: 1,
      revenue: order.total,
      delivered: order.orderStatus === "Delivered" ? 1 : 0,
    });
    incrementMetric(stateMap, order.state, {
      orders: 1,
      revenue: order.total,
      delivered: order.orderStatus === "Delivered" ? 1 : 0,
    });
    if (order.discountCode) {
      incrementMetric(discountMap, order.discountCode, {
        uses: 1,
        discount: order.discount,
        revenue: order.total,
      });
    }
    const codBucket = /cod|cash/i.test(order.paymentMethod) ? "COD" : "Prepaid/Online";
    incrementMetric(codMap, codBucket, {
      orders: 1,
      revenue: order.total,
      pending: order.orderStatus.includes("COD") ? order.total : 0,
      paid: order.orderStatus === "Delivered" ? order.total : 0,
      cancelled: /Cancelled|Refunded/i.test(order.orderStatus) ? order.total : 0,
    });
    const customerKey = order.email || order.phone || order.customerName || order.orderNo;
    if (!customerMap[customerKey]) {
      const parts = order.customerName.split(/\s+/);
      customerMap[customerKey] = {
        customerId: order.customerId ? `'${order.customerId}` : "",
        firstName: parts[0] || "",
        lastName: parts.slice(1).join(" "),
        email: order.email,
        phone: order.phone,
        city: order.city,
        state: order.state,
        country: order.country,
        orders: 0,
        spent: 0,
        emailMarketing: order.emailMarketing,
        smsMarketing: order.smsMarketing,
      };
    }
    customerMap[customerKey].orders += 1;
    customerMap[customerKey].spent += order.total;
  }

  const productRows = sortedEntries(productMap, "revenue").map(([product, value]: [string, any]) => {
    const orderCount = new Set(lines.filter((line) => line.product === product).map((line) => line.orderNo)).size;
    return [
      product,
      number(value.units),
      currency(value.revenue),
      currency(value.units ? value.revenue / value.units : 0),
      number(orderCount),
      number(value.fulfilled),
      number(value.pending),
      number(value.restocked),
      pct(value.units ? value.fulfilled / value.units : 0),
      pct(summary.totalRevenue ? value.revenue / summary.totalRevenue : 0),
    ];
  });

  const orderDetailRows = orders.map((order) => [
    `#${order.orderNo}`,
    order.date ? order.date.toISOString().slice(0, 10) : "",
    order.customerName,
    order.email,
    order.orderStatus,
    order.financialStatus,
    order.fulfillmentStatus,
    currency(order.total),
    order.paymentMethod,
    currency(order.discount),
    currency(order.refunded),
    order.city,
    order.state,
    order.risk,
  ]);

  const customerRows = sortedEntries(customerMap, "spent").map(([, c]: [string, any]) => {
    const segment = segmentCustomer(c.spent, c.orders);
    const reachable = [c.email ? "Email" : "", c.phone ? "SMS" : ""].filter(Boolean).join(", ");
    return [c.customerId, c.firstName, c.lastName, c.email, c.phone, c.city, c.state, c.country, number(c.orders), currency(c.spent), c.emailMarketing, c.smsMarketing, segment, reachable];
  });

  const segmentMap: Record<string, any> = {};
  for (const row of customerRows) {
    const segment = row[12] as string;
    const spent = toNumber((row[9] as any).v ?? row[9]);
    const ordersCount = toNumber((row[8] as any).v ?? row[8]);
    incrementMetric(segmentMap, segment, {
      count: 1,
      spent,
      orders: ordersCount,
      emailOptIn: row[10] === "yes" ? 1 : 0,
      smsOptIn: row[11] === "yes" ? 1 : 0,
      emailReachable: row[3] ? 1 : 0,
      smsReachable: row[4] ? 1 : 0,
      maxSpent: 0,
    });
    segmentMap[segment].maxSpent = Math.max(segmentMap[segment].maxSpent || 0, spent);
  }

  addSheet("Dashboard", [
    [`${fileName} - Shopify Analytics Command Center`],
    [`Rows: ${summary.totalRows.toLocaleString("en-IN")} | Orders: ${summary.totalOrders.toLocaleString("en-IN")} | Customers: ${summary.totalCustomers.toLocaleString("en-IN")} | Products: ${summary.productCount}`],
    [],
    ["KEY PERFORMANCE METRICS"],
    ["Metric", "Value", "", "Metric", "Value"],
    ["Total Revenue", currency(summary.totalRevenue), "", "Units Sold", number(summary.totalUnits)],
    ["Average Order Value", currency(summary.totalOrders ? summary.totalRevenue / summary.totalOrders : 0), "", "Top Product", summary.topProduct],
    ["Unique Customers", number(summary.totalCustomers), "", "Top City", summary.topCity],
    [],
    ["ORDER STATUS MIX"],
    ["Status", "Orders", "% Share"],
    ...sortedEntries(summary.statusCounts).map(([status, count]) => [status, number(count), pct(summary.totalOrders ? count / summary.totalOrders : 0)]),
    [],
    ["CUSTOMER SEGMENTS"],
    ["Segment", "Customers", "% Share"],
    ...sortedEntries(summary.segmentCounts).map(([segment, count]) => [segment, number(count), pct(summary.totalCustomers ? count / summary.totalCustomers : 0)]),
  ], [32, 18, 4, 28, 22]);

  addSheet("Order Status Detail", [
    ["Full Order Status Detail"],
    ["Order #", "Date", "Customer Name", "Email", "Order Status", "Financial Status", "Fulfillment Status", "Total (INR)", "Payment Method", "Discount (INR)", "Refunded (INR)", "City", "State", "Risk Level"],
    ...orderDetailRows,
  ], [14, 12, 24, 30, 24, 16, 18, 14, 24, 14, 14, 18, 10, 12]);

  addSheet("Product Analysis", [
    ["Product-Wise Performance Analysis"],
    ["Product", "Units Sold", "Gross Revenue (INR)", "Avg Price (INR)", "Orders", "Fulfilled Units", "Pending Units", "Restocked", "Fulfillment %", "Revenue Share %"],
    ...productRows,
  ], [44, 12, 18, 16, 12, 14, 14, 12, 14, 14]);

  const statuses = Object.keys(summary.statusCounts).slice(0, 10);
  addSheet("Product x Order Status", [
    ["Product x Order Status Cross-Analysis"],
    ["Product", ...statuses, "Total Units"],
    ...sortedEntries(productMap, "units").map(([product, value]: [string, any]) => [
      product,
      ...statuses.map((status) => number(statusProductMap[product]?.[status] || 0)),
      number(value.units),
    ]),
  ], [42, ...statuses.map(() => 16), 12]);

  addSheet("Monthly Trends", [
    ["Monthly Order & Revenue Trends"],
    ["Month", "Orders", "Revenue (INR)", "Avg Order Value (INR)", "Delivered", "Cancelled", "COD Orders", "Discounts Given (INR)"],
    ...Object.entries(monthMap).map(([month, v]: [string, any]) => [month, number(v.orders), currency(v.revenue), currency(v.orders ? v.revenue / v.orders : 0), number(v.delivered), number(v.cancelled), number(v.cod), currency(v.discounts)]),
  ], [14, 12, 18, 18, 12, 12, 12, 18]);

  addSheet("COD Analysis", [
    ["Cash on Delivery Deep Analysis"],
    ["Payment Bucket", "Orders", "Revenue (INR)", "Pending Collection (INR)", "Delivered/Paid Revenue (INR)", "Cancelled/Refunded Revenue (INR)", "% Orders"],
    ...sortedEntries(codMap, "orders").map(([bucket, v]: [string, any]) => [bucket, number(v.orders), currency(v.revenue), currency(v.pending), currency(v.paid), currency(v.cancelled), pct(summary.totalOrders ? v.orders / summary.totalOrders : 0)]),
  ], [22, 12, 18, 22, 24, 26, 12]);

  addSheet("Geographic", [
    ["Geographic Order Distribution"],
    [],
    ["Top Cities", "", "", "", "", "Top States"],
    ["City", "Orders", "Revenue (INR)", "Avg OV (INR)", "Delivered %", "", "State", "Orders", "Revenue (INR)", "Delivered %"],
    ...Array.from({ length: Math.max(25, Object.keys(stateMap).length) }).map((_, idx) => {
      const city = sortedEntries(cityMap, "orders")[idx];
      const state = sortedEntries(stateMap, "orders")[idx];
      return [
        city?.[0] || "",
        city ? number(city[1].orders) : "",
        city ? currency(city[1].revenue) : "",
        city ? currency(city[1].orders ? city[1].revenue / city[1].orders : 0) : "",
        city ? pct(city[1].orders ? city[1].delivered / city[1].orders : 0) : "",
        "",
        state?.[0] || "",
        state ? number(state[1].orders) : "",
        state ? currency(state[1].revenue) : "",
        state ? pct(state[1].orders ? state[1].delivered / state[1].orders : 0) : "",
      ];
    }),
  ], [18, 10, 16, 16, 12, 4, 14, 10, 16, 12]);

  addSheet("Discount Analysis", [
    ["Discount Code & Promotion Analysis"],
    ["Discount Code", "Uses", "Total Discount (INR)", "Avg Discount (INR)", "Total Revenue (INR)", "Avg OV (INR)"],
    ...sortedEntries(discountMap, "uses").map(([code, v]: [string, any]) => [code, number(v.uses), currency(v.discount), currency(v.uses ? v.discount / v.uses : 0), currency(v.revenue), currency(v.uses ? v.revenue / v.uses : 0)]),
  ], [24, 10, 20, 18, 20, 16]);

  addSheet("Customer Data", [
    ["Customer ID", "First Name", "Last Name", "Email", "Phone", "City", "State", "Country", "Total Orders", "Total Spent (INR)", "Email Marketing", "SMS Marketing", "Segment", "Reachable Via"],
    ...customerRows,
  ], [18, 16, 18, 32, 18, 18, 10, 10, 12, 18, 16, 14, 18, 18]);

  addSheet("Segment Analysis", [
    ["Customer Segment Analysis"],
    ["Segment", "Count", "% of Total", "Total Spent (INR)", "Avg Spend (INR)", "Max Spend (INR)", "Email Opt-in", "SMS Opt-in", "Email Reachable", "SMS Reachable"],
    ...sortedEntries(segmentMap, "count").map(([segment, v]: [string, any]) => [segment, number(v.count), pct(summary.totalCustomers ? v.count / summary.totalCustomers : 0), currency(v.spent), currency(v.count ? v.spent / v.count : 0), currency(v.maxSpent), number(v.emailOptIn), number(v.smsOptIn), number(v.emailReachable), number(v.smsReachable)]),
  ], [18, 10, 12, 18, 18, 18, 14, 14, 16, 16]);

  addSheet("Retargeting Lists", [
    ["Retargeting Contact Lists - Ready for Export"],
    [],
    ["EMAIL REACHABLE CUSTOMERS | Sorted by segment priority and total spend"],
    ["Segment", "First Name", "Last Name", "Email", "Phone", "City", "State", "Total Orders", "Total Spent (INR)"],
    ...customerRows.filter((row) => row[3]).map((row) => [row[12], row[1], row[2], row[3], row[4], row[5], row[6], row[8], row[9]]),
    [],
    ["SMS REACHABLE CUSTOMERS | Sorted by segment priority and total spend"],
    ["Segment", "First Name", "Last Name", "Phone", "City", "State", "Total Orders", "Total Spent (INR)"],
    ...customerRows.filter((row) => row[4]).map((row) => [row[12], row[1], row[2], row[4], row[5], row[6], row[8], row[9]]),
  ], [20, 16, 18, 32, 18, 18, 10, 12, 18]);

  const topProducts = sortedEntries(productMap, "revenue").slice(0, 20);
  addSheet("Product Sheet Index", [
    [`Product-Wise Customer & Order Report | ${topProducts.length} product sheets`],
    ["Product", "Orders", "Units", "Revenue (INR)", "Sheet Name"],
    ...topProducts.map(([product, v], idx) => {
      const safe = `${String(idx + 1).padStart(2, "0")}. ${product}`.slice(0, 31);
      const orderCount = new Set(lines.filter((line) => line.product === product).map((line) => line.orderNo)).size;
      return [product, number(orderCount), number((v as any).units), currency((v as any).revenue), safe];
    }),
  ], [44, 12, 12, 18, 32]);

  topProducts.forEach(([product, v], idx) => {
    const productLines = lines.filter((line) => line.product === product);
    const sheetName = `${String(idx + 1).padStart(2, "0")}. ${product}`.slice(0, 31);
    const orderCount = new Set(productLines.map((line) => line.orderNo)).size;
    addSheet(sheetName, [
      [`${product} | ${orderCount.toLocaleString("en-IN")} Orders | ${(v as any).units.toLocaleString("en-IN")} Units | INR ${Math.round((v as any).revenue).toLocaleString("en-IN")} Revenue`],
      ["Order #", "Order Date", "Order Status", "Customer ID", "First Name", "Last Name", "Email", "Phone", "Ship City", "State", "Address", "Product", "Qty", "Unit Price(INR)", "Revenue(INR)", "Discount Code", "Discount(INR)", "Payment", "Financial Status", "Fulfillment", "Email Mktg", "SMS Mktg", "Risk"],
      ...productLines.map((line) => {
        const parts = line.customerName.split(/\s+/);
        return [`#${line.orderNo}`, line.date ? line.date.toISOString().slice(0, 10) : "", line.orderStatus, line.customerId ? `'${line.customerId}` : "", parts[0] || "", parts.slice(1).join(" "), line.email, line.phone, line.city, line.state, line.address, line.product, number(line.qty), currency(line.unitPrice), currency(line.lineRevenue), line.discountCode, currency(line.discount), line.paymentMethod, line.financialStatus, line.fulfillmentStatus, line.emailMarketing, line.smsMarketing, line.risk];
      }),
    ], [14, 12, 24, 18, 16, 18, 30, 18, 18, 10, 34, 42, 10, 16, 16, 18, 14, 22, 16, 18, 12, 12, 10]);
  });

  // 1. Apriori Association Rules Sheet
  const rulesRows = (summary.aprioriRules || []).map((rule) => [
    rule.itemA,
    rule.itemB,
    pct(rule.support),
    pct(rule.confidence),
    number(rule.lift),
    number(rule.coPurchaseCount),
    rule.lift > 2.5 ? "Highly recommended bundle package opportunity!" : "Active cross-selling package opportunity."
  ]);
  addSheet("Growth Association Rules", [
    ["🧠 E-Commerce Growth Association Rules (Market Basket Analysis)"],
    ["Groups shopify order lines by order ID to isolate products regularly purchased together."],
    [],
    ["Item A", "Item B", "Support (Rule Frequency)", "Confidence (Rule Reliability)", "Lift (Association Strength)", "Co-Purchase Count", "Growth Action Recommendation"],
    ...rulesRows
  ], [36, 36, 24, 26, 26, 20, 42]);

  // 2. RFM Customer Segments Matrix Sheet
  const rfmRows = (summary.rfmMatrix || []).map((r) => {
    const action = r.cohort === "Champions"
      ? "Exclusive VIP discounts & Early access product reveals"
      : r.cohort === "Loyal Shoppers"
      ? "Introduce referral incentives & cross-sell programs"
      : r.cohort === "Recent Starters"
      ? "Welcome onboarding flow & standard discount vouchers"
      : r.cohort === "At Risk"
      ? "High-priority win-back retargeting campaigns"
      : "Standard low-frequency re-engagement newsletters";
    return [
      r.customerName,
      r.email,
      r.phone,
      number(r.recencyDays),
      number(r.frequency),
      currency(r.monetary),
      number(r.rScore),
      number(r.fScore),
      number(r.mScore),
      r.cohort,
      action
    ];
  });
  addSheet("RFM Segment Cohorts", [
    ["👥 Customer RFM Segment Cohort Matrix"],
    ["Grades all customer records chronologically across Recency (R), Frequency (F), and Monetary (M) scores (1 to 3)."],
    [],
    ["Customer Name", "Email", "Phone", "Recency (Days)", "Frequency (Orders)", "Monetary Spent (INR)", "R-Score (Recency)", "F-Score (Frequency)", "M-Score (Monetary)", "RFM Cohort Category", "Targeted Retention Action"],
    ...rfmRows
  ], [24, 30, 18, 16, 18, 18, 16, 16, 16, 22, 48]);

  // 3. Cohort Retention Heatmap Sheet
  const cohortData = calculateCohortRetention(lines);
  const cohortRows = cohortData.map((c) => [
    c.cohortMonth,
    number(c.totalCustomers),
    pct(c.rates[0] / 100),
    pct(c.rates[1] / 100),
    pct(c.rates[2] / 100),
    pct(c.rates[3] / 100),
    pct(c.rates[4] / 100),
    pct(c.rates[5] / 100)
  ]);
  addSheet("Cohort Retention Heatmap", [
    ["👥 Customer Cohort Retention Matrix (N-Month Retention Grid)"],
    ["Tracks customer cohorts monthly repeat transactions to measure long-term brand loyalty and churn."],
    [],
    ["Cohort Month", "Acquired Customers", "Month 0 (Acquisition)", "Month 1 (Repeat)", "Month 2 (Repeat)", "Month 3 (Repeat)", "Month 4 (Repeat)", "Month 5 (Repeat)"],
    ...cohortRows
  ], [16, 20, 24, 20, 20, 20, 20, 20]);

  return wb;
}

