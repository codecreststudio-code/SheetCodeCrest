import * as XLSX from "xlsx";

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

function detectHeaderRowForSheet(raw: any[][], maxRows = 12) {
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

function normalizeHeaders(rawHeaders: any[]) {
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

function buildTable(raw: any[][], headerRow: number) {
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

function chooseBestSheet(wb: XLSX.WorkBook) {
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

function parseWorkbook(wb: XLSX.WorkBook) {
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

  return {
    totalRows,
    totalColumns: headers.length,
    duplicateRows,
    headers,
    columns: columnStats,
    topDuplicateRows,
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

  addSheet("📋 Merged Clean Data", [
    presentCols,
    ...formattedDataRows,
  ], presentCols.map((c) => c === "Product Name" ? 65 : c === "Order ID" ? 12 : 18));

  return wb;
}

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

  return wb;
}

