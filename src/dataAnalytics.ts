import * as XLSX from "xlsx";
import {
  computeColumnStatsAndOutliers,
  computePearsonCorrelation,
  forecastSeries,
  groupTimeSeries,
  ForecastResult,
  OutlierDetail
} from "./utils/mathEngine";
import { parserRegistry } from "./parsers/ParserRegistry";

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
  outliers?: number[] | undefined;
  outliersList?: OutlierDetail[] | undefined;
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

// Helper functions copied from the original excelAnalytics.ts for header detection and table building
// These are needed by the parsers to produce standardized output

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

// Main parsing function that uses the parser registry
// This maintains backward compatibility with the existing parseExcel function
export async function parseExcel(file: File): Promise<{
  data: DataRow[];
  originalRows: number;
  headerRow: number;
  sheetName: string;
  headers: string[];
}> {
  // Use the parser registry to get the appropriate parser
  const parser = await parserRegistry.getParserForFile(file);
  if (!parser) {
    // Fallback to Excel parser if no specific parser is found
    // This maintains backward compatibility
    return await parseExcelFallback(file);
  }

  return parser.parse(file);
}

// Fallback to the original Excel parsing logic
async function parseExcelFallback(file: File): Promise<{
  data: DataRow[];
  originalRows: number;
  headerRow: number;
  sheetName: string;
  headers: string[];
}> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target!.result, { type: "array" });
        const result = parseWorkbook(wb);
        resolve(result);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// Buffer parsing function (for Node.js environments)
export function parseExcelBuffer(buffer: any): {
  data: DataRow[];
  originalRows: number;
  headerRow: number;
  sheetName: string;
  headers: string[];
} {
  const isNodeBuffer = typeof Buffer !== "undefined" && Buffer.isBuffer && Buffer.isBuffer(buffer);
  const wb = XLSX.read(buffer, { type: isNodeBuffer ? "buffer" : "array" });
  return parseWorkbook(wb);
}

// Internal function to parse a workbook (used by both main and fallback parsers)
function parseWorkbook(wb: XLSX.WorkBook): {
  data: DataRow[];
  originalRows: number;
  headerRow: number;
  sheetName: string;
  headers: string[];
} {
  // Use the same logic as before to choose the best sheet
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

  const candidate = best;
  if (candidate) {
    return {
      data: candidate.data,
      originalRows: candidate.data.length,
      headerRow: candidate.headerRow,
      sheetName: candidate.sheetName,
      headers: candidate.headers
    };
  }

  // Fallback to first sheet
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" }) as any[][];
  const headerRow = detectHeaderRowForSheet(raw, 12);
  const table = buildTable(raw, headerRow);
  return {
    data: table.data,
    originalRows: table.data.length,
    headerRow,
    sheetName,
    headers: table.headers
  };
}

function isDateValue(value: any) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return true;
  const str = String(value ?? "").trim();
  if (!str) return false;
  const parsed = Date.parse(str);
  return !Number.isNaN(parsed);
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

// The rest of the file remains unchanged - all the analysis functions are the same
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

// The rest of the functions (buildAnalyticsWorkbook, buildLogisticsWorkbook, etc.) remain unchanged
// ... (I'll include them in the next write operation to avoid making this too large)