export interface ForecastResult {
  forecast?: number[];
  lower?: number[];
  upper?: number[];
  model: string;
  labels?: string[];
  historicalValues?: number[];
  forecastValues?: number[];
  confidenceLower?: number[];
  confidenceUpper?: number[];
}

export interface OutlierDetail {
  index: number;
  val: number;
  type: "iqr" | "zscore";
  explanation: string;
}

export interface ColumnMathSummary {
  q1: number;
  median: number;
  q3: number;
  iqr: number;
  lowerBound: number;
  upperBound: number;
  outliers: number[];
  outliersList: OutlierDetail[];
  mean: number;
  stddev: number;
}

/**
 * Calculates the Pearson Correlation Coefficient between two arrays of numbers.
 */
export function computePearsonCorrelation(x: number[], y: number[]): number {
  const n = x.length;
  if (n <= 1) return 0;
  const meanX = x.reduce((a, b) => a + b, 0) / n;
  const meanY = y.reduce((a, b) => a + b, 0) / n;
  let num = 0;
  let denX = 0;
  let denY = 0;
  for (let i = 0; i < n; i++) {
    const dx = x[i] - meanX;
    const dy = y[i] - meanY;
    num += dx * dy;
    denX += dx * dx;
    denY += dy * dy;
  }
  if (denX === 0 || denY === 0) return 0;
  const r = num / Math.sqrt(denX * denY);
  return Number.isNaN(r) ? 0 : r;
}

/**
 * Computes statistical summaries, interquartile ranges, boxplot parameters, and flags outliers/anomalies.
 */
export function computeColumnStatsAndOutliers(
  numericValues: number[],
  rawValues: any[]
): ColumnMathSummary | null {
  if (numericValues.length === 0) return null;
  const sorted = [...numericValues].sort((a, b) => a - b);
  
  // Percentile helper
  const getPercentile = (p: number, sortedList: number[]): number => {
    if (sortedList.length === 0) return 0;
    const idx = (sortedList.length - 1) * p;
    const base = Math.floor(idx);
    const rest = idx - base;
    if (sortedList[base + 1] !== undefined) {
      return sortedList[base] + rest * (sortedList[base + 1] - sortedList[base]);
    }
    return sortedList[base];
  };

  const q1 = getPercentile(0.25, sorted);
  const median = getPercentile(0.50, sorted);
  const q3 = getPercentile(0.75, sorted);
  const iqr = q3 - q1;
  const lowerBound = q1 - 1.5 * iqr;
  const upperBound = q3 + 1.5 * iqr;

  const mean = numericValues.reduce((a, b) => a + b, 0) / numericValues.length;
  const variance = numericValues.reduce((acc, val) => acc + Math.pow(val - mean, 2), 0) / numericValues.length;
  const stddev = Math.sqrt(variance);

  const outliersList: OutlierDetail[] = [];

  rawValues.forEach((rawVal, idx) => {
    if (rawVal === null || rawVal === undefined || String(rawVal).trim() === "") return;
    const val = Number(String(rawVal).replace(/[,₹\s%]/g, ""));
    if (Number.isNaN(val)) return;

    const z = stddev > 0 ? (val - mean) / stddev : 0;
    const isZAnomaly = Math.abs(z) > 3.0;
    const isIQROutlier = val < lowerBound || val > upperBound;

    if (isIQROutlier || isZAnomaly) {
      let explanation = "";
      if (isIQROutlier && isZAnomaly) {
        explanation = `Outlier & Anomaly: Value is ${z.toFixed(1)}σ from the mean, lying outside the 1.5×IQR boundary [${lowerBound.toFixed(0)}, ${upperBound.toFixed(0)}].`;
      } else if (isIQROutlier) {
        explanation = `IQR Outlier: Value lies outside [${lowerBound.toFixed(0)}, ${upperBound.toFixed(0)}] (1.5×IQR boundary).`;
      } else {
        explanation = `Z-Score Anomaly: Value is ${z.toFixed(1)}σ from the mean (threshold is ±3.0σ).`;
      }
      outliersList.push({
        index: idx,
        val,
        type: isZAnomaly ? "zscore" : "iqr",
        explanation
      });
    }
  });

  const outliers = numericValues.filter(val => val < lowerBound || val > upperBound);

  return {
    q1,
    median,
    q3,
    iqr,
    lowerBound,
    upperBound,
    outliers,
    outliersList,
    mean,
    stddev
  };
}

/**
 * Simple Exponential Smoothing (SES) fallback model
 */
function simpleExponentialSmoothing(history: number[], steps: number): ForecastResult {
  const alpha = 0.3;
  const n = history.length;
  let level = history[0];

  for (let t = 1; t < n; t++) {
    level = alpha * history[t] + (1 - alpha) * level;
  }

  const forecast: number[] = Array(steps).fill(level);
  const avg = history.reduce((a, b) => a + b, 0) / n;
  const variance = history.reduce((acc, val) => acc + Math.pow(val - avg, 2), 0) / n;
  const stddev = Math.sqrt(variance);
  const rmse = stddev > 0 ? stddev : (avg * 0.1);

  const lower: number[] = [];
  const upper: number[] = [];
  for (let h = 1; h <= steps; h++) {
    const margin = 1.96 * rmse * Math.sqrt(h);
    lower.push(Math.max(0, level - margin));
    upper.push(level + margin);
  }

  return {
    forecast,
    lower,
    upper,
    model: "Simple Exponential Smoothing"
  };
}

/**
 * Double Exponential Smoothing (Holt's Linear Trend) fallback model
 */
function doubleExponentialSmoothing(history: number[], steps: number): ForecastResult {
  const alpha = 0.3;
  const beta = 0.1;
  const n = history.length;

  let level = history[0];
  let trend = history[1] - history[0];

  const levels: number[] = [level];
  const trends: number[] = [trend];

  for (let t = 1; t < n; t++) {
    const y = history[t];
    const prevLevel = levels[t - 1];
    const prevTrend = trends[t - 1];

    const curLevel = alpha * y + (1 - alpha) * (prevLevel + prevTrend);
    const curTrend = beta * (curLevel - prevLevel) + (1 - beta) * prevTrend;

    levels.push(curLevel);
    trends.push(curTrend);
  }

  const forecast: number[] = [];
  const lower: number[] = [];
  const upper: number[] = [];

  const lastLevel = levels[n - 1];
  const lastTrend = trends[n - 1];

  let sumSqError = 0;
  let fitCount = 0;
  for (let t = 1; t < n; t++) {
    const fitVal = levels[t - 1] + trends[t - 1];
    sumSqError += Math.pow(history[t] - fitVal, 2);
    fitCount++;
  }
  const rmse = fitCount > 0 ? Math.sqrt(sumSqError / fitCount) : (history[0] * 0.1);

  for (let h = 1; h <= steps; h++) {
    const fVal = lastLevel + h * lastTrend;
    forecast.push(fVal);

    const margin = 1.96 * rmse * Math.sqrt(h);
    lower.push(Math.max(0, fVal - margin));
    upper.push(fVal + margin);
  }

  return {
    forecast,
    lower,
    upper,
    model: "Holt's Linear Trend Model"
  };
}

/**
 * Triple Exponential Smoothing (Additive Holt-Winters Model)
 */
function tripleExponentialSmoothing(
  history: number[],
  steps: number,
  period: number
): ForecastResult {
  const alpha = 0.2;
  const beta = 0.1;
  const gamma = 0.3;
  const n = history.length;
  const L = period;

  // Initialize level
  let level = history.slice(0, L).reduce((a, b) => a + b, 0) / L;

  // Initialize trend
  let trend = 0;
  for (let i = 0; i < L; i++) {
    if (history[i + L] !== undefined) {
      trend += (history[i + L] - history[i]) / L;
    }
  }
  trend = trend / L;

  // Initialize seasonal components
  const seasonals: number[] = [];
  for (let i = 0; i < L; i++) {
    seasonals.push(history[i] - level);
  }

  const levels: number[] = Array(n).fill(0);
  const trends: number[] = Array(n).fill(0);
  const sExt = [...seasonals];

  levels[L - 1] = level;
  trends[L - 1] = trend;

  for (let t = L; t < n; t++) {
    const y = history[t];
    const prevLevel = levels[t - 1];
    const prevTrend = trends[t - 1];
    const prevSeasonal = sExt[t - L];

    const curLevel = alpha * (y - prevSeasonal) + (1 - alpha) * (prevLevel + prevTrend);
    const curTrend = beta * (curLevel - prevLevel) + (1 - beta) * prevTrend;
    const curSeasonal = gamma * (y - curLevel) + (1 - gamma) * prevSeasonal;

    levels[t] = curLevel;
    trends[t] = curTrend;
    sExt.push(curSeasonal);
  }

  const forecast: number[] = [];
  const lower: number[] = [];
  const upper: number[] = [];

  const lastLevel = levels[n - 1];
  const lastTrend = trends[n - 1];

  let sumSqError = 0;
  let fitCount = 0;
  for (let t = L; t < n; t++) {
    const fitVal = levels[t - 1] + trends[t - 1] + sExt[t - L];
    sumSqError += Math.pow(history[t] - fitVal, 2);
    fitCount++;
  }
  const rmse = fitCount > 0 ? Math.sqrt(sumSqError / fitCount) : (history[0] * 0.1);

  for (let h = 1; h <= steps; h++) {
    const seasonalIdx = (n - L + ((h - 1) % L)) % sExt.length;
    const fVal = lastLevel + h * lastTrend + sExt[seasonalIdx];
    forecast.push(fVal);

    const margin = 1.96 * rmse * Math.sqrt(h);
    lower.push(Math.max(0, fVal - margin));
    upper.push(fVal + margin);
  }

  return {
    forecast,
    lower,
    upper,
    model: `Holt-Winters Additive (Period: ${L})`
  };
}

/**
 * Hierarchical Forecasting Engine. Selects the most mathematically sound model based on data quantity.
 */
export function forecastSeries(
  history: number[],
  steps: number = 3,
  period?: number
): ForecastResult {
  const n = history.length;
  if (n === 0) {
    return { forecast: [], lower: [], upper: [], model: "Empty Data" };
  }
  if (n < 3) {
    const avg = history.reduce((a, b) => a + b, 0) / n;
    return {
      forecast: Array(steps).fill(avg),
      lower: Array(steps).fill(avg * 0.8),
      upper: Array(steps).fill(avg * 1.2),
      model: "Historical Average Baseline"
    };
  }

  // Auto-determine seasonal period if not specified
  const L = period || (n >= 24 ? 12 : 7);

  if (n >= 2 * L) {
    return tripleExponentialSmoothing(history, steps, L);
  } else if (n >= 5) {
    return doubleExponentialSmoothing(history, steps);
  } else {
    return simpleExponentialSmoothing(history, steps);
  }
}

/**
 * Group raw values chronologically by automatically resolving time gaps in date-grouped series.
 */
export function groupTimeSeries(
  rawDates: any[],
  rawValues: number[]
): { labels: string[]; values: number[] } {
  const pairs: { date: Date; val: number }[] = [];
  rawDates.forEach((d, i) => {
    if (d === null || d === undefined || String(d).trim() === "") return;
    const dt = new Date(d);
    const val = rawValues[i];
    if (!isNaN(dt.getTime()) && val !== null && val !== undefined && !isNaN(val)) {
      pairs.push({ date: dt, val });
    }
  });

  if (pairs.length < 3) {
    // If not enough chronological values, return step indices
    return {
      labels: rawValues.map((_, i) => `Step ${i + 1}`),
      values: rawValues
    };
  }

  // Sort chronological sequence
  pairs.sort((a, b) => a.date.getTime() - b.date.getTime());

  const minDate = pairs[0].date;
  const maxDate = pairs[pairs.length - 1].date;
  const diffDays = (maxDate.getTime() - minDate.getTime()) / (1000 * 3600 * 24);

  const groups: Record<string, number[]> = {};

  if (diffDays > 180) {
    // Group by Year-Month
    pairs.forEach(p => {
      const key = `${p.date.getFullYear()}-${String(p.date.getMonth() + 1).padStart(2, "0")}`;
      groups[key] = groups[key] || [];
      groups[key].push(p.val);
    });
  } else if (diffDays > 20) {
    // Group by Year-Week
    pairs.forEach(p => {
      const start = new Date(p.date.getFullYear(), 0, 1);
      const diff = p.date.getTime() - start.getTime();
      const weekNum = Math.floor(diff / (1000 * 3600 * 24 * 7)) + 1;
      const key = `${p.date.getFullYear()}-W${String(weekNum).padStart(2, "0")}`;
      groups[key] = groups[key] || [];
      groups[key].push(p.val);
    });
  } else {
    // Group by Date (YYYY-MM-DD)
    pairs.forEach(p => {
      const key = p.date.toISOString().split("T")[0];
      groups[key] = groups[key] || [];
      groups[key].push(p.val);
    });
  }

  const sortedKeys = Object.keys(groups).sort();
  return {
    labels: sortedKeys,
    values: sortedKeys.map(k => {
      const arr = groups[k];
      return arr.reduce((sum, v) => sum + v, 0) / arr.length;
    })
  };
}
