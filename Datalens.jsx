import React, { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import Papa from 'papaparse';
import * as XLSX from 'xlsx';
import {
  BarChart, Bar, LineChart, Line, PieChart, Pie, Cell, XAxis, YAxis,
  CartesianGrid, Tooltip, ResponsiveContainer, Legend,
} from 'recharts';
import {
  Upload, LayoutDashboard, Sparkles, ShieldCheck, MessageSquare, Download,
  Sun, Moon, Lock, AlertTriangle, CheckCircle2, Info, TrendingUp, TrendingDown,
  FileText, X, Loader2, Send, RefreshCw, Table2, Hash, Calendar as CalendarIcon,
  ToggleLeft, Type as TypeIcon, ChevronDown, Sparkle, ScanLine, Beaker,
} from 'lucide-react';
 
/* ============================================================
   THEME TOKENS
   ============================================================ */
const THEMES = {
  dark: {
    bg: '#0A101C', panel: '#111A2B', panelAlt: '#17223A', panelHover: '#1C2A45',
    border: '#243149', borderStrong: '#324262',
    text: '#EAF0FA', textSoft: '#9AACC6', textMuted: '#5D6E8A',
    accent: '#45D9C8', accent2: '#F2A93B', danger: '#E5646E', good: '#5FD08A',
    chartGrid: '#1E2A42',
  },
  light: {
    bg: '#F2F5FA', panel: '#FFFFFF', panelAlt: '#F6F8FC', panelHover: '#EDF1F8',
    border: '#DEE5F0', borderStrong: '#C6D0E2',
    text: '#0F1A2E', textSoft: '#4C5C77', textMuted: '#8494AC',
    accent: '#0C8F82', accent2: '#B9720E', danger: '#C23A4B', good: '#1D8A55',
    chartGrid: '#E4E9F2',
  },
};
const CHART_PALETTE = ['#45D9C8', '#F2A93B', '#7DA6FF', '#E5646E', '#B084F5', '#5FD08A', '#F589B8', '#6FC7E8'];
 
const FONT_STYLE = `
@import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
.dl-root{font-family:'Inter',ui-sans-serif,system-ui,sans-serif;}
.dl-display{font-family:'Space Grotesk',ui-sans-serif,system-ui,sans-serif;}
.dl-mono{font-family:'JetBrains Mono',ui-monospace,SFMono-Regular,Menlo,monospace;}
@keyframes dl-pulse{0%{box-shadow:0 0 0 0 rgba(69,217,200,0.55);}70%{box-shadow:0 0 0 10px rgba(69,217,200,0);}100%{box-shadow:0 0 0 0 rgba(69,217,200,0);}}
@keyframes dl-sweep{0%{transform:translateX(-100%);}100%{transform:translateX(220%);}}
@keyframes dl-fadeup{from{opacity:0;transform:translateY(6px);}to{opacity:1;transform:translateY(0);}}
@keyframes dl-spin{from{transform:rotate(0deg);}to{transform:rotate(360deg);}}
.dl-fadeup{animation:dl-fadeup .35s ease both;}
.dl-scanbar{position:absolute;top:0;left:0;height:100%;width:32%;background:linear-gradient(90deg, transparent, rgba(69,217,200,0.35), transparent);animation:dl-sweep 2.4s linear infinite;}
.dl-scrollbar::-webkit-scrollbar{width:8px;height:8px;}
.dl-scrollbar::-webkit-scrollbar-thumb{background:var(--dl-border-strong);border-radius:8px;}
@media print{
  .dl-noprint{display:none !important;}
  .dl-print-only{display:block !important;}
  body{background:#fff;}
}
.dl-print-only{display:none;}
`;
 
/* ============================================================
   PARSING HELPERS
   ============================================================ */
function toNumber(v) {
  if (v === null || v === undefined) return NaN;
  if (typeof v === 'number') return v;
  const s = String(v).trim().replace(/,/g, '');
  if (s === '' || !/^-?\d+(\.\d+)?$/.test(s)) return NaN;
  return parseFloat(s);
}
 
function isDateLike(v) {
  if (v === null || v === undefined) return false;
  const s = String(v).trim();
  if (s === '') return false;
  if (/^\d{4}-\d{1,2}-\d{1,2}/.test(s)) return true;
  if (/^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(s)) return true;
  if (/[a-zA-Z]{3,}/.test(s) && !isNaN(Date.parse(s))) return true;
  return false;
}
 
function parseDelimitedText(text) {
  const res = Papa.parse(text.trim(), { header: true, skipEmptyLines: true });
  if (res.meta.fields && res.meta.fields.length > 1) {
    return { columns: res.meta.fields, rows: res.data, warning: null };
  }
  const lines = text.split(/\r?\n/).filter((l) => l.trim() !== '');
  return {
    columns: ['line_number', 'text'],
    rows: lines.map((l, i) => ({ line_number: i + 1, text: l })),
    warning: null,
  };
}
 
async function parseFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();
 
  if (ext === 'csv') {
    const text = await file.text();
    const res = Papa.parse(text, { header: true, skipEmptyLines: true });
    return {
      columns: res.meta.fields || [],
      rows: res.data,
      warning: res.errors && res.errors.length ? `${res.errors.length} row(s) had minor parsing issues.` : null,
    };
  }
  if (ext === 'xlsx' || ext === 'xls') {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    const sheetName = wb.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: null });
    return {
      columns: rows.length ? Object.keys(rows[0]) : [],
      rows,
      warning: wb.SheetNames.length > 1 ? `Workbook has ${wb.SheetNames.length} sheets — analyzing "${sheetName}" only.` : null,
    };
  }
  if (ext === 'json') {
    const text = await file.text();
    let data = JSON.parse(text);
    if (!Array.isArray(data)) {
      const arrKey = Object.keys(data).find((k) => Array.isArray(data[k]));
      data = arrKey ? data[arrKey] : [data];
    }
    const colSet = new Set();
    data.slice(0, 200).forEach((row) => Object.keys(row || {}).forEach((k) => colSet.add(k)));
    return { columns: Array.from(colSet), rows: data, warning: null };
  }
  if (ext === 'txt') {
    const text = await file.text();
    return parseDelimitedText(text);
  }
  if (ext === 'pdf') {
    const buf = await file.arrayBuffer();
    const bytes = new Uint8Array(buf);
    let raw = '';
    for (let i = 0; i < bytes.length; i++) raw += String.fromCharCode(bytes[i]);
    const matches = [...raw.matchAll(/\(((?:\\.|[^()\\])*)\)\s*Tj/g)].map((m) => m[1]);
    const text = matches.join(' ').replace(/\\n/g, '\n').replace(/\\r/g, '').replace(/\\\(/g, '(').replace(/\\\)/g, ')');
    if (!text.trim()) {
      return {
        columns: ['text'], rows: [],
        warning: 'This PDF could not be read on-device — it likely uses compressed streams or is a scanned image. Try exporting it as TXT or CSV instead.',
      };
    }
    const parsed = parseDelimitedText(text);
    parsed.warning = 'PDF text extraction is best-effort and works reliably only for simple, uncompressed text-based PDFs.';
    return parsed;
  }
  throw new Error(`Unsupported file type: .${ext}`);
}
 
/* ============================================================
   COLUMN TYPE INFERENCE & STATS
   ============================================================ */
function percentile(sorted, p) {
  if (!sorted.length) return NaN;
  const idx = (p / 100) * (sorted.length - 1);
  const lo = Math.floor(idx), hi = Math.ceil(idx);
  if (lo === hi) return sorted[lo];
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (idx - lo);
}
 
function inferColumnType(values) {
  const nonNull = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== '');
  if (nonNull.length === 0) return 'empty';
  const n = nonNull.length;
  const numCount = nonNull.filter((v) => !isNaN(toNumber(v))).length;
  if (numCount / n > 0.9) return 'numeric';
  const boolSet = new Set(['true', 'false', 'yes', 'no', 'y', 'n']);
  const boolCount = nonNull.filter((v) => boolSet.has(String(v).trim().toLowerCase())).length;
  if (boolCount / n > 0.9) return 'boolean';
  const dateCount = nonNull.filter(isDateLike).length;
  if (dateCount / n > 0.8) return 'date';
  const uniqueCount = new Set(nonNull.map((v) => String(v).trim().toLowerCase())).size;
  if (uniqueCount <= Math.max(20, n * 0.5) && uniqueCount / n < 0.6) return 'categorical';
  return 'text';
}
 
function numericStats(values) {
  const nums = values.map(toNumber).filter((v) => !isNaN(v));
  if (!nums.length) return null;
  const sorted = [...nums].sort((a, b) => a - b);
  const n = sorted.length;
  const mean = nums.reduce((a, b) => a + b, 0) / n;
  const variance = nums.reduce((a, b) => a + (b - mean) ** 2, 0) / n;
  const std = Math.sqrt(variance);
  const q1 = percentile(sorted, 25), median = percentile(sorted, 50), q3 = percentile(sorted, 75);
  const iqr = q3 - q1;
  const lower = q1 - 1.5 * iqr, upper = q3 + 1.5 * iqr;
  const outlierCount = nums.filter((v) => v < lower || v > upper).length;
  return { n, mean, std, min: sorted[0], max: sorted[n - 1], q1, median, q3, iqr, lowerFence: lower, upperFence: upper, outlierCount, values: nums };
}
 
function categoricalStats(values) {
  const nonNull = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== '');
  const freq = {};
  nonNull.forEach((v) => { const k = String(v).trim(); freq[k] = (freq[k] || 0) + 1; });
  const entries = Object.entries(freq).sort((a, b) => b[1] - a[1]);
  return { top: entries.slice(0, 8).map(([value, count]) => ({ value, count })) };
}
 
function computeColumnProfile(name, values, totalRows) {
  const missing = values.filter((v) => v === null || v === undefined || String(v).trim() === '').length;
  const type = inferColumnType(values);
  const nonNullVals = values.filter((v) => v !== null && v !== undefined && String(v).trim() !== '');
  const profile = {
    name, type, count: totalRows, missing,
    missingPct: totalRows ? (missing / totalRows) * 100 : 0,
    unique: new Set(nonNullVals.map((v) => String(v).trim())).size,
  };
  if (type === 'numeric') profile.stats = numericStats(values);
  else profile.stats = categoricalStats(values);
  return profile;
}
 
function pearson(xs, ys) {
  const pairs = [];
  for (let i = 0; i < xs.length; i++) {
    const x = toNumber(xs[i]), y = toNumber(ys[i]);
    if (!isNaN(x) && !isNaN(y)) pairs.push([x, y]);
  }
  const n = pairs.length;
  if (n < 3) return null;
  const mx = pairs.reduce((a, p) => a + p[0], 0) / n;
  const my = pairs.reduce((a, p) => a + p[1], 0) / n;
  let num = 0, dx = 0, dy = 0;
  pairs.forEach(([x, y]) => { num += (x - mx) * (y - my); dx += (x - mx) ** 2; dy += (y - my) ** 2; });
  if (dx === 0 || dy === 0) return null;
  return { r: num / Math.sqrt(dx * dy), n };
}
 
function findDuplicates(rows) {
  const seen = new Map();
  let count = 0;
  rows.forEach((row) => {
    const key = JSON.stringify(row);
    if (seen.has(key)) count++;
    else seen.set(key, true);
  });
  return { count };
}
 
function computeQuality(profiles, rowCount, dupCount) {
  const cols = Object.values(profiles);
  const totalCells = cols.length * rowCount;
  const totalMissing = cols.reduce((a, c) => a + c.missing, 0);
  const completeness = totalCells ? 1 - totalMissing / totalCells : 1;
  const uniqueness = rowCount ? 1 - dupCount / rowCount : 1;
  const numericCols = cols.filter((c) => c.type === 'numeric' && c.stats);
  const totalNumericVals = numericCols.reduce((a, c) => a + (c.stats.n || 0), 0);
  const totalOutliers = numericCols.reduce((a, c) => a + (c.stats.outlierCount || 0), 0);
  const validity = totalNumericVals ? 1 - totalOutliers / totalNumericVals : 1;
  const score = Math.round(100 * (0.4 * completeness + 0.3 * uniqueness + 0.3 * validity));
  return { score: Math.max(0, Math.min(100, score)), completeness: completeness * 100, uniqueness: uniqueness * 100, validity: validity * 100 };
}
 
function generateInsights(profiles, meta, dup, correlations, trends, quality) {
  const out = [];
  out.push({ type: 'info', text: `Dataset contains ${meta.rowCount.toLocaleString()} rows and ${meta.colCount} columns (${meta.numericCols} numeric, ${meta.categoricalCols} categorical, ${meta.dateCols} date, ${meta.booleanCols} boolean, ${meta.textCols} text).` });
  Object.values(profiles).forEach((p) => {
    if (p.missingPct > 30) out.push({ type: 'warning', text: `"${p.name}" is missing ${p.missingPct.toFixed(1)}% of its values — this may significantly affect analysis.` });
    else if (p.missingPct > 5) out.push({ type: 'info', text: `"${p.name}" has ${p.missingPct.toFixed(1)}% missing values.` });
    if (p.type === 'numeric' && p.stats && p.stats.outlierCount > 0) out.push({ type: 'warning', text: `"${p.name}" has ${p.stats.outlierCount} statistical outlier${p.stats.outlierCount > 1 ? 's' : ''} (beyond 1.5× IQR).` });
    if (p.type === 'categorical' && p.stats && p.stats.top[0]) out.push({ type: 'info', text: `In "${p.name}", "${p.stats.top[0].value}" is the most common value (${p.stats.top[0].count} occurrences).` });
  });
  out.push(dup.count > 0
    ? { type: 'warning', text: `Found ${dup.count} duplicate row${dup.count > 1 ? 's' : ''} in the dataset.` }
    : { type: 'success', text: 'No duplicate rows detected.' });
  correlations.slice(0, 3).forEach((c) => out.push({ type: 'info', text: `"${c.a}" and "${c.b}" are ${Math.abs(c.r) > 0.7 ? 'strongly' : 'moderately'} ${c.r > 0 ? 'positively' : 'negatively'} correlated (r = ${c.r.toFixed(2)}).` }));
  trends.slice(0, 3).forEach((t) => out.push({ type: 'info', text: `"${t.col}" shows a ${t.direction} trend across the row order (r = ${t.r.toFixed(2)}).` }));
  if (quality.score >= 85) out.push({ type: 'success', text: `Overall data quality is excellent (${quality.score}/100).` });
  else if (quality.score >= 60) out.push({ type: 'info', text: `Overall data quality is good (${quality.score}/100), with some room for improvement.` });
  else out.push({ type: 'warning', text: `Overall data quality needs attention (${quality.score}/100).` });
  return out;
}
 
function generateRecommendations(profiles, meta, dup, correlations, quality) {
  const out = [];
  Object.values(profiles).forEach((p) => {
    if (p.missingPct > 20) out.push(`Consider imputing or removing "${p.name}" — ${p.missingPct.toFixed(0)}% of its values are missing.`);
    if (p.type === 'numeric' && p.stats && p.stats.outlierCount > 0 && p.stats.n) {
      const pct = (p.stats.outlierCount / p.stats.n) * 100;
      if (pct > 3) out.push(`Investigate ${p.stats.outlierCount} outliers in "${p.name}" — they may be data entry errors or genuine rare events.`);
    }
    if (p.type === 'categorical' && p.unique > 30) out.push(`"${p.name}" has high cardinality (${p.unique} unique values) — consider grouping rare categories.`);
  });
  if (dup.count > 0) out.push(`Remove ${dup.count} duplicate row(s) before downstream analysis to avoid skewed statistics.`);
  correlations.filter((c) => Math.abs(c.r) > 0.85).forEach((c) => out.push(`"${c.a}" and "${c.b}" are highly correlated (r = ${c.r.toFixed(2)}) — one may be redundant for modeling.`));
  if (quality.completeness < 80) out.push('Overall completeness is low — review your data collection process for systematic gaps.');
  if (!out.length) out.push('No major data quality issues were found. This dataset looks ready for analysis.');
  return out;
}
 
function runAnalysis(parsed) {
  const { columns, rows } = parsed;
  const totalRows = rows.length;
  const profiles = {};
  columns.forEach((col) => {
    profiles[col] = computeColumnProfile(col, rows.map((r) => (r ? r[col] : null)), totalRows);
  });
  const dup = findDuplicates(rows);
  const numericCols = columns.filter((c) => profiles[c].type === 'numeric');
  const categoricalCols = columns.filter((c) => profiles[c].type === 'categorical');
 
  const correlations = [];
  for (let i = 0; i < numericCols.length; i++) {
    for (let j = i + 1; j < numericCols.length; j++) {
      const a = numericCols[i], b = numericCols[j];
      const res = pearson(rows.map((r) => r[a]), rows.map((r) => r[b]));
      if (res && Math.abs(res.r) >= 0.3) correlations.push({ a, b, r: res.r, n: res.n });
    }
  }
  correlations.sort((x, y) => Math.abs(y.r) - Math.abs(x.r));
 
  const trends = [];
  numericCols.forEach((c) => {
    const idx = rows.map((_, i) => i);
    const res = pearson(idx, rows.map((r) => r[c]));
    if (res && Math.abs(res.r) >= 0.3) trends.push({ col: c, direction: res.r > 0 ? 'increasing' : 'decreasing', r: res.r });
  });
 
  const quality = computeQuality(profiles, totalRows, dup.count);
  const meta = {
    rowCount: totalRows, colCount: columns.length,
    numericCols: numericCols.length, categoricalCols: categoricalCols.length,
    dateCols: columns.filter((c) => profiles[c].type === 'date').length,
    booleanCols: columns.filter((c) => profiles[c].type === 'boolean').length,
    textCols: columns.filter((c) => profiles[c].type === 'text').length,
    emptyCols: columns.filter((c) => profiles[c].type === 'empty').length,
  };
 
  const corrMatrixCols = numericCols.slice(0, 8);
  const fullCorr = {};
  corrMatrixCols.forEach((a) => {
    fullCorr[a] = {};
    corrMatrixCols.forEach((b) => {
      if (a === b) { fullCorr[a][b] = 1; return; }
      const res = pearson(rows.map((r) => r[a]), rows.map((r) => r[b]));
      fullCorr[a][b] = res ? res.r : null;
    });
  });
 
  const insights = generateInsights(profiles, meta, dup, correlations, trends, quality);
  const recommendations = generateRecommendations(profiles, meta, dup, correlations, quality);
 
  return {
    profiles, dup, correlations, trends, quality, meta, insights, recommendations,
    columns, numericCols, categoricalCols, corrMatrixCols, fullCorr,
  };
}
 
/* ============================================================
   CHAT ENGINE (fully local — no data ever leaves the browser)
   ============================================================ */
function answerQuestion(q, analysis) {
  if (!analysis) return "Upload a dataset first and I'll be able to answer questions about it.";
  const ql = q.toLowerCase();
  const cols = analysis.columns;
  const mentioned = cols.find((c) => ql.includes(c.toLowerCase()));
 
  if (/missing|null|empty/.test(ql)) {
    if (mentioned) {
      const p = analysis.profiles[mentioned];
      return `"${mentioned}" has ${p.missing} missing value${p.missing === 1 ? '' : 's'} out of ${p.count} (${p.missingPct.toFixed(1)}%).`;
    }
    const withMissing = Object.values(analysis.profiles).filter((p) => p.missing > 0).sort((a, b) => b.missingPct - a.missingPct);
    if (!withMissing.length) return 'No missing values were found anywhere in the dataset.';
    return `Columns with missing values: ${withMissing.slice(0, 6).map((p) => `${p.name} (${p.missingPct.toFixed(1)}%)`).join(', ')}.`;
  }
  if (/duplicate/.test(ql)) {
    return analysis.dup.count > 0 ? `I found ${analysis.dup.count} duplicate row(s) in your data.` : 'No duplicate rows were found.';
  }
  if (/correlat|relationship|related/.test(ql)) {
    if (!analysis.correlations.length) return 'No strong correlations (|r| ≥ 0.3) were found between numeric columns.';
    return 'Notable correlations: ' + analysis.correlations.slice(0, 5).map((c) => `${c.a} ↔ ${c.b} (r=${c.r.toFixed(2)})`).join(', ') + '.';
  }
  if (/outlier|anomal/.test(ql)) {
    const withOutliers = Object.values(analysis.profiles).filter((p) => p.type === 'numeric' && p.stats?.outlierCount > 0);
    if (!withOutliers.length) return 'No significant outliers were detected in the numeric columns.';
    return 'Outliers detected in: ' + withOutliers.map((p) => `${p.name} (${p.stats.outlierCount})`).join(', ') + '.';
  }
  if (/trend/.test(ql)) {
    if (!analysis.trends.length) return 'No strong trends were detected across the row order of the dataset.';
    return analysis.trends.map((t) => `"${t.col}" is trending ${t.direction} (r=${t.r.toFixed(2)}).`).join(' ');
  }
  if (/quality|score/.test(ql)) {
    return `The overall data quality score is ${analysis.quality.score}/100 — completeness ${analysis.quality.completeness.toFixed(0)}%, uniqueness ${analysis.quality.uniqueness.toFixed(0)}%, validity ${analysis.quality.validity.toFixed(0)}%.`;
  }
  if (/(mean|average|median|deviation|std|min|max|sum|range)/.test(ql) && mentioned) {
    const p = analysis.profiles[mentioned];
    if (p.type !== 'numeric' || !p.stats) return `"${mentioned}" isn't numeric, so I can't compute stats like mean or median for it.`;
    const s = p.stats;
    return `"${mentioned}": mean ${s.mean.toFixed(2)}, median ${s.median.toFixed(2)}, std dev ${s.std.toFixed(2)}, min ${s.min}, max ${s.max}.`;
  }
  if (mentioned) {
    const p = analysis.profiles[mentioned];
    if (p.type === 'numeric' && p.stats) return `"${mentioned}" is numeric: mean ${p.stats.mean.toFixed(2)}, range ${p.stats.min}–${p.stats.max}, ${p.missing} missing values.`;
    if (p.stats && p.stats.top && p.stats.top[0]) return `"${mentioned}" has ${p.unique} unique values. Most common: "${p.stats.top[0].value}" (${p.stats.top[0].count} times). ${p.missing} missing values.`;
    return `"${mentioned}" is a ${p.type} column with ${p.missing} missing values.`;
  }
  if (/summary|overview|describe|about/.test(ql)) {
    return `This dataset has ${analysis.meta.rowCount.toLocaleString()} rows and ${analysis.meta.colCount} columns. Quality score: ${analysis.quality.score}/100. ${analysis.dup.count} duplicate row(s) found.`;
  }
  return `I can answer questions about missing values, duplicates, correlations, outliers, trends, data quality, or stats for a specific column (e.g. "what's the average of ${cols[0] || 'a column'}?"). You can also check the Insights tab for a full breakdown.`;
}
 
/* ============================================================
   SAMPLE DATA (for quick demo — still 100% local)
   ============================================================ */
function generateSampleData() {
  const regions = ['North', 'South', 'East', 'West'];
  const products = ['Aurora Desk', 'Nimbus Chair', 'Vertex Lamp', 'Halo Monitor', 'Pulse Keyboard'];
  const rows = [];
  const start = new Date('2024-01-01').getTime();
  for (let i = 0; i < 180; i++) {
    const date = new Date(start + i * 86400000).toISOString().slice(0, 10);
    const region = regions[i % regions.length];
    const product = products[(i * 3) % products.length];
    const base = 40 + (i % 30) + product.length * 2;
    const units = Math.max(1, Math.round(base / 5 + (Math.random() * 10 - 5)));
    const price = 30 + (product.charCodeAt(0) % 20);
    let revenue = units * price + (Math.random() * 20 - 10);
    if (i % 47 === 0) revenue *= 6;
    const rating = Math.min(5, Math.max(1, +(3.5 + Math.sin(i / 10) + (Math.random() * 0.6 - 0.3)).toFixed(1)));
    rows.push({
      date, region, product,
      units_sold: units, unit_price: price,
      revenue: Math.round(revenue * 100) / 100,
      customer_rating: rating,
      returned: Math.random() < 0.06 ? 'Yes' : 'No',
    });
  }
  rows[12].customer_rating = null;
  rows[45].unit_price = null;
  rows[90] = { ...rows[89] };
  return { columns: Object.keys(rows[0]), rows };
}
 
/* ============================================================
   EXPORT HELPERS
   ============================================================ */
function downloadBlob(content, filename, type) {
  const blob = new Blob([content], { type });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(a.href);
}
 
function downloadSvgAsPng(containerEl, filename) {
  if (!containerEl) return;
  const svg = containerEl.querySelector('svg');
  if (!svg) return;
  let source = new XMLSerializer().serializeToString(svg);
  if (!/xmlns="http:\/\/www\.w3\.org\/2000\/svg"/.test(source)) {
    source = source.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"');
  }
  const rect = svg.getBoundingClientRect();
  const scale = 2;
  const canvas = document.createElement('canvas');
  canvas.width = Math.max(1, rect.width * scale);
  canvas.height = Math.max(1, rect.height * scale);
  const ctx = canvas.getContext('2d');
  const img = new Image();
  const svgBlob = new Blob([source], { type: 'image/svg+xml;charset=utf-8' });
  const url = URL.createObjectURL(svgBlob);
  img.onload = () => {
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    URL.revokeObjectURL(url);
    canvas.toBlob((blob) => {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      a.click();
      URL.revokeObjectURL(a.href);
    });
  };
  img.src = url;
}
 
function histogramBins(values, binCount = 10) {
  const nums = values.map(toNumber).filter((v) => !isNaN(v));
  if (!nums.length) return [];
  const min = Math.min(...nums), max = Math.max(...nums);
  if (min === max) return [{ bin: `${min}`, count: nums.length }];
  const width = (max - min) / binCount;
  const bins = Array.from({ length: binCount }, (_, i) => ({ start: min + i * width, end: min + (i + 1) * width, count: 0 }));
  nums.forEach((v) => {
    let idx = Math.floor((v - min) / width);
    if (idx >= binCount) idx = binCount - 1;
    if (idx < 0) idx = 0;
    bins[idx].count++;
  });
  return bins.map((b) => ({ bin: `${b.start.toFixed(1)}–${b.end.toFixed(1)}`, count: b.count }));
}
 
function corrColor(r, theme) {
  if (r === null || r === undefined || isNaN(r)) return theme.panelAlt;
  const a = Math.min(1, Math.abs(r));
  const base = r >= 0 ? [69, 217, 200] : [229, 100, 110];
  const mix = base.map((c) => Math.round(theme === THEMES.dark ? c * a + 17 * (1 - a) : c * a + 255 * (1 - a)));
  return `rgb(${mix[0]},${mix[1]},${mix[2]})`;
}
 
/* ============================================================
   SMALL UI ATOMS
   ============================================================ */
function IconBadge({ Icon, color, bg }) {
  return (
    <div className="flex items-center justify-center rounded-lg" style={{ width: 34, height: 34, background: bg, color }}>
      <Icon size={17} />
    </div>
  );
}
 
function Panel({ theme, children, className = '', style = {} }) {
  return (
    <div className={`rounded-2xl border ${className}`} style={{ background: theme.panel, borderColor: theme.border, ...style }}>
      {children}
    </div>
  );
}
 
function SectionTitle({ theme, icon: Icon, title, subtitle, right }) {
  return (
    <div className="flex items-start justify-between gap-3 mb-4">
      <div className="flex items-center gap-2.5">
        {Icon && <Icon size={18} style={{ color: theme.accent }} />}
        <div>
          <h3 className="dl-display text-[15px] font-semibold" style={{ color: theme.text }}>{title}</h3>
          {subtitle && <p className="text-xs mt-0.5" style={{ color: theme.textMuted }}>{subtitle}</p>}
        </div>
      </div>
      {right}
    </div>
  );
}
 
function TypeChip({ type, theme }) {
  const map = {
    numeric: { icon: Hash, color: theme.accent },
    categorical: { icon: TypeIcon, color: theme.accent2 },
    date: { icon: CalendarIcon, color: '#7DA6FF' },
    boolean: { icon: ToggleLeft, color: theme.good },
    text: { icon: FileText, color: theme.textSoft },
    empty: { icon: X, color: theme.danger },
  };
  const m = map[type] || map.text;
  const I = m.icon;
  return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[10px] font-semibold uppercase tracking-wide dl-mono"
      style={{ color: m.color, background: `${m.color}18` }}>
      <I size={11} /> {type}
    </span>
  );
}
 
function Gauge({ value, theme, size = 128 }) {
  const r = (size - 14) / 2;
  const c = 2 * Math.PI * r;
  const pct = Math.max(0, Math.min(100, value));
  const color = pct >= 85 ? theme.good : pct >= 60 ? theme.accent2 : theme.danger;
  return (
    <div className="relative" style={{ width: size, height: size }}>
      <svg width={size} height={size}>
        <circle cx={size / 2} cy={size / 2} r={r} stroke={theme.border} strokeWidth="10" fill="none" />
        <circle cx={size / 2} cy={size / 2} r={r} stroke={color} strokeWidth="10" fill="none"
          strokeDasharray={c} strokeDashoffset={c - (pct / 100) * c} strokeLinecap="round"
          transform={`rotate(-90 ${size / 2} ${size / 2})`} style={{ transition: 'stroke-dashoffset 0.8s ease' }} />
      </svg>
      <div className="absolute inset-0 flex flex-col items-center justify-center">
        <span className="dl-mono text-2xl font-semibold" style={{ color: theme.text }}>{Math.round(pct)}</span>
        <span className="text-[10px]" style={{ color: theme.textMuted }}>/ 100</span>
      </div>
    </div>
  );
}
 
function MetricBar({ label, value, theme, color }) {
  return (
    <div>
      <div className="flex justify-between mb-1 text-xs" style={{ color: theme.textSoft }}>
        <span>{label}</span>
        <span className="dl-mono" style={{ color: theme.text }}>{value.toFixed(0)}%</span>
      </div>
      <div className="h-1.5 rounded-full overflow-hidden" style={{ background: theme.border }}>
        <div className="h-full rounded-full" style={{ width: `${Math.max(2, value)}%`, background: color, transition: 'width 0.8s ease' }} />
      </div>
    </div>
  );
}
 
/* ============================================================
   MAIN APP
   ============================================================ */
const TABS = [
  { id: 'upload', label: 'Upload', icon: Upload },
  { id: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { id: 'insights', label: 'Insights', icon: Sparkles },
  { id: 'quality', label: 'Data Quality', icon: ShieldCheck },
  { id: 'chat', label: 'Ask DataLens', icon: MessageSquare },
  { id: 'export', label: 'Export', icon: Download },
];
 
export default function DataLensApp() {
  const [mode, setMode] = useState('dark');
  const theme = THEMES[mode];
  const [activeTab, setActiveTab] = useState('upload');
  const [fileName, setFileName] = useState(null);
  const [parsed, setParsed] = useState(null);
  const [analysis, setAnalysis] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState('');
  const [dragOver, setDragOver] = useState(false);
  const [error, setError] = useState(null);
  const [parseWarning, setParseWarning] = useState(null);
  const [selectedNumeric, setSelectedNumeric] = useState(null);
  const [selectedCat, setSelectedCat] = useState(null);
  const [chatMessages, setChatMessages] = useState([
    { role: 'assistant', text: "Hi, I'm your on-device data analyst. Upload a file and ask me anything about it — nothing ever leaves your browser." },
  ]);
  const [chatInput, setChatInput] = useState('');
  const fileInputRef = useRef(null);
  const histRef = useRef(null);
  const catRef = useRef(null);
  const missingRef = useRef(null);
  const typesRef = useRef(null);
  const chatEndRef = useRef(null);
 
  useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior: 'smooth' }); }, [chatMessages]);
 
  const finalizeDataset = useCallback((parsedObj, name) => {
    let rows = parsedObj.rows;
    let truncNote = null;
    if (rows.length > 20000) {
      rows = rows.slice(0, 20000);
      truncNote = `Dataset truncated to the first 20,000 rows for in-browser performance.`;
    }
    const cleanParsed = { ...parsedObj, rows };
    const a = runAnalysis(cleanParsed);
    setParsed(cleanParsed);
    setAnalysis(a);
    setFileName(name);
    setSelectedNumeric(a.numericCols[0] || null);
    setSelectedCat(a.categoricalCols[0] || null);
    setParseWarning([parsedObj.warning, truncNote].filter(Boolean).join(' '));
    setChatMessages([
      { role: 'assistant', text: `Loaded "${name}" — ${a.meta.rowCount.toLocaleString()} rows × ${a.meta.colCount} columns. Data quality score: ${a.quality.score}/100. Ask me anything about it, e.g. "any duplicates?" or "summarize this data".` },
    ]);
    setActiveTab('dashboard');
    setLoading(false);
  }, []);
 
  const handleFile = useCallback(async (file) => {
    setError(null);
    setLoading(true);
    setLoadingMsg('Reading file on-device…');
    try {
      const p = await parseFile(file);
      setLoadingMsg('Profiling columns & detecting patterns…');
      setTimeout(() => finalizeDataset(p, file.name), 450);
    } catch (e) {
      setError(e.message || 'Could not parse this file.');
      setLoading(false);
    }
  }, [finalizeDataset]);
 
  const handleSample = useCallback(() => {
    setError(null);
    setLoading(true);
    setLoadingMsg('Generating sample dataset…');
    const p = generateSampleData();
    setTimeout(() => finalizeDataset(p, 'sample-sales-data.csv'), 450);
  }, [finalizeDataset]);
 
  const onDrop = useCallback((e) => {
    e.preventDefault(); setDragOver(false);
    const f = e.dataTransfer.files?.[0];
    if (f) handleFile(f);
  }, [handleFile]);
 
  const sendChat = useCallback(() => {
    const q = chatInput.trim();
    if (!q) return;
    setChatMessages((m) => [...m, { role: 'user', text: q }]);
    setChatInput('');
    setTimeout(() => {
      const ans = answerQuestion(q, analysis);
      setChatMessages((m) => [...m, { role: 'assistant', text: ans }]);
    }, 300);
  }, [chatInput, analysis]);
 
  const suggestedQuestions = useMemo(() => {
    if (!analysis) return [];
    const qs = ['Summarize this data', 'Any duplicates?', 'Which columns have missing values?'];
    if (analysis.correlations.length) qs.push('Show me the strongest correlations');
    if (analysis.numericCols[0]) qs.push(`What's the average ${analysis.numericCols[0]}?`);
    return qs.slice(0, 4);
  }, [analysis]);
 
  /* ---------- export actions ---------- */
  const exportJSON = () => {
    const report = {
      generatedAt: new Date().toISOString(), fileName,
      rows: analysis.meta.rowCount, columns: analysis.meta.colCount,
      qualityScore: analysis.quality.score, quality: analysis.quality,
      duplicates: analysis.dup.count,
      insights: analysis.insights.map((i) => i.text),
      recommendations: analysis.recommendations,
      correlations: analysis.correlations,
      trends: analysis.trends,
      columnProfiles: Object.fromEntries(Object.entries(analysis.profiles).map(([k, v]) => [k, {
        type: v.type, missing: v.missing, missingPct: v.missingPct, unique: v.unique,
        stats: v.type === 'numeric' && v.stats ? { mean: v.stats.mean, median: v.stats.median, std: v.stats.std, min: v.stats.min, max: v.stats.max, outliers: v.stats.outlierCount } : (v.stats?.top || null),
      }])),
    };
    downloadBlob(JSON.stringify(report, null, 2), 'datalens-report.json', 'application/json');
  };
  const exportCSV = () => {
    const seen = new Set(); const unique = [];
    parsed.rows.forEach((r) => { const k = JSON.stringify(r); if (!seen.has(k)) { seen.add(k); unique.push(r); } });
    downloadBlob(Papa.unparse(unique), 'datalens-cleaned-data.csv', 'text/csv');
  };
  const exportPrint = () => window.print();
 
  /* ---------- derived chart data ---------- */
  const typeData = analysis ? [
    { name: 'Numeric', value: analysis.meta.numericCols }, { name: 'Categorical', value: analysis.meta.categoricalCols },
    { name: 'Date', value: analysis.meta.dateCols }, { name: 'Boolean', value: analysis.meta.booleanCols },
    { name: 'Text', value: analysis.meta.textCols },
  ].filter((d) => d.value > 0) : [];
 
  const missingData = analysis ? Object.values(analysis.profiles)
    .filter((p) => p.missing > 0).sort((a, b) => b.missingPct - a.missingPct).slice(0, 10)
    .map((p) => ({ name: p.name, pct: +p.missingPct.toFixed(1) })) : [];
 
  const histData = analysis && selectedNumeric ? histogramBins(parsed.rows.map((r) => r[selectedNumeric])) : [];
  const catData = analysis && selectedCat ? analysis.profiles[selectedCat].stats.top.map((t) => ({ name: t.value, count: t.count })) : [];
 
  /* ---------- layout ---------- */
  return (
    <div className="dl-root min-h-screen w-full flex" style={{ background: theme.bg, color: theme.text, '--dl-border-strong': theme.borderStrong }}>
      <style>{FONT_STYLE}</style>
 
      {/* -------- Sidebar (desktop) -------- */}
      <aside className="dl-noprint hidden md:flex flex-col w-60 shrink-0 border-r p-4" style={{ borderColor: theme.border, background: theme.panel }}>
        <div className="flex items-center gap-2.5 px-1 mb-8">
          <div className="flex items-center justify-center rounded-xl" style={{ width: 36, height: 36, background: `${theme.accent}20` }}>
            <ScanLine size={19} style={{ color: theme.accent }} />
          </div>
          <div>
            <div className="dl-display font-bold text-[16px] leading-none" style={{ color: theme.text }}>DataLens</div>
            <div className="text-[10px] tracking-wide mt-1" style={{ color: theme.textMuted }}>ON-DEVICE ANALYST</div>
          </div>
        </div>
        <nav className="flex flex-col gap-1">
          {TABS.map((t) => {
            const disabled = t.id !== 'upload' && !analysis;
            const active = activeTab === t.id;
            const I = t.icon;
            return (
              <button key={t.id} disabled={disabled} onClick={() => setActiveTab(t.id)}
                className="flex items-center gap-3 px-3 py-2.5 rounded-xl text-sm font-medium transition-colors text-left"
                style={{
                  background: active ? `${theme.accent}18` : 'transparent',
                  color: disabled ? theme.textMuted : active ? theme.accent : theme.textSoft,
                  cursor: disabled ? 'not-allowed' : 'pointer', opacity: disabled ? 0.5 : 1,
                }}>
                <I size={17} /> {t.label}
              </button>
            );
          })}
        </nav>
        <div className="mt-auto pt-4 border-t" style={{ borderColor: theme.border }}>
          <div className="flex items-center gap-2 rounded-xl px-3 py-2.5" style={{ background: theme.panelAlt }}>
            <div className="rounded-full" style={{ width: 8, height: 8, background: theme.good, animation: 'dl-pulse 2s infinite' }} />
            <span className="text-xs font-medium" style={{ color: theme.textSoft }}>All processing is local</span>
          </div>
        </div>
      </aside>
 
      {/* -------- Main -------- */}
      <div className="flex-1 min-w-0 flex flex-col">
        {/* Topbar */}
        <header className="dl-noprint flex items-center justify-between gap-3 px-4 md:px-7 py-4 border-b sticky top-0 z-10" style={{ borderColor: theme.border, background: `${theme.bg}E8`, backdropFilter: 'blur(8px)' }}>
          <div className="flex items-center gap-2 md:hidden">
            <ScanLine size={18} style={{ color: theme.accent }} />
            <span className="dl-display font-bold text-sm">DataLens</span>
          </div>
          <div className="hidden md:block">
            <h1 className="dl-display text-lg font-semibold" style={{ color: theme.text }}>
              {TABS.find((t) => t.id === activeTab)?.label}
            </h1>
            {fileName && <p className="text-xs mt-0.5" style={{ color: theme.textMuted }}>{fileName}</p>}
          </div>
          <div className="flex items-center gap-2.5">
            <div className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 rounded-full border" style={{ borderColor: `${theme.good}55`, background: `${theme.good}12` }}>
              <Lock size={12} style={{ color: theme.good }} />
              <span className="text-[11px] font-semibold" style={{ color: theme.good }}>Nothing leaves your device</span>
            </div>
            <button onClick={() => setMode(mode === 'dark' ? 'light' : 'dark')}
              className="flex items-center justify-center rounded-full border" style={{ width: 34, height: 34, borderColor: theme.border, color: theme.textSoft }}>
              {mode === 'dark' ? <Sun size={15} /> : <Moon size={15} />}
            </button>
          </div>
        </header>
 
        {/* Mobile tab bar */}
        <div className="dl-noprint md:hidden flex overflow-x-auto dl-scrollbar gap-1.5 px-4 py-2 border-b" style={{ borderColor: theme.border }}>
          {TABS.map((t) => {
            const disabled = t.id !== 'upload' && !analysis;
            const I = t.icon;
            return (
              <button key={t.id} disabled={disabled} onClick={() => setActiveTab(t.id)}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium whitespace-nowrap"
                style={{ background: activeTab === t.id ? `${theme.accent}18` : theme.panelAlt, color: activeTab === t.id ? theme.accent : theme.textMuted, opacity: disabled ? 0.4 : 1 }}>
                <I size={13} /> {t.label}
              </button>
            );
          })}
        </div>
 
        <main className="dl-noprint flex-1 overflow-y-auto dl-scrollbar p-4 md:p-7">
          {/* ============ UPLOAD ============ */}
          {activeTab === 'upload' && (
            <div className="max-w-3xl mx-auto dl-fadeup">
              <div className="mb-6">
                <h2 className="dl-display text-2xl font-bold mb-2" style={{ color: theme.text }}>Analyze your data without ever uploading it.</h2>
                <p className="text-sm" style={{ color: theme.textSoft }}>DataLens parses, profiles, and visualizes your file entirely inside this browser tab. No server, no network call, no trace.</p>
              </div>
 
              <div
                onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={onDrop}
                onClick={() => fileInputRef.current?.click()}
                className="relative overflow-hidden rounded-2xl border-2 border-dashed flex flex-col items-center justify-center text-center px-6 py-16 cursor-pointer transition-colors"
                style={{ borderColor: dragOver ? theme.accent : theme.borderStrong, background: dragOver ? `${theme.accent}0C` : theme.panel }}>
                {loading && <div className="dl-scanbar" />}
                <input ref={fileInputRef} type="file" className="hidden" accept=".csv,.xlsx,.xls,.json,.txt,.pdf"
                  onChange={(e) => { const f = e.target.files?.[0]; if (f) handleFile(f); }} />
                {loading ? (
                  <>
                    <Loader2 className="animate-spin mb-4" size={30} style={{ color: theme.accent }} />
                    <p className="text-sm font-medium" style={{ color: theme.text }}>{loadingMsg}</p>
                  </>
                ) : (
                  <>
                    <div className="rounded-2xl flex items-center justify-center mb-4" style={{ width: 60, height: 60, background: `${theme.accent}18` }}>
                      <Upload size={26} style={{ color: theme.accent }} />
                    </div>
                    <p className="font-semibold mb-1" style={{ color: theme.text }}>Drop a file here, or click to browse</p>
                    <p className="text-xs" style={{ color: theme.textMuted }}>CSV · Excel (.xlsx/.xls) · JSON · TXT · PDF — processed entirely on-device</p>
                  </>
                )}
              </div>
 
              {error && (
                <div className="mt-4 flex items-start gap-2 rounded-xl border px-4 py-3 text-sm" style={{ borderColor: `${theme.danger}55`, background: `${theme.danger}12`, color: theme.danger }}>
                  <AlertTriangle size={16} className="shrink-0 mt-0.5" /> {error}
                </div>
              )}
 
              <div className="mt-6 flex items-center gap-3">
                <div className="h-px flex-1" style={{ background: theme.border }} />
                <span className="text-xs" style={{ color: theme.textMuted }}>or</span>
                <div className="h-px flex-1" style={{ background: theme.border }} />
              </div>
              <button onClick={handleSample} className="mt-6 w-full flex items-center justify-center gap-2 rounded-xl border py-3 text-sm font-medium transition-colors"
                style={{ borderColor: theme.border, color: theme.textSoft, background: theme.panel }}
                onMouseEnter={(e) => e.currentTarget.style.background = theme.panelHover}
                onMouseLeave={(e) => e.currentTarget.style.background = theme.panel}>
                <Beaker size={15} /> Try a sample dataset instead
              </button>
 
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mt-8">
                {[
                  { icon: Lock, title: 'Private by design', body: 'Files are parsed with in-browser JavaScript. Nothing is ever sent to a server.' },
                  { icon: Sparkles, title: 'Instant insights', body: 'Column types, missing values, outliers, trends, and correlations — detected automatically.' },
                  { icon: MessageSquare, title: 'Ask in plain English', body: 'Chat with your data using natural language, answered from the local analysis.' },
                ].map((f, i) => (
                  <div key={i} className="rounded-xl border p-4" style={{ borderColor: theme.border, background: theme.panel }}>
                    <f.icon size={17} style={{ color: theme.accent }} className="mb-2" />
                    <div className="text-sm font-semibold mb-1" style={{ color: theme.text }}>{f.title}</div>
                    <div className="text-xs leading-relaxed" style={{ color: theme.textMuted }}>{f.body}</div>
                  </div>
                ))}
              </div>
            </div>
          )}
 
          {/* ============ DASHBOARD ============ */}
          {activeTab === 'dashboard' && analysis && (
            <div className="dl-fadeup space-y-5">
              {parseWarning && (
                <div className="flex items-start gap-2 rounded-xl border px-4 py-3 text-sm" style={{ borderColor: `${theme.accent2}55`, background: `${theme.accent2}12`, color: theme.accent2 }}>
                  <Info size={15} className="shrink-0 mt-0.5" /> {parseWarning}
                </div>
              )}
              {/* KPI row */}
              <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                {[
                  { label: 'Rows', value: analysis.meta.rowCount.toLocaleString(), icon: Table2, color: theme.accent },
                  { label: 'Columns', value: analysis.meta.colCount, icon: Hash, color: theme.accent2 },
                  { label: 'Missing cells', value: `${Object.values(analysis.profiles).reduce((a, p) => a + p.missing, 0).toLocaleString()}`, icon: AlertTriangle, color: theme.danger },
                  { label: 'Duplicates', value: analysis.dup.count.toLocaleString(), icon: RefreshCw, color: theme.accent2 },
                  { label: 'Quality', value: `${analysis.quality.score}/100`, icon: ShieldCheck, color: theme.good },
                ].map((k, i) => (
                  <Panel theme={theme} key={i} className="p-4">
                    <div className="flex items-center justify-between mb-2">
                      <span className="text-[11px] font-medium" style={{ color: theme.textMuted }}>{k.label}</span>
                      <k.icon size={14} style={{ color: k.color }} />
                    </div>
                    <div className="dl-mono text-xl font-semibold" style={{ color: theme.text }}>{k.value}</div>
                  </Panel>
                ))}
              </div>
 
              <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                <Panel theme={theme} className="p-5" style={{ position: 'relative' }} >
                  <div ref={typesRef}>
                    <SectionTitle theme={theme} icon={LayoutDashboard} title="Column Types" subtitle="Detected automatically from your data" />
                    <ResponsiveContainer width="100%" height={220}>
                      <PieChart>
                        <Pie data={typeData} dataKey="value" nameKey="name" innerRadius={55} outerRadius={85} paddingAngle={3}>
                          {typeData.map((_, i) => <Cell key={i} fill={CHART_PALETTE[i % CHART_PALETTE.length]} stroke="none" />)}
                        </Pie>
                        <Tooltip contentStyle={{ background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 10, fontSize: 12 }} />
                        <Legend wrapperStyle={{ fontSize: 11, color: theme.textSoft }} />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                  <button onClick={() => downloadSvgAsPng(typesRef.current, 'column-types.png')} className="absolute top-5 right-5 text-[11px] font-medium px-2.5 py-1 rounded-lg border" style={{ borderColor: theme.border, color: theme.textMuted }}>PNG</button>
                </Panel>
 
                <Panel theme={theme} className="p-5" style={{ position: 'relative' }}>
                  <div ref={missingRef}>
                    <SectionTitle theme={theme} icon={AlertTriangle} title="Missing Values" subtitle={missingData.length ? 'Top columns by % missing' : 'No missing values found'} />
                    {missingData.length > 0 ? (
                      <ResponsiveContainer width="100%" height={220}>
                        <BarChart data={missingData} layout="vertical" margin={{ left: 8 }}>
                          <CartesianGrid stroke={theme.chartGrid} horizontal={false} />
                          <XAxis type="number" tick={{ fontSize: 10, fill: theme.textMuted }} unit="%" />
                          <YAxis type="category" dataKey="name" width={90} tick={{ fontSize: 10, fill: theme.textMuted }} />
                          <Tooltip contentStyle={{ background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 10, fontSize: 12 }} />
                          <Bar dataKey="pct" fill={theme.danger} radius={[0, 5, 5, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    ) : (
                      <div className="h-[220px] flex flex-col items-center justify-center" style={{ color: theme.good }}>
                        <CheckCircle2 size={26} className="mb-2" /><span className="text-sm">Dataset is complete</span>
                      </div>
                    )}
                  </div>
                  {missingData.length > 0 && <button onClick={() => downloadSvgAsPng(missingRef.current, 'missing-values.png')} className="absolute top-5 right-5 text-[11px] font-medium px-2.5 py-1 rounded-lg border" style={{ borderColor: theme.border, color: theme.textMuted }}>PNG</button>}
                </Panel>
 
                {analysis.numericCols.length > 0 && (
                  <Panel theme={theme} className="p-5" style={{ position: 'relative' }}>
                    <div ref={histRef}>
                      <SectionTitle theme={theme} icon={Hash} title="Distribution"
                        right={
                          <select value={selectedNumeric || ''} onChange={(e) => setSelectedNumeric(e.target.value)}
                            className="text-xs rounded-lg px-2 py-1.5 border outline-none" style={{ background: theme.panelAlt, borderColor: theme.border, color: theme.text }}>
                            {analysis.numericCols.map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        } />
                      <ResponsiveContainer width="100%" height={200}>
                        <BarChart data={histData}>
                          <CartesianGrid stroke={theme.chartGrid} vertical={false} />
                          <XAxis dataKey="bin" tick={{ fontSize: 9, fill: theme.textMuted }} interval={1} angle={-20} textAnchor="end" height={45} />
                          <YAxis tick={{ fontSize: 10, fill: theme.textMuted }} />
                          <Tooltip contentStyle={{ background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 10, fontSize: 12 }} />
                          <Bar dataKey="count" fill={theme.accent} radius={[4, 4, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                    <button onClick={() => downloadSvgAsPng(histRef.current, `${selectedNumeric}-distribution.png`)} className="absolute top-5 right-24 text-[11px] font-medium px-2.5 py-1 rounded-lg border" style={{ borderColor: theme.border, color: theme.textMuted }}>PNG</button>
                  </Panel>
                )}
 
                {analysis.categoricalCols.length > 0 && (
                  <Panel theme={theme} className="p-5" style={{ position: 'relative' }}>
                    <div ref={catRef}>
                      <SectionTitle theme={theme} icon={TypeIcon} title="Top Categories"
                        right={
                          <select value={selectedCat || ''} onChange={(e) => setSelectedCat(e.target.value)}
                            className="text-xs rounded-lg px-2 py-1.5 border outline-none" style={{ background: theme.panelAlt, borderColor: theme.border, color: theme.text }}>
                            {analysis.categoricalCols.map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        } />
                      <ResponsiveContainer width="100%" height={200}>
                        <BarChart data={catData}>
                          <CartesianGrid stroke={theme.chartGrid} vertical={false} />
                          <XAxis dataKey="name" tick={{ fontSize: 9, fill: theme.textMuted }} interval={0} angle={-20} textAnchor="end" height={45} />
                          <YAxis tick={{ fontSize: 10, fill: theme.textMuted }} />
                          <Tooltip contentStyle={{ background: theme.panel, border: `1px solid ${theme.border}`, borderRadius: 10, fontSize: 12 }} />
                          <Bar dataKey="count" fill={theme.accent2} radius={[4, 4, 0, 0]} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                    <button onClick={() => downloadSvgAsPng(catRef.current, `${selectedCat}-frequency.png`)} className="absolute top-5 right-24 text-[11px] font-medium px-2.5 py-1 rounded-lg border" style={{ borderColor: theme.border, color: theme.textMuted }}>PNG</button>
                  </Panel>
                )}
              </div>
 
              {analysis.corrMatrixCols.length >= 2 && (
                <Panel theme={theme} className="p-5">
                  <SectionTitle theme={theme} icon={TrendingUp} title="Correlation Matrix" subtitle="Pearson r between numeric columns (up to 8 shown)" />
                  <div className="overflow-x-auto dl-scrollbar">
                    <table className="text-[11px] dl-mono border-separate" style={{ borderSpacing: 3 }}>
                      <thead>
                        <tr>
                          <td />
                          {analysis.corrMatrixCols.map((c) => <td key={c} className="px-1 pb-1 text-center" style={{ color: theme.textMuted, maxWidth: 60 }}>{c.slice(0, 8)}</td>)}
                        </tr>
                      </thead>
                      <tbody>
                        {analysis.corrMatrixCols.map((a) => (
                          <tr key={a}>
                            <td className="pr-2 whitespace-nowrap text-right" style={{ color: theme.textMuted }}>{a.slice(0, 12)}</td>
                            {analysis.corrMatrixCols.map((b) => {
                              const r = analysis.fullCorr[a][b];
                              return (
                                <td key={b} className="rounded-md text-center" style={{ width: 46, height: 30, background: corrColor(r, theme), color: mode === 'dark' ? '#08131F' : '#08131F', fontWeight: 600 }}>
                                  {r === null ? '—' : r.toFixed(2)}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </Panel>
              )}
 
              <Panel theme={theme} className="p-5">
                <SectionTitle theme={theme} icon={Table2} title="Data Preview" subtitle={`Showing first 8 of ${analysis.meta.rowCount.toLocaleString()} rows`} />
                <div className="overflow-x-auto dl-scrollbar">
                  <table className="text-xs w-full border-collapse">
                    <thead>
                      <tr>{analysis.columns.map((c) => (
                        <th key={c} className="text-left px-3 py-2 whitespace-nowrap font-semibold border-b" style={{ color: theme.textSoft, borderColor: theme.border }}>
                          <div className="flex items-center gap-1.5">{c} <TypeChip type={analysis.profiles[c].type} theme={theme} /></div>
                        </th>
                      ))}</tr>
                    </thead>
                    <tbody>
                      {parsed.rows.slice(0, 8).map((row, i) => (
                        <tr key={i}>
                          {analysis.columns.map((c) => (
                            <td key={c} className="px-3 py-2 whitespace-nowrap dl-mono border-b" style={{ color: theme.text, borderColor: theme.border, opacity: 0.9 }}>
                              {row[c] === null || row[c] === undefined || row[c] === '' ? <span style={{ color: theme.danger }}>∅</span> : String(row[c])}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Panel>
            </div>
          )}
 
          {/* ============ INSIGHTS ============ */}
          {activeTab === 'insights' && analysis && (
            <div className="dl-fadeup space-y-5 max-w-4xl mx-auto">
              <Panel theme={theme} className="p-5">
                <SectionTitle theme={theme} icon={Sparkles} title="Automatic Insights" subtitle="Generated from your dataset's statistical profile" />
                <div className="space-y-2.5">
                  {analysis.insights.map((ins, i) => {
                    const conf = { warning: { icon: AlertTriangle, color: theme.danger }, success: { icon: CheckCircle2, color: theme.good }, info: { icon: Info, color: theme.accent } }[ins.type];
                    const I = conf.icon;
                    return (
                      <div key={i} className="flex items-start gap-3 rounded-xl border px-4 py-3" style={{ borderColor: theme.border, background: theme.panelAlt }}>
                        <I size={15} style={{ color: conf.color }} className="shrink-0 mt-0.5" />
                        <span className="text-sm leading-relaxed" style={{ color: theme.text }}>{ins.text}</span>
                      </div>
                    );
                  })}
                </div>
              </Panel>
 
              <Panel theme={theme} className="p-5">
                <SectionTitle theme={theme} icon={Sparkle} title="Smart Recommendations" subtitle="Suggested next steps to improve this dataset" />
                <div className="space-y-2.5">
                  {analysis.recommendations.map((r, i) => (
                    <div key={i} className="flex items-start gap-3 rounded-xl border px-4 py-3" style={{ borderColor: theme.border, background: theme.panelAlt }}>
                      <span className="dl-mono text-[10px] font-bold rounded-md px-1.5 py-0.5 shrink-0 mt-0.5" style={{ color: theme.accent, background: `${theme.accent}18` }}>{String(i + 1).padStart(2, '0')}</span>
                      <span className="text-sm leading-relaxed" style={{ color: theme.text }}>{r}</span>
                    </div>
                  ))}
                </div>
              </Panel>
 
              {analysis.trends.length > 0 && (
                <Panel theme={theme} className="p-5">
                  <SectionTitle theme={theme} icon={TrendingUp} title="Detected Trends" subtitle="Direction across row order (proxy for sequence/time)" />
                  <div className="grid sm:grid-cols-2 gap-3">
                    {analysis.trends.map((t, i) => (
                      <div key={i} className="flex items-center gap-3 rounded-xl border px-4 py-3" style={{ borderColor: theme.border }}>
                        {t.direction === 'increasing' ? <TrendingUp size={18} style={{ color: theme.good }} /> : <TrendingDown size={18} style={{ color: theme.danger }} />}
                        <div>
                          <div className="text-sm font-medium" style={{ color: theme.text }}>{t.col}</div>
                          <div className="text-xs dl-mono" style={{ color: theme.textMuted }}>{t.direction} · r = {t.r.toFixed(2)}</div>
                        </div>
                      </div>
                    ))}
                  </div>
                </Panel>
              )}
            </div>
          )}
 
          {/* ============ QUALITY ============ */}
          {activeTab === 'quality' && analysis && (
            <div className="dl-fadeup space-y-5 max-w-4xl mx-auto">
              <Panel theme={theme} className="p-6 flex flex-col sm:flex-row items-center gap-6">
                <Gauge value={analysis.quality.score} theme={theme} />
                <div className="flex-1 w-full space-y-3">
                  <MetricBar label="Completeness" value={analysis.quality.completeness} theme={theme} color={theme.accent} />
                  <MetricBar label="Uniqueness" value={analysis.quality.uniqueness} theme={theme} color={theme.accent2} />
                  <MetricBar label="Validity" value={analysis.quality.validity} theme={theme} color={theme.good} />
                </div>
              </Panel>
 
              <Panel theme={theme} className="p-5">
                <SectionTitle theme={theme} icon={Table2} title="Column-by-Column Quality" />
                <div className="overflow-x-auto dl-scrollbar">
                  <table className="text-xs w-full border-collapse">
                    <thead>
                      <tr>
                        {['Column', 'Type', 'Missing', 'Unique', 'Outliers'].map((h) => (
                          <th key={h} className="text-left px-3 py-2 font-semibold border-b" style={{ color: theme.textSoft, borderColor: theme.border }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {analysis.columns.map((c) => {
                        const p = analysis.profiles[c];
                        return (
                          <tr key={c}>
                            <td className="px-3 py-2 border-b font-medium" style={{ color: theme.text, borderColor: theme.border }}>{c}</td>
                            <td className="px-3 py-2 border-b" style={{ borderColor: theme.border }}><TypeChip type={p.type} theme={theme} /></td>
                            <td className="px-3 py-2 border-b dl-mono" style={{ color: p.missingPct > 5 ? theme.danger : theme.textMuted, borderColor: theme.border }}>{p.missing} ({p.missingPct.toFixed(1)}%)</td>
                            <td className="px-3 py-2 border-b dl-mono" style={{ color: theme.textMuted, borderColor: theme.border }}>{p.unique}</td>
                            <td className="px-3 py-2 border-b dl-mono" style={{ color: (p.stats?.outlierCount || 0) > 0 ? theme.accent2 : theme.textMuted, borderColor: theme.border }}>{p.type === 'numeric' ? (p.stats?.outlierCount ?? 0) : '—'}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </Panel>
            </div>
          )}
 
          {/* ============ CHAT ============ */}
          {activeTab === 'chat' && analysis && (
            <div className="dl-fadeup max-w-3xl mx-auto flex flex-col" style={{ height: 'calc(100vh - 170px)' }}>
              <Panel theme={theme} className="flex-1 flex flex-col overflow-hidden">
                <div className="flex-1 overflow-y-auto dl-scrollbar p-5 space-y-3">
                  {chatMessages.map((m, i) => (
                    <div key={i} className={`flex ${m.role === 'user' ? 'justify-end' : 'justify-start'}`}>
                      <div className="max-w-[80%] rounded-2xl px-4 py-2.5 text-sm leading-relaxed"
                        style={{
                          background: m.role === 'user' ? theme.accent : theme.panelAlt,
                          color: m.role === 'user' ? '#06201C' : theme.text,
                          borderTopRightRadius: m.role === 'user' ? 4 : 16,
                          borderTopLeftRadius: m.role === 'assistant' ? 4 : 16,
                        }}>
                        {m.text}
                      </div>
                    </div>
                  ))}
                  <div ref={chatEndRef} />
                </div>
                {suggestedQuestions.length > 0 && (
                  <div className="flex gap-2 overflow-x-auto dl-scrollbar px-5 pb-3">
                    {suggestedQuestions.map((q, i) => (
                      <button key={i} onClick={() => { setChatInput(''); setChatMessages((m) => [...m, { role: 'user', text: q }]); setTimeout(() => setChatMessages((m) => [...m, { role: 'assistant', text: answerQuestion(q, analysis) }]), 300); }}
                        className="whitespace-nowrap text-xs px-3 py-1.5 rounded-full border" style={{ borderColor: theme.border, color: theme.textSoft }}>
                        {q}
                      </button>
                    ))}
                  </div>
                )}
                <div className="flex items-center gap-2 p-4 border-t" style={{ borderColor: theme.border }}>
                  <input value={chatInput} onChange={(e) => setChatInput(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter') sendChat(); }}
                    placeholder="Ask about your data…"
                    className="flex-1 rounded-xl px-4 py-2.5 text-sm outline-none border"
                    style={{ background: theme.panelAlt, borderColor: theme.border, color: theme.text }} />
                  <button onClick={sendChat} className="flex items-center justify-center rounded-xl shrink-0" style={{ width: 42, height: 42, background: theme.accent, color: '#06201C' }}>
                    <Send size={16} />
                  </button>
                </div>
              </Panel>
            </div>
          )}
 
          {/* ============ EXPORT ============ */}
          {activeTab === 'export' && analysis && (
            <div className="dl-fadeup max-w-2xl mx-auto space-y-4">
              {[
                { icon: FileText, title: 'Download PDF report', body: 'Opens your browser\'s print dialog with a formatted analysis report — choose "Save as PDF" as the destination.', action: exportPrint, cta: 'Print / Save as PDF' },
                { icon: Download, title: 'Download JSON report', body: 'Full machine-readable report with stats, insights, correlations, and recommendations.', action: exportJSON, cta: 'Download .json' },
                { icon: Table2, title: 'Download cleaned CSV', body: 'Your original data with exact duplicate rows removed.', action: exportCSV, cta: 'Download .csv' },
              ].map((e, i) => (
                <Panel theme={theme} key={i} className="p-5 flex items-center gap-4">
                  <IconBadge Icon={e.icon} color={theme.accent} bg={`${theme.accent}18`} />
                  <div className="flex-1">
                    <div className="text-sm font-semibold" style={{ color: theme.text }}>{e.title}</div>
                    <div className="text-xs mt-0.5" style={{ color: theme.textMuted }}>{e.body}</div>
                  </div>
                  <button onClick={e.action} className="text-xs font-semibold px-3.5 py-2 rounded-lg shrink-0" style={{ background: theme.accent, color: '#06201C' }}>{e.cta}</button>
                </Panel>
              ))}
              <p className="text-xs text-center pt-2" style={{ color: theme.textMuted }}>Charts can also be exported as PNG directly from the Dashboard tab.</p>
            </div>
          )}
 
          {!analysis && activeTab !== 'upload' && (
            <div className="flex flex-col items-center justify-center h-full py-24 text-center">
              <FileText size={28} style={{ color: theme.textMuted }} className="mb-3" />
              <p className="text-sm" style={{ color: theme.textMuted }}>Upload a dataset to unlock this view.</p>
            </div>
          )}
        </main>
      </div>
 
      {/* -------- Print-only report -------- */}
      {analysis && (
        <div className="dl-print-only p-10" style={{ color: '#0F1A2E' }}>
          <h1 className="dl-display" style={{ fontSize: 24, fontWeight: 700 }}>DataLens — Data Analysis Report</h1>
          <p style={{ fontSize: 12, color: '#4C5C77', marginTop: 4 }}>{fileName} · Generated {new Date().toLocaleString()}</p>
          <h2 style={{ fontSize: 16, fontWeight: 700, marginTop: 24 }}>Overview</h2>
          <p style={{ fontSize: 12 }}>{analysis.meta.rowCount.toLocaleString()} rows · {analysis.meta.colCount} columns · Quality score {analysis.quality.score}/100 · {analysis.dup.count} duplicate rows</p>
          <h2 style={{ fontSize: 16, fontWeight: 700, marginTop: 20 }}>Insights</h2>
          <ul style={{ fontSize: 12, lineHeight: 1.7 }}>{analysis.insights.map((i, idx) => <li key={idx}>• {i.text}</li>)}</ul>
          <h2 style={{ fontSize: 16, fontWeight: 700, marginTop: 20 }}>Recommendations</h2>
          <ul style={{ fontSize: 12, lineHeight: 1.7 }}>{analysis.recommendations.map((r, idx) => <li key={idx}>{idx + 1}. {r}</li>)}</ul>
          <h2 style={{ fontSize: 16, fontWeight: 700, marginTop: 20 }}>Column Profiles</h2>
          <table style={{ fontSize: 11, width: '100%', borderCollapse: 'collapse', marginTop: 6 }}>
            <thead><tr>{['Column', 'Type', 'Missing %', 'Unique'].map((h) => <th key={h} style={{ textAlign: 'left', borderBottom: '1px solid #ccc', padding: '4px 6px' }}>{h}</th>)}</tr></thead>
            <tbody>{analysis.columns.map((c) => { const p = analysis.profiles[c]; return (
              <tr key={c}><td style={{ padding: '4px 6px', borderBottom: '1px solid #eee' }}>{c}</td><td style={{ padding: '4px 6px', borderBottom: '1px solid #eee' }}>{p.type}</td><td style={{ padding: '4px 6px', borderBottom: '1px solid #eee' }}>{p.missingPct.toFixed(1)}%</td><td style={{ padding: '4px 6px', borderBottom: '1px solid #eee' }}>{p.unique}</td></tr>
            ); })}</tbody>
          </table>
          <p style={{ fontSize: 10, color: '#8494AC', marginTop: 24 }}>Generated entirely on-device by DataLens. No data was uploaded to any server.</p>
        </div>
      )}
    </div>
  );
}
