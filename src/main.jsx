import React, { useMemo, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import pdfWorker from 'pdfjs-dist/build/pdf.worker.min.mjs?url';
import {
  AlertTriangle, ArrowRight, Check, CheckCircle2, ChevronDown,
  Download, FileSpreadsheet, FileText, Info, LoaderCircle, Lock,
  RotateCcw, ShieldCheck, UploadCloud, X, XCircle
} from 'lucide-react';
import './styles.css';

pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

const DOC_PATTERN = /QC\d{8}-[A-Z0-9]+-(?:DRW|MOD)-[A-Z0-9-]+-\d{5}(?:-(?:PDF|CAD))?/i;
const HYPHENS = /[‐‑‒–—−]/g;

function clean(value) {
  return String(value ?? '').replace(HYPHENS, '-').replace(/\s+/g, ' ').trim();
}

function baseNumber(value) {
  return clean(value).toUpperCase().replace(/-(PDF|CAD)$/i, '');
}

function revision(value) {
  const s = clean(value).replace(/^R/i, '');
  if (!s) return '';
  return /^\d+$/.test(s) ? s.padStart(2, '0') : s.toUpperCase();
}

function normalizedTitle(value) {
  return clean(value).toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
}

function excelDate(value) {
  if (value instanceof Date) return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`;
  if (typeof value === 'number' && value > 20000) {
    const d = XLSX.SSF.parse_date_code(value);
    return d ? `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}` : clean(value);
  }
  const s = clean(value);
  if (!s) return '';
  const d = new Date(s);
  return Number.isNaN(d.valueOf()) ? s : d.toISOString().slice(0, 10);
}

function formatDate(value) {
  const iso = excelDate(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) return iso || '—';
  return new Intl.DateTimeFormat('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }).format(new Date(`${iso}T00:00:00`));
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellDates: true, cellStyles: true });
}

function sheetRows(workbook, sheetName) {
  return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: true });
}

function findHeader(rows, phrases) {
  for (let r = 0; r < Math.min(rows.length, 30); r += 1) {
    const normalized = rows[r].map((v) => clean(v).toLowerCase());
    if (phrases.some((p) => normalized.includes(p))) return r;
  }
  return -1;
}

function headerIndex(headers, aliases) {
  const hs = headers.map((h) => clean(h).toLowerCase());
  return hs.findIndex((h) => aliases.some((a) => h === a || h.includes(a)));
}

function parseMetadata(workbook) {
  const records = [];
  for (const name of workbook.SheetNames) {
    const rows = sheetRows(workbook, name);
    const h = findHeader(rows, ['document no', 'document number']);
    if (h < 0) continue;
    const headers = rows[h];
    const idx = {
      doc: headerIndex(headers, ['document no', 'document number']),
      rev: headerIndex(headers, ['revision']),
      title: headerIndex(headers, ['title']),
      status: headerIndex(headers, ['status']),
      date: headerIndex(headers, ['revision date', 'issue date']),
      model: headerIndex(headers, ['reference', 'model file reference']),
    };
    if (idx.doc < 0) continue;
    rows.slice(h + 1).forEach((row, offset) => {
      const doc = clean(row[idx.doc]);
      if (!DOC_PATTERN.test(doc) || !/-(PDF|CAD)$/i.test(doc)) return;
      records.push({
        sourceRow: h + offset + 2, sheet: name, doc, base: baseNumber(doc),
        revision: idx.rev >= 0 ? revision(row[idx.rev]) : '',
        title: idx.title >= 0 ? clean(row[idx.title]) : '',
        status: idx.status >= 0 ? clean(row[idx.status]) : '',
        issueDate: idx.date >= 0 ? excelDate(row[idx.date]) : '',
        modelRef: idx.model >= 0 ? clean(row[idx.model]) : '',
      });
    });
  }
  return records;
}

function parseMidp(workbook) {
  const records = [];
  // Limit this release to drawing registers. Some non-drawing tabs in real MIDPs
  // have a saved range extending to Excel's maximum column, making a full scan
  // needlessly expensive in the browser.
  for (const name of workbook.SheetNames.filter((n) => /^MIDP-DRW/i.test(n))) {
    const rows = sheetRows(workbook, name);
    rows.forEach((row, r) => row.forEach((cell, c) => {
      const doc = clean(cell);
      if (!DOC_PATTERN.test(doc) || /-(PDF|CAD)$/i.test(doc) || !/-DRW-/i.test(doc)) return;
      // The supplied MIDP uses the selected drawing number followed by title/scope,
      // with revision/date/model-reference in the same record.
      const revisionValue = row[13] ?? '';
      const issueValue = row[16] ?? '';
      const modelValue = row[24] ?? '';
      records.push({
        sourceRow: r + 1, sheet: name, doc, base: baseNumber(doc),
        title: clean(row[c + 1]), revision: revision(revisionValue),
        issueDate: excelDate(issueValue), modelRef: clean(modelValue),
        status: clean(row[10]), scale: clean(row[14]),
      });
    }));
  }
  return records;
}

function valueAfter(text, label, nextLabels = []) {
  const escaped = label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const end = nextLabels.length
    ? `(?=${nextLabels.map((v) => v.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})`
    : '$';
  return clean((text.match(new RegExp(`${escaped}\\s*:?\\s*([\\s\\S]*?)${end}`, 'i')) || [])[1]);
}

function titleFromLayout(items, pageWidth, rotation = 0) {
  // PDF text streams do not necessarily follow visual reading order. Rebuild
  // the right-hand title-block lines from their coordinates first.
  const rotated = rotation === 90 || rotation === 270;
  const words = items
    .map((item) => ({
      text: clean(item.str),
      x: rotated ? item.transform?.[5] : item.transform?.[4],
      y: rotated ? item.transform?.[4] : item.transform?.[5],
    }))
    .filter((item) => item.text && item.x > pageWidth * 0.8)
    .sort((a, b) => Math.abs(a.y - b.y) > 2
      ? (rotated ? a.y - b.y : b.y - a.y)
      : a.x - b.x);
  const lines = [];
  for (const word of words) {
    let line = lines.find((candidate) => Math.abs(candidate.y - word.y) <= 2);
    if (!line) { line = { y: word.y, words: [] }; lines.push(line); }
    line.words.push(word);
  }
  const textLines = lines
    .sort((a, b) => rotated ? a.y - b.y : b.y - a.y)
    .map((line) => line.words.sort((a, b) => a.x - b.x).map((word) => word.text).join(' '));
  const start = textLines.findIndex((line) => /Drawing\s*Title\s*:/i.test(line));
  if (start < 0) return '';
  const relativeEnd = textLines.slice(start + 1).findIndex((line) => /Info\.?\s*Classification/i.test(line));
  const end = relativeEnd < 0 ? Math.min(start + 8, textLines.length) : start + 1 + relativeEnd;
  const titleLines = textLines.slice(start + 1, end)
    .filter((line) => !/^ISSUED FOR (APPROVAL|CONSTRUCTION)$/i.test(line))
    .filter((line) => !/^[A-G]$/i.test(line)); // drawing border zone markers
  const descriptorIndex = titleLines.findIndex((line) => /^(SHOP|GENERAL|DETAIL|SECTION|ELEVATION|PLAN)\b/i.test(line));
  if (descriptorIndex <= 0) return clean(titleLines.join(' '));
  // Title blocks often place building/level context above the descriptive title,
  // while registers store that context at the end. Reorder only those visual lines.
  return clean([...titleLines.slice(descriptorIndex), ...titleLines.slice(0, descriptorIndex)].join(' '));
}

async function parsePdf(file) {
  const data = new Uint8Array(await file.arrayBuffer());
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  let text = '';
  let layoutTitle = '';
  for (let p = 1; p <= pdf.numPages; p += 1) {
    const page = await pdf.getPage(p);
    const content = await page.getTextContent();
    if (p === 1) {
      const viewport = page.getViewport({ scale: 1 });
      layoutTitle = titleFromLayout(content.items, viewport.width, page.rotate);
    }
    text += ` ${content.items.map((i) => i.str).join(' ')}`;
  }
  text = clean(text);
  const allDocs = [...text.matchAll(new RegExp(DOC_PATTERN.source, 'gi'))].map((m) => clean(m[0]));
  const fileDoc = clean(file.name).match(DOC_PATTERN)?.[0];
  const drawingDoc = allDocs.find((d) => /-DRW-/.test(d)) || fileDoc || '';
  const rev = text.match(/\b([A-Z0-9]{1,3})\s+ISSUED FOR (?:APPROVAL|CONSTRUCTION)/i)?.[1]
    || text.match(/Revision No\.?\s*:?\s*(?!Issue\b)([A-Z0-9]+)/i)?.[1] || '';
  const date = text.match(/Issue Date\s*:?\s*(\d{2}[-/.]\d{2}[-/.]\d{4})/i)?.[1]
    || text.match(/ISSUED FOR (?:APPROVAL|CONSTRUCTION)\s+(\d{2}[-/.]\d{2}[-/.]\d{4})/i)?.[1] || '';
  const model = text.match(/QC\d{8}-[A-Z0-9]+-MOD-[A-Z0-9-]+-\d{5}/i)?.[0] || '';
  let title = layoutTitle || valueAfter(text, 'Drawing Title', ['Info. Classification Level', 'Information Classification Level']);
  title = title.replace(/^ISSUED FOR (APPROVAL|CONSTRUCTION)\s*/i, '');
  return {
    fileName: file.name, doc: drawingDoc, base: baseNumber(drawingDoc), title,
    revision: revision(rev), issueDate: excelDate(date.split(/[-/.]/).reverse().join('-')),
    modelRef: clean(model), status: text.match(/ISSUED FOR (APPROVAL|CONSTRUCTION)/i)?.[0] || '',
    scale: text.match(/\b\d+\s*:\s*\d+\s*@\s*A\d\b/i)?.[0] || '', pageCount: pdf.numPages,
  };
}

function compareValue(field, a, b, c) {
  const present = [a, b, c].filter(Boolean);
  if (!present.length) return { status: 'missing', note: 'Not found in any source' };
  if (present.length < 3) return { status: 'warning', note: 'Missing from one or more sources' };
  const normalize = field === 'revision' ? revision
    : field === 'title' ? normalizedTitle
      : field === 'issueDate' ? excelDate : baseNumber;
  const values = [a, b, c].map(normalize);
  return new Set(values).size === 1
    ? { status: 'pass', note: 'Matches across all sources' }
    : { status: 'fail', note: 'Values do not match' };
}

const FIELD_LABELS = {
  doc: 'Drawing number', revision: 'Revision', title: 'Drawing title', issueDate: 'Issue date',
  modelRef: 'Model reference', status: 'Status', scale: 'Scale'
};


function normalizeMidpField(value) {
  return clean(value).toUpperCase().replace(/\\s+/g, ' ').trim();
}

function midpCategory(row, statusIndex, codeIndex) {
  const candidates = [
    statusIndex >= 0 ? normalizeMidpField(row[statusIndex]) : '',
    codeIndex >= 0 ? normalizeMidpField(row[codeIndex]) : '',
  ];
  const combined = candidates.join(' ');
  if (/\\bNYS\\b|NOT\\s+YET\\s+SUBMITTED/i.test(combined)) return 'NYS';
  if (/\\bUR\\b|UNDER\\s+REVIEW/i.test(combined)) return 'UR';
  if (/\\bCODE\\s*D\\b|^D$/i.test(combined)) return 'Code D';
  if (/\\bCODE\\s*C\\b|^C$/i.test(combined)) return 'Code C';
  if (/\\bCODE\\s*B\\b|^B$/i.test(combined)) return 'Code B';
  return candidates.find(Boolean) || '';
}

function normalizeCompareValue(value) {
  return clean(value).replace(/\\s+/g, ' ').trim().toUpperCase();
}

function findDrawingSheet(workbook) {
  // Both TAJ MIDP and GLS TIDP use this drawing-register sheet.
  // Prefer it explicitly; do not rely on generic MIDP/TIDP sheet matching.
  const exact = workbook.SheetNames.find((name) => clean(name).toUpperCase() === 'MIDP-DRW-SWD');
  if (exact) return exact;

  // Fallback only if the workbook uses a slightly different sheet name.
  return workbook.SheetNames.find((name) => {
    const n = clean(name).toUpperCase();
    return n.includes('MIDP-DRW-SWD') || n.includes('MIDP') && n.includes('DRW');
  });
}

function normalizeHex(value) {
  return clean(value).replace(/^#/, '').toUpperCase();
}

function cellFillColor(cell) {
  const fg = cell?.s?.fgColor || cell?.s?.fill?.fgColor;
  if (!fg) return '';
  if (fg.rgb) return normalizeHex(fg.rgb);
  if (fg.indexed != null) return String(fg.indexed);
  return '';
}

function colorActionFromCell(cell) {
  const color = cellFillColor(cell);
  if (!color) return '';

  // Common Excel ARGB/RGB values. Also allow the dominant RGB portion
  // when the workbook stores an alpha channel.
  const rgb = color.length === 8 ? color.slice(2) : color;

  if (
    ['FF0000', 'C00000', 'FF3333', 'F4CCCC', 'FFC7CE'].includes(rgb)
    || /FF0000$|C00000$/.test(color)
  ) return 'REMOVE';

  if (
    ['00FF00', '008000', '70AD47', '92D050', 'C6EFCE'].includes(rgb)
    || /00FF00$|008000$|92D050$/.test(color)
  ) return 'ADD';

  if (
    ['FFFF00', 'FFD966', 'FFF2CC', 'FFC000', 'FFEB9C'].includes(rgb)
    || /FFFF00$|FFD966$|FFC000$/.test(color)
  ) return 'MODIFY';

  return '';
}

function detectRowAction(sheet, rowIndex, columnCount) {
  const actions = [];
  for (let c = 0; c < columnCount; c += 1) {
    const action = colorActionFromCell(sheet?.[XLSX.utils.encode_cell({ r: rowIndex, c })]);
    if (action && !actions.includes(action)) actions.push(action);
  }
  return actions.length === 1 ? actions[0] : actions.length > 1 ? 'MULTIPLE' : '';
}

function parseDrawingRegister(workbook, { filterGLS = false } = {}) {
  const targetSheet = findDrawingSheet(workbook);
  if (!targetSheet) {
    throw new Error(`No MIDP-DRW-SWD sheet was found. Available sheets: ${workbook.SheetNames.join(', ')}`);
  }

  const worksheet = workbook.Sheets[targetSheet];
  const rows = sheetRows(workbook, targetSheet);
  const headerRow = 3; // Excel Row 4
  if (!rows[headerRow]) throw new Error(`${targetSheet}: Excel Row 4 could not be read.`);

  const suitabilityIndex = 3; // D
  const docIndex = 5;         // F
  const titleIndex = 6;       // G
  const revisionIndex = 13;   // N
  const ownerIndex = 25;      // Z

  const rawHeaders = rows[headerRow].map((h, i) => clean(h) || `Column ${i + 1}`);
  const normalizedHeaders = rawHeaders.map((h) => clean(h).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim());

  // Prefer the actual column headed "Status" for Code B / Code C / UR checks.
  // Status is fixed to Column K (11th Excel column) in the drawing register.
  const statusIndex = 10; // K

  const records = [];
  const duplicateKeys = new Set();
  const seen = new Set();

  rows.slice(headerRow + 1).forEach((row, offset) => {
    const doc = clean(row[docIndex]);
    if (!doc) return;

    const key = baseNumber(doc);
    if (!key) return;

    if (seen.has(key)) duplicateKeys.add(key);
    seen.add(key);

    if (filterGLS && clean(row[ownerIndex]).toUpperCase() !== 'GLS') return;

    const values = {};
    row.forEach((value, index) => {
      const header = rawHeaders[index] || `Column ${index + 1}`;
      const cleaned = clean(value);
      if (cleaned) values[header] = cleaned;
    });

    const rowAction = detectRowAction(worksheet, headerRow + 1 + offset, Math.max(rawHeaders.length, row.length));

    records.push({
      sourceRow: headerRow + offset + 2,
      sheet: targetSheet,
      doc,
      base: key,
      title: clean(row[titleIndex]),
      revision: revision(row[revisionIndex]),
      status: clean(row[statusIndex]),
      suitabilityStatus: clean(row[suitabilityIndex]),
      statusColumn: statusIndex + 1,
      category: clean(row[statusIndex]),
      subOwner: clean(row[ownerIndex]),
      requestedAction: rowAction,
      values,
      headers: Object.keys(values),
    });
  });

  if (!records.length) {
    throw new Error(filterGLS
      ? `${targetSheet} was read successfully, but no drawing rows with Sub-Owner = GLS were found.`
      : `${targetSheet} was read successfully, but no drawing rows were found.`);
  }

  const map = new Map();
  records.forEach((record) => {
    if (!map.has(record.base)) map.set(record.base, record);
  });

  return { records: [...map.values()], duplicates: [...duplicateKeys].sort(), totalRows: records.length };
}

function parseTajMidp(workbook) {
  return parseDrawingRegister(workbook, { filterGLS: true });
}

function parseGlsTidp(workbook) {
  return parseDrawingRegister(workbook, { filterGLS: false });
}

function compareMidpRecords(taj, gls) {
  const tajMap = new Map(taj.records.map((r) => [r.base, r]));
  const glsMap = new Map(gls.records.map((r) => [r.base, r]));
  const keys = [...new Set([...tajMap.keys(), ...glsMap.keys()])].sort();

  return keys.map((key) => {
    const a = tajMap.get(key);
    const b = glsMap.get(key);

    if (!a) {
      const requested = b?.requestedAction || '';
      return {
        key, type: 'added', taj: null, gls: b,
        requestedAction: requested,
        implementationStatus: requested === 'ADD' ? 'NOT IMPLEMENTED'
          : requested === 'REMOVE' ? 'IMPLEMENTED'
          : requested === 'MODIFY' ? 'NOT IMPLEMENTED'
          : 'CHECK',
        changes: ['Drawing exists in TIDP but not baseline MIDP'],
        fieldChanges: [],
        revisionChanged: false, suitabilityChanged: false, titleChanged: false,
      };
    }

    if (!b) {
      return {
        key, type: 'removed', taj: a, gls: null,
        requestedAction: '',
        implementationStatus: 'CHECK',
        changes: ['Drawing exists in baseline MIDP but not updated TIDP'],
        fieldChanges: [],
        revisionChanged: false, suitabilityChanged: false, titleChanged: false,
      };
    }

    const headers = [...new Set([...(a.headers || []), ...(b.headers || [])])];
    const fieldChanges = headers.map((field) => ({
      field, taj: a.values?.[field] ?? '', gls: b.values?.[field] ?? '',
    })).filter((x) => normalizeCompareValue(x.taj) !== normalizeCompareValue(x.gls));

    const revisionChanged = normalizeCompareValue(a.revision) !== normalizeCompareValue(b.revision);
    const suitabilityChanged = normalizeCompareValue(a.status) !== normalizeCompareValue(b.status);
    const titleChanged = normalizeCompareValue(a.title) !== normalizeCompareValue(b.title);

    const requestedAction = b.requestedAction || '';
    let implementationStatus = 'NO REQUESTED COLOUR';
    if (requestedAction === 'REMOVE') {
      implementationStatus = 'NOT IMPLEMENTED';
    } else if (requestedAction === 'ADD') {
      implementationStatus = 'IMPLEMENTED';
    } else if (requestedAction === 'MODIFY') {
      implementationStatus = (revisionChanged || suitabilityChanged || titleChanged || fieldChanges.length)
        ? 'IMPLEMENTED'
        : 'NOT IMPLEMENTED';
    } else if (revisionChanged || suitabilityChanged || titleChanged) {
      implementationStatus = 'NOT MATCHING';
    } else {
      implementationStatus = 'MATCHING';
    }

    return {
      key, type: fieldChanges.length ? 'changed' : 'unchanged',
      taj: a, gls: b, requestedAction, implementationStatus,
      changes: fieldChanges.map((x) => x.field), fieldChanges,
      revisionChanged, suitabilityChanged, titleChanged,
    };
  });
}

function midpSummary(records) {
  const summary = {
    total: records.length,
    UR: 0, 'Code B': 0, 'Code C': 0, 'Code D': 0, NYS: 0,
    ADD: 0, REMOVE: 0, MODIFY: 0, Other: 0
  };

  records.forEach((r) => {
    const status = normalizeCompareValue(r.status);
    if (/^UR$|UNDER REVIEW/.test(status)) summary.UR += 1;
    else if (/CODE\\s*B|^B$/.test(status)) summary['Code B'] += 1;
    else if (/CODE\\s*C|^C$/.test(status)) summary['Code C'] += 1;
    else if (/CODE\\s*D|^D$/.test(status)) summary['Code D'] += 1;
    else if (/NYS|NOT YET SUBMITTED/.test(status)) summary.NYS += 1;
    else summary.Other += 1;

    if (r.requestedAction === 'ADD') summary.ADD += 1;
    else if (r.requestedAction === 'REMOVE') summary.REMOVE += 1;
    else if (r.requestedAction === 'MODIFY') summary.MODIFY += 1;
  });

  return summary;
}

function MidpSummaryCard({ title, summary }) {
  const items = [['Total Drawings', summary.total], ['UR', summary.UR], ['Code B', summary['Code B']], ['Code C', summary['Code C']], ['Code D', summary['Code D']], ['NYS', summary.NYS]];
  return <div className="midp-summary-card"><h3>{title}</h3><div className="midp-count-grid">{items.map(([label, value]) => <div key={label}><small>{label}</small><strong>{value}</strong></div>)}</div></div>;
}

function createChecks(midpRecords, metadataRecords, pdfRecords) {
  // A check session is defined by the drawing PDFs the user selected. The MIDP
  // and metadata may contain thousands of unrelated historical records.
  const keys = [...new Set(pdfRecords.map((r) => r.base).filter(Boolean))];
  return keys.map((key) => {
    const midp = midpRecords.find((r) => r.base === key);
    const metadata = metadataRecords.find((r) => r.base === key && /-PDF$/i.test(r.doc))
      || metadataRecords.find((r) => r.base === key);
    const pdf = pdfRecords.find((r) => r.base === key);
    const rows = ['doc', 'revision', 'title', 'issueDate', 'modelRef'].map((field) => {
      const m = field === 'doc' ? midp?.base : midp?.[field];
      const d = field === 'doc' ? metadata?.base : metadata?.[field];
      const p = field === 'doc' ? pdf?.base : pdf?.[field];
      return { field, label: FIELD_LABELS[field], midp: m || '', metadata: d || '', pdf: p || '', ...compareValue(field, m, d, p) };
    });
    const fail = rows.filter((r) => r.status === 'fail').length;
    const warning = rows.filter((r) => ['warning', 'missing'].includes(r.status)).length;
    return { key, midp, metadata, pdf, rows, fail, warning, status: fail ? 'fail' : warning ? 'warning' : 'pass' };
  }).sort((a, b) => (a.status === 'fail' ? -1 : 1) - (b.status === 'fail' ? -1 : 1));
}

function DropZone({ title, subtitle, icon: Icon, accept, multiple, files, onFiles, color }) {
  const id = `upload-${title.replace(/\s/g, '-')}`;
  return (
    <div className={`drop-card ${files.length ? 'has-file' : ''}`} style={{ '--accent': color }}>
      <div className="drop-heading"><span className="file-icon"><Icon size={19} /></span><div><h3>{title}</h3><p>{subtitle}</p></div></div>
      <label className="drop-target" htmlFor={id}>
        <UploadCloud size={25} />
        <span>{multiple ? 'Drop PDFs here or browse' : 'Drop file here or browse'}</span>
        <small>{accept.includes('pdf') ? 'PDF files' : 'XLSX or XLS files'}</small>
      </label>
      <input id={id} type="file" accept={accept} multiple={multiple} onChange={(e) => onFiles([...e.target.files])} />
      {files.length > 0 && <div className="file-list">{files.slice(0, 3).map((f) => <div key={f.name}><Check size={14} /><span>{f.name}</span></div>)}{files.length > 3 && <small>+{files.length - 3} more files</small>}</div>}
    </div>
  );
}

function StatusBadge({ status, count }) {
  const Icon = status === 'pass' ? CheckCircle2 : status === 'fail' ? XCircle : AlertTriangle;
  return <span className={`status ${status}`}><Icon size={14} />{count != null ? count : status}</span>;
}

function App() {
  const [mode, setMode] = useState('midp');
  const [tajMidpFiles, setTajMidpFiles] = useState([]);
  const [glsMidpFiles, setGlsMidpFiles] = useState([]);
  const [midpComparison, setMidpComparison] = useState(null);
  const [midpLoading, setMidpLoading] = useState(false);
  const [midpError, setMidpError] = useState('');
  const [midpFiles, setMidpFiles] = useState([]);
  const [metadataFiles, setMetadataFiles] = useState([]);
  const [pdfFiles, setPdfFiles] = useState([]);
  const [checks, setChecks] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [open, setOpen] = useState({});
  const ready = midpFiles.length && metadataFiles.length && pdfFiles.length;
  const totals = useMemo(() => ({
    drawings: checks.length,
    pass: checks.reduce((n, c) => n + c.rows.filter((r) => r.status === 'pass').length, 0),
    warning: checks.reduce((n, c) => n + c.rows.filter((r) => ['warning', 'missing'].includes(r.status)).length, 0),
    fail: checks.reduce((n, c) => n + c.rows.filter((r) => r.status === 'fail').length, 0),
  }), [checks]);

  async function runMidpComparison() {
    setMidpLoading(true); setMidpError(''); setMidpComparison(null);
    try {
      const [tajBook, glsBook] = await Promise.all([readWorkbook(tajMidpFiles[0]), readWorkbook(glsMidpFiles[0])]);
      const taj = parseTajMidp(tajBook);
      const gls = parseGlsTidp(glsBook);
      const rows = compareMidpRecords(taj, gls);

      setMidpComparison({
        taj: taj.records,
        gls: gls.records,
        tajDuplicates: taj.duplicates,
        glsDuplicates: gls.duplicates,
        rows,
        tajSummary: midpSummary(taj.records),
        glsSummary: midpSummary(gls.records),
        counts: {
          added: rows.filter((r) => r.type === 'added').length,
          removed: rows.filter((r) => r.type === 'removed').length,
          changed: rows.filter((r) => r.type === 'changed').length,
          unchanged: rows.filter((r) => r.type === 'unchanged').length,
        },
        filterLabel: 'Baseline MIDP: Column Z (Sub-Owner) = GLS; Updated TIDP: no Sub-Owner filter; TIDP colours = red Remove, green Add, yellow Modify'
      });
    } catch (e) {
      setMidpError(e.message || 'The MIDP/TIDP files could not be compared.');
    }
    finally { setMidpLoading(false); }
  }

  function downloadMidpReport() {
    if (!midpComparison) return;

    const makeSheet = (rows) => XLSX.utils.json_to_sheet(rows.length ? rows : [{ 'DRAWING NUMBER': 'No records' }]);

    const fullSummary = midpComparison.rows.map((r) => ({
      'DRAWING NUMBER': r.key,
      'Requested Action (TIDP Colour)': r.requestedAction || '—',
      'Implementation Status': r.implementationStatus,
      'MIDP DRAWING TITLE': r.taj?.title || '',
      'TIDP DRAWING TITLE': r.gls?.title || '',
      'MIDP Revision': r.taj?.revision || '',
      'TIDP Revision': r.gls?.revision || '',
      'MIDP Status': r.taj?.status || '',
      'TIDP Status': r.gls?.status || '',
      'Revision Changed': r.revisionChanged ? 'YES' : 'NO',
      'Status Changed': r.suitabilityChanged ? 'YES' : 'NO',
      'Title Changed': r.titleChanged ? 'YES' : 'NO',
      'Comparison Status': r.type === 'added' ? 'ADDED'
        : r.type === 'removed' ? 'REMOVED'
        : r.revisionChanged || r.suitabilityChanged ? 'REVISION / STATUS CHANGED'
        : r.titleChanged ? 'TITLE CHANGED' : 'UNCHANGED',
    }));

    const added = midpComparison.rows.filter((r) => r.type === 'added' || r.requestedAction === 'ADD').map((r) => ({
      'DRAWING NUMBER': r.key,
      'Requested Action': r.requestedAction || 'ADD',
      'Implementation Status': r.implementationStatus,
      'DRAWING TITLE': r.gls?.title || r.taj?.title || '',
      'Revision': r.gls?.revision || r.taj?.revision || '',
      'Status': r.gls?.status || r.taj?.status || '',
    }));

    const revisionMismatches = midpComparison.rows
      .filter((r) => r.revisionChanged || r.suitabilityChanged || r.titleChanged)
      .map((r) => ({
        'DRAWING NUMBER': r.key,
        'Requested Action': r.requestedAction || '',
        'Implementation Status': r.implementationStatus,
        'MIDP Revision': r.taj?.revision || '',
        'TIDP Revision': r.gls?.revision || '',
        'MIDP Status': r.taj?.status || '',
        'TIDP Status': r.gls?.status || '',
        'Revision Changed': r.revisionChanged ? 'YES' : 'NO',
        'Status Changed': r.suitabilityChanged ? 'YES' : 'NO',
        'Title Changed': r.titleChanged ? 'YES' : 'NO',
        'MIDP Title': r.taj?.title || '',
        'TIDP Title': r.gls?.title || '',
      }));

    const missing = midpComparison.rows.filter((r) => r.type === 'removed' || r.requestedAction === 'REMOVE').map((r) => ({
      'DRAWING NUMBER': r.key,
      'Requested Action': r.requestedAction || 'REMOVE',
      'Implementation Status': r.implementationStatus,
      'DRAWING TITLE': r.taj?.title || r.gls?.title || '',
      'Revision': r.taj?.revision || r.gls?.revision || '',
      'Status': r.taj?.status || r.gls?.status || '',
    }));

    const duplicates = [
      ...midpComparison.tajDuplicates.map((d) => ({ Source: 'Baseline MIDP', 'DRAWING NUMBER': d })),
      ...midpComparison.glsDuplicates.map((d) => ({ Source: 'Updated TIDP', 'DRAWING NUMBER': d })),
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, makeSheet(fullSummary), 'Full Summary');
    XLSX.utils.book_append_sheet(wb, makeSheet(added), 'Added Drawings');
    XLSX.utils.book_append_sheet(wb, makeSheet(revisionMismatches), 'Revision Mismatches');
    XLSX.utils.book_append_sheet(wb, makeSheet(missing), 'Missing Drawings');
    if (duplicates.length) XLSX.utils.book_append_sheet(wb, makeSheet(duplicates), 'Duplicate Drawings');

    XLSX.writeFile(wb, `MIDP_vs_TIDP_Comparison_Report_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  async function runCheck() {
    setLoading(true); setError('');
    try {
      const [midpBook, metadataBook] = await Promise.all([readWorkbook(midpFiles[0]), readWorkbook(metadataFiles[0])]);
      const [midp, metadata, pdfs] = await Promise.all([
        Promise.resolve(parseMidp(midpBook)), Promise.resolve(parseMetadata(metadataBook)),
        Promise.all(pdfFiles.map(parsePdf)),
      ]);
      const output = createChecks(midp, metadata, pdfs);
      if (!output.length) throw new Error('No drawing numbers could be matched. Check that these files use the expected MIDP and metadata formats.');
      setChecks(output); setOpen(Object.fromEntries(output.map((c) => [c.key, true])));
    } catch (e) { setError(e.message || 'The files could not be processed.'); }
    finally { setLoading(false); }
  }

  function reset() { setMidpFiles([]); setMetadataFiles([]); setPdfFiles([]); setChecks([]); setError(''); }

  function downloadReport() {
    const data = checks.flatMap((c) => c.rows.map((r) => ({
      'Drawing Number': c.key, Field: r.label, MIDP: r.midp, Metadata: r.metadata,
      'PDF Title Block': r.pdf, Result: r.status.toUpperCase(), Note: r.note,
    })));
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{ wch: 50 }, { wch: 22 }, { wch: 55 }, { wch: 55 }, { wch: 55 }, { wch: 12 }, { wch: 30 }];
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Validation Results');
    XLSX.writeFile(wb, `DrawingCheck_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  if (mode === 'midp') {
    return (
      <div className="app">
        <header><div className="brand"><span className="brand-mark"><ShieldCheck /></span><div><strong>DrawingCheck</strong><small>BIM document validation</small></div></div><span className="privacy"><Lock size={13} /> Files stay in your browser</span></header>
        <main>
          {!midpComparison ? <section className="workspace">
            <div className="section-title"><div><span>01</span><div><h2>TAJ MIDP vs GLS TIDP</h2><p>Compare drawing deliverables from MIDP-DRW-SWD. Baseline MIDP = Sub-Owner GLS; updated TIDP = no Sub-Owner filter. TIDP colour marks are used as the weekly requested-change instruction.</p></div></div><span className="secure"><Lock size={14}/> Local processing</span></div>
            <div className="drop-grid midp-upload-grid">
              <DropZone title="TAJ MIDP" subtitle="TAJ Master Information Delivery Plan" icon={FileSpreadsheet} accept=".xlsx,.xls" files={tajMidpFiles} onFiles={(f) => setTajMidpFiles(f.slice(0,1))} color="#c77645" />
              <DropZone title="GLS TIDP" subtitle="Glassline Task Information Delivery Plan" icon={FileSpreadsheet} accept=".xlsx,.xls" files={glsMidpFiles} onFiles={(f) => setGlsMidpFiles(f.slice(0,1))} color="#507e79" />
            </div>
            {midpError && <div className="error"><XCircle size={18}/>{midpError}</div>}
            <button className="primary" disabled={!tajMidpFiles.length || !glsMidpFiles.length || midpLoading} onClick={runMidpComparison}>{midpLoading ? <><LoaderCircle className="spin"/>Comparing MIDPs…</> : <>Compare TAJ vs GLS <ArrowRight size={18}/></>}</button>
            <p className="helper"><Info size={13}/> The comparison is processed locally in your browser.</p>
          </section> : <section className="results midp-results">
            <div className="results-top"><div><span className="eyebrow">MIDP COMPARISON COMPLETE</span><h2>TAJ MIDP vs GLS TIDP</h2><p>Weekly implementation audit: TIDP colour marks requested ADD / REMOVE / MODIFY actions; the current TAJ MIDP is checked to confirm whether those requests were implemented.</p></div><div className="actions"><button className="secondary" onClick={() => setMidpComparison(null)}><RotateCcw size={16}/>Change files</button><button className="primary compact" onClick={downloadMidpReport}><Download size={17}/>Export comparison</button></div></div>
            <div className="gls-filter-note"><strong>Result:</strong> Drawing numbers are matched between GLS TIDP and TAJ MIDP. GLS colour = expected action; TAJ = actual implementation.</div>
            <div className="comparison-head midp-table-head"><span>Drawing Number</span><span>TAJ Drawing Title</span><span>TIDP Drawing Title</span><span>TAJ Status</span><span>TIDP Status</span><span>Requested</span><span>Implementation</span></div>
            <div className="checks">{midpComparison.rows.map((r) => <article className="check" key={r.key}>
              <div className="midp-result-row"><strong>{r.key}</strong><span title={r.taj?.title || ''}>{r.taj?.title || '—'}</span><span title={r.gls?.title || ''}>{r.gls?.title || '—'}</span><span>{r.taj?.status || '—'}</span><span>{r.gls?.status || '—'}</span><span>{r.requestedAction || '—'}</span><span title={r.implementationStatus}>{r.implementationStatus}</span></div>
              {r.type === 'changed' && <div className="midp-field-changes">
                {r.fieldChanges.map((c) => <div className="midp-field-change" key={c.field}><strong>{c.field}</strong><span title={c.taj || ''}>{c.taj || '—'}</span><span title={c.gls || ''}>{c.gls || '—'}</span></div>)}
              </div>}
            </article>)}</div>
            <div className="next-check"><div><span className="eyebrow">NEXT STEP</span><h3>MIDP vs Metadata vs PDF</h3><p>Continue with the existing drawing validation workflow using the GLS TIDP.</p></div><button className="primary" onClick={() => { setMidpFiles(glsMidpFiles); setMode('validator'); }}><ArrowRight size={18}/>Continue to drawing checker</button></div>
          </section>}
        </main>
        <footer><span>DrawingCheck <b>0.2</b></span><span>Designed for controlled BIM / Document Control review</span></footer>
      </div>
    );
  }

  return (
    <div className="app">
      <header><div className="brand"><span className="brand-mark"><ShieldCheck /></span><div><strong>DrawingCheck</strong><small>BIM document validation</small></div></div><span className="privacy"><Lock size={13} /> Files stay in your browser</span></header>
      <main>
        {!checks.length ? <section className="workspace">
          <div className="section-title"><div><span>02</span><div><h2>MIDP vs Metadata vs PDF</h2><p>Validate the GLS MIDP against metadata and issued drawing PDFs.</p></div></div><span className="secure"><Lock size={14}/> Local processing</span></div>
          <div className="drop-grid">
            <DropZone title="MIDP workbook" subtitle="Master Information Delivery Plan" icon={FileSpreadsheet} accept=".xlsx,.xls" files={midpFiles} onFiles={(f) => setMidpFiles(f.slice(0,1))} color="#c77645" />
            <DropZone title="Metadata export" subtitle="Document Control data" icon={FileSpreadsheet} accept=".xlsx,.xls" files={metadataFiles} onFiles={(f) => setMetadataFiles(f.slice(0,1))} color="#507e79" />
            <DropZone title="Drawing PDFs" subtitle="Issued title blocks" icon={FileText} accept=".pdf" multiple files={pdfFiles} onFiles={setPdfFiles} color="#8b7961" />
          </div>
          {error && <div className="error"><XCircle size={18}/>{error}</div>}
          <button className="secondary back-to-midp" onClick={() => setMode('midp')}><RotateCcw size={16}/>TAJ vs GLS TIDP</button><button className="primary" disabled={!ready || loading} onClick={runCheck}>{loading ? <><LoaderCircle className="spin"/>Reading and comparing files…</> : <>Run validation <ArrowRight size={18}/></>}</button>
          <p className="helper"><Info size={13}/> Digitally generated PDFs are supported. Scanned title blocks will require the OCR add-on.</p>
        </section> : <section className="results">
          <div className="results-top"><div><span className="eyebrow">VALIDATION COMPLETE</span><h2>Submission review</h2><p>{totals.drawings} drawing{totals.drawings !== 1 ? 's' : ''} checked across three sources.</p></div><div className="actions"><button className="secondary" onClick={() => setMode('midp')}><RotateCcw size={16}/>TAJ vs GLS</button><button className="secondary" onClick={reset}><RotateCcw size={16}/>New check</button><button className="primary compact" onClick={downloadReport}><Download size={17}/>Export report</button></div></div>
          <div className="summary-grid"><div><small>Drawings</small><strong>{totals.drawings}</strong></div><div className="green"><small>Matching fields</small><strong>{totals.pass}</strong></div><div className="amber"><small>Warnings</small><strong>{totals.warning}</strong></div><div className="red"><small>Mismatches</small><strong>{totals.fail}</strong></div></div>
          <div className="comparison-head"><span>Drawing / field</span><span>MIDP</span><span>Metadata</span><span>PDF title block</span><span>Result</span></div>
          <div className="checks">{checks.map((c) => <article className="check" key={c.key}>
            <button className="check-title" onClick={() => setOpen((o) => ({...o, [c.key]: !o[c.key]}))}><div><ChevronDown className={open[c.key] ? 'rotated' : ''} size={18}/><span>{c.key}</span></div><StatusBadge status={c.status} count={c.fail ? `${c.fail} mismatch${c.fail > 1 ? 'es' : ''}` : c.warning ? `${c.warning} warning${c.warning > 1 ? 's' : ''}` : 'All matched'} /></button>
            {open[c.key] && <div className="check-rows">{c.rows.map((r) => <div className="check-row" key={r.field}><strong>{r.label}</strong><span title={r.midp}>{r.field === 'issueDate' ? formatDate(r.midp) : r.midp || '—'}</span><span title={r.metadata}>{r.field === 'issueDate' ? formatDate(r.metadata) : r.metadata || '—'}</span><span title={r.pdf}>{r.field === 'issueDate' ? formatDate(r.pdf) : r.pdf || '—'}</span><StatusBadge status={r.status}/></div>)}</div>}
          </article>)}</div>
        </section>}
      </main>
      <footer><span>DrawingCheck <b>0.2</b></span><span>Designed for controlled BIM / Document Control review</span></footer>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
