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

function normalizeMetadataDescription(value) {
  return clean(value).toUpperCase().replace(/[^A-Z0-9]+/g, ' ').trim();
}

function parseMetadataComparison(workbook, { filterGLS = false } = {}) {
  const sheetName = 'MIDP-DRW-SWD';
  if (!workbook.SheetNames.includes(sheetName)) {
    throw new Error(`The required sheet "${sheetName}" was not found.`);
  }

  const rows = sheetRows(workbook, sheetName);
  const headerRow = 3; // Excel row 4
  const headers = rows[headerRow] || [];
  const docIndex = 5; // F = Drawing Number
  const titleIndex = 6; // G = Drawing Description / Title
  const statusIndex = 10; // K = Status
  const ownerIndex = 25; // Z = Sub-Owner

  const records = [];
  rows.slice(headerRow + 1).forEach((row, offset) => {
    const doc = clean(row[docIndex]);
    if (!doc || !DOC_PATTERN.test(doc)) return;

    if (filterGLS && clean(row[ownerIndex]).toUpperCase() !== 'GLS') return;

    records.push({
      sourceRow: headerRow + offset + 2,
      sheet: sheetName,
      doc,
      base: baseNumber(doc),
      description: clean(row[titleIndex]),
      status: clean(row[statusIndex]),
      subOwner: clean(row[ownerIndex]),
    });
  });

  const map = new Map();
  const duplicates = [];
  records.forEach((record) => {
    if (map.has(record.base)) {
      duplicates.push(record.base);
      return;
    }
    map.set(record.base, record);
  });

  if (!map.size) {
    throw new Error(`No drawing records were found in "${sheetName}".`);
  }

  return { records: [...map.values()], duplicates };
}

function compareMetadataVersions(gls, taj) {
  const glsMap = new Map(gls.records.map((r) => [r.base, r]));
  const tajMap = new Map(taj.records.map((r) => [r.base, r]));

  const rows = [...glsMap.keys()].sort().map((key) => {
    const glsRecord = glsMap.get(key);
    const tajRecord = tajMap.get(key);
    const existsInTaj = Boolean(tajRecord);
    const descriptionMatch = existsInTaj &&
      normalizeMetadataDescription(glsRecord.description) === normalizeMetadataDescription(tajRecord.description);
    const statusMatch = existsInTaj &&
      normalizeCompareValue(glsRecord.status) === normalizeCompareValue(tajRecord.status);

    return {
      key,
      gls: glsRecord,
      taj: tajRecord || null,
      existsInTaj,
      descriptionMatch,
      statusMatch,
      overallMatch: existsInTaj && descriptionMatch && statusMatch,
      result: !existsInTaj ? 'NOT IN TAJ' : (descriptionMatch && statusMatch ? 'MATCHING' : 'NOT MATCHING'),
    };
  });

  const statusTotals = {};
  gls.records.forEach((record) => {
    const status = clean(record.status) || 'Blank';
    if (!statusTotals[status]) statusTotals[status] = { gls: 0, taj: 0 };
    statusTotals[status].gls += 1;
  });
  taj.records.forEach((record) => {
    const status = clean(record.status) || 'Blank';
    if (!statusTotals[status]) statusTotals[status] = { gls: 0, taj: 0 };
    statusTotals[status].taj += 1;
  });

  return {
    rows,
    summary: {
      glsTotal: gls.records.length,
      tajTotal: taj.records.length,
      inTaj: rows.filter((r) => r.existsInTaj).length,
      notInTaj: rows.filter((r) => !r.existsInTaj).length,
      descriptionMatching: rows.filter((r) => r.existsInTaj && r.descriptionMatch).length,
      descriptionNotMatching: rows.filter((r) => r.existsInTaj && !r.descriptionMatch).length,
      statusMatching: rows.filter((r) => r.existsInTaj && r.statusMatch).length,
      statusNotMatching: rows.filter((r) => r.existsInTaj && !r.statusMatch).length,
      overallMatching: rows.filter((r) => r.overallMatch).length,
      overallNotMatching: rows.filter((r) => !r.overallMatch).length,
    },
    statusTotals: Object.entries(statusTotals).sort(([a],[b]) => a.localeCompare(b)),
  };
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
  const statusIndexes = [10, 11, 12]; // K, L, M

  const rawHeaders = rows[headerRow].map((h, i) => clean(h) || `Column ${i + 1}`);
  const normalizedHeaders = rawHeaders.map((h) => clean(h).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim());

  // Status / suitability codes are checked across Columns K, L and M.
  // K is the primary status column, while L/M may contain additional codes.
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
    const statusCodes = statusIndexes
      .map((index) => clean(row[index]))
      .filter(Boolean);
    const combinedStatus = statusCodes.join(' | ');

    records.push({
      sourceRow: headerRow + offset + 2,
      sheet: targetSheet,
      doc,
      base: key,
      title: clean(row[titleIndex]),
      revision: revision(row[revisionIndex]),
      status: combinedStatus,
      statusCodes,
      suitabilityStatus: clean(row[suitabilityIndex]),
      statusColumn: statusIndex + 1,
      category: midpCategory(row, statusIndexes[0], statusIndexes[1]),
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
    const tajRecord = tajMap.get(key);
    const glsRecord = glsMap.get(key);
    const requestedAction = glsRecord?.requestedAction || 'NO CHANGE';

    // ADD: requested drawing must exist in the TAJ MIDP.
    if (!tajRecord && glsRecord) {
      return {
        key,
        type: 'added',
        taj: null,
        gls: glsRecord,
        requestedAction,
        implementationStatus: requestedAction === 'ADD' ? 'NO' : requestedAction === 'REMOVE' ? 'YES' : 'NO',
        changes: [],
        fieldChanges: [],
        revisionChanged: false,
        suitabilityChanged: false,
        titleChanged: false,
        statusChanged: false,
      };
    }

    // REMOVE: requested drawing must no longer exist in the TAJ MIDP.
    if (tajRecord && !glsRecord) {
      return {
        key,
        type: 'removed',
        taj: tajRecord,
        gls: null,
        requestedAction: 'NO CHANGE',
        implementationStatus: 'N/A',
        changes: [],
        fieldChanges: [],
        revisionChanged: false,
        suitabilityChanged: false,
        titleChanged: false,
        statusChanged: false,
      };
    }

    if (!tajRecord || !glsRecord) return null;

    const tajTitle = normalizedTitle(tajRecord.title);
    const glsTitle = normalizedTitle(glsRecord.title);
    const tajStatus = normalizeCompareValue(statusFromColumnK(tajRecord));
    const glsStatus = normalizeCompareValue(statusFromColumnK(glsRecord));

    const titleChanged = tajTitle !== glsTitle;
    const statusChanged = tajStatus !== glsStatus;
    const revisionChanged = normalizeCompareValue(tajRecord.revision) !== normalizeCompareValue(glsRecord.revision);

    // MODIFY is implemented only when the TAJ MIDP matches the updated
    // GLS TIDP title AND status/suitability codes.
    let implementationStatus = 'N/A';
    if (requestedAction === 'ADD') {
      implementationStatus = 'YES';
    } else if (requestedAction === 'REMOVE') {
      implementationStatus = 'NO';
    } else if (requestedAction === 'MODIFY') {
      implementationStatus = !titleChanged && !statusChanged ? 'YES' : 'NO';
    }

    const type = titleChanged || statusChanged || revisionChanged ? 'changed' : 'unchanged';

    return {
      key,
      type,
      taj: tajRecord,
      gls: glsRecord,
      requestedAction,
      implementationStatus,
      changes: [],
      fieldChanges: [
        ...(titleChanged ? [{ field: 'DRAWING TITLE', taj: tajRecord.title, gls: glsRecord.title }] : []),
        ...(statusChanged ? [{ field: 'STATUS / SUITABILITY CODES', taj: tajRecord.status, gls: glsRecord.status }] : []),
        ...(revisionChanged ? [{ field: 'REVISION', taj: tajRecord.revision, gls: glsRecord.revision }] : []),
      ],
      revisionChanged,
      suitabilityChanged: statusChanged,
      titleChanged,
      statusChanged,
    };
  }).filter(Boolean);
}

function statusFromColumnK(record) {
  // Column K is the authoritative drawing status.
  return clean(record?.statusCodes?.[0] || record?.status || '');
}

function classifyStatus(value) {
  const status = normalizeCompareValue(value);
  if (/\\bUR\\b|UNDER REVIEW/.test(status)) return 'UR';
  if (/\\bCODE\\s*B\\b|^B$/.test(status)) return 'Code B';
  if (/\\bCODE\\s*C\\b|^C$/.test(status)) return 'Code C';
  if (/\\bCODE\\s*D\\b|^D$/.test(status)) return 'Code D';
  if (/\\bNYS\\b|NOT YET SUBMITTED/.test(status)) return 'NYS';
  return 'Other';
}

function midpSummary(records) {
  const summary = {
    total: records.length,
    UR: 0,
    'Code B': 0,
    'Code C': 0,
    'Code D': 0,
    NYS: 0,
    Other: 0,
  };

  records.forEach((r) => {
    const category = classifyStatus(statusFromColumnK(r));
    summary[category] = (summary[category] || 0) + 1;
  });

  return summary;
}

function implementationSummary(rows) {
  return {
    totalRequests: rows.filter((r) => ['ADD', 'REMOVE', 'MODIFY'].includes(r.requestedAction)).length,
    implemented: rows.filter((r) => r.implementationStatus === 'YES').length,
    notImplemented: rows.filter((r) => r.implementationStatus === 'NO').length,
  };
}

function buildStatusComparison(tajRecords, glsRecords) {
  const taj = midpSummary(tajRecords);
  const gls = midpSummary(glsRecords);
  return {
    labels: ['Total Drawings', 'Code B', 'Code C', 'UR'],
    rows: [
      ['Total Drawings', gls.total, taj.total],
      ['Code B', gls['Code B'], taj['Code B']],
      ['Code C', gls['Code C'], taj['Code C']],
      ['UR', gls.UR, taj.UR],
    ]
  };
}

function MidpSummaryCard({ title, summary }) {
  const items = [['Total Drawings', summary.total], ['UR', summary.UR], ['Code B', summary['Code B']], ['Code C', summary['Code C']], ['Code D', summary['Code D']], ['NYS', summary.NYS]];
  return <div className="midp-summary-card"><h3>{title}</h3><div className="midp-count-grid">{items.map(([label, value]) => <div key={label}><small>{label}</small><strong>{value}</strong></div>)}</div></div>;
}

function createChecks(midpRecords, TIDP/MIDPRecords, pdfRecords) {
  // A check session is defined by the drawing PDFs the user selected. The MIDP
  // and TIDP/MIDP may contain thousands of unrelated historical records.
  const keys = [...new Set(pdfRecords.map((r) => r.base).filter(Boolean))];
  return keys.map((key) => {
    const midp = midpRecords.find((r) => r.base === key);
    const TIDP/MIDP = TIDP/MIDPRecords.find((r) => r.base === key && /-PDF$/i.test(r.doc))
      || TIDP/MIDPRecords.find((r) => r.base === key);
    const pdf = pdfRecords.find((r) => r.base === key);
    const rows = ['doc', 'revision', 'title', 'issueDate', 'modelRef'].map((field) => {
      const m = field === 'doc' ? midp?.base : midp?.[field];
      const d = field === 'doc' ? TIDP/MIDP?.base : TIDP/MIDP?.[field];
      const p = field === 'doc' ? pdf?.base : pdf?.[field];
      return { field, label: FIELD_LABELS[field], midp: m || '', TIDP/MIDP: d || '', pdf: p || '', ...compareValue(field, m, d, p) };
    });
    const fail = rows.filter((r) => r.status === 'fail').length;
    const warning = rows.filter((r) => ['warning', 'missing'].includes(r.status)).length;
    return { key, midp, TIDP/MIDP, pdf, rows, fail, warning, status: fail ? 'fail' : warning ? 'warning' : 'pass' };
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
  const [TIDP/MIDPFiles, setMetadataFiles] = useState([]);
  const [pdfFiles, setPdfFiles] = useState([]);
  const [checks, setChecks] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [open, setOpen] = useState({});
  const ready = midpFiles.length && TIDP/MIDPFiles.length && pdfFiles.length;
  const totals = useMemo(() => ({
    drawings: checks.length,
    pass: checks.reduce((n, c) => n + c.rows.filter((r) => r.status === 'pass').length, 0),
    warning: checks.reduce((n, c) => n + c.rows.filter((r) => ['warning', 'missing'].includes(r.status)).length, 0),
    fail: checks.reduce((n, c) => n + c.rows.filter((r) => r.status === 'fail').length, 0),
  }), [checks]);

  async function runMidpComparison() {
    setMidpLoading(true); setMidpError(''); setMidpComparison(null);
    try {
      const [tidpBook, midpBook] = await Promise.all([
        readWorkbook(glsMidpFiles[0]),
        readWorkbook(tajMidpFiles[0]),
      ]);
      // TIDP = GLS file (no Sub-Owner filter). MIDP = TAJ file (only GLS rows from Sub-Owner column Z).
      const tidp = parseMetadataComparison(tidpBook, { filterGLS: false });
      const midp = parseMetadataComparison(midpBook, { filterGLS: true });
      const comparison = compareMetadataVersions(tidp, midp);
      setMidpComparison({ ...comparison, glsDuplicates: tidp.duplicates, tajDuplicates: midp.duplicates });
    } catch (e) {
      setMidpError(e.message || 'The GLS TIDP and TAJ MIDP files could not be compared.');
    } finally {
      setMidpLoading(false);
    }
  }


  function downloadMidpReport() {
    if (!midpComparison) return;
    const makeSheet = (rows) => XLSX.utils.json_to_sheet(rows.length ? rows : [{ 'DRAWING NUMBER': 'No records' }]);

    const details = midpComparison.rows.map((r) => ({
      'DRAWING NUMBER': r.key,
      'GLS DESCRIPTION': r.gls?.description || '',
      'TAJ DESCRIPTION': r.taj?.description || '',
      'IN TAJ': r.existsInTaj ? 'YES' : 'NO',
      'DESCRIPTION MATCH': r.existsInTaj ? (r.descriptionMatch ? 'YES' : 'NO') : '—',
      'GLS STATUS': r.gls?.status || '',
      'TAJ STATUS': r.taj?.status || '',
      'STATUS MATCH': r.existsInTaj ? (r.statusMatch ? 'YES' : 'NO') : '—',
      'OVERALL': r.result,
    }));

    const addedToGls = midpComparison.rows.map((r) => ({
      'DRAWING NUMBER': r.key,
      'GLS DESCRIPTION': r.gls?.description || '',
      'TAJ DESCRIPTION': r.taj?.description || '',
      'ALREADY IN TAJ': r.existsInTaj ? 'YES' : 'NO',
      'DESCRIPTION MATCH': r.existsInTaj ? (r.descriptionMatch ? 'YES' : 'NO') : '—',
      'GLS STATUS': r.gls?.status || '',
      'TAJ STATUS': r.taj?.status || '',
      'STATUS MATCH': r.existsInTaj ? (r.statusMatch ? 'YES' : 'NO') : '—',
      'RESULT': r.result,
    }));

    const statusTotals = midpComparison.statusTotals.map(([status, values]) => ({
      STATUS: status,
      'GLS TOTAL': values.gls,
      'TAJ TOTAL': values.taj,
    }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, makeSheet(details), 'GLS vs TAJ');
    XLSX.utils.book_append_sheet(wb, makeSheet(addedToGls), 'GLS Drawings');
    XLSX.utils.book_append_sheet(wb, makeSheet(statusTotals), 'Status Totals');
    XLSX.writeFile(wb, `GLS_Metadata_vs_TAJ_Metadata_${new Date().toISOString().slice(0,10)}.xlsx`);
  }

  async function runCheck() {
    setLoading(true); setError('');
    try {
      const [midpBook, TIDP/MIDPBook] = await Promise.all([readWorkbook(midpFiles[0]), readWorkbook(TIDP/MIDPFiles[0])]);
      const [midp, TIDP/MIDP, pdfs] = await Promise.all([
        Promise.resolve(parseMidp(midpBook)), Promise.resolve(parseMetadata(TIDP/MIDPBook)),
        Promise.all(pdfFiles.map(parsePdf)),
      ]);
      const output = createChecks(midp, TIDP/MIDP, pdfs);
      if (!output.length) throw new Error('No drawing numbers could be matched. Check that these files use the expected MIDP and TIDP/MIDP formats.');
      setChecks(output); setOpen(Object.fromEntries(output.map((c) => [c.key, true])));
    } catch (e) { setError(e.message || 'The files could not be processed.'); }
    finally { setLoading(false); }
  }

  function reset() { setMidpFiles([]); setMetadataFiles([]); setPdfFiles([]); setChecks([]); setError(''); }

  function downloadReport() {
    const data = checks.flatMap((c) => c.rows.map((r) => ({
      'Drawing Number': c.key, Field: r.label, MIDP: r.midp, Metadata: r.TIDP/MIDP,
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
        <header><div className="brand"><span className="brand-mark"><ShieldCheck /></span><div><strong>DrawingCheck</strong><small>TIDP vs MIDP comparison</small></div></div><span className="privacy"><Lock size={13} /> Files stay in your browser</span></header>
        <main>
          {!midpComparison ? <section className="workspace">
            <div className="section-title"><div><span>01</span><div><h2>GLS TIDP vs TAJ MIDP</h2><p>Compare GLS drawing descriptions and status against the TAJ TIDP/MIDP.</p></div></div><span className="secure"><Lock size={14}/> Local processing</span></div>
            <div className="drop-grid midp-upload-grid">
              <DropZone title="GLS TIDP" subtitle="Glassline TIDP/MIDP export" icon={FileSpreadsheet} accept=".xlsx,.xls" files={glsMidpFiles} onFiles={(f) => setGlsMidpFiles(f.slice(0,1))} color="#507e79" />
              <DropZone title="TAJ MIDP" subtitle="TAJ MIDP" icon={FileSpreadsheet} accept=".xlsx,.xls" files={tajMidpFiles} onFiles={(f) => setTajMidpFiles(f.slice(0,1))} color="#c77645" />
            </div>
            {midpError && <div className="error"><XCircle size={18}/>{midpError}</div>}
            <button className="primary" disabled={!glsMidpFiles.length || !tajMidpFiles.length || midpLoading} onClick={runMidpComparison}>{midpLoading ? <><LoaderCircle className="spin"/>Comparing TIDP/MIDP…</> : <>Compare GLS vs TAJ <ArrowRight size={18}/></>}</button>
            <p className="helper"><Info size={13}/> The comparison is processed locally in your browser.</p>
          </section> : <section className="results midp-results">
            <div className="results-top"><div><span className="eyebrow">METADATA COMPARISON COMPLETE</span><h2>GLS TIDP vs TAJ MIDP</h2><p>Each GLS drawing is checked to see whether it exists in TAJ, whether the drawing description matches, and whether the status matches.</p></div><div className="actions"><button className="secondary" onClick={() => setMidpComparison(null)}><RotateCcw size={16}/>Change files</button><button className="primary compact" onClick={downloadMidpReport}><Download size={17}/>Export comparison</button></div></div>

            <div className="midp-change-grid">
              <div><small>GLS Total</small><strong>{midpComparison.summary.glsTotal}</strong></div>
              <div><small>TAJ Total</small><strong>{midpComparison.summary.tajTotal}</strong></div>
              <div><small>Already in TAJ</small><strong>{midpComparison.summary.inTaj}</strong></div>
              <div><small>Added to GLS / Not in TAJ</small><strong>{midpComparison.summary.notInTaj}</strong></div>
            </div>

            <div className="simple-match-summary">
              <div className="simple-match-card matching">
                <div className="simple-match-title">Matching</div>
                <strong>{midpComparison.summary.overallMatching}</strong>
                <small>Description + Status match</small>
              </div>
              <div className="simple-match-card not-matching">
                <div className="simple-match-title">Not Matching</div>
                <strong>{midpComparison.summary.overallNotMatching}</strong>
                <small>Missing in TAJ or Description / Status mismatch</small>
              </div>
            </div>

            <div className="comparison-head midp-table-head"><span>Drawing Number</span><span>GLS TIDP</span><span>TAJ MIDP</span><span>In TAJ</span><span>Description</span><span>Status</span><span>Overall</span></div>
            <div className="checks">{midpComparison.rows.map((r) => <article className="check" key={r.key}>
              <div className="midp-result-row TIDP/MIDP-compare-row">
                <strong>{r.key}</strong>
                <span title={r.gls?.description || ''}>{r.gls?.description || '—'}</span>
                <span title={r.taj?.description || ''}>{r.taj?.description || '—'}</span>
                <span className={r.existsInTaj ? 'implementation-yes' : 'implementation-no'}>{r.existsInTaj ? 'YES' : 'NO'}</span>
                <span className={r.existsInTaj && r.descriptionMatch ? 'implementation-yes' : 'implementation-no'}>{r.existsInTaj ? (r.descriptionMatch ? 'MATCH' : 'NOT MATCH') : '—'}</span>
                <span className={r.existsInTaj && r.statusMatch ? 'implementation-yes' : 'implementation-no'}>{r.existsInTaj ? (r.statusMatch ? 'MATCH' : 'NOT MATCH') : '—'}</span>
                <span className={r.overallMatch ? 'implementation-yes' : 'implementation-no'}>{r.result}</span>
              </div>
            </article>)}</div>

            <div className="status-audit-table">
              <div className="status-audit-head"><span>Status</span><span>GLS Total</span><span>TAJ Total</span></div>
              {midpComparison.statusTotals.map(([status, values]) => (
                <div className="status-audit-row" key={status}><strong>{status}</strong><span>{values.gls}</span><span>{values.taj}</span></div>
              ))}
            </div>

            <div className="next-check"><div><span className="eyebrow">NEXT TOOL</span><h3>MIDP vs Metadata vs PDF</h3><p>The second checker remains unchanged.</p></div><button className="primary" onClick={() => { setMidpFiles(glsMidpFiles); setMode('validator'); }}><ArrowRight size={18}/>Continue to drawing checker</button></div>
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
          <div className="section-title"><div><span>02</span><div><h2>MIDP vs Metadata vs PDF</h2><p>Validate the GLS MIDP against TIDP/MIDP and issued drawing PDFs.</p></div></div><span className="secure"><Lock size={14}/> Local processing</span></div>
          <div className="drop-grid">
            <DropZone title="MIDP workbook" subtitle="Master Information Delivery Plan" icon={FileSpreadsheet} accept=".xlsx,.xls" files={midpFiles} onFiles={(f) => setMidpFiles(f.slice(0,1))} color="#c77645" />
            <DropZone title="Metadata export" subtitle="Document Control data" icon={FileSpreadsheet} accept=".xlsx,.xls" files={TIDP/MIDPFiles} onFiles={(f) => setMetadataFiles(f.slice(0,1))} color="#507e79" />
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
            {open[c.key] && <div className="check-rows">{c.rows.map((r) => <div className="check-row" key={r.field}><strong>{r.label}</strong><span title={r.midp}>{r.field === 'issueDate' ? formatDate(r.midp) : r.midp || '—'}</span><span title={r.TIDP/MIDP}>{r.field === 'issueDate' ? formatDate(r.TIDP/MIDP) : r.TIDP/MIDP || '—'}</span><span title={r.pdf}>{r.field === 'issueDate' ? formatDate(r.pdf) : r.pdf || '—'}</span><StatusBadge status={r.status}/></div>)}</div>}
          </article>)}</div>
        </section>}
      </main>
      <footer><span>DrawingCheck <b>0.2</b></span><span>Designed for controlled BIM / Document Control review</span></footer>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
