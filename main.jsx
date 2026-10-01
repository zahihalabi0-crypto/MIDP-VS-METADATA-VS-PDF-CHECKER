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
  return XLSX.read(buffer, { type: 'array', cellDates: true });
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

  return (
    <div className="app">
      <header><div className="brand"><span className="brand-mark"><ShieldCheck /></span><div><strong>DrawingCheck</strong><small>BIM document validation</small></div></div><span className="privacy"><Lock size={13} /> Files stay in your browser</span></header>
      <main>
        {!checks.length ? <section className="workspace">
          <div className="section-title"><div><span>01</span><div><h2>Add submission files</h2><p>One MIDP, one metadata export, and the issued drawing PDFs.</p></div></div><span className="secure"><Lock size={14}/> Local processing</span></div>
          <div className="drop-grid">
            <DropZone title="MIDP workbook" subtitle="Master Information Delivery Plan" icon={FileSpreadsheet} accept=".xlsx,.xls" files={midpFiles} onFiles={(f) => setMidpFiles(f.slice(0,1))} color="#c77645" />
            <DropZone title="Metadata export" subtitle="Document Control data" icon={FileSpreadsheet} accept=".xlsx,.xls" files={metadataFiles} onFiles={(f) => setMetadataFiles(f.slice(0,1))} color="#507e79" />
            <DropZone title="Drawing PDFs" subtitle="Issued title blocks" icon={FileText} accept=".pdf" multiple files={pdfFiles} onFiles={setPdfFiles} color="#8b7961" />
          </div>
          {error && <div className="error"><XCircle size={18}/>{error}</div>}
          <button className="primary" disabled={!ready || loading} onClick={runCheck}>{loading ? <><LoaderCircle className="spin"/>Reading and comparing files…</> : <>Run validation <ArrowRight size={18}/></>}</button>
          <p className="helper"><Info size={13}/> Digitally generated PDFs are supported. Scanned title blocks will require the OCR add-on.</p>
        </section> : <section className="results">
          <div className="results-top"><div><span className="eyebrow">VALIDATION COMPLETE</span><h2>Submission review</h2><p>{totals.drawings} drawing{totals.drawings !== 1 ? 's' : ''} checked across three sources.</p></div><div className="actions"><button className="secondary" onClick={reset}><RotateCcw size={16}/>New check</button><button className="primary compact" onClick={downloadReport}><Download size={17}/>Export report</button></div></div>
          <div className="summary-grid"><div><small>Drawings</small><strong>{totals.drawings}</strong></div><div className="green"><small>Matching fields</small><strong>{totals.pass}</strong></div><div className="amber"><small>Warnings</small><strong>{totals.warning}</strong></div><div className="red"><small>Mismatches</small><strong>{totals.fail}</strong></div></div>
          <div className="comparison-head"><span>Drawing / field</span><span>MIDP</span><span>Metadata</span><span>PDF title block</span><span>Result</span></div>
          <div className="checks">{checks.map((c) => <article className="check" key={c.key}>
            <button className="check-title" onClick={() => setOpen((o) => ({...o, [c.key]: !o[c.key]}))}><div><ChevronDown className={open[c.key] ? 'rotated' : ''} size={18}/><span>{c.key}</span></div><StatusBadge status={c.status} count={c.fail ? `${c.fail} mismatch${c.fail > 1 ? 'es' : ''}` : c.warning ? `${c.warning} warning${c.warning > 1 ? 's' : ''}` : 'All matched'} /></button>
            {open[c.key] && <div className="check-rows">{c.rows.map((r) => <div className="check-row" key={r.field}><strong>{r.label}</strong><span title={r.midp}>{r.field === 'issueDate' ? formatDate(r.midp) : r.midp || '—'}</span><span title={r.metadata}>{r.field === 'issueDate' ? formatDate(r.metadata) : r.metadata || '—'}</span><span title={r.pdf}>{r.field === 'issueDate' ? formatDate(r.pdf) : r.pdf || '—'}</span><StatusBadge status={r.status}/></div>)}</div>}
          </article>)}</div>
        </section>}
      </main>
      <footer><span>DrawingCheck <b>0.1</b></span><span>Designed for controlled BIM / Document Control review</span></footer>
    </div>
  );
}

createRoot(document.getElementById('root')).render(<App />);
