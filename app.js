import * as pdfjsLib from './vendor/pdf.min.mjs';
import { PDFDocument, StandardFonts, rgb } from './vendor/pdf-lib.esm.min.js';

pdfjsLib.GlobalWorkerOptions.workerSrc = './vendor/pdf.worker.min.mjs';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const MAX_ROWS = 500;
const MAX_COLS = 60;
const EDITABLE = ['sheet', 'docx', 'pptx', 'pdf', 'text'];

const KINDS = {
  sheet: ['xlsx', 'xlsm', 'xls', 'ods', 'csv', 'tsv'],
  docx: ['docx', 'docm', 'dotx'],
  pptx: ['pptx', 'pptm', 'potx'],
  pdf: ['pdf'],
  text: ['txt', 'md', 'markdown', 'json', 'xml', 'html', 'htm', 'css', 'js', 'log', 'yaml', 'yml', 'ini'],
  image: ['png', 'jpg', 'jpeg', 'gif', 'webp'],
  legacy: ['doc', 'dot', 'ppt', 'pot', 'pps'],
};

const MIME = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls: 'application/vnd.ms-excel',
  ods: 'application/vnd.oasis.opendocument.spreadsheet',
  csv: 'text/csv;charset=utf-8',
  tsv: 'text/tab-separated-values;charset=utf-8',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  pdf: 'application/pdf',
};

const $ = (sel) => document.querySelector(sel);
const el = {
  open: $('#open-btn'),
  save: $('#save-btn'),
  name: $('#file-name'),
  tools: $('#tools'),
  view: $('#view'),
  status: $('#status'),
  input: $('#file-input'),
};
let current = null;

const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const setStatus = (msg) => { el.status.textContent = msg; };
const kindOf = (ext) => Object.keys(KINDS).find((k) => KINDS[k].includes(ext)) || 'unknown';
const outName = (base, ext) => (ext ? `${base}-edited.${ext}` : `${base}-edited`);
const mimeFor = (ext) => MIME[ext] || 'application/octet-stream';

function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}

function parseXml(text) {
  const dom = new DOMParser().parseFromString(text, 'application/xml');
  if (dom.getElementsByTagName('parsererror').length) {
    throw new Error('Could not read the XML inside this file. It may be damaged.');
  }
  return dom;
}

function showError(err) {
  console.error(err);
  setStatus(`Something went wrong: ${err?.message || err}`);
}

function showNotice(msg) {
  el.view.innerHTML = `<div class="notice"><p>${esc(msg)}</p></div>`;
  setStatus('');
}

// ---------- event wiring ----------

el.open.addEventListener('click', () => el.input.click());
el.input.addEventListener('change', () => {
  const file = el.input.files[0];
  el.input.value = '';
  if (file) openFile(file).catch(showError);
});
el.save.addEventListener('click', () => saveCurrent().catch(showError));

el.view.addEventListener('change', (e) => {
  if (e.target.matches('input[data-addr]')) setCell(e.target.dataset.addr, e.target.value);
});
el.view.addEventListener('input', (e) => {
  if (e.target.matches('textarea.para')) {
    current.items[Number(e.target.dataset.i)].value = e.target.value;
    current.dirty = true;
  } else if (e.target.matches('textarea.plain')) {
    current.dirty = true;
  }
});
el.view.addEventListener('keydown', (e) => {
  // one paragraph per box: Enter would need a real paragraph break, so block it
  if (e.key === 'Enter' && e.target.matches('textarea.para')) e.preventDefault();
});
el.view.addEventListener('click', (e) => {
  const tab = e.target.closest('[data-sheet]');
  if (tab) { current.sheet = tab.dataset.sheet; renderSheet(); }
});
el.tools.addEventListener('click', (e) => {
  const b = e.target.closest('button');
  if (!b) return;
  if (b.dataset.tool) setTool(b.dataset.tool);
  else if (b.id === 'undo-btn') { current.overlays.pop(); drawOverlays(); }
  else if (b.id === 'text-btn') showPdfText().catch(showError);
});
window.addEventListener('beforeunload', (e) => {
  if (current?.dirty) { e.preventDefault(); e.returnValue = ''; }
});
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}

// ---------- opening ----------

async function openFile(file) {
  const dot = file.name.lastIndexOf('.');
  const ext = dot > 0 ? file.name.slice(dot + 1).toLowerCase() : '';
  const base = dot > 0 ? file.name.slice(0, dot) : file.name;
  const bytes = await file.arrayBuffer();

  let kind = kindOf(ext);
  if (kind === 'unknown' && !new Uint8Array(bytes, 0, Math.min(bytes.byteLength, 4096)).includes(0)) {
    kind = 'text'; // unknown extension but plain text: open it as text
  }

  current = { name: file.name, base, ext, bytes, kind, dirty: false, overlays: [], tool: 'text', items: [], parts: {}, wb: null, sheet: null };
  el.name.textContent = file.name;
  el.view.innerHTML = '';
  el.tools.innerHTML = '';
  el.tools.hidden = true;
  el.save.disabled = !EDITABLE.includes(kind);
  setStatus('Opening…');

  switch (kind) {
    case 'sheet': openSheet(); break;
    case 'docx': await openDocx(); break;
    case 'pptx': await openPptx(); break;
    case 'pdf': await openPdf(); break;
    case 'text': openText(); break;
    case 'image': openImage(); break;
    case 'legacy':
      showNotice(`.${ext} is the old binary Office format, which this app can't read or edit. In Word, PowerPoint or LibreOffice use Save As to convert it to .docx or .pptx (Google Drive can convert it too), then open the new file here.`);
      break;
    default:
      showNotice(`This app can't open ${ext ? `.${ext}` : 'this file'}. It handles PDF, Word, PowerPoint, Excel, CSV, plain text and common images.`);
  }
}

// ---------- spreadsheets (xlsx, xls, ods, csv, tsv) ----------

function openSheet() {
  current.wb = XLSX.read(new Uint8Array(current.bytes), { type: 'array', cellFormula: true, cellNF: true });
  current.sheet = current.wb.SheetNames[0];
  renderSheet();
  setStatus(`${current.wb.SheetNames.length} sheet(s). Tap a cell to edit it, then Save. Formulas are replaced by their typed value when you change that cell.`);
}

function renderSheet() {
  const { wb, sheet } = current;
  const ws = wb.Sheets[sheet] || {};
  const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  const rows = Math.min(ref.e.r, MAX_ROWS - 1);
  const cols = Math.min(ref.e.c, MAX_COLS - 1);

  const tabs = wb.SheetNames.map((n) =>
    `<button class="tab${n === sheet ? ' on' : ''}" data-sheet="${esc(n)}">${esc(n)}</button>`).join('');
  let h = `<div class="tabs">${tabs}</div><div class="grid-wrap"><table class="grid"><thead><tr><th></th>`;
  for (let c = 0; c <= cols; c++) h += `<th>${XLSX.utils.encode_col(c)}</th>`;
  h += '</tr></thead><tbody>';
  for (let r = 0; r <= rows; r++) {
    h += `<tr><th>${r + 1}</th>`;
    for (let c = 0; c <= cols; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      const cell = ws[addr];
      const shown = cell ? (cell.w ?? cell.v ?? '') : '';
      h += `<td><input data-addr="${addr}" value="${esc(shown)}"></td>`;
    }
    h += '</tr>';
  }
  h += '</tbody></table></div>';
  if (ref.e.r > rows || ref.e.c > cols) {
    h += `<p class="hint">Showing ${rows + 1} rows by ${cols + 1} columns. Everything else in the workbook is kept and saved unchanged.</p>`;
  }
  el.view.innerHTML = h;
}

function setCell(addr, raw) {
  const ws = current.wb.Sheets[current.sheet];
  const text = String(raw);
  const trimmed = text.trim();
  if (text === '') {
    delete ws[addr];
  } else {
    const isNumber = /^-?\d+(\.\d+)?$/.test(trimmed) && !/^-?0\d/.test(trimmed);
    ws[addr] = isNumber ? { t: 'n', v: Number(trimmed) } : { t: 's', v: text };
    expandRange(ws, addr);
  }
  current.dirty = true;
}

function expandRange(ws, addr) {
  const { r, c } = XLSX.utils.decode_cell(addr);
  const ref = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : { s: { r, c }, e: { r, c } };
  ref.s.r = Math.min(ref.s.r, r);
  ref.s.c = Math.min(ref.s.c, c);
  ref.e.r = Math.max(ref.e.r, r);
  ref.e.c = Math.max(ref.e.c, c);
  ws['!ref'] = XLSX.utils.encode_range(ref);
}

function writeWorkbook(bookType) {
  if (bookType === 'csv') return XLSX.write(current.wb, { bookType, type: 'string' });
  if (bookType === 'txt') return XLSX.write(current.wb, { bookType, FS: '\t', type: 'string' });
  return XLSX.write(current.wb, { bookType, type: 'array' });
}

function saveSheet() {
  const ext = current.ext === 'xlsm' ? 'xlsx' : current.ext; // macros are not kept
  const types = { xlsx: 'xlsx', xls: 'xls', ods: 'ods', csv: 'csv', tsv: 'txt' };
  let bookType = types[ext] || 'xlsx';
  let outExt = ext;
  let data;
  let note = '';
  try {
    data = writeWorkbook(bookType);
  } catch (err) {
    console.warn(err);
    bookType = 'xlsx';
    outExt = 'xlsx';
    data = writeWorkbook('xlsx');
    note = ` This format can't be written here, so it was saved as .xlsx.`;
  }
  if ((bookType === 'csv' || bookType === 'txt') && current.wb.SheetNames.length > 1) {
    note += ' CSV and TSV hold one sheet, so only the first sheet was saved.';
  }
  download(new Blob([data], { type: mimeFor(outExt) }), outName(current.base, outExt));
  current.dirty = false;
  setStatus(`Saved ${outName(current.base, outExt)}.${note}`);
}

// ---------- Word (docx) and PowerPoint (pptx) ----------

function nearestP(node, ns) {
  let n = node.parentNode;
  while (n && !(n.nodeType === 1 && n.localName === 'p' && n.namespaceURI === ns)) n = n.parentNode;
  return n;
}

// text elements that belong to this paragraph itself, not to a nested text box
function textNodes(p, ns) {
  return Array.from(p.getElementsByTagNameNS(ns, 't')).filter((t) => nearestP(t, ns) === p);
}

function paraText(p, ns) {
  return textNodes(p, ns).map((t) => t.textContent).join('');
}

// Puts the new text into the first text run and clears the rest, so the first run's formatting is kept.
function writeParaText(p, ns, prefix, text) {
  const ts = textNodes(p, ns);
  const doc = p.ownerDocument;
  if (!ts.length) {
    if (!text) return;
    const r = doc.createElementNS(ns, `${prefix}r`);
    const t = doc.createElementNS(ns, `${prefix}t`);
    if (prefix === 'w:') t.setAttribute('xml:space', 'preserve');
    t.textContent = text;
    r.appendChild(t);
    const end = Array.from(p.childNodes).find((n) => n.nodeType === 1 && n.localName === 'endParaRPr');
    p.insertBefore(r, end || null); // a:r must come before a:endParaRPr
    return;
  }
  ts[0].textContent = text;
  if (prefix === 'w:') ts[0].setAttribute('xml:space', 'preserve');
  ts.slice(1).forEach((t) => { t.textContent = ''; });
}

async function openDocx() {
  const zip = await JSZip.loadAsync(current.bytes);
  const entry = zip.file('word/document.xml');
  if (!entry) throw new Error('This is not a standard Word file (word/document.xml is missing).');
  const dom = parseXml(await entry.async('string'));
  current.zip = zip;
  current.parts = { 'word/document.xml': dom };
  current.items = Array.from(dom.getElementsByTagNameNS(W_NS, 'p')).map((p) => {
    const text = paraText(p, W_NS);
    return { part: 'word/document.xml', p, ns: W_NS, prefix: 'w:', orig: text, value: text, group: '' };
  });
  renderItems();
  setStatus(`Word document, ${current.items.length} paragraphs. Each box is one paragraph. Adding or deleting whole paragraphs needs Word itself.`);
}

const slideNo = (name) => Number(name.match(/slide(\d+)\.xml$/)[1]);

async function openPptx() {
  const zip = await JSZip.loadAsync(current.bytes);
  const names = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => slideNo(a) - slideNo(b));
  if (!names.length) throw new Error('No slides were found in this file.');
  current.zip = zip;
  current.parts = {};
  current.items = [];
  for (const name of names) {
    const dom = parseXml(await zip.file(name).async('string'));
    current.parts[name] = dom;
    for (const p of Array.from(dom.getElementsByTagNameNS(A_NS, 'p'))) {
      const text = paraText(p, A_NS);
      if (!text.trim()) continue;
      current.items.push({ part: name, p, ns: A_NS, prefix: 'a:', orig: text, value: text, group: `Slide ${slideNo(name)}` });
    }
  }
  renderItems();
  setStatus(`${names.length} slides, ${current.items.length} text paragraphs. Speaker notes aren't shown yet.`);
}

function renderItems() {
  let html = '';
  let last = null;
  current.items.forEach((it, i) => {
    if (it.group && it.group !== last) {
      html += `<h3>${esc(it.group)}</h3>`;
      last = it.group;
    }
    const rows = Math.min(8, Math.max(1, Math.ceil(it.value.length / 70)));
    html += `<textarea class="para" data-i="${i}" rows="${rows}">${esc(it.value)}</textarea>`;
  });
  el.view.innerHTML = html || '<p class="hint">No editable text was found in this file.</p>';
}

async function saveOffice() {
  const touched = new Set();
  for (const it of current.items) {
    if (it.value === it.orig) continue;
    writeParaText(it.p, it.ns, it.prefix, it.value);
    it.orig = it.value;
    touched.add(it.part);
  }
  for (const part of touched) {
    current.zip.file(part, XML_HEAD + new XMLSerializer().serializeToString(current.parts[part]));
  }
  const blob = await current.zip.generateAsync({ type: 'blob', mimeType: mimeFor(current.ext), compression: 'DEFLATE' });
  const name = outName(current.base, current.ext);
  download(blob, name);
  current.dirty = false;
  setStatus(`Saved ${name}.`);
}

// ---------- PDF ----------

async function openPdf() {
  current.pdfjs = await pdfjsLib.getDocument({ data: new Uint8Array(current.bytes.slice(0)) }).promise;
  current.pages = [];
  el.tools.hidden = false;
  el.tools.innerHTML = `
    <button data-tool="text" class="on">Add text</button>
    <button data-tool="white">Whiteout</button>
    <button id="undo-btn">Undo</button>
    <button id="text-btn">Show text</button>`;

  const readout = document.createElement('div');
  readout.id = 'readout';
  readout.className = 'readout';
  readout.hidden = true;
  const pagesEl = document.createElement('div');
  pagesEl.id = 'pages';
  el.view.append(readout, pagesEl);

  const avail = Math.min(el.view.clientWidth - 16, 900);
  const dpr = window.devicePixelRatio || 1;
  for (let n = 1; n <= current.pdfjs.numPages; n++) {
    const page = await current.pdfjs.getPage(n);
    const base = page.getViewport({ scale: 1 });
    const vp = page.getViewport({ scale: avail / base.width });

    const wrap = document.createElement('div');
    wrap.className = 'page';
    wrap.style.width = `${vp.width}px`;
    wrap.style.height = `${vp.height}px`;
    const canvas = document.createElement('canvas');
    canvas.width = Math.floor(vp.width * dpr);
    canvas.height = Math.floor(vp.height * dpr);
    canvas.style.width = `${vp.width}px`;
    canvas.style.height = `${vp.height}px`;
    const layer = document.createElement('div');
    layer.className = 'layer';
    wrap.append(canvas, layer);
    pagesEl.append(wrap);

    const params = { canvasContext: canvas.getContext('2d'), viewport: vp };
    if (dpr !== 1) params.transform = [dpr, 0, 0, dpr, 0, 0];
    await page.render(params).promise;

    current.pages[n - 1] = { layer, w: vp.width, h: vp.height, ptW: base.width };
    bindLayer(layer, n);
  }
  setTool('text');
  setStatus(`PDF, ${current.pdfjs.numPages} pages. Add text or whiteout, then Save. Existing text can be covered or added to, not edited in place.`);
}

function setTool(tool) {
  current.tool = tool;
  el.tools.querySelectorAll('[data-tool]').forEach((b) => b.classList.toggle('on', b.dataset.tool === tool));
  document.querySelectorAll('.layer').forEach((l) => l.classList.toggle('white-mode', tool === 'white'));
}

function bindLayer(layer, n) {
  layer.addEventListener('click', (e) => {
    if (current.tool !== 'text') return;
    const text = window.prompt('Text to add on this page:');
    if (!text) return;
    const r = layer.getBoundingClientRect();
    const pg = current.pages[n - 1];
    current.overlays.push({
      page: n, kind: 'text', text,
      fx: (e.clientX - r.left) / pg.w,
      fy: (e.clientY - r.top) / pg.h,
    });
    drawOverlays();
  });

  layer.addEventListener('pointerdown', (e) => {
    if (current.tool !== 'white') return;
    e.preventDefault();
    layer.setPointerCapture(e.pointerId);
    const r = layer.getBoundingClientRect();
    const x0 = e.clientX - r.left;
    const y0 = e.clientY - r.top;
    const box = document.createElement('div');
    box.className = 'white preview';
    layer.append(box);

    const move = (ev) => {
      const x1 = ev.clientX - r.left;
      const y1 = ev.clientY - r.top;
      box.style.left = `${Math.min(x0, x1)}px`;
      box.style.top = `${Math.min(y0, y1)}px`;
      box.style.width = `${Math.abs(x1 - x0)}px`;
      box.style.height = `${Math.abs(y1 - y0)}px`;
    };
    const finish = (ev) => {
      layer.removeEventListener('pointermove', move);
      layer.removeEventListener('pointerup', finish);
      layer.removeEventListener('pointercancel', finish);
      box.remove();
      if (ev.type !== 'pointerup') return;
      const x1 = ev.clientX - r.left;
      const y1 = ev.clientY - r.top;
      const w = Math.abs(x1 - x0);
      const h = Math.abs(y1 - y0);
      if (w > 4 && h > 4) {
        const pg = current.pages[n - 1];
        current.overlays.push({
          page: n, kind: 'white',
          fx: Math.min(x0, x1) / pg.w, fy: Math.min(y0, y1) / pg.h,
          fw: w / pg.w, fh: h / pg.h,
        });
        drawOverlays();
      }
    };
    layer.addEventListener('pointermove', move);
    layer.addEventListener('pointerup', finish);
    layer.addEventListener('pointercancel', finish);
  });
}

function drawOverlays() {
  current.pages.forEach((pg) => { pg.layer.textContent = ''; });
  for (const o of current.overlays) {
    const pg = current.pages[o.page - 1];
    const d = document.createElement('div');
    if (o.kind === 'white') {
      d.className = 'white';
      d.style.cssText = `left:${o.fx * pg.w}px;top:${o.fy * pg.h}px;width:${o.fw * pg.w}px;height:${o.fh * pg.h}px`;
    } else {
      d.className = 'txt';
      d.style.cssText = `left:${o.fx * pg.w}px;top:${o.fy * pg.h}px;font-size:${12 * (pg.w / pg.ptW)}px`;
      d.textContent = o.text;
    }
    pg.layer.append(d);
  }
}

async function showPdfText() {
  const box = document.getElementById('readout');
  if (!box.hidden) { box.hidden = true; return; }
  setStatus('Reading text…');
  let out = '';
  for (let n = 1; n <= current.pdfjs.numPages; n++) {
    const page = await current.pdfjs.getPage(n);
    const tc = await page.getTextContent();
    out += `--- Page ${n} ---\n${tc.items.map((i) => (i.str ?? '') + (i.hasEOL ? '\n' : ' ')).join('')}\n`;
  }
  box.textContent = out.trim() || 'No text found. This may be a scanned PDF made of images.';
  box.hidden = false;
  setStatus('Extracted text is shown above. Tap Show text again to hide it.');
}

async function savePdf() {
  let doc;
  let font;
  try {
    doc = await PDFDocument.load(current.bytes, { ignoreEncryption: true });
    font = await doc.embedFont(StandardFonts.Helvetica);
    for (const o of current.overlays) {
      const page = doc.getPage(o.page - 1);
      const { width: W, height: H } = page.getSize();
      if (o.kind === 'white') {
        page.drawRectangle({ x: o.fx * W, y: H - (o.fy + o.fh) * H, width: o.fw * W, height: o.fh * H, color: rgb(1, 1, 1) });
      } else {
        page.drawText(o.text, { x: o.fx * W, y: H - o.fy * H - 12, size: 12, font, color: rgb(0, 0, 0) });
      }
    }
  } catch (err) {
    console.error(err);
    throw new Error("Added text uses characters the built-in PDF font can't write (for example non-Latin letters). Retype it with Latin characters.");
  }
  const bytes = await doc.save();
  download(new Blob([bytes], { type: MIME.pdf }), outName(current.base, 'pdf'));
  current.dirty = false;
  setStatus(`Saved ${outName(current.base, 'pdf')}.`);
}

// ---------- plain text and images ----------

function openText() {
  const ta = document.createElement('textarea');
  ta.className = 'plain';
  ta.spellcheck = false;
  ta.value = new TextDecoder('utf-8').decode(current.bytes);
  el.view.append(ta);
  setStatus('Text file. Edit freely, then Save.');
}

function saveText() {
  const ta = el.view.querySelector('textarea.plain');
  const name = outName(current.base, current.ext);
  download(new Blob([ta.value], { type: 'text/plain;charset=utf-8' }), name);
  current.dirty = false;
  setStatus(`Saved ${name}.`);
}

function openImage() {
  const type = current.ext === 'jpg' ? 'image/jpeg' : `image/${current.ext}`;
  const img = document.createElement('img');
  img.className = 'preview-img';
  img.alt = current.name;
  img.src = URL.createObjectURL(new Blob([current.bytes], { type }));
  el.view.append(img);
  setStatus('Image preview. Images can be viewed here but not edited.');
}

// ---------- save dispatch ----------

async function saveCurrent() {
  switch (current.kind) {
    case 'sheet': return saveSheet();
    case 'docx':
    case 'pptx': return saveOffice();
    case 'pdf': return savePdf();
    case 'text': return saveText();
    default: return undefined;
  }
}
