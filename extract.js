// Text extraction for common document formats. No npm dependencies:
// Office files are zip archives, read here with Node's built-in zlib.
const zlib = require('zlib');
const path = require('path');
const { pathToFileURL } = require('url');

const KINDS = {
  pdf: 'pdf',
  jpg: 'image', jpeg: 'image', png: 'image', gif: 'image', webp: 'image',
  docx: 'docx', docm: 'docx', dotx: 'docx',
  pptx: 'pptx', pptm: 'pptx', ppsx: 'pptx', potx: 'pptx',
  xlsx: 'xlsx', xlsm: 'xlsx', xltx: 'xlsx',
  odt: 'odf', odp: 'odf', ods: 'odf',
  txt: 'text', md: 'text', markdown: 'text', csv: 'text', tsv: 'text', json: 'text',
  xml: 'text', yaml: 'text', yml: 'text', log: 'text', tex: 'text',
  html: 'html', htm: 'html',
  rtf: 'rtf',
  doc: 'legacy', ppt: 'legacy', xls: 'legacy',
};
const IMAGE_TYPES = { jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp' };

const extOf = (name) => (String(name).split('.').pop() || '').toLowerCase();
const kindOf = (name) => KINDS[extOf(name)] || null;
const imageTypeOf = (name) => IMAGE_TYPES[extOf(name)] || null;

function readZip(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('not a valid document (could not open it as an Office/zip file).');
  const total = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  const entries = {};
  for (let n = 0; n < total; n++) {
    if (off + 46 > buf.length || buf.readUInt32LE(off) !== 0x02014b50) break;
    const method = buf.readUInt16LE(off + 10);
    const csize = buf.readUInt32LE(off + 20);
    const nlen = buf.readUInt16LE(off + 28), elen = buf.readUInt16LE(off + 30), clen = buf.readUInt16LE(off + 32);
    const lho = buf.readUInt32LE(off + 42);
    entries[buf.toString('utf8', off + 46, off + 46 + nlen)] = { method, csize, lho };
    off += 46 + nlen + elen + clen;
  }
  return {
    names: Object.keys(entries),
    read(name) {
      const e = entries[name];
      if (!e) return null;
      const nlen = buf.readUInt16LE(e.lho + 26), elen = buf.readUInt16LE(e.lho + 28);
      const start = e.lho + 30 + nlen + elen;
      const data = buf.subarray(start, start + e.csize);
      if (e.method === 0) return data;
      if (e.method === 8) return zlib.inflateRawSync(data, { maxOutputLength: 200 * 1024 * 1024 });
      throw new Error('uses an unsupported compression method.');
    },
    text(name) { const b = this.read(name); return b ? b.toString('utf8') : null; },
  };
}

function decode(s) {
  return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos|nbsp);/g, (m, e) => {
    if (e[0] === '#') {
      const code = e[1] === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      try { return String.fromCodePoint(code); } catch (_) { return ''; }
    }
    return { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' }[e];
  });
}

// Collect the text runs of each paragraph in an OOXML part.
function paragraphs(xml, pTag, tTag) {
  const out = [];
  const re = new RegExp('<' + tTag + '(?:\\s[^>]*)?>([^<]*)</' + tTag + '>|<w:tab/>', 'g');
  for (const part of xml.split('</' + pTag + '>')) {
    let line = '', m;
    re.lastIndex = 0;
    while ((m = re.exec(part))) line += m[1] !== undefined ? decode(m[1]) : '\t';
    if (line.trim()) out.push(line);
  }
  return out;
}

// Word body in reading order: headings become "# ", list items "- ", table rows "a | b | c".
function docxBody(xml) {
  const out = [];
  const re = /<w:tbl>[\s\S]*?<\/w:tbl>|<w:p[ >][\s\S]*?<\/w:p>/g;
  let m;
  while ((m = re.exec(xml))) {
    const blk = m[0];
    if (blk.startsWith('<w:tbl>')) {
      const rows = [];
      for (const tr of blk.match(/<w:tr[ >][\s\S]*?<\/w:tr>/g) || []) {
        const cells = (tr.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || []).map((tc) => paragraphs(tc, 'w:p', 'w:t').join(' ').replace(/\s+/g, ' ').trim());
        if (cells.some(Boolean)) rows.push(cells.join(' | '));
      }
      if (rows.length) out.push(rows.join('\n'));
      continue;
    }
    const text = (paragraphs(blk, 'w:p', 'w:t')[0] || '').trim();
    if (!text) continue;
    const style = (blk.match(/<w:pStyle w:val="([^"]+)"/) || [])[1] || '';
    const level = (blk.match(/<w:ilvl w:val="(\d+)"/) || [])[1];
    const hm = style.match(/^Heading\s*(\d)/i);
    if (/^Title$/i.test(style)) out.push('# ' + text);
    else if (hm) out.push('#'.repeat(Math.min(+hm[1] + 1, 4)) + ' ' + text);
    else if (/<w:numPr>/.test(blk) || /^List/i.test(style)) out.push('  '.repeat(Math.min(+level || 0, 3)) + '- ' + text);
    else out.push(text);
  }
  return out.join('\n');
}

function fromDocx(buf) {
  const z = readZip(buf);
  const doc = z.text('word/document.xml');
  if (!doc) throw new Error('is not a Word document.');
  const parts = [docxBody(doc)];
  for (const n of z.names.filter((n) => /^word\/(footnotes|endnotes)\.xml$/.test(n)).sort()) {
    const t = paragraphs(z.text(n), 'w:p', 'w:t').join('\n');
    if (t) parts.push('Notes:\n' + t);
  }
  return parts.join('\n\n');
}

const IMG_MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg' };
// Pictures stored inside an Office file (skips tiny icons and formats the vision model cannot read).
function officeImages(z, prefix, labelFor) {
  const out = [];
  for (const n of z.names) {
    if (!n.startsWith(prefix)) continue;
    const mime = IMG_MIME[extOf(n)];
    if (!mime) continue;
    let data;
    try { data = z.read(n); } catch (_) { continue; }
    if (!data || data.length < 20 * 1024 || data.length > 3.5 * 1024 * 1024) continue;
    out.push({ label: labelFor ? labelFor(n) : 'picture in the document', mime, data, order: n });
  }
  return out;
}

function pptxSlides(z) {
  return z.names
    .map((n) => ({ n, m: n.match(/^ppt\/slides\/slide(\d+)\.xml$/) }))
    .filter((x) => x.m)
    .sort((a, b) => +a.m[1] - +b.m[1]);
}

function fromPptx(buf) {
  const z = readZip(buf);
  const slides = pptxSlides(z);
  if (!slides.length) throw new Error('is not a PowerPoint file.');
  const out = [];
  slides.forEach((s, i) => {
    const xml = z.text(s.n) || '';
    const tables = (xml.match(/<a:tbl>[\s\S]*?<\/a:tbl>/g) || []).map((t) =>
      (t.match(/<a:tr[ >][\s\S]*?<\/a:tr>/g) || [])
        .map((tr) => (tr.match(/<a:tc[ >][\s\S]*?<\/a:tc>/g) || []).map((tc) => paragraphs(tc, 'a:p', 'a:t').join(' ').replace(/\s+/g, ' ').trim()).join(' | '))
        .filter((r) => r.replace(/[\s|]/g, '')).join('\n'));
    const rest = xml.replace(/<a:tbl>[\s\S]*?<\/a:tbl>/g, '');
    let title = '';
    const body = [];
    for (const sp of rest.match(/<p:sp>[\s\S]*?<\/p:sp>/g) || []) {
      const lines = paragraphs(sp, 'a:p', 'a:t');
      if (/<p:ph[^>]*type="(title|ctrTitle)"/.test(sp)) title = title || lines.join(' ');
      else body.push(...lines);
    }
    let notes = [];
    const rels = z.text(`ppt/slides/_rels/slide${s.m[1]}.xml.rels`) || '';
    const nm = rels.match(/notesSlides\/(notesSlide\d+\.xml)/);
    if (nm) {
      const nx = z.text('ppt/notesSlides/' + nm[1]);
      // Notes pages repeat the slide number as a field; keep only real text.
      if (nx) notes = paragraphs(nx, 'a:p', 'a:t').filter((l) => !/^\d+$/.test(l.trim()));
    }
    if (!title && !body.length && !tables.length && !notes.length) return;
    out.push(`--- Slide ${i + 1}${title ? ': ' + title : ''} ---\n${body.join('\n')}` + (tables.length ? '\n' + tables.join('\n\n') : '') + (notes.length ? `\nSpeaker notes: ${notes.join(' ')}` : ''));
  });
  return out.join('\n\n');
}

function pptxImages(z) {
  const out = [];
  const seen = new Set();
  pptxSlides(z).forEach((s, i) => {
    const rels = z.text(`ppt/slides/_rels/slide${s.m[1]}.xml.rels`) || '';
    for (const r of rels.matchAll(/Target="\.\.\/media\/([^"]+)"/g)) {
      const name = 'ppt/media/' + r[1];
      if (seen.has(name)) continue;
      seen.add(name);
      for (const im of officeImages({ names: [name], read: (n) => z.read(n) }, 'ppt/media/', () => `picture on slide ${i + 1}`)) out.push({ ...im, order: String(i + 1).padStart(4, '0') + name });
    }
  });
  return out;
}

function colIndex(ref) {
  const letters = (ref.match(/^[A-Z]+/) || ['A'])[0];
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

function fromXlsx(buf) {
  const z = readZip(buf);
  const shared = [];
  const ss = z.text('xl/sharedStrings.xml');
  if (ss) {
    for (const si of ss.match(/<si>[\s\S]*?<\/si>|<si\/>/g) || []) {
      shared.push(decode((si.match(/<t(?:\s[^>]*)?>[^<]*<\/t>/g) || []).map((t) => t.replace(/<[^>]+>/g, '')).join('')));
    }
  }
  const wb = z.text('xl/workbook.xml') || '';
  const rels = z.text('xl/_rels/workbook.xml.rels') || '';
  const relMap = {};
  for (const r of rels.match(/<Relationship\b[^>]*>/g) || []) {
    const id = (r.match(/\bId="([^"]+)"/) || [])[1];
    const target = (r.match(/\bTarget="([^"]+)"/) || [])[1];
    if (id && target) relMap[id] = target.replace(/^\/?(xl\/)?/, 'xl/');
  }
  const sheets = [];
  for (const s of wb.match(/<sheet\b[^>]*>/g) || []) {
    const name = decode((s.match(/\bname="([^"]*)"/) || [])[1] || 'Sheet');
    const rid = (s.match(/\br:id="([^"]+)"/) || [])[1];
    if (relMap[rid]) sheets.push({ name, path: relMap[rid] });
  }
  if (!sheets.length) throw new Error('is not an Excel workbook.');
  const out = [];
  for (const sh of sheets) {
    const xml = z.text(sh.path);
    if (!xml) continue;
    const rows = [];
    for (const row of (xml.match(/<row\b[\s\S]*?<\/row>/g) || []).slice(0, 3000)) {
      const cells = [];
      const cre = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
      let m;
      while ((m = cre.exec(row))) {
        const attrs = m[1], inner = m[2] || '';
        const ref = (attrs.match(/\br="([A-Z]+\d+)"/) || [])[1];
        const t = (attrs.match(/\bt="(\w+)"/) || [])[1];
        let val = '';
        if (t === 's') val = shared[+((inner.match(/<v>(\d+)<\/v>/) || [])[1])] || '';
        else if (t === 'inlineStr') val = decode((inner.match(/<t(?:\s[^>]*)?>([^<]*)<\/t>/g) || []).map((x) => x.replace(/<[^>]+>/g, '')).join(''));
        else val = decode((inner.match(/<v>([^<]*)<\/v>/) || [])[1] || '');
        if (val !== '' && ref) cells[colIndex(ref)] = val;
      }
      if (cells.length) rows.push(Array.from(cells, (c) => c || '').join('\t'));
    }
    if (rows.length) out.push(`--- Sheet: ${sh.name} ---\n${rows.join('\n')}`);
  }
  return out.join('\n\n');
}

function fromOdf(buf) {
  const z = readZip(buf);
  let xml = z.text('content.xml');
  if (!xml) throw new Error('is not an OpenDocument file.');
  xml = xml
    .replace(/<text:(tab)\/>/g, '\t')
    .replace(/<text:line-break\/>/g, '\n')
    .replace(/<\/draw:page>/g, '\n\n--- next slide ---\n\n')
    .replace(/<\/(text:p|text:h|table:table-row)>/g, '\n')
    .replace(/<\/table:table-cell>/g, '\t')
    .replace(/<[^>]+>/g, '');
  return decode(xml).replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n');
}

function fromHtml(buf) {
  return decode(
    buf.toString('utf8')
      .replace(/<(script|style|noscript)[\s\S]*?<\/\1>/gi, '')
      .replace(/<br\s*\/?>|<\/(p|div|li|tr|h[1-6]|section|article)>/gi, '\n')
      .replace(/<\/t[dh]>/gi, '\t')
      .replace(/<[^>]+>/g, '')
  ).replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n');
}

function fromRtf(buf) {
  return buf.toString('latin1')
    .replace(/\{\\\*[^{}]*\}/g, '')
    .replace(/\\par[d]?/g, '\n')
    .replace(/\\u(-?\d+)\??/g, (m, n) => String.fromCharCode(+n < 0 ? +n + 65536 : +n))
    .replace(/\\'([0-9a-fA-F]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/\\[a-zA-Z]+-?\d* ?/g, '')
    .replace(/[{}]/g, '')
    .replace(/\n{3,}/g, '\n\n');
}


// ---- PDF text (best effort, no dependencies). Handles Flate streams, object streams and ToUnicode maps.
function ascii85(b) {
  const t = b.toString('latin1').replace(/\s+/g, '').replace(/^<~/, '').replace(/~>.*$/, '');
  const out = [];
  let grp = [];
  for (const ch of t) {
    if (ch === 'z' && !grp.length) { out.push(0, 0, 0, 0); continue; }
    const c = ch.charCodeAt(0) - 33;
    if (c < 0 || c > 84) continue;
    grp.push(c);
    if (grp.length === 5) { let v = 0; for (const x of grp) v = v * 85 + x; out.push((v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255); grp = []; }
  }
  if (grp.length > 1) { const n = grp.length; while (grp.length < 5) grp.push(84); let v = 0; for (const x of grp) v = v * 85 + x; const bytes = [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255]; out.push(...bytes.slice(0, n - 1)); }
  return Buffer.from(out);
}

function pdfObjects(buf) {
  const latin = buf.toString('latin1');
  const objs = {};
  const inflate = (b) => { try { return zlib.inflateSync(b); } catch (_) { try { return zlib.inflateRawSync(b); } catch (_) { return null; } } };
  const re = /(\d+)\s+(\d+)\s+obj\b/g;
  let m;
  while ((m = re.exec(latin))) {
    const start = re.lastIndex;
    const sIdx = latin.indexOf('stream', start);
    const eIdx = latin.indexOf('endobj', start);
    if (eIdx < 0) break;
    let dict, data = null, jpeg = false;
    if (sIdx >= 0 && sIdx < eIdx) {
      dict = latin.slice(start, sIdx);
      let b = sIdx + 6;
      if (latin[b] === '\r') b++;
      if (latin[b] === '\n') b++;
      let e = latin.indexOf('endstream', b);
      if (e < 0) e = eIdx;
      let raw = buf.subarray(b, e);
      const fm = dict.match(/\/Filter\s*(\[[^\]]*\]|\/\w+)/);
      const filters = fm ? [...fm[1].matchAll(/\/(\w+)/g)].map((x) => x[1]) : [];
      for (const fl of filters) {
        if (!raw) break;
        if (fl === 'FlateDecode' || fl === 'Fl') raw = inflate(raw);
        else if (fl === 'ASCII85Decode' || fl === 'A85') raw = ascii85(raw);
        else if (fl === 'ASCIIHexDecode' || fl === 'AHx') raw = Buffer.from(raw.toString('latin1').replace(/[^0-9a-fA-F]/g, ''), 'hex');
        else if (fl === 'DCTDecode' || fl === 'DCT') { jpeg = true; break; } // a JPEG: keep the bytes as they are
        else raw = null; // a filter this reader does not handle
      }
      data = raw;
      re.lastIndex = Math.max(re.lastIndex, e);
    } else dict = latin.slice(start, eIdx);
    objs[m[1]] = { dict, data, jpeg };
  }
  // compressed object streams
  for (const k of Object.keys(objs)) {
    const o = objs[k];
    if (!/\/Type\s*\/ObjStm/.test(o.dict) || !o.data) continue;
    const n = +((o.dict.match(/\/N\s+(\d+)/) || [])[1] || 0), first = +((o.dict.match(/\/First\s+(\d+)/) || [])[1] || 0);
    const txt = o.data.toString('latin1');
    const nums = txt.slice(0, first).trim().split(/\s+/).map(Number);
    for (let i = 0; i < n; i++) {
      const id = nums[i * 2], off = nums[i * 2 + 1];
      const next = i + 1 < n ? nums[(i + 1) * 2 + 1] : txt.length - first;
      if (!(id in objs)) objs[id] = { dict: txt.slice(first + off, first + next), data: null };
    }
  }
  return objs;
}

function parseCMap(text) {
  const map = {};
  const u16 = (h) => { let s = ''; for (let i = 0; i + 3 < h.length + 0; i += 4) s += String.fromCharCode(parseInt(h.substr(i, 4), 16)); if (h.length === 2) s = String.fromCharCode(parseInt(h, 16)); return s; };
  for (const blk of text.match(/beginbfchar[\s\S]*?endbfchar/g) || []) {
    for (const x of blk.matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g)) map[parseInt(x[1], 16)] = u16(x[2]);
  }
  for (const blk of text.match(/beginbfrange[\s\S]*?endbfrange/g) || []) {
    for (const x of blk.matchAll(/<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(<[0-9a-fA-F]+>|\[[^\]]*\])/g)) {
      const lo = parseInt(x[1], 16), hi = Math.min(parseInt(x[2], 16), lo + 65535);
      if (x[3][0] === '[') { const arr = [...x[3].matchAll(/<([0-9a-fA-F]+)>/g)].map((a) => u16(a[1])); for (let c = lo; c <= hi && c - lo < arr.length; c++) map[c] = arr[c - lo]; }
      else { const base = x[3].slice(1, -1); const start = parseInt(base, 16); const len = base.length; for (let c = lo; c <= hi; c++) map[c] = u16((start + c - lo).toString(16).padStart(len, '0')); }
    }
  }
  return map;
}

const LIGATURES = { 0x0b: 'ff', 0x0c: 'fi', 0x0d: 'fl', 0x0e: 'ffi', 0x0f: 'ffl', 0x1b: 'ff', 0x1c: 'fi', 0x1d: 'fl', 0x1e: 'ffi', 0x1f: 'ffl' }; // LaTeX fonts without a text map

function fromPdf(buf) {
  const objs = pdfObjects(buf);
  const cmapOf = {};
  const fontMap = (num) => {
    if (num in cmapOf) return cmapOf[num];
    const o = objs[num]; let map = null;
    if (o) {
      const tu = (o.dict.match(/\/ToUnicode\s+(\d+)\s+\d+\s+R/) || [])[1];
      if (tu && objs[tu] && objs[tu].data) map = parseCMap(objs[tu].data.toString('latin1'));
    }
    return (cmapOf[num] = map);
  };
  const resolveDict = (d) => { const r = d.match(/^\s*(\d+)\s+\d+\s+R\s*$/); return r && objs[r[1]] ? objs[r[1]].dict : d; };
  const grabDict = (src, key) => { // returns the text of << ... >> after /Key, or an indirect object's dict
    const i = src.search(new RegExp('/' + key + '\\b'));
    if (i < 0) return null;
    let rest = src.slice(i + key.length + 1).trimStart();
    const ref = rest.match(/^(\d+)\s+\d+\s+R/);
    if (ref) return objs[ref[1]] ? objs[ref[1]].dict : null;
    if (rest.startsWith('<<')) { let depth = 0, j = 0; for (; j < rest.length; j++) { if (rest.startsWith('<<', j)) { depth++; j++; } else if (rest.startsWith('>>', j)) { depth--; j++; if (!depth) break; } } return rest.slice(0, j + 1); }
    return null;
  };
  const pageFonts = (dict, seen = 0) => {
    const res = grabDict(dict, 'Resources');
    const fonts = {};
    if (res) {
      const fd = grabDict(res, 'Font');
      if (fd) for (const x of fd.matchAll(/\/([^\s\/<>\[\]()]+)\s+(\d+)\s+\d+\s+R/g)) fonts[x[1]] = x[2];
    } else if (seen < 6) {
      const par = (dict.match(/\/Parent\s+(\d+)\s+\d+\s+R/) || [])[1];
      if (par && objs[par]) return pageFonts(objs[par].dict, seen + 1);
    }
    return fonts;
  };
  const pages = [];
  const contentFont = {}; // content stream object -> font resource map
  const order = [];
  for (const k of Object.keys(objs).map(Number).sort((a, b) => a - b)) {
    const d = objs[k].dict;
    if (!/\/Type\s*\/Page\b(?!s)/.test(d)) continue;
    const fonts = pageFonts(d);
    const c = d.match(/\/Contents\s*(\[[^\]]*\]|\d+\s+\d+\s+R)/);
    if (!c) continue;
    const refs = [...c[1].matchAll(/(\d+)\s+\d+\s+R/g)].map((x) => x[1]);
    pages.push({ refs, fonts });
  }
  const unesc = (s) => s.replace(/\\([nrtbf()\\]|[0-7]{1,3}|\r?\n)/g, (m, e) => {
    if (e === 'n') return '\n'; if (e === 'r') return '\r'; if (e === 't') return '\t'; if (e === 'b' || e === 'f') return '';
    if (/^[0-7]+$/.test(e)) return String.fromCharCode(parseInt(e, 8) & 255);
    if (e[0] === '\r' || e[0] === '\n') return '';
    return e;
  });
  const out = [];
  for (const pg of pages) {
    let text = '';
    for (const ref of pg.refs) {
      const o = objs[ref];
      if (!o || !o.data) continue;
      const s = o.data.toString('latin1');
      let cmap = null;
      const tokens = /BT|ET|\/([^\s\/<>\[\]()]+)\s+[-\d.]+\s+Tf|\[((?:[^\]\\]|\\.)*)\]\s*TJ|\(((?:[^()\\]|\\.|\((?:[^()\\]|\\.)*\))*)\)\s*(?:Tj|'|")|<([0-9a-fA-F\s]+)>\s*Tj|T\*|(-?[\d.]+)\s+(-?[\d.]+)\s+T[dD]/g;
      const dec = (hexOrLit, isHex) => {
        if (isHex) {
          const h = hexOrLit.replace(/\s+/g, '');
          let r = '';
          if (cmap) {
            const step = Object.keys(cmap).some((c) => +c > 255) || h.length % 4 === 0 ? 4 : 2;
            for (let i = 0; i + step <= h.length; i += step) { const ch = cmap[parseInt(h.substr(i, step), 16)]; if (ch !== undefined) r += ch; }
          } else for (let i = 0; i + 2 <= h.length; i += 2) r += String.fromCharCode(parseInt(h.substr(i, 2), 16));
          return r;
        }
        const raw = unesc(hexOrLit);
        if (cmap) { let r = ''; for (const ch of raw) { const v = cmap[ch.charCodeAt(0)]; r += v !== undefined ? v : ch; } return r; }
        return raw.replace(/[\x0b-\x0f\x1b-\x1f]/g, (c) => LIGATURES[c.charCodeAt(0)] || '');
      };
      let m;
      while ((m = tokens.exec(s))) {
        const t = m[0];
        if (m[1] !== undefined) { const fnum = pg.fonts[m[1]]; cmap = fnum ? fontMap(fnum) : null; }
        else if (m[2] !== undefined) {
          for (const x of m[2].matchAll(/\(((?:[^()\\]|\\.)*)\)|<([0-9a-fA-F\s]+)>|(-?[\d.]+)/g)) {
            if (x[1] !== undefined) text += dec(x[1], false);
            else if (x[2] !== undefined) text += dec(x[2], true);
            else if (x[3] !== undefined && parseFloat(x[3]) < -200) text += ' ';
          }
        } else if (m[3] !== undefined) text += dec(m[3], false);
        else if (m[4] !== undefined) text += dec(m[4], true);
        else if (t === 'ET' || t === 'T*') text += '\n';
        else if (m[6] !== undefined) { if (Math.abs(parseFloat(m[6])) > 2) text += '\n'; else if (parseFloat(m[5]) > 2 && !/\s$/.test(text)) text += ' '; }
      }
    }
    text = text.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    if (text) out.push(text);
  }
  const joined = out.map((t, i) => `--- Page ${i + 1} ---\n${t}`).join('\n\n');
  if (joined.replace(/--- Page \d+ ---|\s/g, '').length < 40) throw new Error('has no text the server can read (it may be a scan, or use an unusual encoding). Paste the text instead, or export it as a .docx.');
  return joined;
}

// ---- PDF through pdf.js (bundled in vendor/pdfjs; needs Node 22.13+). Returns null if it cannot run here.
let pdfjsLib;
function loadPdfjs() {
  if (!pdfjsLib) {
    pdfjsLib = (async () => {
      try {
        const dir = path.join(__dirname, 'vendor', 'pdfjs');
        // pdf.js expects browser classes that plain Node lacks. Reading text never draws, so empty stand-ins are enough.
        if (typeof globalThis.DOMMatrix === 'undefined') globalThis.DOMMatrix = class DOMMatrix { constructor(m) { [this.a, this.b, this.c, this.d, this.e, this.f] = Array.isArray(m) && m.length === 6 ? m : [1, 0, 0, 1, 0, 0]; } multiplySelf() { return this; } preMultiplySelf() { return this; } translate() { return this; } scale() { return this; } invertSelf() { return this; } };
        if (typeof globalThis.Path2D === 'undefined') globalThis.Path2D = class Path2D { };
        if (typeof globalThis.ImageData === 'undefined') globalThis.ImageData = class ImageData { constructor(d, w, h) { this.data = d; this.width = w; this.height = h; } };
        const so = { log: console.log, warn: console.warn }; // pdf.js prints a harmless notice about an optional drawing package
        const quiet = (f) => (...a) => { if (!/napi-rs\/canvas/.test(String(a[0]))) f.apply(console, a); };
        console.log = quiet(so.log); console.warn = quiet(so.warn);
        let mod;
        try { mod = await import(pathToFileURL(path.join(dir, 'pdf.min.mjs')).href); }
        finally { console.log = so.log; console.warn = so.warn; }
        mod.GlobalWorkerOptions.workerSrc = pathToFileURL(path.join(dir, 'pdf.worker.min.mjs')).href;
        return mod;
      } catch (_) { return null; }
    })();
  }
  return pdfjsLib;
}

async function pdfPagesModern(buf) {
  const lib = await loadPdfjs();
  if (!lib) return null;
  let doc;
  try {
    doc = await lib.getDocument({ data: new Uint8Array(buf), isEvalSupported: false, disableFontFace: true, useSystemFonts: false, verbosity: 0 }).promise;
  } catch (e) {
    if (e && e.name === 'PasswordException') throw new Error('is password protected. Remove the password and upload it again.');
    return null;
  }
  const pages = [];
  try {
    for (let i = 1; i <= doc.numPages && i <= 400; i++) {
      const tc = await (await doc.getPage(i)).getTextContent();
      let text = '', lineY = null, lineH = 10, start = true;
      for (const it of tc.items) {
        if (typeof it.str !== 'string') continue;
        if (it.str) {
          const y = it.transform ? it.transform[5] : null, h = it.height || lineH;
          if (start && lineY !== null && y !== null && lineY - y > 1.7 * h) text += '\n'; // gap between paragraphs
          if (y !== null) { lineY = y; lineH = h; }
          text += it.str;
          start = false;
        }
        if (it.hasEOL) { text += '\n'; start = true; }
      }
      pages.push(text.replace(/(\w)-\n(\p{Ll})/gu, '$1$2').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim());
    }
  } finally { try { await doc.destroy(); } catch (_) { } }
  return pages;
}

// JPEG pictures stored in a PDF (page scans are usually these). Skips CMYK, which vision models read poorly.
function pdfJpegs(buf) {
  const out = [];
  const objs = pdfObjects(buf);
  for (const k of Object.keys(objs).map(Number).sort((a, b) => a - b)) {
    const o = objs[k];
    if (!o.jpeg || !o.data || !/\/Subtype\s*\/Image/.test(o.dict) || o.data[0] !== 0xff || o.data[1] !== 0xd8) continue;
    const w = +((o.dict.match(/\/Width\s+(\d+)/) || [])[1] || 0), h = +((o.dict.match(/\/Height\s+(\d+)/) || [])[1] || 0);
    if (/\/DeviceCMYK/.test(o.dict) || /\/SMask/.test(o.dict) && w < 400) continue;
    out.push({ w, h, data: Buffer.from(o.data), mime: 'image/jpeg', order: String(k).padStart(8, '0') });
  }
  return out;
}

async function fromPdfAll(buf) {
  let pages = null;
  try { pages = await pdfPagesModern(buf); } catch (e) { if (/password/.test(e.message)) throw e; }
  const modern = !!pages;
  if (!pages) { try { pages = fromPdf(buf).split(/--- Page \d+ ---\n/).slice(1); } catch (_) { pages = []; } }
  const chars = pages.reduce((n, p) => n + p.replace(/\s/g, '').length, 0);
  const scanned = pages.length === 0 || chars < 60 * Math.max(1, pages.length) * 0.4 || chars < 80;
  const jpegs = pdfJpegs(buf).filter((im) => im.w >= 300 && im.h >= 300 && im.data.length >= 15 * 1024 && im.data.length <= 3.5 * 1024 * 1024);
  let images;
  if (scanned) images = jpegs.map((im, i) => ({ ...im, label: `scanned page picture ${i + 1}`, scan: true }));
  else images = jpegs.filter((im) => im.data.length >= 40 * 1024).sort((a, b) => b.data.length - a.data.length).map((im, i) => ({ ...im, label: `figure ${i + 1} in the PDF` }));
  const text = pages.map((p, i) => (p ? `--- Page ${i + 1} ---\n${p}` : '')).filter(Boolean).join('\n\n');
  return { text, images, scanned, modern };
}

// Text decoding that copes with UTF-16 and old Windows encodings.
function decodeText(buf) {
  if (buf[0] === 0xff && buf[1] === 0xfe) return buf.subarray(2).toString('utf16le');
  if (buf[0] === 0xfe && buf[1] === 0xff) return Buffer.from(buf.subarray(2)).swap16().toString('utf16le');
  try { return new TextDecoder('utf-8', { fatal: true }).decode(buf).replace(/^﻿/, ''); }
  catch (_) { try { return new TextDecoder('windows-1252').decode(buf); } catch (e) { return buf.toString('latin1'); } }
}

// Everything the server needs from a file: its text, plus pictures worth sending to the vision model.
async function extractAll(kind, buf) {
  if (kind === 'pdf') return fromPdfAll(buf);
  if (kind === 'docx') { const z = readZip(buf); return { text: fromDocx(buf), images: officeImages(z, 'word/media/', () => 'picture in the document').sort((a, b) => b.data.length - a.data.length) }; }
  if (kind === 'pptx') { const z = readZip(buf); return { text: fromPptx(buf), images: pptxImages(z).sort((a, b) => (a.order < b.order ? -1 : 1)) }; }
  return { text: extractText(kind, buf), images: [] };
}

// kind is one of the text-based kinds from kindOf(). Returns plain text.
function extractText(kind, buf) {
  switch (kind) {
    case 'pdf': return fromPdf(buf);
    case 'docx': return fromDocx(buf);
    case 'pptx': return fromPptx(buf);
    case 'xlsx': return fromXlsx(buf);
    case 'odf': return fromOdf(buf);
    case 'html': return fromHtml(buf);
    case 'rtf': return fromRtf(buf);
    case 'text': return decodeText(buf);
    default: throw new Error('is not a supported file type.');
  }
}

module.exports = { kindOf, imageTypeOf, extractText, extractAll };
