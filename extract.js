// Text extraction for common document formats. No npm dependencies:
// Office files are zip archives, read here with Node's built-in zlib.
const zlib = require('zlib');

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

function fromDocx(buf) {
  const z = readZip(buf);
  const doc = z.text('word/document.xml');
  if (!doc) throw new Error('is not a Word document.');
  const parts = [paragraphs(doc, 'w:p', 'w:t').join('\n')];
  for (const n of z.names.filter((n) => /^word\/(footnotes|endnotes)\.xml$/.test(n)).sort()) {
    const t = paragraphs(z.text(n), 'w:p', 'w:t').join('\n');
    if (t) parts.push('Notes:\n' + t);
  }
  return parts.join('\n\n');
}

function fromPptx(buf) {
  const z = readZip(buf);
  const slides = z.names
    .map((n) => ({ n, m: n.match(/^ppt\/slides\/slide(\d+)\.xml$/) }))
    .filter((x) => x.m)
    .sort((a, b) => +a.m[1] - +b.m[1]);
  if (!slides.length) throw new Error('is not a PowerPoint file.');
  const out = [];
  slides.forEach((s, i) => {
    const lines = paragraphs(z.text(s.n), 'a:p', 'a:t');
    let notes = [];
    const rels = z.text(`ppt/slides/_rels/slide${s.m[1]}.xml.rels`) || '';
    const nm = rels.match(/notesSlides\/(notesSlide\d+\.xml)/);
    if (nm) {
      const nx = z.text('ppt/notesSlides/' + nm[1]);
      // Notes pages repeat the slide number as a field; keep only real text.
      if (nx) notes = paragraphs(nx, 'a:p', 'a:t').filter((l) => !/^\d+$/.test(l.trim()));
    }
    if (!lines.length && !notes.length) return;
    out.push(`--- Slide ${i + 1} ---\n${lines.join('\n')}` + (notes.length ? `\nSpeaker notes: ${notes.join(' ')}` : ''));
  });
  return out.join('\n\n');
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
    for (const row of xml.match(/<row\b[\s\S]*?<\/row>/g) || []) {
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
    .replace(/\\'([0-9a-fA-F]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/\\[a-zA-Z]+-?\d* ?/g, '')
    .replace(/[{}]/g, '')
    .replace(/\n{3,}/g, '\n\n');
}


// ---- PDF text (best effort, no dependencies). Handles Flate streams, object streams and ToUnicode maps.
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
    let dict, data = null;
    if (sIdx >= 0 && sIdx < eIdx) {
      dict = latin.slice(start, sIdx);
      let b = sIdx + 6;
      if (latin[b] === '\r') b++;
      if (latin[b] === '\n') b++;
      let e = latin.indexOf('endstream', b);
      if (e < 0) e = eIdx;
      let raw = buf.subarray(b, e);
      if (/\/FlateDecode/.test(dict)) raw = inflate(raw);
      else if (/\/Filter/.test(dict)) raw = null;
      data = raw;
      re.lastIndex = Math.max(re.lastIndex, e);
    } else dict = latin.slice(start, eIdx);
    objs[m[1]] = { dict, data };
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
        return raw;
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
        else if (m[6] !== undefined && Math.abs(parseFloat(m[6])) > 2) text += '\n';
      }
    }
    text = text.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    if (text) out.push(text);
  }
  const joined = out.map((t, i) => `--- Page ${i + 1} ---\n${t}`).join('\n\n');
  if (joined.replace(/--- Page \d+ ---|\s/g, '').length < 40) throw new Error('has no text the server can read (it may be a scan, or use an unusual encoding). Paste the text instead, or export it as a .docx.');
  return joined;
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
    case 'text': return buf.toString('utf8').replace(/^﻿/, '');
    default: throw new Error('is not a supported file type.');
  }
}

module.exports = { kindOf, imageTypeOf, extractText };
