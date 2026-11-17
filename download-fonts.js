// Downloads the website's fonts once, so the site never needs the internet for fonts again.
// Usage:  node download-fonts.js        (needs Node 18+ and an internet connection, once)
//
// It saves the files into ./fonts and writes ./fonts.css, which the page already links to.
// Fonts: Lexend, JetBrains Mono, Pinyon Script, Emilys Candy, Google Sans (all open-licence Google Fonts).
// It also saves GoogleSans-Regular.ttf and GoogleSans-Bold.ttf, which the PDF export embeds.
const fs = require('fs');
const path = require('path');

const CSS_URL = process.env.FONTS_CSS_URL || 'https://fonts.googleapis.com/css2'; // overridable for testing
const WEB_UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';
const TTF_UA = 'curl/8.5.0'; // an unknown browser gets plain TrueType files, which a PDF can embed
const OUT = path.join(__dirname, 'fonts');

const FAMILIES = [
  { name: 'Lexend', weights: '300;400;500;600;700' },
  { name: 'JetBrains Mono', weights: '400;500' },
  { name: 'Pinyon Script', weights: null },
  { name: 'Emilys Candy', weights: null },
  { name: 'Google Sans', weights: '400;500;600;700' },
];
const PDF_FONTS = [
  { family: 'Google Sans', weight: '400', file: 'GoogleSans-Regular.ttf' },
  { family: 'Google Sans', weight: '700', file: 'GoogleSans-Bold.ttf' },
];

const slug = (s) => s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
const cssUrl = (family, weights) => `${CSS_URL}?family=${encodeURIComponent(family).replace(/%20/g, '+')}${weights ? ':wght@' + weights : ''}&display=swap`;

async function getText(url, ua) {
  const r = await fetch(url, { headers: { 'User-Agent': ua } });
  if (!r.ok) throw new Error(`${r.status} ${r.statusText} for ${url}`);
  return r.text();
}
async function getBytes(url) {
  const r = await fetch(url);
  if (!r.ok) throw new Error(`${r.status} ${r.statusText} for ${url}`);
  return Buffer.from(await r.arrayBuffer());
}

// Reads every @font-face block (with the "/* latin */" comment Google puts before it).
function parseFaces(css) {
  const faces = [];
  const re = /(?:\/\*\s*([\w-]+)\s*\*\/\s*)?@font-face\s*\{([^}]*)\}/g;
  let m;
  while ((m = re.exec(css))) {
    const body = m[2];
    const get = (prop) => ((body.match(new RegExp(prop + '\\s*:\\s*([^;]+);')) || [])[1] || '').trim();
    const src = body.match(/url\(\s*['"]?([^'")]+)['"]?\s*\)/);
    if (!src) continue;
    faces.push({
      subset: m[1] || 'all',
      family: get('font-family').replace(/['"]/g, ''),
      style: get('font-style') || 'normal',
      weight: get('font-weight') || '400',
      range: get('unicode-range'),
      url: src[1],
    });
  }
  return faces;
}

function hasTable(buf, tag) {
  const n = buf.readUInt16BE(4);
  for (let i = 0; i < n; i++) if (buf.toString('latin1', 12 + i * 16, 16 + i * 16) === tag) return true;
  return false;
}

async function main() {
  fs.mkdirSync(OUT, { recursive: true });
  const cssBlocks = [];
  const failures = [];
  const byUrl = new Map(); // one download per distinct file (variable fonts share a file across weights)

  for (const fam of FAMILIES) {
    try {
      const faces = parseFaces(await getText(cssUrl(fam.name, fam.weights), WEB_UA));
      if (!faces.length) throw new Error('the font service returned no font files');
      let count = 0;
      for (const f of faces) {
        let file = byUrl.get(f.url);
        if (!file) {
          file = `${slug(fam.name)}-${f.subset}-${f.weight.replace(/\s+/g, '-')}.woff2`;
          fs.writeFileSync(path.join(OUT, file), await getBytes(f.url));
          byUrl.set(f.url, file);
          count++;
        }
        cssBlocks.push(
          `@font-face {\n  font-family: '${fam.name}';\n  font-style: ${f.style};\n  font-weight: ${f.weight};\n  font-display: swap;\n  src: url('/fonts/${file}') format('woff2');` +
          (f.range ? `\n  unicode-range: ${f.range};` : '') + '\n}'
        );
      }
      console.log(`ok    ${fam.name} (${faces.length} styles, ${count} files)`);
    } catch (e) {
      failures.push(fam.name);
      console.log(`FAIL  ${fam.name}: ${e.message}`);
    }
  }

  for (const p of PDF_FONTS) {
    try {
      const faces = parseFaces(await getText(cssUrl(p.family, p.weight), TTF_UA));
      const face = faces.find((f) => /\.ttf(\?|$)/i.test(f.url)) || faces[0];
      if (!face) throw new Error('the font service returned no font file');
      const bytes = await getBytes(face.url);
      const magic = bytes.readUInt32BE(0);
      if (magic !== 0x00010000 && bytes.toString('latin1', 0, 4) !== 'true') throw new Error('did not receive a TrueType file');
      fs.writeFileSync(path.join(OUT, p.file), bytes);
      console.log(`ok    ${p.file}${hasTable(bytes, 'fvar') ? '  (note: variable font file, bold may look regular in the PDF)' : ''}`);
    } catch (e) {
      failures.push(p.file);
      console.log(`FAIL  ${p.file}: ${e.message} (the PDF will use its built-in font instead)`);
    }
  }

  if (cssBlocks.length) {
    fs.writeFileSync(path.join(__dirname, 'fonts.css'), '/* Generated by download-fonts.js. Local copies of the fonts, no internet needed. */\n' + cssBlocks.join('\n') + '\n');
    console.log(`\nWrote fonts.css and ${fs.readdirSync(OUT).length} files in fonts/.`);
  }
  if (failures.length) { console.log('\nSome downloads failed: ' + failures.join(', ') + '. Check your connection and run it again.'); process.exit(1); }
  console.log('Done. Restart the server (node server.js) and reload the page.');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
