// SmoothStudy server: no npm dependencies (Node 18+).
// The AI key stays here on the server and is never sent to the browser. The AI is Groq running openai/gpt-oss-120b.
const http = require('http');
const fs = require('fs');
const path = require('path');
const { kindOf, extractText } = require('./extract');
const demo = require('./demo');

// Minimal .env loader
try {
  fs.readFileSync(path.join(__dirname, '.env'), 'utf8').split(/\r?\n/).forEach((line) => {
    const m = line.match(/^\s*([A-Z_][A-Z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  });
} catch (_) {}

// Demo mode: no AI calls, built-in sample summary and quiz. Start with `node server.js --demo` or DEMO_MODE=1.
const DEMO = process.argv.includes('--demo') || /^(1|true|yes|on)$/i.test(process.env.DEMO_MODE || '');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const KEY = process.env.GROQ_API_KEY; // Groq key only (starts with gsk_)
const OLD_NAME = !KEY && !!process.env.ANTHROPIC_API_KEY; // an old .env that still uses the previous variable name
const MODEL = process.env.AI_MODEL || 'openai/gpt-oss-120b';
const API_URL = process.env.GROQ_API_URL || 'https://api.groq.com/openai/v1/chat/completions';
const EFFORT = process.env.REASONING_EFFORT || 'low'; // low | medium | high
const PORT = process.env.PORT || 3000;
const MAX_BODY = 360 * 1024 * 1024; // 5 files x 50 MB, base64-encoded

const STATIC = {
  '/': ['index.html', 'text/html; charset=utf-8'],
  '/index.html': ['index.html', 'text/html; charset=utf-8'],
  '/styles.css': ['styles.css', 'text/css; charset=utf-8'],
  '/app.js': ['app.js', 'text/javascript; charset=utf-8'],
  '/pdf.js': ['pdf.js', 'text/javascript; charset=utf-8'],
  '/fonts.css': ['fonts.css', 'text/css; charset=utf-8'],
};
const FONT_TYPES = { '.woff2': 'font/woff2', '.ttf': 'font/ttf' };

const SUMMARY_SYSTEM = `You are SmoothStudy, an expert study-notes condenser. You receive a student's notes or textbook material (text extracted from documents).
Produce a summary that keeps ALL the important information and every key concept, but scaled down (roughly 30-40% of the original length) so the student can finish studying faster. Do not invent facts that are not supported by the material.

Return ONLY valid JSON (no markdown fences, no commentary) with this exact shape:
{
  "title": "short title for the material",
  "overview": "2-3 sentence big-picture overview",
  "sections": [
    {
      "heading": "section heading",
      "points": ["concise study point, may use **bold** for key terms and \`code\` for formulas or technical tokens", "..."],
      "diagram": { "title": "diagram title", "mermaid": "mermaid source" } 
    }
  ],
  "technical": ["formulas, equations, code or exact technical facts worth memorising, one per entry"],
  "keyTerms": [ { "term": "term", "definition": "short definition" } ]
}

Rules:
- "diagram" is optional per section; include it (as a flowchart, sequence, mindmap or timeline) only where it genuinely helps understanding, at most 4 diagrams in total. Otherwise use null.
- Mermaid rules: valid Mermaid syntax only. Start with e.g. "flowchart TD". Wrap every node label in double quotes, e.g. A["Label text"]. Avoid parentheses, quotes, colons and special characters inside labels. Keep diagrams small (under 12 nodes).
- "technical" and "keyTerms" may be empty arrays.
- Write in clear student-friendly language in the same language as the source material.`;

const QUIZ_SYSTEM = `You are SmoothStudy's quiz generator. Using ONLY the study summary provided, write multiple choice questions that test understanding.
Return ONLY valid JSON (no markdown fences, no commentary) with this exact shape:
{ "questions": [ { "question": "text", "options": ["A text", "B text", "C text", "D text"], "answer": 0, "explanation": "why the answer is right, one or two sentences", "difficulty": "Easy" | "Medium" | "Hard" } ] }
Rules:
- Exactly 4 options per question, exactly one correct; "answer" is the zero-based index of the correct option.
- Vary the position of the correct answer. Distractors must be plausible, not silly.
- Easy = direct recall of a fact or term. Medium = understanding and applying a concept. Hard = multi-step reasoning, comparing concepts, or edge cases.
- If the requested difficulty is Mixed, use a roughly even spread of Easy, Medium and Hard. Otherwise every question uses the requested difficulty.
- Do not repeat questions. Cover the whole summary rather than one section.`;

function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', (c) => {
      size += c.length;
      if (size > MAX_BODY) { reject(new Error('Upload too large')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

function send(res, status, obj) {
  res.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
  res.end(JSON.stringify(obj));
}

function extractJSON(text) {
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start === -1 || end === -1) throw new Error('The AI did not return structured data. Please try again.');
  return JSON.parse(text.slice(start, end + 1));
}

async function callAI(system, userText, maxTokens) {
  if (OLD_NAME) throw new Error('.env still uses ANTHROPIC_API_KEY. Rename it to GROQ_API_KEY (keep your Groq key as the value), then restart the server.');
  if (!KEY) throw new Error('Server has no AI key. Put your Groq key in .env as GROQ_API_KEY. See README.md.');
  if (/^sk-ant/i.test(KEY)) throw new Error('.env still holds an Anthropic key. Put your Groq key (it starts with gsk_) in .env as GROQ_API_KEY.');
  const body = {
    model: MODEL,
    messages: [{ role: 'system', content: system }, { role: 'user', content: userText }],
    temperature: 0.3,
    max_completion_tokens: maxTokens,
    reasoning_effort: EFFORT,
    response_format: { type: 'json_object' },
  };
  const post = () => fetch(API_URL, {
    method: 'POST',
    headers: { 'content-type': 'application/json', authorization: 'Bearer ' + KEY },
    body: JSON.stringify(body),
  });
  let r = await post();
  let data = await r.json().catch(() => ({}));
  if (r.status === 400 && /response_format|json_object|reasoning_effort/i.test((data.error && data.error.message) || '')) {
    delete body.response_format; delete body.reasoning_effort; // an older model setting: retry plainly
    r = await post();
    data = await r.json().catch(() => ({}));
  }
  if (!r.ok) {
    const msg = (data.error && data.error.message) || `AI request failed (${r.status})`;
    if (r.status === 401) throw new Error('The AI key was rejected. Check GROQ_API_KEY in .env.');
    if (r.status === 413 || r.status === 429) throw new Error('The AI service says the request is too big or too frequent for your plan: ' + msg);
    throw new Error(msg);
  }
  const text = (((data.choices || [])[0] || {}).message || {}).content;
  if (!text) throw new Error('The AI returned an empty answer. Please try again.');
  return text;
}

const TEXT_CAP_PER_FILE = 200000; // characters; the model reads about 131k tokens in total
const TEXT_CAP_TOTAL = 300000;

async function handleSummarize(body) {
  const parts = [];
  let textTotal = 0, fileCount = 0;
  for (const f of body.files || []) {
    if (!f || typeof f.data !== 'string') continue;
    const name = String(f.name || 'file');
    const kind = kindOf(name);
    if (!kind) throw new Error(`${name}: this file type isn't supported.`);
    if (kind === 'legacy') throw new Error(`${name}: old .doc, .ppt and .xls files can't be read. Open it and save as .docx, .pptx or .xlsx, then upload again.`);
    if (kind === 'image') throw new Error(`${name}: the AI model reads text only, so pictures can't be read. Type or paste the text instead, or upload a PDF, Word or PowerPoint file.`);
    fileCount++;
    const buf = Buffer.from(f.data, 'base64');
    let text;
    try { text = extractText(kind, buf).trim(); }
    catch (e) { throw new Error(`${name}: ${e.message}`); }
    if (!text) throw new Error(`${name}: no readable text found (a deck made only of pictures, for example). Paste the text instead.`);
    text = text.slice(0, Math.min(TEXT_CAP_PER_FILE, Math.max(0, TEXT_CAP_TOTAL - textTotal)));
    textTotal += text.length;
    if (text) parts.push(`Content of "${name}":\n\n${text}`);
  }
  const typed = (body.text || '').toString().slice(0, 100000);
  if (typed.trim()) parts.push(`Student's typed notes:\n\n${typed}`);
  if (!parts.length) throw new Error('Nothing to summarise. Add notes or upload a file.');
  if (DEMO) {
    await sleep(900); // so the loading spinner is visible
    return demo.summary({ files: fileCount, chars: textTotal + typed.length });
  }
  return extractJSON(await callAI(SUMMARY_SYSTEM, parts.join('\n\n') + '\n\nCreate the study summary JSON now.', 8000));
}

async function handleQuiz(body) {
  const count = [5, 10, 20].includes(body.count) ? body.count : 10;
  const difficulty = ['Easy', 'Medium', 'Hard', 'Mixed'].includes(body.difficulty) ? body.difficulty : 'Mixed';
  if (DEMO) {
    await sleep(700);
    return demo.quiz(count, difficulty);
  }
  const summary = JSON.stringify(body.summary || {}).slice(0, 150000);
  const prompt = `Study summary (JSON):\n${summary}\n\nWrite exactly ${count} multiple choice questions. Difficulty: ${difficulty}.`;
  const out = extractJSON(await callAI(QUIZ_SYSTEM, prompt, 8000));
  const qs = (out.questions || []).filter(
    (q) => q && q.question && Array.isArray(q.options) && q.options.length >= 2 && Number.isInteger(q.answer) && q.options[q.answer] !== undefined
  );
  if (!qs.length) throw new Error('Could not build a quiz from that summary. Please try again.');
  return { questions: qs.slice(0, count) };
}

http
  .createServer(async (req, res) => {
    try {
      const url = req.url.split('?')[0];
      if (req.method === 'GET' && (STATIC[url] || /^\/fonts\/[\w.-]+\.(woff2|ttf)$/.test(url))) {
        const [file, type] = STATIC[url] || ['fonts/' + url.slice(7), FONT_TYPES[path.extname(url)]];
        let data;
        try { data = fs.readFileSync(path.join(__dirname, file)); }
        catch (_) { res.writeHead(404); return res.end('Not found'); }
        res.writeHead(200, { 'Content-Type': type, 'Cache-Control': file.startsWith('fonts') ? 'public, max-age=86400' : 'no-cache' });
        return res.end(data);
      }
      if (req.method === 'POST' && (url === '/api/summarize' || url === '/api/quiz')) {
        const body = JSON.parse((await readBody(req)) || '{}');
        const result = url === '/api/summarize' ? await handleSummarize(body) : await handleQuiz(body);
        return send(res, 200, result);
      }
      res.writeHead(404);
      res.end('Not found');
    } catch (e) {
      send(res, 500, { error: e.message || 'Something went wrong' });
    }
  })
  .listen(PORT, () => {
    console.log(`SmoothStudy running at http://localhost:${PORT}`);
    if (DEMO) console.log('DEMO MODE: no AI calls are made; summaries and quizzes are built-in sample data.');
    else if (OLD_NAME) console.log('Warning: .env uses the old name ANTHROPIC_API_KEY. Rename it to GROQ_API_KEY (same Groq key as the value).');
    else if (!KEY) console.log('Warning: no AI key. Copy .env.example to .env and add your Groq key, or run with --demo.');
    else if (/^sk-ant/i.test(KEY)) console.log('Warning: .env holds an Anthropic key. Put your Groq key (gsk_...) in GROQ_API_KEY.');
    if (!fs.existsSync(path.join(__dirname, 'fonts.css'))) console.log('Fonts: not downloaded yet, so system fonts are used. Run `node download-fonts.js` once (needs internet) to save them locally.');
  });
