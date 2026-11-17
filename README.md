# SmoothStudy

Upload notes (PDF, Word, PowerPoint and more) or type them, get a condensed summary with diagrams, then quiz yourself.

## Run it

Requires Node 18 or newer. There is nothing to install.

1. Copy `.env.example` to `.env` and paste your Groq API key after `GROQ_API_KEY=` (it starts with `gsk_`). Only a Groq key works, and the variable must be named `GROQ_API_KEY`. If your `.env` still says `ANTHROPIC_API_KEY=`, rename that to `GROQ_API_KEY=` and keep the same value.
2. Run `node download-fonts.js` once. It needs internet and saves every font into `fonts/` (see "Fonts" below).
3. Run `node server.js`
4. Open http://localhost:3000

The AI is Groq running `openai/gpt-oss-120b`. Optional settings in `.env`: `AI_MODEL`, `REASONING_EFFORT` (low, medium or high; default low), `PORT`.

The key is only read by `server.js`. It is never sent to the browser, so do not paste it into `app.js` or `index.html`.

## What the AI can read

The model reads text only. The server turns files into text first:

- Works: PDF (text PDFs), Word (.docx), PowerPoint (.pptx), Excel (.xlsx), OpenDocument (.odt/.odp/.ods), TXT, Markdown, CSV, JSON, HTML, RTF.
- Does not work: pictures (JPG, PNG, GIF, WebP) and scanned PDFs, because they hold no text. Paste the text instead.
- Old .doc/.ppt/.xls files must be re-saved in the newer format.

PDF text is read by a small built-in reader. Ordinary PDFs work. PDFs that are scans, or that use unusual font encodings, give a clear message asking you to paste the text.

Very large uploads can exceed your Groq plan's per-minute token limit. The error from Groq is shown on the page if that happens.

## Demo mode (no API key needed)

Run `node server.js --demo` (or set `DEMO_MODE=1`) to skip the AI. Uploads are still read for real, but the summary and quiz come from built-in sample data about photosynthesis.

## Fonts

The fonts are Lexend, JetBrains Mono, Pinyon Script, Emilys Candy and Google Sans. Nothing is loaded from the internet while you use the site. `node download-fonts.js` saves them once into `fonts/` and writes `fonts.css`. Until you run it, the page falls back to system fonts.

The PDF download embeds Google Sans from `fonts/GoogleSans-Regular.ttf` and `fonts/GoogleSans-Bold.ttf`. If those files are missing, the PDF uses Helvetica instead.

## PDF download

The summary page has a download button that saves the notes as a PDF (text, diagrams, technical box and key terms). It is generated in the browser by `pdf.js`, with no extra libraries. If the notes contain characters the PDF fonts cannot show (for example Cyrillic or Chinese), the button opens the browser's print dialog instead, where you choose "Save as PDF". Diagrams are pictures drawn by Mermaid, so their labels use a standard sans-serif rather than Google Sans.

Diagrams use Mermaid, which is loaded from a CDN, so diagrams need internet access.
