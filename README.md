# SmoothStudy

Upload notes (PDF, Word, PowerPoint and more) or type them, get a condensed summary with diagrams, then quiz yourself.

## Run it

Requires Node 22.13 or newer for best PDF reading (Node 18 still runs it, with a simpler PDF reader). There is nothing to install.

1. Open `.env` and paste your Groq API key after `GROQ_API_KEY=` (it starts with `gsk_`).
2. Run `node download-fonts.js` once. It needs internet and saves every font into `fonts/` (see "Fonts" below).
3. Run `node server.js`
4. Open http://localhost:3000

The AI is Groq running `openai/gpt-oss-120b`. Optional settings in `.env`: `AI_MODEL`, `VISION_MODEL`, `VISION_MAX`, `REASONING_EFFORT` (low, medium or high; default low), `PORT`.

The key is only read by `server.js`. It is never sent to the browser, so do not paste it into `app.js` or `index.html`.

## What the AI can read

The main model (`openai/gpt-oss-120b`) reads text only, so the server first turns every file into text, and a second Groq model reads pictures.

- Text: PDF, Word (.docx), PowerPoint (.pptx, including tables and speaker notes), Excel (.xlsx), OpenDocument (.odt/.odp/.ods), TXT, Markdown, CSV, JSON, HTML, RTF and similar.
- Pictures: JPG, PNG, GIF and WebP are shrunk in the browser, then read by a vision model (text, handwriting, diagrams, charts). Scanned PDFs are read the same way, page by page. Figures inside PDFs, Word and PowerPoint files are read too.
- Old .doc/.ppt/.xls files must be re-saved in the newer format.

PDFs are read with pdf.js (bundled in `vendor/pdfjs/`, nothing to install), which handles Word, LibreOffice, LaTeX and browser-made PDFs. It needs Node 22.13 or newer. On older Node the site falls back to a simpler built-in reader that can miss words, so please use a recent Node.

Picture reading uses `VISION_MODEL` (default `qwen/qwen3.8-27b`) and reads up to `VISION_MAX` pictures per summary (default 6). Set `VISION_MAX=0` to switch it off. If Groq renames or retires that model, set `VISION_MODEL` in `.env` to another Groq vision model.

Very large uploads can exceed your Groq plan's per-minute token limit. The error from Groq is shown on the page if that happens.

## Demo mode (no API key needed)

Run `node server.js --demo` (or set `DEMO_MODE=1`) to skip the AI. Uploads are still read for real, but the summary and quiz come from built-in sample data about photosynthesis.

## Fonts

The fonts are Outfit (all main text), Lexend (only the theme buttons, scroll buttons and the top-left brand button), JetBrains Mono, Pinyon Script, Emilys Candy and Google Sans. Nothing is loaded from the internet while you use the site. `node download-fonts.js` saves them once into `fonts/` and writes `fonts.css`. Until you run it, the page falls back to system fonts.

The PDF download embeds Google Sans from `fonts/GoogleSans-Regular.ttf` and `fonts/GoogleSans-Bold.ttf`. If those files are missing, the PDF uses Helvetica instead.

## PDF download

The summary page has a download button that saves the notes as a PDF (text, diagrams, technical box and key terms). It is generated in the browser by `pdf.js`, with no extra libraries. If the notes contain characters the PDF fonts cannot show (for example Cyrillic or Chinese), the button opens the browser's print dialog instead, where you choose "Save as PDF". Diagrams are pictures drawn by Mermaid, so their labels use a standard sans-serif rather than Google Sans.

Diagrams use Mermaid, which is loaded from a CDN, so diagrams need internet access.
