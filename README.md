# Kiarii Tools Suite

**Private, browser-only document tools.**  
PDF → Word • Word → PDF • OCR  
No servers. No uploads. No accounts. No watermarks.

Your files never leave your device.

---

## Live Demo

Host these files on **GitHub Pages**, Netlify, or any static host.

```
https://yourusername.github.io/kiarii-suite/
```

---

## Features (v1)

| Tool           | Description                                      |
|----------------|--------------------------------------------------|
| **PDF → Word** | Extracts text and builds a clean `.docx`         |
| **Word → PDF** | Converts `.docx` via HTML intermediate           |
| **OCR**        | Recognizes text from images & scanned PDFs       |

All processing happens **inside the browser** using:

- PDF.js
- Tesseract.js
- docx
- Mammoth.js
- html2pdf.js

---

## Privacy

- There is **no backend** and **no upload endpoint**.
- Files are read with the File API and processed in memory.
- OCR language models are downloaded once from a public CDN and then cached by your browser.
- No analytics, no tracking scripts, no cookies for tracking.

You can inspect every line of code.

---

## File Structure (ultra simple)

```
kiarii-suite/
├── index.html
├── styles.css
├── app.js
├── bg-portrait.jpg
└── README.md
```

Just push these 5 files to a GitHub repository and enable **GitHub Pages**  
(Settings → Pages → Deploy from branch → main → `/` root).

No branches needed. No build step. No folders.

---

## Local Preview

```bash
python -m http.server 8000
```

Then open `http://localhost:8000`

---

## Roadmap

- [x] PDF → Word
- [x] Word → PDF
- [x] OCR (images + PDF)
- [ ] Merge / Split PDF
- [ ] Compress PDF
- [ ] Batch mode
- [ ] PWA (installable + offline)

---

## License

MIT — free to use, modify, and host anywhere.

**Kiarii Tools Suite**
