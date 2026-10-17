/**
 * Kiarii Tools Suite — Browser-only document tools
 * Everything runs in the user's browser. No uploads.
 */

// ---------- PDF.js worker setup ----------
import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';

// ---------- DOM helpers ----------
const $ = (sel, ctx = document) => ctx.querySelector(sel);
const $$ = (sel, ctx = document) => [...ctx.querySelectorAll(sel)];

// ---------- Particles (lightweight motion) ----------
function createParticles() {
  const container = $('#particles');
  if (!container) return;
  const count = window.innerWidth < 768 ? 12 : 22;
  for (let i = 0; i < count; i++) {
    const p = document.createElement('div');
    p.className = 'particle';
    p.style.left = Math.random() * 100 + '%';
    p.style.animationDuration = 12 + Math.random() * 18 + 's';
    p.style.animationDelay = Math.random() * 10 + 's';
    p.style.width = p.style.height = 2 + Math.random() * 3 + 'px';
    container.appendChild(p);
  }
}

// ---------- Header scroll ----------
function initHeader() {
  const header = $('.header');
  const onScroll = () => {
    header.classList.toggle('scrolled', window.scrollY > 40);
  };
  window.addEventListener('scroll', onScroll, { passive: true });
  onScroll();
}

// ---------- Tool tabs ----------
function initTabs() {
  const tabs = $$('.tool-tab');
  const panels = $$('.tool-panel');

  tabs.forEach(tab => {
    tab.addEventListener('click', () => {
      const tool = tab.dataset.tool;
      tabs.forEach(t => t.classList.remove('active'));
      panels.forEach(p => p.classList.remove('active'));
      tab.classList.add('active');
      $(`#panel-${tool}`).classList.add('active');
    });
  });
}

// ---------- Dropzone logic ----------
function setupDropzone(dropId, inputId, onFile) {
  const drop = $(dropId);
  const input = $(inputId);
  if (!drop || !input) return;

  const openPicker = () => input.click();

  drop.addEventListener('click', openPicker);
  drop.addEventListener('dragover', e => {
    e.preventDefault();
    drop.classList.add('dragover');
  });
  drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
  drop.addEventListener('drop', e => {
    e.preventDefault();
    drop.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) onFile(file);
  });

  input.addEventListener('change', () => {
    if (input.files[0]) onFile(input.files[0]);
  });
}

// ---------- Progress helpers ----------
function showProgress(id, text = 'Working…') {
  const wrap = $(`#progress-${id}`);
  if (!wrap) return;
  wrap.hidden = false;
  $('.progress-text', wrap).textContent = text;
  $('.progress-fill', wrap).style.width = '8%';
}

function updateProgress(id, pct, text) {
  const wrap = $(`#progress-${id}`);
  if (!wrap) return;
  $('.progress-fill', wrap).style.width = Math.min(100, pct) + '%';
  if (text) $('.progress-text', wrap).textContent = text;
}

function hideProgress(id) {
  const wrap = $(`#progress-${id}`);
  if (wrap) wrap.hidden = true;
}

function showResult(id, html) {
  const el = $(`#result-${id}`);
  if (!el) return;
  el.hidden = false;
  el.innerHTML = html;
}

function clearResult(id) {
  const el = $(`#result-${id}`);
  if (el) {
    el.hidden = true;
    el.innerHTML = '';
  }
}

// ---------- Download helper ----------
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

// ========== PDF → Word ==========
async function pdfToWord(file) {
  const id = 'pdf2word';
  clearResult(id);
  showProgress(id, 'Reading PDF…');

  try {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    const total = pdf.numPages;
    let fullText = '';

    for (let i = 1; i <= total; i++) {
      updateProgress(id, (i / total) * 70, `Extracting page ${i}/${total}…`);
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const strings = content.items.map(item => item.str).join(' ');
      fullText += strings + '\n\n';
    }

    // If almost no text, it is likely a scanned PDF → suggest OCR
    if (fullText.trim().length < 40) {
      hideProgress(id);
      showResult(id, `
        <p>This PDF appears to be scanned (little selectable text found).</p>
        <p style="margin-top:0.5rem;color:var(--text-muted)">Switch to the <strong>OCR</strong> tool for better results.</p>
        <div class="result-actions">
          <button class="btn-secondary" onclick="document.querySelector('[data-tool=ocr]').click()">Go to OCR</button>
        </div>
      `);
      return;
    }

    updateProgress(id, 85, 'Building Word document…');

    // Use docx library
    const { Document, Packer, Paragraph, TextRun } = window.docx;
    const paragraphs = fullText
      .split(/\n+/)
      .filter(line => line.trim())
      .map(line => new Paragraph({
        children: [new TextRun({ text: line, size: 22 })]
      }));

    const doc = new Document({
      sections: [{
        properties: {},
        children: paragraphs.length ? paragraphs : [
          new Paragraph({ children: [new TextRun('No text extracted.')] })
        ]
      }]
    });

    const blob = await Packer.toBlob(doc);
    updateProgress(id, 100, 'Done!');
    hideProgress(id);

    const name = file.name.replace(/\.pdf$/i, '') + '.docx';
    showResult(id, `
      <p><strong>Conversion complete</strong> — ${total} page(s) processed.</p>
      <p style="color:var(--text-muted);font-size:0.9rem;margin-top:0.35rem">Layout is simplified (text extraction). Complex designs may need manual cleanup.</p>
      <div class="result-actions">
        <button class="btn-download" id="dl-pdf2word">Download .docx</button>
        <button class="btn-secondary" onclick="location.reload()">Convert another</button>
      </div>
    `);

    $('#dl-pdf2word').addEventListener('click', () => downloadBlob(blob, name));
  } catch (err) {
    console.error(err);
    hideProgress(id);
    showResult(id, `<p style="color:#f87171">Error: ${err.message || 'Could not process PDF'}</p>`);
  }
}

// ========== Word → PDF ==========
async function wordToPdf(file) {
  const id = 'word2pdf';
  clearResult(id);
  showProgress(id, 'Reading Word file…');

  try {
    const arrayBuffer = await file.arrayBuffer();
    updateProgress(id, 30, 'Converting to HTML…');

    const result = await mammoth.convertToHtml({ arrayBuffer });
    const html = result.value;

    updateProgress(id, 60, 'Generating PDF…');

    // Temporary element for html2pdf
    const container = document.createElement('div');
    container.innerHTML = html;
    container.style.cssText = 'padding:40px;font-family:Georgia,serif;font-size:14px;line-height:1.6;color:#111;';
    document.body.appendChild(container);

    const opt = {
      margin: 12,
      filename: file.name.replace(/\.docx?$/i, '') + '.pdf',
      image: { type: 'jpeg', quality: 0.95 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    const worker = html2pdf().set(opt).from(container);
    const blob = await worker.outputPdf('blob');

    document.body.removeChild(container);
    updateProgress(id, 100, 'Done!');
    hideProgress(id);

    showResult(id, `
      <p><strong>Conversion complete</strong></p>
      <p style="color:var(--text-muted);font-size:0.9rem;margin-top:0.35rem">Converted via HTML intermediate. Complex formatting may differ slightly from original.</p>
      <div class="result-actions">
        <button class="btn-download" id="dl-word2pdf">Download .pdf</button>
        <button class="btn-secondary" onclick="location.reload()">Convert another</button>
      </div>
    `);

    $('#dl-word2pdf').addEventListener('click', () => downloadBlob(blob, opt.filename));
  } catch (err) {
    console.error(err);
    hideProgress(id);
    showResult(id, `<p style="color:#f87171">Error: ${err.message || 'Could not process Word file'}</p>`);
  }
}

// ========== OCR ==========
async function runOcr(file) {
  const id = 'ocr';
  clearResult(id);
  showProgress(id, 'Preparing…');

  const lang = $('#ocr-lang')?.value || 'eng';
  $('#ocr-options').hidden = false;

  try {
    let images = [];

    if (file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf')) {
      updateProgress(id, 10, 'Rendering PDF pages…');
      const arrayBuffer = await file.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const total = Math.min(pdf.numPages, 15); // safety limit

      for (let i = 1; i <= total; i++) {
        updateProgress(id, 10 + (i / total) * 30, `Rendering page ${i}/${total}…`);
        const page = await pdf.getPage(i);
        const viewport = page.getViewport({ scale: 2 });
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
        images.push(canvas);
      }
    } else {
      // Image
      images = [file];
    }

    updateProgress(id, 45, 'Loading OCR engine…');
    const worker = await Tesseract.createWorker(lang, 1, {
      logger: m => {
        if (m.status === 'recognizing text') {
          updateProgress(id, 50 + m.progress * 45, `Recognizing… ${Math.round(m.progress * 100)}%`);
        }
      }
    });

    let fullText = '';
    for (let i = 0; i < images.length; i++) {
      updateProgress(id, 50 + (i / images.length) * 45, `OCR page ${i + 1}/${images.length}…`);
      const { data: { text } } = await worker.recognize(images[i]);
      fullText += text + '\n\n---\n\n';
    }

    await worker.terminate();
    updateProgress(id, 100, 'Done!');
    hideProgress(id);

    const text = fullText.trim() || '(No text detected)';
    showResult(id, `
      <p><strong>OCR complete</strong></p>
      <div class="ocr-text" id="ocr-output">${escapeHtml(text)}</div>
      <div class="result-actions">
        <button class="btn-download" id="dl-ocr-txt">Download .txt</button>
        <button class="btn-secondary" id="copy-ocr">Copy text</button>
        <button class="btn-secondary" onclick="location.reload()">New file</button>
      </div>
    `);

    $('#dl-ocr-txt').addEventListener('click', () => {
      const blob = new Blob([text], { type: 'text/plain' });
      downloadBlob(blob, file.name.replace(/\.[^.]+$/, '') + '-ocr.txt');
    });

    $('#copy-ocr').addEventListener('click', async () => {
      await navigator.clipboard.writeText(text);
      $('#copy-ocr').textContent = 'Copied!';
      setTimeout(() => { $('#copy-ocr').textContent = 'Copy text'; }, 1500);
    });
  } catch (err) {
    console.error(err);
    hideProgress(id);
    showResult(id, `<p style="color:#f87171">Error: ${err.message || 'OCR failed'}</p>`);
  }
}

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}


// ========== Background Remover ==========
let removeBackgroundFn = null;
let lastTransparentBlob = null;

async function loadBgRemover() {
  if (removeBackgroundFn) return removeBackgroundFn;
  // Dynamic import of @imgly/background-removal (ESM)
  const mod = await import('https://cdn.jsdelivr.net/npm/@imgly/background-removal@1.5.5/+esm');
  removeBackgroundFn = mod.removeBackground || mod.default;
  return removeBackgroundFn;
}

async function runBgRemove(file) {
  const id = 'bgremove';
  clearResult(id);
  showProgress(id, 'Loading AI model… (first time ~40MB, then cached)');

  try {
    const removeBackground = await loadBgRemover();

    updateProgress(id, 20, 'Removing background…');

    const transparentBlob = await removeBackground(file, {
      model: 'medium',
      output: { format: 'image/png', quality: 0.9 },
      progress: (key, current, total) => {
        if (total) {
          const pct = 20 + Math.round((current / total) * 70);
          updateProgress(id, pct, `Processing… ${Math.round((current / total) * 100)}%`);
        }
      }
    });

    lastTransparentBlob = transparentBlob;
    updateProgress(id, 100, 'Done!');
    hideProgress(id);

    const url = URL.createObjectURL(transparentBlob);
    const nameBase = file.name.replace(/\.[^.]+$/, '');

    showResult(id, `
      <div class="bg-preview-wrap">
        <img src="${url}" alt="Result" class="bg-preview" id="bg-preview-img" />
        <div class="bg-controls">
          <label>
            Background
            <input type="color" id="bg-color" value="#14b8a6" />
          </label>
          <div class="bg-presets">
            <button class="bg-preset" data-color="transparent" title="Transparent" style="background:linear-gradient(45deg,#334155 25%,#1e293b 25%,#1e293b 50%,#334155 50%,#334155 75%,#1e293b 75%)"></button>
            <button class="bg-preset" data-color="#ffffff" style="background:#fff"></button>
            <button class="bg-preset" data-color="#000000" style="background:#000"></button>
            <button class="bg-preset" data-color="#14b8a6" style="background:#14b8a6"></button>
            <button class="bg-preset" data-color="#0ea5e9" style="background:#0ea5e9"></button>
            <button class="bg-preset" data-color="#f43f5e" style="background:#f43f5e"></button>
          </div>
        </div>
        <div class="result-actions">
          <button class="btn-download" id="dl-bg-transparent">Download Transparent PNG</button>
          <button class="btn-download" id="dl-bg-color">Download with Color</button>
          <button class="btn-secondary" onclick="location.reload()">New image</button>
        </div>
      </div>
    `);

    // Wire color controls
    const preview = $('#bg-preview-img');
    const colorInput = $('#bg-color');

    function applyColor(hex) {
      if (hex === 'transparent') {
        preview.style.backgroundColor = 'transparent';
        preview.style.backgroundImage = '';
        return;
      }
      preview.style.backgroundImage = 'none';
      preview.style.backgroundColor = hex;
      colorInput.value = hex;
    }

    colorInput.addEventListener('input', () => applyColor(colorInput.value));

    $$('.bg-preset').forEach(btn => {
      btn.addEventListener('click', () => {
        $$('.bg-preset').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        applyColor(btn.dataset.color);
      });
    });

    $('#dl-bg-transparent').addEventListener('click', () => {
      downloadBlob(transparentBlob, nameBase + '-nobg.png');
    });

    $('#dl-bg-color').addEventListener('click', async () => {
      const color = colorInput.value;
      const img = new Image();
      img.src = url;
      await img.decode();

      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth;
      canvas.height = img.naturalHeight;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = color;
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0);

      canvas.toBlob(blob => {
        downloadBlob(blob, nameBase + '-bg.png');
      }, 'image/png');
    });

  } catch (err) {
    console.error(err);
    hideProgress(id);
    showResult(id, `<p style="color:#f87171">Error: ${err.message || 'Background removal failed. Try a smaller image or refresh.'}</p>`);
  }
}

// ---------- Init ----------
function init() {
  createParticles();
  initHeader();
  initTabs();

  setupDropzone('#drop-pdf2word', '#input-pdf2word', pdfToWord);
  setupDropzone('#drop-word2pdf', '#input-word2pdf', wordToPdf);
  setupDropzone('#drop-ocr', '#input-ocr', runOcr);
  setupDropzone('#drop-bgremove', '#input-bgremove', runBgRemove);

  // GSAP entrance animations (lightweight)
  if (window.gsap) {
    gsap.registerPlugin(ScrollTrigger);
    gsap.from('.hero-badge, .hero-title, .hero-sub, .hero-actions', {
      y: 30,
      opacity: 0,
      duration: 0.8,
      stagger: 0.12,
      ease: 'power3.out',
      delay: 0.15
    });

    gsap.utils.toArray('.privacy-card').forEach((card, i) => {
      gsap.from(card, {
        scrollTrigger: { trigger: card, start: 'top 88%' },
        y: 40,
        opacity: 0,
        duration: 0.6,
        delay: i * 0.08
      });
    });
  }
}

document.addEventListener('DOMContentLoaded', init);
