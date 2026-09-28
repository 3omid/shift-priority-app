// Turns the printable Shift Exchange form (the exact HTML that the Print
// button opens — buildSwapFormHtml in App.jsx) into real files the person can
// save or send from their phone: a one-page Letter PDF and a PNG image.
//
// No libraries and nothing leaves the device. The browser lays the form out
// once in a hidden iframe (so the files always look exactly like the printed
// page — one layout, not two), then we copy every word and every border line
// from that layout onto a high-resolution canvas and wrap the canvas in a
// minimal PDF. Drawing the text ourselves (instead of drawing the HTML into
// the canvas via SVG) matters on iPhone: Safari refuses to export a canvas
// that had HTML drawn into it.

const PAGE_W = 816; // US Letter, 8.5in at 96 CSS px/in
const PAGE_H = 1056; // 11in
const MARGIN_X = 48; // 0.5in — same as the form's @page margin
const MARGIN_Y = 43.2; // 0.45in
const SCALE = 3; // 288 dpi — crisp when printed or zoomed

function loadIntoHiddenFrame(html) {
  return new Promise((resolve, reject) => {
    const iframe = document.createElement("iframe");
    iframe.setAttribute("aria-hidden", "true");
    iframe.tabIndex = -1;
    iframe.style.cssText = `position:fixed; left:-12000px; top:0; width:${PAGE_W}px; height:${PAGE_H}px; border:0; opacity:0; pointer-events:none;`;
    // Screen-only bits of the print page (toolbar, preview padding) off, so
    // the layout is exactly what ends up on paper.
    const printLike = "<style>.toolbar{display:none!important} html,body{margin:0!important;padding:0!important} .page{margin:0!important;padding:0!important}</style>";
    iframe.srcdoc = html.replace("</head>", `${printLike}</head>`);
    const timer = setTimeout(() => { iframe.remove(); reject(new Error("form layout timed out")); }, 15000);
    iframe.onload = async () => {
      clearTimeout(timer);
      try {
        const doc = iframe.contentDocument;
        if (doc.fonts && doc.fonts.ready) await doc.fonts.ready;
        resolve(iframe);
      } catch (e) {
        iframe.remove();
        reject(e);
      }
    };
    document.body.appendChild(iframe);
  });
}

function drawLayoutOnCanvas(iframe) {
  const win = iframe.contentWindow;
  const doc = iframe.contentDocument;
  const root = doc.querySelector(".page") || doc.body;
  const origin = root.getBoundingClientRect();
  const ox = MARGIN_X - origin.left;
  const oy = MARGIN_Y - origin.top;

  const canvas = document.createElement("canvas");
  canvas.width = PAGE_W * SCALE;
  canvas.height = PAGE_H * SCALE;
  const ctx = canvas.getContext("2d");
  ctx.scale(SCALE, SCALE);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, PAGE_W, PAGE_H);

  const visible = (cs) => cs.display !== "none" && cs.visibility !== "hidden" && cs.opacity !== "0";

  // 1) Every border line / box outline (underlines, driver boxes, checkboxes).
  for (const el of root.querySelectorAll("*")) {
    const cs = win.getComputedStyle(el);
    if (!visible(cs)) continue;
    const r = el.getBoundingClientRect();
    if (r.width === 0 && r.height === 0) continue;
    const side = (name) => {
      const w = parseFloat(cs[`border${name}Width`]) || 0;
      const style = cs[`border${name}Style`];
      return w > 0 && style !== "none" && style !== "hidden" ? { w, color: cs[`border${name}Color`] } : null;
    };
    const x = r.left + ox, y = r.top + oy;
    const top = side("Top"), right = side("Right"), bottom = side("Bottom"), left = side("Left");
    if (top) { ctx.fillStyle = top.color; ctx.fillRect(x, y, r.width, top.w); }
    if (bottom) { ctx.fillStyle = bottom.color; ctx.fillRect(x, y + r.height - bottom.w, r.width, bottom.w); }
    if (left) { ctx.fillStyle = left.color; ctx.fillRect(x, y, left.w, r.height); }
    if (right) { ctx.fillStyle = right.color; ctx.fillRect(x + r.width - right.w, y, right.w, r.height); }
  }

  // 2) Every word, at exactly the position the browser laid it out.
  const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const range = doc.createRange();
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.data;
    if (!/[^\s ]/.test(text)) continue;
    const parent = node.parentElement;
    if (!parent) continue;
    const cs = win.getComputedStyle(parent);
    if (!visible(cs)) continue;
    const size = parseFloat(cs.fontSize) || 12;
    ctx.font = `${cs.fontStyle} ${cs.fontWeight} ${cs.fontSize} ${cs.fontFamily}`;
    ctx.fillStyle = cs.color;
    ctx.textBaseline = "alphabetic";
    ctx.textAlign = "left";
    const spaced = cs.letterSpacing && cs.letterSpacing !== "normal" && parseFloat(cs.letterSpacing) !== 0;
    const underline = (cs.textDecorationLine || cs.textDecoration || "").includes("underline");
    const draw = (str, start, end) => {
      range.setStart(node, start);
      range.setEnd(node, end);
      const rects = range.getClientRects();
      const r = rects[0];
      if (!r || r.width === 0) return;
      const m = ctx.measureText(str);
      const asc = m.fontBoundingBoxAscent || size * 0.905;
      const desc = m.fontBoundingBoxDescent || size * 0.212;
      const baseline = r.top + oy + (r.height - (asc + desc)) / 2 + asc;
      ctx.fillText(str, r.left + ox, baseline);
      if (underline) ctx.fillRect(r.left + ox, baseline + size * 0.12, r.width, Math.max(0.6, size / 14));
    };
    const wordRe = /[^\s ]+/g;
    let m;
    while ((m = wordRe.exec(text))) {
      if (spaced) {
        for (let i = 0; i < m[0].length; i++) draw(m[0][i], m.index + i, m.index + i + 1);
      } else {
        draw(m[0], m.index, m.index + m[0].length);
      }
    }
  }
  return canvas;
}

function canvasToBlob(canvas, type, quality) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("canvas export failed"))), type, quality);
  });
}

// A minimal, valid one-page PDF (US Letter) whose only content is the given
// JPEG stretched over the whole page.
export function jpegToPdf(jpegBytes, pxW, pxH) {
  const enc = new TextEncoder();
  const chunks = [];
  const offsets = [];
  let length = 0;
  const push = (data) => {
    const bytes = typeof data === "string" ? enc.encode(data) : data;
    chunks.push(bytes);
    length += bytes.length;
  };
  const obj = (n, body) => {
    offsets[n] = length;
    push(`${n} 0 obj\n${body}\nendobj\n`);
  };
  push(new Uint8Array([0x25, 0x50, 0x44, 0x46, 0x2d, 0x31, 0x2e, 0x34, 0x0a, 0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a])); // %PDF-1.4 + binary marker
  obj(1, "<< /Type /Catalog /Pages 2 0 R >>");
  obj(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");
  obj(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /XObject << /Im0 4 0 R >> >> /Contents 5 0 R >>");
  offsets[4] = length;
  push(`4 0 obj\n<< /Type /XObject /Subtype /Image /Width ${pxW} /Height ${pxH} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${jpegBytes.length} >>\nstream\n`);
  push(jpegBytes);
  push("\nendstream\nendobj\n");
  const content = "q 612 0 0 792 0 0 cm /Im0 Do Q";
  obj(5, `<< /Length ${content.length} >>\nstream\n${content}\nendstream`);
  const xrefAt = length;
  let xref = "xref\n0 6\n0000000000 65535 f \n";
  for (let n = 1; n <= 5; n++) xref += `${String(offsets[n]).padStart(10, "0")} 00000 n \n`;
  push(`${xref}trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`);
  const out = new Uint8Array(length);
  let at = 0;
  for (const c of chunks) { out.set(c, at); at += c.length; }
  return out;
}

// html -> { pdf: Blob, png: Blob }
export async function buildExchangeFormFiles(html) {
  const iframe = await loadIntoHiddenFrame(html);
  try {
    const canvas = drawLayoutOnCanvas(iframe);
    const jpeg = await canvasToBlob(canvas, "image/jpeg", 0.92);
    const png = await canvasToBlob(canvas, "image/png");
    const pdfBytes = jpegToPdf(new Uint8Array(await jpeg.arrayBuffer()), canvas.width, canvas.height);
    return { pdf: new Blob([pdfBytes], { type: "application/pdf" }), png };
  } finally {
    iframe.remove();
  }
}

// On a phone: the system share sheet (Save to Files / Save Image / Messages /
// Mail / Print …). On a computer, or if sharing isn't possible: a download.
export async function shareOrDownloadFile(blob, filename, title) {
  const file = new File([blob], filename, { type: blob.type });
  const touch = typeof window.matchMedia === "function" && window.matchMedia("(pointer: coarse)").matches;
  if (touch && navigator.canShare && navigator.share) {
    let can = false;
    try { can = navigator.canShare({ files: [file] }); } catch { can = false; }
    if (can) {
      try {
        await navigator.share({ files: [file], title });
        return;
      } catch (e) {
        if (e && e.name === "AbortError") return; // person closed the share sheet
      }
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.rel = "noopener";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
