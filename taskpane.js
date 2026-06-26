"use strict";

const state = {
  canInsert: false,
  renderedHtml: "",
  mathCache: new Map()
};

const sampleText = `and:
\\[
\\mu_{\\rm line}=\\mu_{\\rm eff}A_{\\rm eff}=C\\mu_{\\rm eff}l_P^2
\\]
Because both scale by the same \\(C\\), the wave speed stays the same:
\\[
v=\\sqrt{\\frac{T_{\\rm eff}}{\\mu_{\\rm line}}}
=
\\sqrt{
\\frac{CB l_P^2}{C\\mu_{\\rm eff}l_P^2}
}
=
\\sqrt{\\frac{B}{\\mu_{\\rm eff}}}
=c
\\]
So packing factor changes force and line inertia, but not wave speed.
Using:
\\[
B\\approx4.63\\times10^{113}\\ {\\rm J/m^3}
\\]
\\[
\\mu_{\\rm eff}\\approx5.15\\times10^{96}\\ {\\rm kg/m^3}
\\]
\\[
l_P^2\\approx2.61\\times10^{-70}\\ {\\rm m^2}
\\]
Results:
\\[
\\begin{array}{c|c|c}
C & T_{\\rm eff}=BA & \\mu_{\\rm line}=\\mu A\\\\
\\hline
1 & 1.21\\times10^{44}\\ {\\rm N} & 1.35\\times10^{27}\\ {\\rm kg/m}\\\\
\\pi & 3.80\\times10^{44}\\ {\\rm N} & 4.23\\times10^{27}\\ {\\rm kg/m}\\\\
4\\pi & 1.52\\times10^{45}\\ {\\rm N} & 1.69\\times10^{28}\\ {\\rm kg/m}
\\end{array}
\\]
Interpretation:
\\[
\\boxed{C=1}
\\]
means one effective Planck-area channel.`;

const els = {};

document.addEventListener("DOMContentLoaded", () => {
  els.source = document.getElementById("sourceText");
  els.preview = document.getElementById("preview");
  els.status = document.getElementById("statusText");
  els.host = document.getElementById("hostState");
  els.render = document.getElementById("renderButton");
  els.insert = document.getElementById("insertButton");
  els.copy = document.getElementById("copyButton");
  els.clear = document.getElementById("clearButton");
  els.sample = document.getElementById("sampleButton");

  els.source.value = sampleText;
  els.render.addEventListener("click", renderCurrentNote);
  els.insert.addEventListener("click", insertIntoOneNote);
  els.copy.addEventListener("click", copyRenderedHtml);
  els.clear.addEventListener("click", clearAll);
  els.sample.addEventListener("click", () => {
    els.source.value = sampleText;
    renderCurrentNote();
  });

  bootOffice();
  renderCurrentNote();
});

function bootOffice() {
  if (!window.Office) {
    setHostState("Browser preview", "status-warn");
    return;
  }

  Office.onReady((info) => {
    state.canInsert = Boolean(info && info.host === Office.HostType.OneNote);
    els.insert.disabled = !state.canInsert;
    setHostState(state.canInsert ? "OneNote Online" : "Browser preview", state.canInsert ? "status-good" : "status-warn");
  });
}

async function renderCurrentNote() {
  setStatus("Rendering", "");
  els.render.disabled = true;
  els.insert.disabled = true;
  try {
    await waitForMathJax();
    const blocks = parseBlocks(els.source.value);
    const html = await renderBlocks(blocks);
    state.renderedHtml = `<div class="insert-note">${html}</div>`;
    els.preview.innerHTML = html || emptyPreview();
    setStatus("Ready", "status-good");
  } catch (error) {
    state.renderedHtml = "";
    els.preview.innerHTML = `<div class="math-error">${escapeHtml(error.message || String(error))}</div>`;
    setStatus("Render error", "status-bad");
  } finally {
    els.render.disabled = false;
    els.insert.disabled = !state.canInsert || !state.renderedHtml;
  }
}

function parseBlocks(source) {
  const blocks = [];
  const pattern = /(\\\[[\s\S]*?\\\]|\$\$[\s\S]*?\$\$)/g;
  let lastIndex = 0;
  let match;

  while ((match = pattern.exec(source)) !== null) {
    pushTextBlocks(blocks, source.slice(lastIndex, match.index));
    const raw = match[0];
    const tex = raw.startsWith("$$") ? raw.slice(2, -2) : raw.slice(2, -2);
    blocks.push({ type: isArrayMath(tex) ? "array" : "math", tex: tex.trim() });
    lastIndex = pattern.lastIndex;
  }

  pushTextBlocks(blocks, source.slice(lastIndex));
  return blocks;
}

function pushTextBlocks(blocks, text) {
  text
    .replace(/\r\n/g, "\n")
    .split(/\n{2,}/)
    .map((part) => part.trim())
    .filter(Boolean)
    .forEach((part) => blocks.push({ type: "text", text: part }));
}

function isArrayMath(tex) {
  return /\\begin\{array\}/.test(tex);
}

async function renderBlocks(blocks) {
  const html = [];
  for (const block of blocks) {
    if (block.type === "text") {
      html.push(`<p>${await renderInlineMath(block.text)}</p>`);
    } else if (block.type === "array") {
      html.push(await renderArrayTable(block.tex));
    } else {
      html.push(await renderDisplayMath(block.tex));
    }
  }
  return html.join("");
}

async function renderInlineMath(text) {
  const html = [];
  const pattern = /\\\(([\s\S]*?)\\\)/g;
  let lastIndex = 0;
  let match;

  while ((match = pattern.exec(text)) !== null) {
    html.push(escapeHtml(text.slice(lastIndex, match.index)).replace(/\n/g, "<br>"));
    const image = await texToPng(match[1].trim(), false);
    html.push(`<img class="math-inline" src="${image.src}" alt="${escapeHtml(match[1].trim())}">`);
    lastIndex = pattern.lastIndex;
  }

  html.push(escapeHtml(text.slice(lastIndex)).replace(/\n/g, "<br>"));
  return html.join("");
}

async function renderDisplayMath(tex) {
  try {
    const image = await texToPng(tex, true);
    return `<div class="math-line"><img src="${image.src}" width="${image.width}" height="${image.height}" alt="${escapeHtml(tex)}"></div>`;
  } catch (error) {
    return `<div class="math-error">${escapeHtml(tex)}\n\n${escapeHtml(error.message || String(error))}</div>`;
  }
}

async function renderArrayTable(tex) {
  const model = parseArray(tex);
  if (!model.rows.length) {
    return renderDisplayMath(tex);
  }

  const rows = [];
  for (let rowIndex = 0; rowIndex < model.rows.length; rowIndex += 1) {
    const cells = [];
    for (const cellTex of model.rows[rowIndex]) {
      const content = await renderTableCell(cellTex);
      cells.push(`<td${rowIndex === 0 ? ' class="table-head"' : ""}>${content}</td>`);
    }
    rows.push(`<tr>${cells.join("")}</tr>`);
  }

  return `<table class="note-table"><tbody>${rows.join("")}</tbody></table>`;
}

function parseArray(tex) {
  const body = tex
    .replace(/\\begin\{array\}\{[^}]*\}/, "")
    .replace(/\\end\{array\}/, "")
    .replace(/\\hline/g, "")
    .trim();

  const rows = splitTexRows(body)
    .map((row) => row.trim())
    .filter(Boolean)
    .map((row) => splitTexColumns(row).map((cell) => cell.trim()));

  return { rows };
}

function splitTexRows(body) {
  const rows = [];
  let buffer = "";
  for (let index = 0; index < body.length; index += 1) {
    if (body[index] === "\\" && body[index + 1] === "\\") {
      rows.push(buffer);
      buffer = "";
      index += 1;
    } else {
      buffer += body[index];
    }
  }
  if (buffer.trim()) {
    rows.push(buffer);
  }
  return rows;
}

function splitTexColumns(row) {
  const cells = [];
  let buffer = "";
  let depth = 0;

  for (let index = 0; index < row.length; index += 1) {
    const char = row[index];
    if (char === "{") depth += 1;
    if (char === "}") depth = Math.max(0, depth - 1);
    if (char === "&" && depth === 0) {
      cells.push(buffer);
      buffer = "";
    } else {
      buffer += char;
    }
  }

  cells.push(buffer);
  return cells;
}

async function renderTableCell(tex) {
  if (/^[A-Za-z0-9.,()+\-/*\s]+$/.test(tex) && !/[\\_^]/.test(tex)) {
    return escapeHtml(tex);
  }

  const image = await texToPng(tex, false);
  const width = Math.max(24, image.width);
  const height = Math.max(16, image.height);
  return `<img class="math-inline" src="${image.src}" width="${width}" height="${height}" alt="${escapeHtml(tex)}">`;
}

async function texToPng(tex, display) {
  const key = `${display ? "d" : "i"}:${tex}`;
  if (state.mathCache.has(key)) {
    return state.mathCache.get(key);
  }

  await waitForMathJax();
  const node = await MathJax.tex2svgPromise(tex, { display });
  const svg = node.querySelector("svg");
  if (!svg) {
    throw new Error("MathJax did not produce SVG.");
  }

  const image = await svgToPng(svg, display);
  state.mathCache.set(key, image);
  return image;
}

async function svgToPng(svg, display) {
  const clone = svg.cloneNode(true);
  const holder = document.createElement("div");
  holder.style.position = "fixed";
  holder.style.left = "-10000px";
  holder.style.top = "0";
  holder.style.background = "#fff";
  holder.appendChild(clone);
  document.body.appendChild(holder);

  const rect = clone.getBoundingClientRect();
  const width = Math.max(8, Math.ceil(rect.width + (display ? 24 : 8)));
  const height = Math.max(8, Math.ceil(rect.height + (display ? 18 : 6)));
  document.body.removeChild(holder);

  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  clone.setAttribute("width", String(width));
  clone.setAttribute("height", String(height));
  clone.setAttribute("style", "background:#ffffff;color:#172033;");
  clone.style.color = "#172033";

  const source = new XMLSerializer().serializeToString(clone);
  const svgUrl = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(source)}`;
  const img = await loadImage(svgUrl);
  const scale = display ? 2.4 : 2.8;
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(width * scale);
  canvas.height = Math.ceil(height * scale);

  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

  return {
    src: canvas.toDataURL("image/png"),
    width,
    height
  };
}

function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Could not convert equation image."));
    img.src = src;
  });
}

async function insertIntoOneNote() {
  if (!state.renderedHtml) {
    await renderCurrentNote();
  }
  if (!state.canInsert) {
    setStatus("Open inside OneNote", "status-warn");
    return;
  }

  setStatus("Inserting", "");
  els.insert.disabled = true;
  try {
    await OneNote.run(async (context) => {
      const page = context.application.getActivePage();
      page.addOutline(40, 80, inlineStylesForOneNote(state.renderedHtml));
      await context.sync();
    });
    setStatus("Inserted", "status-good");
  } catch (error) {
    setStatus("Insert failed", "status-bad");
    console.error(error);
  } finally {
    els.insert.disabled = false;
  }
}

function inlineStylesForOneNote(html) {
  const template = document.createElement("template");
  template.innerHTML = html;

  template.content.querySelectorAll("p").forEach((node) => {
    node.setAttribute("style", "font-family:Segoe UI,Arial,sans-serif;font-size:14px;line-height:1.55;margin:0 0 10px;color:#172033;");
  });
  template.content.querySelectorAll(".math-line").forEach((node) => {
    node.setAttribute("style", "margin:12px 0;text-align:center;");
  });
  template.content.querySelectorAll("table").forEach((node) => {
    node.setAttribute("style", "border-collapse:collapse;margin:12px 0;font-family:Segoe UI,Arial,sans-serif;font-size:13px;");
  });
  template.content.querySelectorAll("td").forEach((node) => {
    node.setAttribute("style", "border:1px solid #cbd2df;padding:7px 8px;text-align:center;vertical-align:middle;");
  });
  template.content.querySelectorAll(".table-head").forEach((node) => {
    node.setAttribute("style", `${node.getAttribute("style")}background:#eef3fb;font-weight:650;`);
  });
  template.content.querySelectorAll("img").forEach((node) => {
    if (node.classList.contains("math-inline")) {
      node.setAttribute("style", "display:inline-block;vertical-align:-0.28em;margin:0 2px;");
    } else {
      node.setAttribute("style", "max-width:100%;height:auto;");
    }
  });

  return template.innerHTML;
}

async function copyRenderedHtml() {
  if (!state.renderedHtml) {
    await renderCurrentNote();
  }

  try {
    await navigator.clipboard.writeText(inlineStylesForOneNote(state.renderedHtml));
    setStatus("Copied", "status-good");
  } catch (error) {
    setStatus("Copy failed", "status-bad");
  }
}

function clearAll() {
  els.source.value = "";
  els.preview.innerHTML = emptyPreview();
  state.renderedHtml = "";
  state.mathCache.clear();
  els.insert.disabled = true;
  setStatus("Ready", "");
}

function waitForMathJax() {
  if (window.MathJax && MathJax.startup && MathJax.startup.promise) {
    return MathJax.startup.promise;
  }

  return new Promise((resolve, reject) => {
    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (window.MathJax && MathJax.startup && MathJax.startup.promise) {
        clearInterval(timer);
        MathJax.startup.promise.then(resolve).catch(reject);
      } else if (attempts > 100) {
        clearInterval(timer);
        reject(new Error("MathJax did not load."));
      }
    }, 100);
  });
}

function setStatus(text, className) {
  els.status.textContent = text;
  els.status.className = className || "";
}

function setHostState(text, className) {
  els.host.textContent = text;
  els.host.className = className || "";
}

function emptyPreview() {
  return '<p style="color:#5d6576;">Empty</p>';
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}
