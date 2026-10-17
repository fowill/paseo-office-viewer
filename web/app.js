import { getDocument, GlobalWorkerOptions, TextLayer } from "pdfjs-dist";
import { renderAsync } from "docx-preview";
import * as XLSX from "xlsx";

// pptxviewjs 1.1.9 overwrites FileReader. Restore the browser API before other
// formats load embedded images or fonts through it.
const NativeFileReader = window.FileReader;
const { PPTXViewer } = await import("pptxviewjs");
window.FileReader = NativeFileReader;

GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

const $ = (id) => document.getElementById(id);
const hash = new URLSearchParams(location.hash.slice(1));
const tokenFromUrl = hash.get("token");
if (tokenFromUrl) {
  sessionStorage.setItem("paseo-office-token", tokenFromUrl);
  sessionStorage.setItem("paseo-office-embed", hash.get("embed") === "1" ? "1" : "0");
  history.replaceState(null, "", location.pathname + location.search);
}
const token = tokenFromUrl || sessionStorage.getItem("paseo-office-token");
const embedded = sessionStorage.getItem("paseo-office-embed") === "1";
document.body.classList.toggle("embed", embedded);
if (embedded) $("empty-prompt").textContent = "在右侧 Office Viewer 面板中点击文件开始预览";
const state = { directory: "", entries: [], path: "", document: null, generation: 0, listGeneration: 0, controller: null, zoom: 1, fit: true, page: 1, pages: 1, sheet: 0, row: 0, column: 0, busy: false };
const rowsPerPage = 100;
const columnsPerPage = 30;
let pptQueue = Promise.resolve();
let events;
let selectedSequence = -1;

// The PPT renderer has shared chart state. Keep its operations in order.
function runPpt(operation) {
  const result = pptQueue.then(operation);
  pptQueue = result.catch(() => {});
  return result;
}

function element(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text !== undefined) node.textContent = text;
  return node;
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

async function api(route, path, signal) {
  const response = await fetch(`${route}?path=${encodeURIComponent(path)}`, { headers: { Authorization: `Bearer ${token || ""}` }, signal });
  if (!response.ok) {
    const body = await response.json().catch(() => ({}));
    if (response.status === 401) throw new Error("预览会话已失效，请从 Paseo 插件重新打开。");
    throw new Error(body.error || `读取失败（HTTP ${response.status}）。`);
  }
  return response;
}

function showError(error) {
  $("preview-status").textContent = "";
  $("preview-error").textContent = `无法预览：${error.message || error}`;
  $("preview-error").title = $("preview-error").textContent;
  $("preview-error").hidden = false;
  $("loading-progress").hidden = true;
}

function applyTheme(theme) {
  if (!theme) return;
  const variables = { surface0: "bg", surface1: "panel", surface2: "control", border: "line", foreground: "fg", foregroundMuted: "muted", accent: "accent", statusDanger: "danger" };
  for (const [color, variable] of Object.entries(variables)) document.documentElement.style.setProperty(`--${variable}`, theme[color]);
  // Resolve both hexadecimal and rgb() host colors through the browser.
  const probe = element("span");
  probe.style.color = theme.surface0;
  document.body.append(probe);
  const rgb = getComputedStyle(probe).color.match(/[\d.]+/g).map(Number);
  probe.remove();
  const dark = .2126 * rgb[0] + .7152 * rgb[1] + .0722 * rgb[2] < 128;
  document.documentElement.style.colorScheme = dark ? "dark" : "light";
  document.documentElement.style.setProperty("--page-shadow", `0 1px 3px rgba(0,0,0,${dark ? ".4" : ".12"})`);
}

function selectFromSession(selected) {
  if (!selected || selected.seq < selectedSequence) return;
  selectedSequence = selected.seq;
  void openFile(selected.path);
}

async function readSession() {
  const session = await (await api("/api/session", "")).json();
  applyTheme(session.theme);
  selectFromSession(session.selected);
}

function connectSession() {
  // EventSource cannot set Authorization. Only this endpoint accepts a query token.
  events = new EventSource(`/api/events?token=${encodeURIComponent(token)}`);
  events.addEventListener("select", event => selectFromSession(JSON.parse(event.data)));
  events.addEventListener("theme", event => applyTheme(JSON.parse(event.data)));
  // Reading after connection also recovers changes missed during a reconnect.
  events.addEventListener("open", () => void readSession().catch(showError));
  events.addEventListener("error", () => {
    $("preview-status").textContent = "预览连接中断，正在重连…";
  });
  void readSession().catch(showError);
}

function renderFileList() {
  const search = $("file-filter").value.trim().toLocaleLowerCase();
  const entries = state.entries.filter((entry) => entry.name.toLocaleLowerCase().includes(search));
  const fragment = document.createDocumentFragment();
  for (const entry of entries) {
    const button = element("button", `file-row${entry.path === state.path ? " active" : ""}`);
    button.dataset.path = entry.path;
    button.title = entry.name;
    button.setAttribute("aria-label", `${entry.directory ? "打开文件夹" : "预览"} ${entry.name}`);
    const kind = entry.directory ? "▸" : entry.name.split(".").pop().toUpperCase();
    button.append(element("span", `file-symbol${entry.directory ? " folder" : ""}`, kind));
    const details = element("span", "file-details");
    details.append(element("strong", "", entry.name), element("small", "", entry.directory ? "文件夹" : formatSize(entry.size)));
    button.append(details);
    button.addEventListener("click", () => entry.directory ? loadDirectory(entry.path) : openFile(entry.path));
    fragment.append(button);
  }
  $("file-list").replaceChildren(fragment);
  $("list-status").textContent = entries.length ? `${entries.length} 个项目` : search ? "没有匹配的文件。" : "此文件夹没有可预览文件。";
}

async function loadDirectory(path) {
  const generation = ++state.listGeneration;
  $("list-status").textContent = "正在读取文件列表…";
  try {
    const result = await (await api("/api/files", path)).json();
    if (generation !== state.listGeneration) return;
    state.directory = result.path;
    state.entries = result.entries;
    $("workspace-name").textContent = result.workspaceName;
    $("file-filter").value = "";
    const crumbs = document.createDocumentFragment();
    const root = element("button", "", "项目根目录");
    root.addEventListener("click", () => loadDirectory(""));
    crumbs.append(root);
    const parts = result.path.split("/").filter(Boolean);
    parts.forEach((part, index) => {
      const button = element("button", "", part);
      button.addEventListener("click", () => loadDirectory(parts.slice(0, index + 1).join("/")));
      crumbs.append(element("span", "", "/"), button);
    });
    $("breadcrumbs").replaceChildren(crumbs);
    renderFileList();
  } catch (error) {
    if (generation === state.listGeneration) $("list-status").textContent = error.message;
  }
}

function releaseDocument(doc) {
  if (!doc) return;
  doc.disposed = true;
  doc.renderTask?.cancel();
  doc.textLayer?.cancel();
  if (doc.loadingTask) doc.loadingTask.destroy().catch(() => {});
  if (doc.ppt && !doc.rendering) {
    doc.ppt.processor?.clearCaches();
    if (window.currentProcessor?.processor === doc.ppt.processor) {
      delete window.currentProcessor;
      delete window.currentZipData;
    }
    doc.ppt.destroy();
  }
}

function updateControls() {
  const doc = state.document;
  const paged = doc?.kind === "pdf" || doc?.kind === "pptx";
  $("page-controls").hidden = !paged;
  $("page-number").value = state.page;
  $("page-number").max = state.pages;
  $("page-number").disabled = state.busy;
  $("page-count").textContent = `/ ${state.pages} 页`;
  $("previous-page").disabled = state.busy || state.page <= 1;
  $("next-page").disabled = state.busy || state.page >= state.pages;
  $("zoom-value").textContent = `${Math.round(state.zoom * 100)}%`;
  $("zoom-out").disabled = state.busy || state.zoom <= 0.5;
  $("zoom-in").disabled = state.busy || state.zoom >= 3;
  $("zoom-reset").disabled = state.busy;
  $("zoom-reset").textContent = doc?.kind === "xlsx" ? "100%" : "适应";
  $("zoom-reset").title = doc?.kind === "xlsx" ? "恢复默认缩放" : "适应宽度";
  $("zoom-reset").setAttribute("aria-pressed", String(state.fit && doc?.kind !== "xlsx"));
  $("loading-progress").hidden = !state.busy;
}

async function openFile(path, force = false) {
  if (!force && path === state.path && (state.busy || state.document)) return;
  const generation = ++state.generation;
  state.controller?.abort();
  state.controller = new AbortController();
  releaseDocument(state.document);
  state.document = null;
  state.path = path;
  state.zoom = 1;
  state.fit = true;
  state.page = 1;
  state.busy = true;
  const name = path.split("/").pop();
  const kind = name.split(".").pop().toLowerCase();
  $("file-name").textContent = name;
  $("file-kind").textContent = kind.toUpperCase();
  $("file-kind").hidden = false;
  $("file-meta").textContent = path;
  $("file-meta").title = path;
  $("file-name").title = path;
  $("preview").className = "preview";
  $("preview").replaceChildren();
  $("preview").scrollTop = 0;
  $("preview-error").hidden = true;
  $("preview-status").textContent = "正在读取文件…";
  $("loading-progress").hidden = false;
  $("reload-file").disabled = false;
  for (const id of ["toolbar", "sheet-tabs", "sheet-controls", "cell-inspector", "format-note"]) $(id).hidden = true;
  renderFileList();
  let doc;
  try {
    const response = await api("/api/file", path, state.controller.signal);
    const bytes = await response.arrayBuffer();
    if (generation !== state.generation) return;
    if (["docx", "xlsx", "pptx"].includes(kind)) {
      const signature = new Uint8Array(bytes, 0, Math.min(4, bytes.byteLength));
      if (signature[0] !== 0x50 || signature[1] !== 0x4b || signature[2] !== 0x03 || signature[3] !== 0x04) {
        throw new Error(`文件不是有效的 ${kind.toUpperCase()} 文档，或文件已加密。请使用 Office 检查后另存为未加密文件。`);
      }
    }
    $("file-meta").textContent = `${path} · ${formatSize(bytes.byteLength)}`;
    $("preview-status").textContent = "正在生成预览…";
    doc = { kind, generation, disposed: false };
    state.document = doc;

    if (kind === "pdf") {
      doc.loadingTask = getDocument({ data: bytes, cMapUrl: "/cmaps/", cMapPacked: true, standardFontDataUrl: "/standard_fonts/", wasmUrl: "/wasm/", isEvalSupported: false });
      doc.pdf = await doc.loadingTask.promise;
      if (generation !== state.generation) return;
      state.pages = doc.pdf.numPages;
      state.busy = false;
      if (await renderPage() === false) return;
    } else if (kind === "docx") {
      doc.host = element("div", "docx-host");
      await renderAsync(bytes, doc.host, doc.host, { inWrapper: true, breakPages: true, useBase64URL: true, renderAltChunks: false, renderComments: false });
      if (generation !== state.generation) return;
      // Hyperlinks remain readable; document content cannot navigate this viewer.
      doc.host.querySelectorAll("a").forEach((link) => link.removeAttribute("href"));
      $("preview").replaceChildren(doc.host);
      applyDocumentZoom();
      $("format-note").textContent = "Word 网页预览。复杂分页、浮动对象和公式的显示可能与 Word 不同。";
      $("format-note").hidden = false;
    } else if (kind === "xlsx") {
      doc.workbook = XLSX.read(bytes, { type: "array", cellStyles: true, cellNF: true, cellFormula: true, cellText: true });
      if (!doc.workbook.SheetNames.length) throw new Error("工作簿中没有工作表。");
      $("format-note").textContent = "显示已保存的值；公式不重算。图表、图片和条件格式暂不显示。点击单元格可查看公式。";
      $("format-note").hidden = false;
      $("preview").className = "preview spreadsheet";
      $("sheet-tabs").hidden = false;
      $("sheet-controls").hidden = false;
      $("cell-inspector").hidden = false;
      selectSheet(0);
    } else if (kind === "pptx") {
      doc.ppt = new PPTXViewer({ enableThumbnails: false, autoChartRerenderDelayMs: 0, autoRenderFirstSlide: false });
      doc.rendering = true;
      try { await runPpt(async () => { if (!doc.disposed) await doc.ppt.loadFile(new File([bytes], name)); }); }
      finally { doc.rendering = false; }
      if (generation !== state.generation) { releaseDocument(doc); return; }
      state.pages = doc.ppt.getSlideCount();
      if (!state.pages) throw new Error("演示文稿中没有幻灯片。");
      $("format-note").textContent = "静态幻灯片预览。复杂母版、字体和动画的显示可能与 PowerPoint 不同。";
      $("format-note").hidden = false;
      state.busy = false;
      if (await renderPage() === false) return;
    } else {
      throw new Error("V1 支持 PDF、DOCX、XLSX 和 PPTX。请先转换旧版 Office 格式。");
    }
    if (generation !== state.generation) return;
    $("toolbar").hidden = false;
    $("format-note").title = $("format-note").textContent;
    $("preview-status").textContent = kind === "pdf" || kind === "pptx" ? `已加载 · 第 ${state.page} / ${state.pages} 页 · 只读预览` : "已加载 · 只读预览";
    state.busy = false;
    updateControls();
  } catch (error) {
    if (generation !== state.generation || error.name === "AbortError") return;
    state.busy = false;
    releaseDocument(doc);
    state.document = null;
    $("toolbar").hidden = true;
    showError(error);
  }
}

async function renderPage() {
  const doc = state.document;
  if (!doc || state.busy) return;
  state.busy = true;
  doc.rendering = true;
  updateControls();
  $("preview-error").hidden = true;
  $("preview-status").textContent = "正在绘制页面…";
  const stage = element("div", `page-stage ${doc.kind === "pdf" ? "pdf-page" : "ppt-page"}`);
  const canvas = element("canvas");
  canvas.setAttribute("aria-label", `第 ${state.page} 页`);
  stage.append(canvas);
  const dpr = Math.min(window.devicePixelRatio || 1, 2);
  try {
    if (doc.kind === "pdf") {
      const page = await doc.pdf.getPage(state.page);
      if (doc.disposed) return;
      if (state.fit) {
        lastWidth = availableWidth();
        state.zoom = Math.max(.5, Math.min(3, lastWidth / page.getViewport({ scale: 1 }).width));
      }
      const viewport = page.getViewport({ scale: state.zoom });
      stage.style.width = `${viewport.width}px`;
      stage.style.height = `${viewport.height}px`;
      stage.style.setProperty("--total-scale-factor", viewport.scale);
      canvas.width = Math.ceil(viewport.width * dpr);
      canvas.height = Math.ceil(viewport.height * dpr);
      canvas.style.width = `${viewport.width}px`;
      canvas.style.height = `${viewport.height}px`;
      doc.renderTask = page.render({ canvasContext: canvas.getContext("2d"), viewport, transform: dpr === 1 ? null : [dpr, 0, 0, dpr, 0, 0] });
      await doc.renderTask.promise;
      if (doc.disposed) return;
      const textContainer = element("div", "textLayer");
      stage.append(textContainer);
      doc.textLayer = new TextLayer({ textContentSource: page.streamTextContent(), container: textContainer, viewport });
      await doc.textLayer.render();
    } else {
      const dimensions = doc.ppt.processor.getSlideDimensions();
      if (state.fit) {
        lastWidth = availableWidth();
        state.zoom = Math.max(480, Math.min(1920, lastWidth)) / 960;
      }
      const width = Math.round(960 * state.zoom);
      const height = Math.round(width * dimensions.cy / dimensions.cx);
      canvas.width = Math.round(width * dpr);
      canvas.height = Math.round(height * dpr);
      canvas.style.width = `${width}px`;
      canvas.style.height = `${height}px`;
      const slideIndex = state.page - 1;
      await runPpt(async () => { if (!doc.disposed) await doc.ppt.render(canvas, { slideIndex }); });
    }
    if (doc.disposed || doc.generation !== state.generation) return;
    $("preview").replaceChildren(stage);
    $("preview").scrollTop = 0;
    $("preview-status").textContent = `已加载 · 第 ${state.page} / ${state.pages} 页 · 只读预览`;
    return true;
  } catch (error) {
    if (!doc.disposed && doc.generation === state.generation) showError(error);
    return false;
  } finally {
    doc.rendering = false;
    if (doc.disposed) releaseDocument(doc);
    if (doc.generation === state.generation) { state.busy = false; updateControls(); }
  }
}

function selectSheet(index) {
  const workbook = state.document.workbook;
  state.sheet = index;
  const sheet = workbook.Sheets[workbook.SheetNames[index]];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  state.row = range.s.r;
  state.column = range.s.c;
  const tabs = workbook.SheetNames.map((name, sheetIndex) => {
    const button = element("button", "sheet-tab", name);
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", String(sheetIndex === index));
    button.addEventListener("click", () => selectSheet(sheetIndex));
    return button;
  });
  $("sheet-tabs").replaceChildren(...tabs);
  $("selected-cell").textContent = "单元格";
  $("cell-value").textContent = "点击单元格查看完整内容及公式。";
  renderSheet();
}

function inspectCell(address, cell, td) {
  $("preview").querySelector("td.selected")?.classList.remove("selected");
  td.classList.add("selected");
  $("selected-cell").textContent = address;
  const value = cell ? (cell.w ?? XLSX.utils.format_cell(cell)) : "";
  $("cell-value").textContent = cell?.f ? `公式：=${cell.f} · 已保存的值：${cell.v === undefined ? "无缓存结果，无法计算" : value}` : value || "（空白）";
  $("cell-value").title = $("cell-value").textContent;
  $("cell-address").value = address;
}

function applyCellStyle(td, cell) {
  if (!cell) return;
  td.style.textAlign = cell.t === "n" ? "right" : "left";
  const style = cell.s;
  if (!style) return;
  const color = style.fgColor?.rgb;
  if (typeof color === "string" && /^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(color)) td.style.backgroundColor = `#${color.slice(-6)}`;
  if (style.alignment?.horizontal && ["left", "center", "right"].includes(style.alignment.horizontal)) td.style.textAlign = style.alignment.horizontal;
  if (style.font?.bold) td.style.fontWeight = "bold";
  if (style.font?.italic) td.style.fontStyle = "italic";
}

function renderSheet() {
  const workbook = state.document.workbook;
  const sheet = workbook.Sheets[workbook.SheetNames[state.sheet]];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const endRow = Math.min(state.row + rowsPerPage - 1, range.e.r);
  const endColumn = Math.min(state.column + columnsPerPage - 1, range.e.c);
  $("row-range").textContent = `${state.row + 1}–${endRow + 1} / ${range.e.r + 1} 行`;
  $("column-range").textContent = `${XLSX.utils.encode_col(state.column)}–${XLSX.utils.encode_col(endColumn)} / ${XLSX.utils.encode_col(range.e.c)} 列`;
  $("previous-rows").disabled = state.row <= range.s.r;
  $("next-rows").disabled = endRow >= range.e.r;
  $("previous-columns").disabled = state.column <= range.s.c;
  $("next-columns").disabled = endColumn >= range.e.c;
  if (!sheet["!ref"]) {
    $("preview").replaceChildren(element("p", "sheet-empty", "此工作表为空。"));
    return;
  }
  const table = element("table", "sheet-grid");
  table.setAttribute("aria-label", workbook.SheetNames[state.sheet]);
  table.style.zoom = state.zoom;
  const colgroup = element("colgroup");
  const rowNumbers = element("col");
  rowNumbers.style.width = "44px";
  colgroup.append(rowNumbers);
  const header = element("tr");
  header.append(element("th", "", ""));
  for (let column = state.column; column <= endColumn; column++) {
    const col = element("col");
    const storedWidth = sheet["!cols"]?.[column];
    col.style.width = `${Math.min(720, Math.max(72, storedWidth?.wpx || (storedWidth?.wch || 13) * 7 + 12))}px`;
    colgroup.append(col);
    const th = element("th", "", XLSX.utils.encode_col(column));
    th.scope = "col";
    header.append(th);
  }
  const thead = element("thead");
  thead.append(header);
  const tbody = element("tbody");
  const visibleMerges = (sheet["!merges"] || []).filter((merge) => merge.e.r >= state.row && merge.s.r <= endRow && merge.e.c >= state.column && merge.s.c <= endColumn);
  const mergedCells = new Map();
  for (const merge of visibleMerges) {
    const firstRow = Math.max(merge.s.r, state.row);
    const firstColumn = Math.max(merge.s.c, state.column);
    for (let row = firstRow; row <= Math.min(merge.e.r, endRow); row++) {
      for (let column = firstColumn; column <= Math.min(merge.e.c, endColumn); column++) {
        mergedCells.set(`${row}:${column}`, row === firstRow && column === firstColumn ? merge : false);
      }
    }
  }
  for (let row = state.row; row <= endRow; row++) {
    const tr = element("tr");
    const th = element("th", "", String(row + 1));
    th.scope = "row";
    tr.append(th);
    const height = sheet["!rows"]?.[row]?.hpt;
    if (height) tr.style.height = `${Math.max(22, height)}pt`;
    for (let column = state.column; column <= endColumn; column++) {
      const merge = mergedCells.get(`${row}:${column}`);
      if (merge === false) continue;
      const address = XLSX.utils.encode_cell(merge ? merge.s : { r: row, c: column });
      const cell = sheet[address];
      const td = element("td", "", cell?.f && cell.v === undefined ? "（无缓存值）" : cell ? (cell.w ?? XLSX.utils.format_cell(cell)) : "");
      td.dataset.address = address;
      td.tabIndex = 0;
      if (merge) {
        td.rowSpan = Math.min(merge.e.r, endRow) - row + 1;
        td.colSpan = Math.min(merge.e.c, endColumn) - column + 1;
      }
      applyCellStyle(td, cell);
      td.addEventListener("click", () => inspectCell(address, cell, td));
      td.addEventListener("keydown", (event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); inspectCell(address, cell, td); } });
      tr.append(td);
    }
    tbody.append(tr);
  }
  table.append(colgroup, thead, tbody);
  $("preview").replaceChildren(table);
  $("preview").scrollTop = 0;
  $("preview").scrollLeft = 0;
}

function moveSheet(rows, columns) {
  const workbook = state.document.workbook;
  const sheet = workbook.Sheets[workbook.SheetNames[state.sheet]];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  state.row = Math.max(range.s.r, Math.min(range.e.r, state.row + rows));
  state.column = Math.max(range.s.c, Math.min(range.e.c, state.column + columns));
  renderSheet();
}

async function changePage(page) {
  if (state.busy) return;
  state.page = Math.max(1, Math.min(state.pages, Number.isFinite(page) ? page : 1));
  await renderPage();
}

function availableWidth() {
  const style = getComputedStyle($("preview"));
  return $("preview").clientWidth - parseFloat(style.paddingLeft) - parseFloat(style.paddingRight);
}

function applyDocumentZoom() {
  const doc = state.document;
  if (doc?.kind === "docx") {
    const page = doc.host.querySelector("section.docx");
    if (state.fit && page) {
      lastWidth = availableWidth();
      state.zoom = Math.min(1, lastWidth / page.offsetWidth);
    }
    doc.host.style.zoom = state.zoom;
  } else if (doc?.kind === "xlsx") {
    const table = $("preview").querySelector(".sheet-grid");
    if (table) {
      const oldZoom = Number(table.style.zoom) || 1;
      const top = $("preview").scrollTop / oldZoom;
      const left = $("preview").scrollLeft / oldZoom;
      table.style.zoom = state.zoom;
      $("preview").scrollTop = top * state.zoom;
      $("preview").scrollLeft = left * state.zoom;
    }
  }
}

function changeZoom(zoom, fit = false) {
  if (state.busy || !state.document) return;
  state.fit = fit;
  state.zoom = Math.max(.5, Math.min(3, zoom));
  if (["docx", "xlsx"].includes(state.document.kind)) applyDocumentZoom();
  else void renderPage();
  updateControls();
}

$("refresh-files").addEventListener("click", () => loadDirectory(state.directory));
$("file-filter").addEventListener("input", renderFileList);
$("reload-file").addEventListener("click", () => openFile(state.path, true));
$("previous-page").addEventListener("click", () => changePage(state.page - 1));
$("next-page").addEventListener("click", () => changePage(state.page + 1));
$("page-number").addEventListener("change", () => changePage(Math.trunc(Number($("page-number").value))));
$("zoom-in").addEventListener("click", () => changeZoom(state.zoom * 1.1));
$("zoom-out").addEventListener("click", () => changeZoom(state.zoom / 1.1));
$("zoom-reset").addEventListener("click", () => changeZoom(1, true));
$("previous-rows").addEventListener("click", () => moveSheet(-rowsPerPage, 0));
$("next-rows").addEventListener("click", () => moveSheet(rowsPerPage, 0));
$("previous-columns").addEventListener("click", () => moveSheet(0, -columnsPerPage));
$("next-columns").addEventListener("click", () => moveSheet(0, columnsPerPage));
$("go-to-cell").addEventListener("click", () => {
  const address = $("cell-address").value.trim().toUpperCase();
  if (!/^[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(address)) { $("cell-value").textContent = "请输入单元格地址，例如 A1 或 AB120。"; return; }
  const workbook = state.document.workbook;
  const sheet = workbook.Sheets[workbook.SheetNames[state.sheet]];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
  const target = XLSX.utils.decode_cell(address);
  if (target.r < range.s.r || target.r > range.e.r || target.c < range.s.c || target.c > range.e.c) { $("cell-value").textContent = "此单元格不在工作表的已用区域内。"; return; }
  state.row = Math.max(range.s.r, Math.floor(target.r / rowsPerPage) * rowsPerPage);
  state.column = Math.max(range.s.c, Math.floor(target.c / columnsPerPage) * columnsPerPage);
  renderSheet();
  const td = $("preview").querySelector(`[data-address="${address}"]`);
  if (td) { inspectCell(address, sheet[address], td); td.scrollIntoView({ block: "nearest", inline: "nearest" }); }
});
$("cell-address").addEventListener("keydown", (event) => { if (event.key === "Enter") $("go-to-cell").click(); });
let resizeTimer;
let lastWidth = 0;
new ResizeObserver(() => {
  clearTimeout(resizeTimer);
  resizeTimer = setTimeout(function fitAfterResize() {
    if (!state.fit || !state.document) return;
    if (state.busy) { resizeTimer = setTimeout(fitAfterResize, 200); return; }
    const width = availableWidth();
    if (Math.abs(width - lastWidth) <= 2) return;
    lastWidth = width;
    changeZoom(state.zoom, true);
  }, 200);
}).observe($("preview"));
window.addEventListener("pagehide", () => { events?.close(); releaseDocument(state.document); });
window.addEventListener("hashchange", () => {
  const next = new URLSearchParams(location.hash.slice(1));
  const nextToken = next.get("token");
  if (!nextToken) return;
  // A host may navigate its existing browser to the same URL with a new fragment.
  // Reload after saving the fragment so initialization also switches embed mode.
  sessionStorage.setItem("paseo-office-token", nextToken);
  sessionStorage.setItem("paseo-office-embed", next.get("embed") === "1" ? "1" : "0");
  history.replaceState(null, "", location.pathname + location.search);
  location.reload();
});

if (token) {
  if (embedded) connectSession();
  else loadDirectory("");
}
else { $("list-status").textContent = "请从 Paseo 插件打开此页面。"; showError(new Error("缺少本地预览会话。请从 Paseo 插件或启动命令提供的完整链接打开。")); }
