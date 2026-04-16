const CTRL_SUFFIX = "_Ctrl-Configuration";
const DEFAULT_TEMPLATE_PATH = "./default_template.json";

const workbookInput = document.getElementById("workbookFile");
const syscadInput = document.getElementById("syscadFile");
const templateInput = document.getElementById("templateFile");
const convertBtn = document.getElementById("convertBtn");
const logEl = document.getElementById("log");

function log(message) {
  logEl.textContent += `${message}\n`;
}

function clearLog() {
  logEl.textContent = "";
}

function text(value) {
  if (value === undefined || value === null) {
    return "";
  }
  return String(value).trim();
}

function num(value) {
  const v = text(value);
  if (!v) {
    return 0;
  }
  const parsed = Number(v);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizedToken(line) {
  let token = line.replace(/;+$/, "").trim();
  if (token.toUpperCase().startsWith("CDLG-")) {
    const parts = token.split("-");
    if (parts.length >= 4) {
      parts.splice(parts.length - 2, 1);
      token = parts.join("-");
    }
  }
  return token;
}

function deriveConfigId(fileName) {
  const stem = fileName.replace(/\.[^.]+$/, "");
  if (stem.endsWith(CTRL_SUFFIX)) {
    return stem.slice(0, -CTRL_SUFFIX.length);
  }
  return stem;
}

function getCellValue(sheet, row1, col1) {
  const ref = XLSX.utils.encode_cell({ r: row1 - 1, c: col1 - 1 });
  const cell = sheet[ref];
  return cell ? cell.v : "";
}

function setCellValue(sheet, row1, col1, value) {
  const ref = XLSX.utils.encode_cell({ r: row1 - 1, c: col1 - 1 });
  sheet[ref] = { t: typeof value === "number" ? "n" : "s", v: value };
}

function sheetMaxRow(sheet) {
  if (!sheet["!ref"]) {
    return 1;
  }
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  return range.e.r + 1;
}

function buildRowLookup(sheet1) {
  const out = new Map();
  const max = sheetMaxRow(sheet1);
  for (let row = 1; row <= max; row += 1) {
    const key = text(getCellValue(sheet1, row, 4)).toLowerCase();
    if (key && !out.has(key)) {
      out.set(key, row);
    }
  }
  return out;
}

function buildReadLookup(readSheet) {
  const out = new Map();
  const max = sheetMaxRow(readSheet);
  for (let row = 1; row <= max; row += 1) {
    const key = text(getCellValue(readSheet, row, 3)).toLowerCase();
    if (!key) {
      continue;
    }
    const idx = Math.trunc(num(getCellValue(readSheet, row, 4)));
    if (idx > 0 && !out.has(key)) {
      out.set(key, idx);
    }
  }
  return out;
}

function runSeqLogic(seqSheet) {
  const total = Math.trunc(num(getCellValue(seqSheet, 1, 10)));

  for (let r = 1; r <= 10; r += 1) {
    setCellValue(seqSheet, r, 15, "");
  }

  if (total <= 0) {
    return;
  }

  const rows = [];
  const max = sheetMaxRow(seqSheet);
  let k = 2;
  let i = 1;
  while (i < total + 1 && k <= max) {
    const col7 = text(getCellValue(seqSheet, k, 7));
    if (col7) {
      rows.push([Math.trunc(num(col7)), text(getCellValue(seqSheet, k, 8)), text(getCellValue(seqSheet, k, 9))]);
      i += 1;
    }
    k += 1;
  }

  rows.sort((a, b) => a[0] - b[0]);
  rows.forEach((item, idx) => {
    setCellValue(seqSheet, idx + 1, 15, item[0]);
  });
}

function convertRawValue(raw) {
  if (raw.length >= 2 && raw.startsWith('"') && raw.endsWith('"')) {
    return raw.slice(1, -1);
  }

  const lower = raw.toLowerCase();
  if (lower === "true") {
    return true;
  }
  if (lower === "false") {
    return false;
  }
  if (raw === "#N/A") {
    return raw;
  }

  const parsed = Number(raw);
  if (!Number.isNaN(parsed)) {
    return raw.includes(".") || raw.includes("e") || raw.includes("E") ? parsed : Math.trunc(parsed);
  }

  return raw;
}

function applyUpdates(templateObj, updates) {
  const variables = templateObj?.applicationSettings?.Variables;
  if (!Array.isArray(variables)) {
    throw new Error("Template JSON missing applicationSettings.Variables array");
  }

  const index = new Map();
  variables.forEach((item, i) => {
    if (item && typeof item === "object") {
      const keys = Object.keys(item);
      if (keys.length === 1) {
        index.set(keys[0], i);
      }
    }
  });

  let changed = 0;
  const missing = [];

  updates.forEach((raw, key) => {
    const idx = index.get(key);
    if (idx === undefined) {
      missing.push(key);
      return;
    }

    if (raw === "#N/A") {
      return;
    }

    const current = variables[idx][key];
    let newValue = convertRawValue(raw);

    if (typeof current === "boolean" && typeof newValue === "string") {
      const lower = newValue.toLowerCase();
      if (lower === "true" || lower === "false") {
        newValue = lower === "true";
      }
    }

    variables[idx][key] = newValue;
    changed += 1;
  });

  return { changed, missing };
}

async function readFileAsArrayBuffer(file) {
  return file.arrayBuffer();
}

async function readFileAsText(file) {
  return file.text();
}

async function loadTemplate() {
  if (templateInput.files && templateInput.files[0]) {
    const raw = await readFileAsText(templateInput.files[0]);
    return JSON.parse(raw);
  }

  const response = await fetch(DEFAULT_TEMPLATE_PATH, { cache: "no-store" });
  if (!response.ok) {
    throw new Error("Default template could not be loaded. Upload template JSON manually.");
  }
  return response.json();
}

function buildRows(workbook, syscadText, syscadFileName) {
  const readSheet = workbook.Sheets.Read;
  const sheet1 = workbook.Sheets.Sheet1;
  const seqSheet = workbook.Sheets.Seq;

  if (!readSheet || !sheet1 || !seqSheet) {
    throw new Error("Workbook must contain sheets: Read, Sheet1, Seq");
  }

  const warnings = [];
  const lines = syscadText
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);

  const sourceFileStem = syscadFileName.replace(/\.[^.]+$/, "");
  setCellValue(readSheet, 1, 8, syscadFileName);
  setCellValue(readSheet, 1, 9, sourceFileStem);

  const rowLookup = buildRowLookup(sheet1);
  const readLookup = buildReadLookup(readSheet);

  const output = [];

  let i = 1;
  const maxSheet1 = sheetMaxRow(sheet1);
  while (true) {
    const marker = text(getCellValue(sheet1, i, 18));
    if (marker === "#end") {
      break;
    }

    output.push({
      signal_name: "",
      description: text(getCellValue(sheet1, i + 1, 18)),
      target_key: "",
      target_value: "",
    });

    i += 1;
    if (i > maxSheet1) {
      warnings.push("Reached end of Sheet1 before pre-setup #end marker.");
      break;
    }
  }

  let x = Math.trunc(num(getCellValue(sheet1, i, 19)));
  if (!x) {
    warnings.push("Could not parse post-setup start row from Sheet1 column S.");
  }

  runSeqLogic(seqSheet);
  for (let z = 2; z < 22; z += 1) {
    output.push({
      signal_name: "",
      description: "",
      target_key: text(getCellValue(seqSheet, z, 19)),
      target_value: text(getCellValue(seqSheet, z, 20)),
    });
  }

  lines.forEach((line) => {
    const src = normalizedToken(line);
    if (!src || src.slice(0, 4).toLowerCase() === "calc") {
      return;
    }

    const key = src.toLowerCase();
    let k = readLookup.get(key) || rowLookup.get(key);
    if (!k) {
      warnings.push(`No Sheet1 match found for token: ${src}`);
      return;
    }

    output.push({
      signal_name: src,
      description: text(getCellValue(sheet1, k, 18)),
      target_key: text(getCellValue(sheet1, k, 19)),
      target_value: text(getCellValue(sheet1, k, 20)),
    });

    while (k + 1 <= maxSheet1 && text(getCellValue(sheet1, k + 1, 4)) === "") {
      output.push({
        signal_name: src,
        description: text(getCellValue(sheet1, k + 1, 18)),
        target_key: text(getCellValue(sheet1, k + 1, 19)),
        target_value: text(getCellValue(sheet1, k + 1, 20)),
      });
      k += 1;
    }
  });

  if (x > 0) {
    while (x <= maxSheet1 && text(getCellValue(sheet1, x, 18)) !== "#end") {
      output.push({
        signal_name: "",
        description: text(getCellValue(sheet1, x, 18)),
        target_key: "",
        target_value: "",
      });
      x += 1;
    }
  }

  return { rows: output, warnings };
}

function rowsToUpdates(rows) {
  const updates = new Map();
  rows.forEach((row) => {
    if (row.target_key) {
      updates.set(row.target_key, String(row.target_value).replace(/,/g, "."));
    }
  });
  return updates;
}

function downloadJson(filename, objectValue) {
  const blob = new Blob([JSON.stringify(objectValue)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

convertBtn.addEventListener("click", async () => {
  clearLog();

  if (!workbookInput.files || !workbookInput.files[0]) {
    log("Please upload an Excel workbook.");
    return;
  }
  if (!syscadInput.files || !syscadInput.files[0]) {
    log("Please upload a SysCAD txt file.");
    return;
  }

  convertBtn.disabled = true;
  try {
    log("Loading template...");
    const templateObj = await loadTemplate();

    log("Reading workbook...");
    const workbookBuf = await readFileAsArrayBuffer(workbookInput.files[0]);
    const wb = XLSX.read(workbookBuf, { type: "array", cellFormula: false, cellNF: false, cellHTML: false });

    log("Reading SysCAD txt...");
    const syscadText = await readFileAsText(syscadInput.files[0]);

    log("Running conversion...");
    const { rows, warnings } = buildRows(wb, syscadText, syscadInput.files[0].name);
    const updates = rowsToUpdates(rows);
    const { changed, missing } = applyUpdates(templateObj, updates);

    const configId = deriveConfigId(syscadInput.files[0].name);
    downloadJson(`${configId}.json`, templateObj);

    log(`Config ID: ${configId}`);
    log(`Rows generated: ${rows.length}`);
    log(`Updates requested: ${updates.size}`);
    log(`Updates applied: ${changed}`);
    log(`Missing keys: ${missing.length}`);
    if (warnings.length) {
      log(`Warnings (${warnings.length}):`);
      warnings.slice(0, 20).forEach((w) => log(`- ${w}`));
    }
    log("Done. JSON downloaded.");
  } catch (error) {
    log(`Error: ${error.message}`);
  } finally {
    convertBtn.disabled = false;
  }
});
