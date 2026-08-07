const els = {
  search: document.getElementById("search"),
  refreshBtn: document.getElementById("refreshBtn"),
  status: document.getElementById("status"),
  tableWrap: document.getElementById("tableWrap"),
};

let headers = [];
let rows = [];

function buildUrl() {
  const { sheetId, sheetName } = SHEET_CONFIG;
  const params = new URLSearchParams({ tqx: "out:csv", sheet: sheetName });
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?${params.toString()}`;
}

// Parser CSV simple que soporta comillas, comas y saltos de línea dentro de celdas.
function parseCsv(text) {
  const result = [];
  let row = [];
  let field = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const next = text[i + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        field += '"';
        i++;
      } else if (char === '"') {
        inQuotes = false;
      } else {
        field += char;
      }
    } else if (char === '"') {
      inQuotes = true;
    } else if (char === ",") {
      row.push(field);
      field = "";
    } else if (char === "\n") {
      row.push(field);
      result.push(row);
      row = [];
      field = "";
    } else if (char === "\r") {
      // ignorar, lo maneja el \n siguiente
    } else {
      field += char;
    }
  }
  if (field.length > 0 || row.length > 0) {
    row.push(field);
    result.push(row);
  }
  return result;
}

function renderTable(filterText) {
  const filter = (filterText || "").trim().toLowerCase();
  const filteredRows = filter
    ? rows.filter((r) => r.some((cell) => cell.toLowerCase().includes(filter)))
    : rows;

  if (headers.length === 0) {
    els.tableWrap.innerHTML = "<p>Sin datos.</p>";
    return;
  }

  const thead = `<thead><tr>${headers
    .map((h) => `<th>${escapeHtml(h)}</th>`)
    .join("")}</tr></thead>`;

  const tbody = `<tbody>${filteredRows
    .map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`)
    .join("")}</tbody>`;

  els.tableWrap.innerHTML = `<table>${thead}${tbody}</table>`;

  els.status.textContent = `${filteredRows.length} de ${rows.length} filas · Actualizado ${new Date().toLocaleTimeString()}`;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

async function loadData() {
  els.status.textContent = "Actualizando...";
  try {
    const res = await fetch(buildUrl(), { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const text = await res.text();
    const parsed = parseCsv(text).filter((r) => r.length > 1 || r[0] !== "");

    if (parsed.length === 0) throw new Error("Respuesta vacía");

    headers = parsed[0];
    rows = parsed.slice(1);

    renderTable(els.search.value);
  } catch (err) {
    els.status.textContent = `Error al cargar los datos: ${err.message}`;
    console.error(err);
  }
}

els.search.addEventListener("input", () => renderTable(els.search.value));
els.refreshBtn.addEventListener("click", loadData);

loadData();
if (SHEET_CONFIG.refreshIntervalMs > 0) {
  setInterval(loadData, SHEET_CONFIG.refreshIntervalMs);
}
