const EXCEL_FILE = "TORNEO-2026.xlsx";
const SHEET_NAME = "ESCUELITAS";

// Índices de columna dentro de la hoja (0-based), según la estructura real del Excel.
const COL = {
  NAME: 2,
  TROFEOS: 3,
  TORNEO_CANT: 4,
  TORNEO_IMP: 5,
  LIBERADOS: 6,
  BONIF: 7,
  ALMUERZOS_CANT: 8,
  ALMUERZOS_IMP: 9,
  ALOJ_CANT: 10,
  ALOJ_IMP: 11,
  PROFES_CANT: 12,
  PROFES_IMP: 13,
  TOTAL: 14,
  SALDO: 15,
};

// Campos de datos, en el mismo orden en que se muestran las columnas.
const FIELDS = [
  "trofeos",
  "torneoCant",
  "torneoImp",
  "liberados",
  "bonif",
  "almuerzosCant",
  "almuerzosImp",
  "alojCant",
  "alojImp",
  "profesCant",
  "profesImp",
  "total",
  "saldo",
];

const FIELD_COL = {
  trofeos: COL.TROFEOS,
  torneoCant: COL.TORNEO_CANT,
  torneoImp: COL.TORNEO_IMP,
  liberados: COL.LIBERADOS,
  bonif: COL.BONIF,
  almuerzosCant: COL.ALMUERZOS_CANT,
  almuerzosImp: COL.ALMUERZOS_IMP,
  alojCant: COL.ALOJ_CANT,
  alojImp: COL.ALOJ_IMP,
  profesCant: COL.PROFES_CANT,
  profesImp: COL.PROFES_IMP,
  total: COL.TOTAL,
  saldo: COL.SALDO,
};

// Reglas de formato condicional que tiene la columna SALDO en el Excel
// (rango P14:P44 -> filas de datos desde la 3ra escuelita en adelante),
// tomadas directamente de conditionalFormatting/dxfs del archivo.
const SALDO_CF_FIRST_EXCEL_ROW = 14;
const SALDO_CF_LAST_EXCEL_ROW = 44;

const els = {
  refreshBtn: document.getElementById("refreshBtn"),
  saldoFilter: document.getElementById("saldoFilter"),
  status: document.getElementById("status"),
  tableWrap: document.getElementById("tableWrap"),
};

let escuelitas = [];

function money(value) {
  if (value === "" || value === null || value === undefined) return "";
  const n = Number(value);
  if (Number.isNaN(n)) return String(value);
  return n.toLocaleString("en-US");
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Lee el color de relleno "sólido" que tiene una celda en el Excel, si tiene.
function cellFill(sheet, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = sheet[addr];
  if (!cell || !cell.s || cell.s.patternType !== "solid") return null;
  const rgb = cell.s.fgColor && cell.s.fgColor.rgb;
  if (!rgb) return null;
  return `#${rgb.slice(-6)}`;
}

// Reproduce el formato condicional real de la columna SALDO (rojo/verde/blanco
// según el importe), tomado de los dxf del archivo. Devuelve null si no aplica.
function saldoConditionalStyle(value, excelRow) {
  if (typeof value !== "number") return null;
  if (excelRow < SALDO_CF_FIRST_EXCEL_ROW || excelRow > SALDO_CF_LAST_EXCEL_ROW) return null;
  if (value > 0) return { background: "#FF0000", color: "#FFFFFF", bold: true };
  if (value === 0) return { background: "#4EA72E", color: "#FFFFFF", bold: true };
  return { background: "#FFFFFF", color: "#4EA72E", bold: false };
}

function cellStyleAttr(bg, extra) {
  const style = { ...(extra || {}) };
  if (bg && !style.background) style.background = bg;
  const parts = [];
  if (style.background) parts.push(`background-color:${style.background}`);
  if (style.color) parts.push(`color:${style.color}`);
  if (style.bold) parts.push("font-weight:bold");
  return parts.length ? ` style="${parts.join(";")}"` : "";
}

function filterEscuelitas() {
  const mode = els.saldoFilter.value;
  if (mode === "all") return escuelitas;
  return escuelitas.filter((e) => {
    if (typeof e.saldo !== "number") return false;
    if (mode === "gt0") return e.saldo > 0;
    if (mode === "eq0") return e.saldo === 0;
    if (mode === "lt0") return e.saldo < 0;
    return true;
  });
}

function renderTable() {
  const filtered = filterEscuelitas();

  const thead = `
    <thead>
      <tr>
        <th rowspan="2">ESCUELITA</th>
        <th rowspan="2">TROFEOS</th>
        <th colspan="2">TORNEO</th>
        <th rowspan="2">LIBERADOS</th>
        <th rowspan="2">BONIF.</th>
        <th colspan="2">Almuerzos 30,000</th>
        <th colspan="2">Torneo con Alojamiento 100,000</th>
        <th colspan="2">ALMUERZO PROFES</th>
        <th rowspan="2">TOTAL</th>
        <th rowspan="2">SALDO</th>
      </tr>
      <tr>
        <th>Cant</th><th>Importe</th>
        <th>Cant</th><th>Importe</th>
        <th>Cant</th><th>Importe</th>
        <th>Cant</th><th>Importe</th>
      </tr>
    </thead>`;

  const tbody = `<tbody>${filtered
    .map((e) => {
      const cells = [`<td${cellStyleAttr(e.bg.nombre)}>${escapeHtml(e.nombre)}</td>`];
      for (const field of FIELDS) {
        let style;
        if (field === "saldo") {
          style = saldoConditionalStyle(e.saldo, e.excelRow) || { background: e.bg.saldo };
        } else {
          style = { background: e.bg[field] };
        }
        cells.push(`<td${cellStyleAttr(null, style)}>${money(e[field])}</td>`);
      }
      return `<tr>${cells.join("")}</tr>`;
    })
    .join("")}</tbody>`;

  els.tableWrap.innerHTML = `<table>${thead}${tbody}</table>`;
  els.status.textContent = `${filtered.length} de ${escuelitas.length} escuelitas · Actualizado ${new Date().toLocaleTimeString()}`;
}

async function loadData() {
  els.status.textContent = "Actualizando...";
  try {
    const res = await fetch(`${EXCEL_FILE}?t=${Date.now()}`, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buffer = await res.arrayBuffer();
    const wb = XLSX.read(buffer, { type: "array", cellStyles: true });
    const sheet = wb.Sheets[SHEET_NAME];
    if (!sheet) throw new Error(`No se encontró la hoja "${SHEET_NAME}"`);

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const headerRowIndex = rows.findIndex((r) => r[COL.NAME] === "ESCUELITA");
    if (headerRowIndex === -1) throw new Error("No se encontró el encabezado ESCUELITA");

    const dataStart = headerRowIndex + 2;
    escuelitas = [];
    for (let i = dataStart; i < rows.length; i++) {
      const r = rows[i];
      const nombre = r[COL.NAME];
      if (!nombre) break;

      const bg = { nombre: cellFill(sheet, i, COL.NAME) };
      for (const field of FIELDS) {
        bg[field] = cellFill(sheet, i, FIELD_COL[field]);
      }

      escuelitas.push({
        nombre,
        trofeos: r[COL.TROFEOS],
        torneoCant: r[COL.TORNEO_CANT],
        torneoImp: r[COL.TORNEO_IMP],
        liberados: r[COL.LIBERADOS],
        bonif: r[COL.BONIF],
        almuerzosCant: r[COL.ALMUERZOS_CANT],
        almuerzosImp: r[COL.ALMUERZOS_IMP],
        alojCant: r[COL.ALOJ_CANT],
        alojImp: r[COL.ALOJ_IMP],
        profesCant: r[COL.PROFES_CANT],
        profesImp: r[COL.PROFES_IMP],
        total: r[COL.TOTAL],
        saldo: r[COL.SALDO],
        excelRow: i + 1,
        bg,
      });
    }

    renderTable();
  } catch (err) {
    els.status.textContent = `Error al cargar los datos: ${err.message}`;
    console.error(err);
  }
}

els.refreshBtn.addEventListener("click", loadData);
els.saldoFilter.addEventListener("change", renderTable);

loadData();
