// Configuración de la fuente de datos (Google Sheets)
// 1. Abrí tu Google Sheet y copiá el ID que aparece en la URL:
//    https://docs.google.com/spreadsheets/d/ESTE_ES_EL_ID/edit
// 2. Poné el nombre exacto de la pestaña (hoja) que querés leer.
// 3. Compartí el Sheet como "Cualquier persona con el enlace" -> "Lector"
//    (Archivo > Compartir > Acceso general), si no, la consulta va a fallar.
const SHEET_CONFIG = {
  sheetId: "TU_SHEET_ID_AQUI",
  sheetName: "Hoja1",
  // Refresco automático en milisegundos. Poner 0 para desactivarlo.
  refreshIntervalMs: 30000,
};
