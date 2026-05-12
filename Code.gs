/* ═══════════════════════════════════════════════════════════════
   GASTOS MANU & BARBI — Google Apps Script (Code.gs)

   INSTRUCCIONES:
   1. Abrí script.google.com → Nuevo proyecto
   2. Borrá todo el contenido y pegá este código
   3. Guardá (Ctrl+S), poné cualquier nombre al proyecto
   4. Clic en "Implementar" → "Nueva implementación"
   5. Tipo: Aplicación web
   6. Ejecutar como: Yo (tu cuenta)
   7. Quién tiene acceso: Cualquier persona
   8. Clic en "Implementar" → autorizás → copiás la URL
   9. Pegá esa URL en la app (pestaña Config)
═══════════════════════════════════════════════════════════════ */

const SHEET_NAME   = "Gastos";
const REGLAS_SHEET = "Reglas";

const COLS = [
  "Fecha", "Tipo", "Descripcion", "Moneda", "Monto",
  "Compartido", "MontoIndividual", "MetodoPago",
  "Categoria", "Quien", "Observaciones", "ID"
];

const REGLAS_COLS = [
  "patron", "tipoPatron", "categoria", "quien",
  "compartido", "metodoPago", "prioridad", "hits", "id"
];

function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const result = handleAction(data);
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, msg: "Gastos M&B API activa" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAction(data) {
  switch (data.action) {

    /* ── Gastos ── */
    case "save": {
      const sheet = getOrCreateSheet(SHEET_NAME, COLS);
      const g = data.gasto;
      sheet.appendRow([
        g.fecha, g.tipo, g.descripcion, g.moneda,
        Number(g.monto), g.compartido, Number(g.montoIndividual),
        g.metodoPago, g.categoria, g.quien,
        g.observaciones || "", g.id
      ]);
      return { ok: true };
    }

    case "update": {
      const sheet = getOrCreateSheet(SHEET_NAME, COLS);
      const g    = data.gasto;
      const rows = sheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (String(rows[i][11]) === String(g.id)) {
          sheet.getRange(i + 1, 1, 1, COLS.length).setValues([[
            g.fecha, g.tipo, g.descripcion, g.moneda,
            Number(g.monto), g.compartido, Number(g.montoIndividual),
            g.metodoPago, g.categoria, g.quien,
            g.observaciones || "", g.id
          ]]);
          return { ok: true };
        }
      }
      sheet.appendRow([
        g.fecha, g.tipo, g.descripcion, g.moneda,
        Number(g.monto), g.compartido, Number(g.montoIndividual),
        g.metodoPago, g.categoria, g.quien,
        g.observaciones || "", g.id
      ]);
      return { ok: true };
    }

    case "delete": {
      const sheet = getOrCreateSheet(SHEET_NAME, COLS);
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length - 1; i >= 1; i--) {
        if (String(rows[i][11]) === String(data.id)) {
          sheet.deleteRow(i + 1);
          return { ok: true };
        }
      }
      return { ok: true };
    }

    case "list": {
      const sheet = getOrCreateSheet(SHEET_NAME, COLS);
      const all = sheet.getDataRange().getValues();
      const gastos = all.slice(1)
        .filter(r => r[11])
        .map(r => ({
          fecha:           formatFecha(r[0]),
          tipo:            String(r[1]),
          descripcion:     String(r[2]),
          moneda:          String(r[3]),
          monto:           Number(r[4]),
          compartido:      String(r[5]),
          montoIndividual: Number(r[6]),
          metodoPago:      String(r[7]),
          categoria:       String(r[8]),
          quien:           String(r[9]),
          observaciones:   String(r[10] || ""),
          id:              String(r[11])
        }))
        .reverse();
      return { ok: true, gastos };
    }

    /* Inserta múltiples gastos en un solo request (usado por el importador) */
    case "saveBatch": {
      const sheet = getOrCreateSheet(SHEET_NAME, COLS);
      const gastosList = data.gastos || [];
      for (const g of gastosList) {
        sheet.appendRow([
          g.fecha, g.tipo, g.descripcion, g.moneda,
          Number(g.monto), g.compartido, Number(g.montoIndividual),
          g.metodoPago, g.categoria, g.quien,
          g.observaciones || "", g.id
        ]);
      }
      return { ok: true, saved: gastosList.length };
    }

    /* ── Reglas ── */
    case "listReglas": {
      const sheet = getOrCreateSheet(REGLAS_SHEET, REGLAS_COLS);
      return handleListReglas(sheet);
    }

    case "saveReglas": {
      const sheet = getOrCreateSheet(REGLAS_SHEET, REGLAS_COLS);
      return handleSaveReglas(data.reglas || [], sheet);
    }

    default:
      return { ok: false, error: "Accion desconocida: " + data.action };
  }
}

/* ── Reglas helpers ── */

function handleListReglas(sheet) {
  const all = sheet.getDataRange().getValues();
  if (all.length <= 1) return { ok: true, reglas: [] };
  const reglas = all.slice(1)
    .filter(r => r[8])
    .map(r => ({
      patron:     String(r[0] || ""),
      tipoPatron: String(r[1] || "contiene"),
      categoria:  String(r[2] || ""),
      quien:      String(r[3] || ""),
      compartido: String(r[4] || ""),
      metodoPago: String(r[5] || ""),
      prioridad:  Number(r[6] || 0),
      hits:       Number(r[7] || 0),
      id:         String(r[8])
    }));
  return { ok: true, reglas };
}

function handleSaveReglas(reglas, sheet) {
  if (!reglas.length) return { ok: true };

  const all = sheet.getDataRange().getValues();
  const existingRow = {};
  for (let i = 1; i < all.length; i++) {
    if (all[i][8]) existingRow[String(all[i][8])] = i + 1;
  }

  for (const r of reglas) {
    const id  = r.id || uid_gs();
    const row = [
      r.patron     || "",
      r.tipoPatron || "contiene",
      r.categoria  || "",
      r.quien      || "",
      r.compartido || "",
      r.metodoPago || "",
      Number(r.prioridad || 0),
      Number(r.hits      || 0),
      id
    ];
    if (existingRow[String(r.id)]) {
      sheet.getRange(existingRow[String(r.id)], 1, 1, REGLAS_COLS.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
  }
  return { ok: true };
}

/* ── Sheet helper ── */

function getOrCreateSheet(name, cols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh   = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
    const header = sh.getRange(1, 1, 1, cols.length);
    header.setValues([cols]);
    header.setBackground("#1a1a2e");
    header.setFontColor("#ffffff");
    header.setFontWeight("bold");
    sh.setFrozenRows(1);
    if (name === SHEET_NAME) sh.setColumnWidth(3, 200);
  }

  return sh;
}

function uid_gs() {
  return Math.random().toString(36).slice(2) + new Date().getTime().toString(36);
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* Convierte la celda de fecha a "YYYY-MM-DD" sin importar si
   Sheets la guardó como Date, número de serie, o string */
function formatFecha(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, "0");
    const d = String(val.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + d;
  }
  const s = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) {
    const y = parsed.getFullYear();
    const m = String(parsed.getMonth() + 1).padStart(2, "0");
    const d = String(parsed.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + d;
  }
  return s;
}
