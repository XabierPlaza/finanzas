/* Motor de datos local de «Control de finanzas».
 *
 * Reproduce en el navegador las operaciones del servidor del PC (src/app.py) para que la app funcione
 * en el móvil sin PC: mismas rutas (/api/...), mismas validaciones y mismas respuestas.
 * Los datos se guardan en el propio dispositivo (IndexedDB) y se pueden importar/exportar como el mismo
 * Excel que usa el PC. La sincronización opcional con Google Drive vive en drive.js y usa los ganchos
 * `Finanzas.onChange` / `Finanzas.replaceModel` / `Finanzas.pending`.
 */
(function (global) {
  "use strict";

  const CATEGORIAS = ["Pago de Deudas", "Ahorro e Inversión", "Gastos Fijos", "Gastos Variables"];
  const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto",
    "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  const MODOS_AHORRO = ["Porcentaje", "Importe"];
  const RECURRENTES = ["Ahorro e Inversión", "Gastos Fijos"];
  const DB_NAME = "Base_Datos_Finanzas.xlsx";
  const EUR_FMT = '#,##0.00 "€"';

  // estructura del Excel (idéntica a crear_base_datos.py)
  const SHEETS = {
    periodos: { sheet: "Periodos", table: "tPeriodos", key: "Periodo",
      headers: ["Periodo", "Año", "Mes", "Nómina", "Modo_Ahorro", "Pct_Ahorro", "Importe_Ahorro", "Notas"],
      widths: [11, 8, 13, 13, 14, 12, 15, 45], formats: { "Nómina": EUR_FMT, Pct_Ahorro: "0.0%", Importe_Ahorro: EUR_FMT } },
    ingresos: { sheet: "Ingresos", table: "tIngresos", key: "ID", id: ["I", 4],
      headers: ["ID", "Periodo", "Concepto", "Importe", "Notas"],
      widths: [9, 11, 35, 13, 45], formats: { Importe: EUR_FMT } },
    movimientos: { sheet: "Movimientos", table: "tMovimientos", key: "ID", id: ["M", 4],
      headers: ["ID", "Periodo", "Categoría", "Concepto", "Acreedor", "ID_Deuda", "Importe", "Notas"],
      widths: [9, 11, 20, 55, 14, 11, 13, 55], formats: { Importe: EUR_FMT } },
    deudas: { sheet: "Deudas", table: "tDeudas", key: "ID_Deuda", id: ["D", 3],
      headers: ["ID_Deuda", "Acreedor", "Descripción", "Importe_Total", "Pagado_Previo", "Mes_Cuota_Final", "Notas"],
      widths: [11, 14, 42, 15, 15, 17, 55], formats: { Importe_Total: EUR_FMT, Pagado_Previo: EUR_FMT } },
    ajustes: { sheet: "Ajustes", table: "tAjustes", key: "Clave",
      headers: ["Clave", "Valor"], widths: [22, 12], formats: {} },
  };
  // opciones de la app guardadas en el Excel (valen para todos los dispositivos); igual que AJUSTES en app.py
  const AJUSTES = { boton_flotante: ["Sí", "No"] };
  const NUMERIC = new Set(["Nómina", "Pct_Ahorro", "Importe_Ahorro", "Importe", "Importe_Total", "Pagado_Previo", "Año"]);
  const INSTRUCCIONES = [
    "",
    "Este libro es la fuente de datos del Dashboard (Dashboard_Finanzas.exe y la app del móvil).",
    "Cada hoja contiene una tabla de Excel. Añade filas al final de la tabla y se ampliará sola.",
    "",
    "Periodos     → una fila por mes: nómina y objetivo de ahorro (Modo_Ahorro = Porcentaje o Importe).",
    "Ingresos     → ingresos extra opcionales de cada mes (se suman a la nómina).",
    "Movimientos  → importes del mes: gastos fijos, ahorro, pagos de deudas y gastos variables.",
    "Deudas       → registro de deudas. 'Pagado_Previo' es lo pagado antes de usar esta base de datos.",
    "Listas       → valores permitidos para los desplegables (categorías, meses, acreedores).",
    "Ajustes      → opciones de la app (p. ej. boton_flotante = Sí/No: botón «+» único en el móvil).",
    "",
    "Reglas:",
    " • Periodo con formato AAAA-MM (p. ej. 2026-10).",
    " • Ingresos del mes = Nómina + ingresos extra de ese periodo.",
    " • Ahorro objetivo = (Ingresos − Gastos fijos − Pagos de deudas) × Pct_Ahorro (modo Porcentaje) o Importe_Ahorro en € (modo Importe), sin superar esa cantidad.",
    " • Presupuesto de gastos variables = Ingresos − Fijos − Ahorro objetivo − Pagos de deudas.",
    " • Al crear un mes nuevo desde el dashboard se copian los gastos fijos y el ahorro/inversión del último mes.",
    " • En 'Pago de Deudas', si indicas el ID_Deuda, ese pago se suma al 'Total pagado' de la deuda.",
    "   Total pagado = Pagado_Previo + pagos de Movimientos con ese ID_Deuda.",
  ];

  class UserError extends Error {}

  /* ------------------------------------------------------------ validación (igual que app.py) */
  const isEmpty = v => v === null || v === undefined || v === "";
  function num(v, field, required = true) {
    if (isEmpty(v)) { if (required) throw new UserError(`Falta el campo ${field}.`); return null; }
    if (typeof v === "number") return v;
    const x = Number(String(v).replace("€", "").trim().replace(",", "."));
    if (!isFinite(x)) throw new UserError(`${field} debe ser un número.`);
    return x;
  }
  function pctv(v) {
    let p = num(v, "% Ahorro");
    if (p > 1) p /= 100;  // se espera una fracción (0,602); >1 se interpreta como %
    if (!(p >= 0 && p <= 1)) throw new UserError("El % de ahorro debe estar entre 0 y 100.");
    return Math.round(p * 10000) / 10000;
  }
  function periodo(v) {
    v = String(v ?? "").trim();
    if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(v)) throw new UserError("El periodo debe tener formato AAAA-MM.");
    return v;
  }
  function txt(v, field, required = false) {
    if (typeof v === "string") v = v.trim();
    if (required && !v) throw new UserError(`Falta el campo ${field}.`);
    return v || null;
  }
  function modo(v) { if (!MODOS_AHORRO.includes(v)) throw new UserError("Modo de ahorro no válido (Porcentaje o Importe)."); return v; }
  function categoria(v) { if (!CATEGORIAS.includes(v)) throw new UserError("Categoría no válida."); return v; }

  const EDITABLE = {
    periodos: { "Nómina": v => num(v, "Nómina"), Modo_Ahorro: modo, Pct_Ahorro: pctv,
      Importe_Ahorro: v => num(v, "Importe de ahorro", false), Notas: v => txt(v, "Notas") },
    ingresos: { Periodo: periodo, Concepto: v => txt(v, "Concepto", true), Importe: v => num(v, "Importe"), Notas: v => txt(v, "Notas") },
    movimientos: { Periodo: periodo, "Categoría": categoria, Concepto: v => txt(v, "Concepto", true), Acreedor: v => txt(v, "Acreedor"),
      ID_Deuda: v => txt(v, "ID_Deuda"), Importe: v => num(v, "Importe"), Notas: v => txt(v, "Notas") },
    deudas: { Acreedor: v => txt(v, "Acreedor", true), "Descripción": v => txt(v, "Descripción", true),
      Importe_Total: v => num(v, "Importe total"), Pagado_Previo: v => num(v, "Pagado previo", false) || 0,
      Mes_Cuota_Final: v => txt(v, "Mes cuota final"), Notas: v => txt(v, "Notas") },
  };

  /* ------------------------------------------------------------ modelo */
  const emptyModel = () => ({ periodos: [], ingresos: [], movimientos: [], deudas: [], ajustes: [], extra: {}, updated: 0, name: DB_NAME });
  let model = emptyModel();
  let pending = [];          // operaciones aún no subidas a Google Drive (las usa drive.js)
  const listeners = new Set();

  function nextId(m, tabla) {
    const [prefix, digits] = SHEETS[tabla].id, key = SHEETS[tabla].key;
    const re = new RegExp(`^${prefix}(\\d+)$`);
    const max = m[tabla].reduce((a, r) => { const x = re.exec(String(r[key] ?? "")); return x ? Math.max(a, +x[1]) : a; }, 0);
    return `${prefix}${String(max + 1).padStart(digits, "0")}`;
  }
  function recurrentesPara(per, movs, cats = RECURRENTES) {
    const out = [];
    for (const cat of cats) {
      if (movs.some(m => m.Periodo === per && m["Categoría"] === cat)) continue;
      const meses = [...new Set(movs.filter(m => m["Categoría"] === cat && m.Periodo && m.Periodo !== per).map(m => m.Periodo))].sort();
      if (!meses.length) continue;
      const previos = meses.filter(p => p < per), origen = previos.length ? previos.at(-1) : meses.at(-1);
      for (const m of movs) if (m.Periodo === origen && m["Categoría"] === cat)
        out.push({ Periodo: per, "Categoría": cat, Concepto: m.Concepto ?? null, Acreedor: m.Acreedor ?? null, ID_Deuda: null,
          Importe: m.Importe ?? null, Notas: `Copiado de ${origen}` });
    }
    return out;
  }
  function find(m, tabla, id) {
    if (!SHEETS[tabla]) throw new UserError("Tabla no válida.");
    if (isEmpty(id)) throw new UserError("Falta el identificador del registro.");
    const key = SHEETS[tabla].key, i = m[tabla].findIndex(r => String(r[key]) === String(id));
    if (i < 0) throw new UserError(`No se encuentra el registro ${id}.`);
    return i;
  }
  function add(m, tabla, rec) {
    const key = SHEETS[tabla].key;
    if (SHEETS[tabla].id && isEmpty(rec[key])) rec[key] = nextId(m, tabla);
    m[tabla].push(rec);
    return rec;
  }

  // Cada operación modifica el modelo `m` y devuelve el mismo «registro» que devolvería el servidor.
  const OPS = {
    "/api/movimiento": (m, p) => add(m, "movimientos", {
      ID: null, Periodo: periodo(p.Periodo), "Categoría": categoria(txt(p["Categoría"], "Categoría", true)),
      Concepto: txt(p.Concepto, "Concepto", true), Acreedor: txt(p.Acreedor, "Acreedor"), ID_Deuda: txt(p.ID_Deuda, "ID_Deuda"),
      Importe: num(p.Importe, "Importe"), Notas: txt(p.Notas, "Notas") }),
    "/api/ingreso": (m, p) => add(m, "ingresos", {
      ID: null, Periodo: periodo(p.Periodo), Concepto: txt(p.Concepto, "Concepto", true), Importe: num(p.Importe, "Importe"), Notas: txt(p.Notas, "Notas") }),
    "/api/deuda": (m, p) => add(m, "deudas", {
      ID_Deuda: null, Acreedor: txt(p.Acreedor, "Acreedor", true), "Descripción": txt(p["Descripción"], "Descripción", true),
      Importe_Total: num(p.Importe_Total, "Importe total"), Pagado_Previo: num(p.Pagado_Previo, "Pagado previo", false) || 0,
      Mes_Cuota_Final: txt(p.Mes_Cuota_Final, "Mes cuota final"), Notas: txt(p.Notas, "Notas") }),
    "/api/periodo": (m, p) => {
      const per = periodo(p.Periodo), md = modo(p.Modo_Ahorro || "Porcentaje");
      const rec = { Periodo: per, "Año": +per.slice(0, 4), Mes: MESES[+per.slice(5) - 1], "Nómina": num(p["Nómina"], "Nómina"), Modo_Ahorro: md,
        Pct_Ahorro: isEmpty(p.Pct_Ahorro) ? null : pctv(p.Pct_Ahorro),
        Importe_Ahorro: num(p.Importe_Ahorro, "Importe de ahorro", md === "Importe"), Notas: txt(p.Notas, "Notas") };
      if (md === "Porcentaje" && rec.Pct_Ahorro === null) throw new UserError("Falta el % de ahorro.");
      if (m.periodos.some(r => r.Periodo === per)) throw new UserError(`El mes ${per} ya existe.`);
      m.periodos.push(rec);
      // qué copiar del último mes: gastos fijos y/o ahorro e inversión («copiar_recurrentes» = ambos, por compatibilidad)
      const todo = p.copiar_recurrentes ?? true;
      const cats = [["Gastos Fijos", p.copiar_fijos], ["Ahorro e Inversión", p.copiar_ahorro]].filter(([, v]) => v ?? todo).map(([c]) => c);
      const copia = cats.length ? recurrentesPara(per, m.movimientos, cats) : [];
      copia.forEach(c => add(m, "movimientos", { ID: null, ...c }));
      return { ...rec, copiados: copia.length };
    },
    "/api/periodo/recurrentes": (m, p) => {
      const per = periodo(p.Periodo), copia = recurrentesPara(per, m.movimientos);
      copia.forEach(c => add(m, "movimientos", { ID: null, ...c }));
      return { Periodo: per, copiados: copia.length };
    },
    "/api/actualizar": (m, p) => {
      const tabla = p.tabla, rules = EDITABLE[tabla];
      if (!rules) throw new UserError("Tabla no válida.");
      const changes = {};
      for (const [k, v] of Object.entries(p.campos || {})) if (rules[k]) changes[k] = rules[k](v);
      if (!Object.keys(changes).length) throw new UserError("Nada que actualizar.");
      const i = find(m, tabla, p.id);
      Object.assign(m[tabla][i], changes);
      return { [SHEETS[tabla].key]: String(p.id), ...changes };
    },
    "/api/ajuste": (m, p) => {
      const { clave, valor } = p;
      if (!AJUSTES[clave]?.includes(valor)) throw new UserError("Ajuste no válido.");
      m.ajustes ||= [];
      const r = m.ajustes.find(a => a.Clave === clave);
      if (r) r.Valor = valor; else m.ajustes.push({ Clave: clave, Valor: valor });
      return { Clave: clave, Valor: valor };
    },
    "/api/borrar": (m, p) => {
      const tabla = p.tabla, i = find(m, tabla, p.id), id = String(p.id);
      m[tabla].splice(i, 1);
      let borrados = 1;
      if (tabla === "periodos") for (const dep of ["movimientos", "ingresos"]) {
        const before = m[dep].length; m[dep] = m[dep].filter(r => r.Periodo !== id); borrados += before - m[dep].length;
      }
      if (tabla === "deudas") m.movimientos.forEach(r => { if (r.ID_Deuda === id) r.ID_Deuda = null; });
      return { borrados };
    },
  };

  /* ------------------------------------------------------------ persistencia en el dispositivo (IndexedDB) */
  const IDB_NAME = "control-finanzas", STORE = "kv";
  function idb() {
    return new Promise((res, rej) => {
      const r = indexedDB.open(IDB_NAME, 1);
      r.onupgradeneeded = () => r.result.createObjectStore(STORE);
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    });
  }
  async function kvGet(k) {
    const db = await idb();
    return new Promise((res, rej) => { const t = db.transaction(STORE).objectStore(STORE).get(k); t.onsuccess = () => res(t.result); t.onerror = () => rej(t.error); });
  }
  async function kvSet(k, v) {
    const db = await idb();
    return new Promise((res, rej) => { const t = db.transaction(STORE, "readwrite"); t.objectStore(STORE).put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); });
  }
  async function persist() {
    try { await kvSet("model", model); await kvSet("pending", pending); }
    catch (e) { console.warn("No se pudo guardar en el dispositivo (¿navegación privada?)", e); }
  }
  let loaded = null;
  function ensureLoaded() {
    return loaded ||= (async () => {
      try {
        model = (await kvGet("model")) || emptyModel();
        pending = (await kvGet("pending")) || [];
        if (navigator.storage?.persist) navigator.storage.persist().catch(() => {});  // pide que el navegador no borre los datos
      } catch (e) { console.warn("Almacenamiento local no disponible", e); }
    })();
  }

  /* ------------------------------------------------------------ API compatible con el servidor del PC */
  const clone = o => JSON.parse(JSON.stringify(o));
  const fmtDate = ts => { const d = new Date(ts || Date.now()), z = n => String(n).padStart(2, "0");
    return `${z(d.getDate())}/${z(d.getMonth() + 1)}/${d.getFullYear()} ${z(d.getHours())}:${z(d.getMinutes())}`; };
  function meta() {
    const drive = global.FinanzasDrive?.status?.() || null;
    return { file: model.name || DB_NAME, path: drive?.connected ? "Google Drive" : "Este dispositivo", mtime: model.updated,
      modified: model.updated ? fmtDate(model.updated) : "—", locked: false, local: true, pending: pending.length, drive };
  }
  async function api(path, body) {
    await ensureLoaded();
    if (body === undefined) {
      if (path === "/api/version") return meta();
      if (path === "/api/data") return { meta: meta(), ...clone({ periodos: model.periodos, ingresos: model.ingresos, movimientos: model.movimientos, deudas: model.deudas, ajustes: model.ajustes || [] }) };
      throw new Error("No encontrado");
    }
    if (path === "/api/abrir-excel") { await downloadExcel(); return { ok: true }; }
    const op = OPS[path];
    if (!op) throw new Error("No encontrado");
    const work = clone(model);
    let registro;
    try { registro = op(work, clone(body)); }
    catch (e) { throw new Error(e.message); }
    model = { ...work, updated: Date.now() };
    pending.push({ path, body: clone(body), result: clone(registro), at: Date.now() });
    await persist();
    listeners.forEach(fn => { try { fn({ path, body, registro }); } catch (e) { console.error(e); } });
    return { ok: true, registro };
  }

  /* Reaplica operaciones pendientes sobre otro modelo (p. ej. el recién bajado de Drive).
     Si un alta obtiene otro ID, las operaciones siguientes que lo usan se reescriben. */
  function replay(base, ops) {
    const m = clone(base), remap = {}, kept = [];
    const fix = v => (v != null && remap[v]) ? remap[v] : v;
    for (const o of ops) {
      const body = clone(o.body);
      if (body.id) body.id = fix(body.id);
      if (body.ID_Deuda) body.ID_Deuda = fix(body.ID_Deuda);
      if (body.campos?.ID_Deuda) body.campos.ID_Deuda = fix(body.campos.ID_Deuda);
      try {
        const r = OPS[o.path](m, body);
        const key = { "/api/movimiento": "ID", "/api/ingreso": "ID", "/api/deuda": "ID_Deuda" }[o.path];
        if (key && o.result?.[key] && r[key] !== o.result[key]) remap[o.result[key]] = r[key];
        kept.push({ ...o, body, result: r });
      } catch (e) {
        console.warn("Operación descartada al sincronizar:", o.path, e.message);  // p. ej. se borró ese registro desde el PC
      }
    }
    return { model: m, ops: kept };
  }

  /* ------------------------------------------------------------ Excel
     Lectura con SheetJS (robusta con cualquier .xlsx, también el que escribe el PC) y escritura con ExcelJS
     (genera tablas, formatos y desplegables). Ambas se cargan solo cuando hacen falta. */
  const libs = {};
  function loadLib(globalName, src) {
    if (global[globalName]) return Promise.resolve(global[globalName]);
    return libs[src] ||= new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = src; s.onload = () => res(global[globalName]); s.onerror = () => rej(new Error("No se pudo cargar el módulo de Excel."));
      document.head.appendChild(s);
    });
  }
  const loadExcelJS = () => loadLib("ExcelJS", "exceljs.min.js");
  const loadSheetJS = () => loadLib("XLSX", "xlsx.full.min.js");
  function cellValue(v) {
    if (v === null || v === undefined) return null;
    if (v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}`;
    if (typeof v === "string") { const t = v.trim(); return t || null; }
    return v;
  }
  async function parseExcel(buffer, name = DB_NAME) {
    const XLSX = await loadSheetJS();
    const wb = XLSX.read(buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : buffer, { type: "array", cellDates: true });
    if (!wb.Sheets.Periodos) throw new UserError("Ese archivo no parece la base de datos de finanzas (falta la hoja «Periodos»).");
    const m = emptyModel(); m.name = name;
    for (const [tabla, def] of Object.entries(SHEETS)) {
      const ws = wb.Sheets[def.sheet];
      if (!ws) continue;
      const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: false });
      if (!grid.length) continue;
      const headers = grid[0].map(h => h == null ? "" : String(h).trim());
      const extra = headers.filter(h => h && !def.headers.includes(h) && !["Cuenta", "Saldo_Cuenta"].includes(h));
      if (extra.length) m.extra[tabla] = extra;
      for (let r = 1; r < grid.length; r++) {
        const row = grid[r], rec = {};
        let any = false;
        headers.forEach((h, col) => {
          if (!h) return;
          let v = cellValue(row[col]);
          if (NUMERIC.has(h) && typeof v === "string") { const x = Number(v.replace(/[€%\s]/g, "").replace(/\./g, "").replace(",", ".")); v = isFinite(x) ? x : null; }
          if (v !== null) any = true;
          rec[h] = v;
        });
        if (!any) continue;
        if (tabla === "periodos" && !rec.Modo_Ahorro) rec.Modo_Ahorro = "Porcentaje";
        for (const h of def.headers) if (!(h in rec)) rec[h] = null;
        m[tabla].push(rec);
      }
    }
    return m;
  }
  async function buildExcel(m = model) {
    const ExcelJS = await loadExcelJS();
    const wb = new ExcelJS.Workbook();
    wb.creator = "Control de finanzas";
    const AZUL = "FF1F4E79";
    const ins = wb.addWorksheet("Instrucciones");
    ins.getColumn(1).width = 110;
    ins.getCell("A1").value = "BASE DE DATOS DE FINANZAS PERSONALES";
    ins.getCell("A1").font = { bold: true, size: 14, color: { argb: AZUL } };
    INSTRUCCIONES.forEach((t, i) => { ins.getCell(`A${i + 2}`).value = t; });
    for (const [tabla, def] of Object.entries(SHEETS)) {
      const ws = wb.addWorksheet(def.sheet, { views: [{ state: "frozen", ySplit: 1 }] });
      const headers = [...def.headers, ...(m.extra?.[tabla] || [])];
      const rows = (m[tabla] || []).map(r => headers.map(h => r[h] ?? null));  // modelos antiguos sin «ajustes»
      ws.addTable({ name: def.table, ref: "A1", headerRow: true, totalsRow: false,
        style: { theme: "TableStyleMedium2", showRowStripes: true },
        columns: headers.map(h => ({ name: h, filterButton: true })),
        rows: rows.length ? rows : [headers.map(() => null)] });
      headers.forEach((h, i) => {
        const c = ws.getColumn(i + 1);
        c.width = def.widths[i] || 15;
        if (def.formats[h]) c.numFmt = def.formats[h];
      });
    }
    const li = wb.addWorksheet("Listas");
    const acreedores = [...new Set(m.deudas.map(d => d.Acreedor).filter(Boolean))].sort();
    li.addRow(["Categorias", "Meses", "Acreedores"]);
    for (let i = 0; i < Math.max(CATEGORIAS.length, MESES.length, acreedores.length); i++) li.addRow([CATEGORIAS[i] ?? null, MESES[i] ?? null, acreedores[i] ?? null]);
    li.getRow(1).eachCell(c => { c.font = { bold: true, color: { argb: "FFFFFFFF" } }; c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL } }; });
    ["A", "B", "C"].forEach(k => { li.getColumn(k).width = 22; });
    const lista = (sheet, range, formula) => { try { wb.getWorksheet(sheet).dataValidations.add(range, { type: "list", allowBlank: true, formulae: [formula] }); } catch (e) { /* opcional */ } };
    lista("Movimientos", "C2:C2000", "Listas!$A$2:$A$5");
    lista("Movimientos", "E2:E2000", "Listas!$C$2:$C$50");
    lista("Periodos", "C2:C500", "Listas!$B$2:$B$13");
    lista("Periodos", "E2:E500", '"Porcentaje,Importe"');
    lista("Deudas", "F2:F500", "Listas!$B$2:$B$13");
    return wb.xlsx.writeBuffer();
  }
  async function downloadExcel() {
    await ensureLoaded();
    const buf = await buildExcel();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const file = new File([blob], model.name || DB_NAME, { type: blob.type });
    // en el móvil, el menú «Compartir» permite guardarlo en Archivos/Drive o enviarlo; si no, se descarga
    if (navigator.canShare?.({ files: [file] })) { try { await navigator.share({ files: [file], title: "Base de datos de finanzas" }); return; } catch (e) { if (e.name === "AbortError") return; } }
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = file.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  async function importExcel(file) {
    const m = await parseExcel(await file.arrayBuffer(), file.name || DB_NAME);
    await replaceModel(m, { keepPending: false });
    return { periodos: m.periodos.length, movimientos: m.movimientos.length };
  }
  async function replaceModel(m, { keepPending = false, ops = null } = {}) {
    model = { ...m, updated: m.updated || Date.now() };
    pending = keepPending ? (ops ?? pending) : [];
    await persist();
  }

  global.Finanzas = {
    api, parseExcel, buildExcel, downloadExcel, importExcel, replaceModel, replay, ensureLoaded,
    get model() { return model; }, get pending() { return pending; },
    clearPending: async () => { pending = []; await persist(); },
    onChange: fn => listeners.add(fn), kvGet, kvSet, DB_NAME, recurrentesPara, UserError,
  };
})(typeof window !== "undefined" ? window : globalThis);
