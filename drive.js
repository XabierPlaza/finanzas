/* Sincronización opcional de la base de datos con Google Drive.
 *
 * - La app sigue funcionando siempre en local (datos-local.js); Drive solo es un sitio donde guardar
 *   el Excel y compartirlo entre dispositivos (móvil, ordenador…).
 * - Permiso mínimo: drive.file → la app solo ve el archivo que el usuario elige con el selector de Drive
 *   o el que la propia app crea. No puede ver el resto de tu Drive.
 * - Antes de subir se compara la versión del archivo en Drive. Si cambió (p. ej. desde otro dispositivo), se baja,
 *   se reaplican encima los cambios pendientes de este dispositivo y después se sube: no se pisa nada.
 */
(function (global) {
  "use strict";
  const F = global.Finanzas;
  const SCOPE = "https://www.googleapis.com/auth/drive.file";
  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const API = "https://www.googleapis.com/drive/v3/files";
  const UPLOAD = "https://www.googleapis.com/upload/drive/v3/files";

  // Proyecto «Finanzas» de Google Cloud. No son secretos: el ID de cliente solo admite el origen
  // https://xabierplaza.github.io y la clave solo sirve para Google Picker desde ese sitio.
  const DEFAULT_CFG = location.hostname === "xabierplaza.github.io" ? {
    clientId: "73774599492-08m9ak2susj02mrhs0fo1bis8sfcm2t1.apps.googleusercontent.com",
    apiKey: "AIzaSyBYqxncRj0APWWigN5ESr7GMhKJAWtgAUA",
    appId: "73774599492",
  } : null;
  let cfg = null;                            // { clientId, apiKey, appId } (DEFAULT_CFG o los del usuario)
  let state = { connected: false };          // { connected, fileId, name, version, lastSync, error }
  let token = null, tokenExp = 0, tokenClient = null, syncing = null, timer = null;
  const listeners = new Set();
  const emit = () => listeners.forEach(fn => { try { fn(status()); } catch (e) { console.error(e); } });
  const save = () => F.kvSet("driveState", state).catch(() => {});

  const scripts = {};
  function loadScript(src) {
    return scripts[src] ||= new Promise((res, rej) => {
      const s = document.createElement("script");
      s.src = src; s.async = true; s.onload = res;
      s.onerror = () => { delete scripts[src]; rej(new Error("No se puede contactar con Google. ¿Hay conexión?")); };
      document.head.appendChild(s);
    });
  }

  /* ---------------- autenticación (Google Identity Services, sin contraseñas en la app) */
  async function getToken(interactive) {
    if (token && Date.now() < tokenExp - 60000) return token;
    if (!cfg?.clientId) throw new Error("Falta configurar el acceso a Google Drive.");
    if (!interactive) throw Object.assign(new Error("Pulsa «Sincronizar ahora» para volver a conectar con Google."), { needsAuth: true });
    await loadScript("https://accounts.google.com/gsi/client");
    return new Promise((res, rej) => {
      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: cfg.clientId, scope: SCOPE,
        callback: r => {
          if (r.error) return rej(new Error(r.error_description || "Google no ha dado acceso."));
          token = r.access_token; tokenExp = Date.now() + (r.expires_in || 3600) * 1000;
          try { sessionStorage.setItem("gdToken", JSON.stringify({ token, tokenExp })); } catch (e) { /* opcional */ }
          res(token);
        },
        error_callback: e => rej(new Error(e?.type === "popup_closed" ? "Se cerró la ventana de Google." : "No se pudo abrir el inicio de sesión de Google.")),
      });
      tokenClient.requestAccessToken({ prompt: state.connected ? "" : "consent" });
    });
  }
  async function gfetch(url, opts = {}, interactive = false) {
    const t = await getToken(interactive);
    const r = await fetch(url, { ...opts, headers: { ...(opts.headers || {}), Authorization: `Bearer ${t}` } });
    if (r.status === 401) { token = null; throw Object.assign(new Error("La sesión de Google ha caducado. Pulsa «Sincronizar ahora»."), { needsAuth: true }); }
    if (r.status === 404) throw new Error("No se encuentra el archivo en Drive (¿se borró o se movió a la papelera?).");
    if (!r.ok) throw new Error(`Google Drive respondió ${r.status}.`);
    return r;
  }
  const meta = (id, i) => gfetch(`${API}/${id}?fields=id,name,version,modifiedTime,trashed`, {}, i).then(r => r.json());
  const download = (id, i) => gfetch(`${API}/${id}?alt=media`, {}, i).then(r => r.arrayBuffer());
  const upload = async (id, buf, i) => (await gfetch(`${UPLOAD}/${id}?uploadType=media&fields=id,name,version`, {
    method: "PATCH", headers: { "Content-Type": XLSX_MIME }, body: new Blob([buf], { type: XLSX_MIME }) }, i)).json();

  async function createFile(buf, name, i) {
    // carpeta «Finanzas» en Mi unidad
    const q = encodeURIComponent("name='Finanzas' and mimeType='application/vnd.google-apps.folder' and trashed=false");
    const found = await (await gfetch(`${API}?q=${q}&fields=files(id)`, {}, i)).json();
    let folderId = found.files?.[0]?.id;
    if (!folderId) folderId = (await (await gfetch(`${API}?fields=id`, { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ name: "Finanzas", mimeType: "application/vnd.google-apps.folder" }) }, i)).json()).id;
    const boundary = "fz" + Math.random().toString(36).slice(2);
    const body = new Blob([
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n`,
      JSON.stringify({ name, parents: [folderId], mimeType: XLSX_MIME }),
      `\r\n--${boundary}\r\nContent-Type: ${XLSX_MIME}\r\n\r\n`, buf, `\r\n--${boundary}--`]);
    return (await gfetch(`${UPLOAD}?uploadType=multipart&fields=id,name,version`, { method: "POST",
      headers: { "Content-Type": `multipart/related; boundary=${boundary}` }, body }, i)).json();
  }

  function pick() {
    return new Promise(async (res, rej) => {
      try {
        const t = await getToken(true);
        await loadScript("https://apis.google.com/js/api.js");
        await new Promise(r => gapi.load("picker", r));
        const view = new google.picker.DocsView(google.picker.ViewId.DOCS).setMimeTypes(XLSX_MIME).setIncludeFolders(true);
        const b = new google.picker.PickerBuilder().addView(view).setOAuthToken(t).setLocale("es")
          .setTitle("Elige tu Base_Datos_Finanzas.xlsx").setCallback(d => {
            if (d.action === google.picker.Action.PICKED) res(d.docs[0]);
            else if (d.action === google.picker.Action.CANCEL) res(null);
          });
        if (cfg.apiKey) b.setDeveloperKey(cfg.apiKey);
        if (cfg.appId) b.setAppId(cfg.appId);
        b.build().setVisible(true);
      } catch (e) { rej(e); }
    });
  }

  /* ---------------- sincronización */
  async function sync(interactive = false) {
    if (!state.connected) return status();
    if (!navigator.onLine) { state.error = "Sin conexión: los cambios se guardan en el teléfono y se subirán después."; emit(); return status(); }
    if (syncing) return syncing;
    syncing = (async () => {
      emit();
      try {
        for (let intento = 0; intento < 3; intento++) {
          const m = await meta(state.fileId, interactive);
          if (m.trashed) throw new Error("El archivo de Drive está en la papelera.");
          if (String(m.version) !== String(state.version)) {
            // cambió en Drive (p. ej. desde otro dispositivo): bajar y reaplicar encima los cambios pendientes
            const remote = await F.parseExcel(await download(state.fileId, interactive), m.name);
            const rp = F.replay(remote, F.pending);
            await F.replaceModel(rp.model, { keepPending: true, ops: rp.ops });
            state.version = m.version; state.name = m.name;
          }
          if (!F.pending.length) break;
          const buf = await F.buildExcel();
          const fresh = await meta(state.fileId, interactive);       // nadie lo ha cambiado mientras preparábamos el archivo
          if (String(fresh.version) !== String(state.version)) continue;
          const res = await upload(state.fileId, buf, interactive);
          state.version = res.version;
          await F.clearPending();
          break;
        }
        state.lastSync = Date.now(); state.error = null;
      } catch (e) {
        state.error = e.message; state.needsAuth = !!e.needsAuth;
        console.warn("Sincronización con Drive:", e);
      } finally {
        await save(); syncing = null; emit();
      }
      return status();
    })();
    return syncing;
  }
  function schedule(ms = 1500) { clearTimeout(timer); timer = setTimeout(() => sync(false), ms); }

  /* ---------------- conectar / desconectar */
  async function connectExisting() {
    const doc = await pick();
    if (!doc) return null;
    const m = await meta(doc.id, true);
    const remote = await F.parseExcel(await download(doc.id, true), m.name);
    await F.replaceModel(remote, { keepPending: false });
    state = { connected: true, fileId: doc.id, name: m.name, version: m.version, lastSync: Date.now(), error: null };
    await save(); emit();
    return status();
  }
  async function connectNew() {
    const res = await createFile(await F.buildExcel(), F.model.name || F.DB_NAME, true);
    await F.clearPending();
    state = { connected: true, fileId: res.id, name: res.name, version: res.version, lastSync: Date.now(), error: null };
    await save(); emit();
    return status();
  }
  async function disconnect() {
    try { if (token) google.accounts.oauth2.revoke(token, () => {}); } catch (e) { /* opcional */ }
    token = null; state = { connected: false }; await save(); emit();
  }

  function status() {
    return { configured: !!cfg?.clientId, connected: !!state.connected, name: state.name || null, lastSync: state.lastSync || null,
      error: state.error || null, needsAuth: !!state.needsAuth, syncing: !!syncing, pending: F.pending.length, online: navigator.onLine };
  }
  async function init() {
    const saved = await F.kvGet("driveConfig").catch(() => null);
    cfg = saved?.clientId ? saved : DEFAULT_CFG;
    state = (await F.kvGet("driveState").catch(() => null)) || { connected: false };
    try { const t = JSON.parse(sessionStorage.getItem("gdToken") || "null"); if (t && Date.now() < t.tokenExp) ({ token, tokenExp } = t); } catch (e) { /* opcional */ }
    F.onChange(() => { if (state.connected) schedule(1500); });
    addEventListener("online", () => schedule(500));
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") schedule(300); });
    setInterval(() => { if (document.visibilityState === "visible") sync(false); }, 60000);
    if (state.connected) schedule(300);
    emit();
  }

  global.FinanzasDrive = {
    init, sync: () => sync(true), status, connectExisting, connectNew, disconnect,
    getConfig: () => ({ ...(cfg || {}) }),
    setConfig: async c => { cfg = { clientId: (c.clientId || "").trim(), apiKey: (c.apiKey || "").trim(), appId: (c.appId || "").trim() }; await F.kvSet("driveConfig", cfg); emit(); },
    onStatus: fn => listeners.add(fn),
  };
})(window);
