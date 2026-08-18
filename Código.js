// ════════════════════════════════════════════════════════════════
// BODA ROCÍO & DAVID — Apps Script v3.7
// ════════════════════════════════════════════════════════════════
//
// CAMBIOS v3.6 — Alojamiento con guardado propio:
//   · Nueva acción POST "save_aloj": guarda el alojamiento propio (columna R)
//     desde la pestaña Alojamiento, sin necesidad de confirmar.
//   · El RSVP ya no pide ni sobrescribe R (solo lo escribe si viene con valor).
//   · El email a los novios lee R de la hoja (actualizado por save_aloj).
//
// CAMBIOS v3.5 — Formulario ampliado + buscador Spotify:
//   · Mensaje → columna J (antes T). Las alergias dejan de ir en J.
//   · Datos por persona y logística de grupo a partir de la S.
//   · Nueva acción GET "song_search" (buscador de canciones Spotify).
//
//   MAPA DE COLUMNAS (1-based):
//     H(8)  total confirmados        I(9)  nº menús infantiles
//     J(10) mensaje para los novios  R(18) alojamiento propio (si L=FALSE)
//     S(19) nombre completo 1        T(20) ¿asiste 1? (Sí/No)
//     U(21) restricción 1            V(22) nombre completo 2
//     W(23) ¿asiste 2? (Sí/No)       X(24) restricción 2
//     Y(25) email                    Z(26) teléfono
//     AA(27) día de llegada          AB(28) día de salida
//     AC(29) cómo viajan             AD(30) ciudad de llegada
//     AE(31) canción 1               AF(32) canción 2
//     AG(33) niños (resumen)
//   Intactas: K,L,M,N,O,P,Q (reserva, flag aloj, web, user, pass, género, hotel).
// ════════════════════════════════════════════════════════════════

const SHEET_ID    = '14kSEOScPo3WSUk9AitH2BZaieUj-o_Z8';
const NOVIOS_MAIL = 'rocioetdavid@gmail.com';
const TOTAL_COLS  = 33; // hasta AG

// ── Spotify (rellena estos dos con los datos de tu app) ───────
const SPOTIFY_CLIENT_ID     = '37d9ecff60d346848a2d707ed9a43d57';
const SPOTIFY_CLIENT_SECRET = '7a90c40c5b324bbfb33e80e13c136c69';

// ── Utilidades ────────────────────────────────────────────────

function normalize(str) {
  return (str || '')
    .toString()
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim();
}

function primerNombre(str) {
  return (str || '').toString().trim().split(' ')[0];
}

function ensureColumns(sheet, needed) {
  const max = sheet.getMaxColumns();
  if (max < needed) sheet.insertColumnsAfter(max, needed - max);
}

// ── Rate limiting ─────────────────────────────────────────────

const RATE_LIMIT_MAX      = 10;
const RATE_LIMIT_WINDOW_S = 300;

function isRateLimited(key) {
  const props   = PropertiesService.getScriptProperties();
  const propKey = 'rl_' + key.replace(/[^a-z0-9_]/gi, '_').toLowerCase();
  const now     = Math.floor(Date.now() / 1000);
  let data;
  try { data = JSON.parse(props.getProperty(propKey) || 'null'); } catch(e) { data = null; }
  if (!data || now - data.start > RATE_LIMIT_WINDOW_S) {
    props.setProperty(propKey, JSON.stringify({ start: now, count: 1 }));
    return false;
  }
  if (data.count >= RATE_LIMIT_MAX) return true;
  data.count++;
  props.setProperty(propKey, JSON.stringify(data));
  return false;
}

// ── Punto de entrada POST ─────────────────────────────────────

function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return jsonResponse({ ok: false, error: 'no_body' });
  }
  let data;
  try { data = JSON.parse(e.postData.contents); }
  catch(err) { return jsonResponse({ ok: false, error: 'bad_request' }); }

  const action = (data.action || '').toLowerCase();
  if (action === 'login')     return handleLogin(data.u, data.p);
  if (action === 'rsvp')      return handleRsvp(data);
  if (action === 'save_aloj') return handleSaveAloj(data);
  return jsonResponse({ ok: false, error: 'unknown_action' });
}

// ── Punto de entrada GET ──────────────────────────────────────

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = (params.action || '').toLowerCase();

  if (action === 'login')       return handleLogin(params.u, params.p);
  if (action === 'song_search') return handleSongSearch(params.q);

  if (action === 'ping') {
    try {
      const ss  = SpreadsheetApp.openById(SHEET_ID);
      const inv = ss.getSheetByName('Invitados');
      return jsonResponse({
        ok: true, status: 'online', version: '3.6',
        sheet_found: !!inv,
      });
    } catch(err) {
      return jsonResponse({ ok: false, status: 'sheet_error', error: err.toString() });
    }
  }

  return jsonResponse({ ok: true, status: 'online', version: '3.6' });
}

// ── LOGIN ─────────────────────────────────────────────────────
//
// Rango leído: E:AG (29 columnas, índices base 0 desde E):
//   0 E nombre            1 F acompañante     2 G hijos
//   3 H total confirm.    4 I menú infantil   5 J mensaje
//   6 K reserva           7 L aloj asignado   8 M web
//   9 N usuario          10 O password       11 P género
//  12 Q hotel asignado   13 R aloj propio     14 S nombre completo 1
//  15 T asiste 1         16 U restricción 1  17 V nombre completo 2
//  18 W asiste 2         19 X restricción 2  20 Y email
//  21 Z teléfono         22 AA día llegada   23 AB día salida
//  24 AC transporte      25 AD ciudad        26 AE canción 1
//  27 AF canción 2       28 AG niños resumen
//
// Bloques:  Rocío: filas 17-84   ·   David: filas 88-156

function handleLogin(usuario, password) {
  try {
    if (!usuario || !password) {
      return jsonResponse({ ok: false, error: 'missing_credentials' });
    }
    if (isRateLimited('login_' + normalize(usuario))) {
      return jsonResponse({ ok: false, error: 'rate_limited' });
    }

    const ss  = SpreadsheetApp.openById(SHEET_ID);
    const inv = ss.getSheetByName('Invitados');
    if (!inv) return jsonResponse({ ok: false, error: 'sheet_not_found' });

    // Lectura defensiva: si la hoja aún no tiene todas las columnas,
    // leemos hasta donde llegue y el resto se rellena con ''.
    const maxCols = inv.getMaxColumns();
    const readW   = Math.min(29, Math.max(1, maxCols - 4));

    const bloques = [
      { rango: inv.getRange(17, 5, 68, readW), offset: 17 },
      { rango: inv.getRange(88, 5, 69, readW), offset: 88 },
    ];

    const usuarioNorm = normalize(usuario);
    const passwordStr = (password || '').toString().trim();

    for (const bloque of bloques) {
      const valores = bloque.rango.getValues();

      for (let i = 0; i < valores.length; i++) {
        const fila = valores[i];
        const g = idx => (idx < fila.length && fila[idx] != null) ? fila[idx] : '';

        const userCol = normalize(g(9));
        if (!userCol) continue;
        if (userCol !== usuarioNorm) continue;

        const pwCol = (g(10) || '').toString().trim();
        if (pwCol !== passwordStr) {
          return jsonResponse({ ok: false, error: 'wrong_password' });
        }

        const filaExcel = bloque.offset + i;
        const nombre    = (g(0) || '').toString().trim();
        const acompRaw  = fila[1];
        const acomp     = acompRaw !== null && acompRaw !== undefined
                          ? String(acompRaw).trim() : '';
        const hijos     = g(2);
        const tipoWeb   = (() => { const v = parseInt(g(8)); return isNaN(v) ? 1 : v; })();
        const confirmado = (() => {
          const v = g(3);
          if (v === null || v === '' || v === false || v === undefined) return 0;
          const n = parseInt(v); return isNaN(n) ? (v ? 1 : 0) : n;
        })();

        const alojAsignado = g(7) === true || g(7) === 'TRUE' || g(7) === 'true' || g(7) === 1;
        const alojHotel  = (g(12) || '').toString().trim();  // Q
        const alojPropio = (g(13) || '').toString().trim();  // R
        const genero      = parseInt(g(11)) === 1 ? 1 : 0;    // P
        const mensajePrev = (g(5)  || '').toString().trim();  // J (mensaje)
        const tieneHijos  = hijos !== null && hijos !== '' &&
                            hijos !== false && hijos !== 0 && hijos !== undefined;

        const tieneAcomp = (() => {
          if (tipoWeb === 2) return 2;
          if (acomp === '1' || acompRaw === 1 || acompRaw === 1.0) return 1;
          if (acomp !== '') return 1;
          return 0;
        })();

        // Valores previos del RSVP (para editar): S..AG
        const asiste = v => {
          const s = (v || '').toString().trim().toLowerCase();
          if (s === 'sí' || s === 'si') return 'si';
          if (s === 'no') return 'no';
          return '';
        };

        return jsonResponse({
          ok:              true,
          usuario:         userCol,
          nombre:          nombre,
          nombre_pareja:   tipoWeb === 2 ? primerNombre(acomp) : '',
          confirmado:      confirmado > 0 ? 1 : 0,
          tipo_web:        tipoWeb,
          tiene_acomp:     tieneAcomp,
          tiene_hijos:     tieneHijos,
          aloj_asignado:   alojAsignado,
          aloj_hotel:      alojHotel,
          aloj_propio:     alojPropio,
          genero:          genero,
          mensaje:         mensajePrev,
          fila_excel:      filaExcel,
          // Prefill del formulario ampliado:
          nombre_completo1: (g(14) || '').toString().trim(),
          asiste1:          asiste(g(15)),
          restriccion1:     (g(16) || '').toString().trim(),
          nombre_completo2: (g(17) || '').toString().trim(),
          asiste2:          asiste(g(18)),
          restriccion2:     (g(19) || '').toString().trim(),
          email:            (g(20) || '').toString().trim(),
          telefono:         (g(21) || '').toString().trim(),
          dia_llegada:      (g(22) || '').toString().trim(),
          dia_salida:       (g(23) || '').toString().trim(),
          transporte:       (g(24) || '').toString().trim(),
          ciudad_llegada:   (g(25) || '').toString().trim(),
          cancion1:         (g(26) || '').toString().trim(),
          cancion2:         (g(27) || '').toString().trim(),
          ninos_resumen:    (g(28) || '').toString().trim(),
        });
      }
    }

    return jsonResponse({ ok: false });

  } catch(err) {
    console.error('handleLogin error: ' + err.toString());
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ── Buscador de canciones (Spotify) ───────────────────────────

function getSpotifyToken() {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('sp_token');
  if (cached) return cached;
  const resp = UrlFetchApp.fetch('https://accounts.spotify.com/api/token', {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(SPOTIFY_CLIENT_ID + ':' + SPOTIFY_CLIENT_SECRET) },
    payload: { grant_type: 'client_credentials' },
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText() || '{}');
  if (data.access_token) {
    cache.put('sp_token', data.access_token, 3400);
    return data.access_token;
  }
  throw new Error('spotify_auth_failed');
}

function handleSongSearch(q) {
  try {
    q = (q || '').toString().trim();
    if (q.length < 2) return jsonResponse({ ok: true, tracks: [] });
    const token = getSpotifyToken();
    const url = 'https://api.spotify.com/v1/search?type=track&limit=6&market=ES&q=' + encodeURIComponent(q);
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
    });
    const data  = JSON.parse(resp.getContentText() || '{}');
    const items = (data.tracks && data.tracks.items) ? data.tracks.items : [];
    const tracks = items.map(t => ({
      id:     t.id,
      name:   t.name,
      artist: (t.artists || []).map(a => a.name).join(', '),
      img:    (t.album && t.album.images && t.album.images.length)
              ? t.album.images[t.album.images.length - 1].url : '',
      url:    (t.external_urls && t.external_urls.spotify) ? t.external_urls.spotify : '',
    }));
    return jsonResponse({ ok: true, tracks: tracks });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.toString(), tracks: [] });
  }
}

// ── Test manual ───────────────────────────────────────────────

function testLogin() {
  const resultado = handleLogin('isabel.caballero', 'ess4di');
  console.log('Resultado testLogin:', resultado.getContent());
}

function testSpotify() {
  const r = handleSongSearch('bohemian rhapsody');
  console.log(r.getContent());
}

// ── RSVP ─────────────────────────────────────────────────────

function handleRsvp(data) {
  try {
    if ((data.hp_name || '').trim() !== '') {
      return jsonResponse({ success: true });
    }

    const ss  = SpreadsheetApp.openById(SHEET_ID);
    const inv = ss.getSheetByName('Invitados');
    if (!inv) throw new Error('Hoja Invitados no encontrada');

    const fila = parseInt(data.fila_excel) || 0;
    if (fila < 17 || fila > 156) throw new Error('Fila fuera de rango: ' + fila);

    ensureColumns(inv, TOTAL_COLS);

    const sn = v => v === 'si' ? 'Sí' : (v === 'no' ? 'No' : '');

    // Género se lee de la hoja (P = col 16), no del cliente.
    const genero = parseInt(inv.getRange(fila, 16).getValue()) === 1 ? 1 : 0;

    // ── Escrituras fijas ──
    inv.getRange(fila, 8 ).setValue(parseInt(data.total_confirmados) || 0);  // H
    inv.getRange(fila, 9 ).setValue(parseInt(data.menu_infantil)     || 0);  // I
    inv.getRange(fila, 10).setValue(data.mensaje || '');                     // J

    // ── Alojamiento propio → R(18) SOLO si no le hemos asignado hotel ──
    const lVal = inv.getRange(fila, 12).getValue();
    const yaAsignado = lVal === true || lVal === 'TRUE' || lVal === 'true' || lVal === 1;
    // El alojamiento propio (R) ahora se guarda desde la pestaña Alojamiento
    // (acción save_aloj). Aquí solo lo escribimos si viniera con valor, para
    // no borrar lo que el invitado ya haya indicado allí.
    if (!yaAsignado && (data.aloj_propio || '').toString().trim() !== '') {
      inv.getRange(fila, 18).setValue(data.aloj_propio);
    }

    // ── Bloque nuevo S:AG (19..33) en una sola escritura ──
    inv.getRange(fila, 19, 1, 15).setValues([[
      data.nombre_completo1 || '',   // S
      sn(data.asiste1),              // T
      data.restriccion1 || '',       // U
      data.nombre_completo2 || '',   // V
      sn(data.asiste2),              // W
      data.restriccion2 || '',       // X
      data.email || '',              // Y
      data.telefono || '',           // Z
      data.dia_llegada || '',        // AA
      data.dia_salida || '',         // AB
      data.transporte || '',         // AC
      data.ciudad_llegada || '',     // AD
      data.cancion1 || '',           // AE
      data.cancion2 || '',           // AF
      data.ninos_resumen || '',      // AG
    ]]);

    // Alojamiento a mostrar en el email: Q si asignado, si no lo declarado.
    const alojEmail = yaAsignado
      ? (inv.getRange(fila, 17).getValue() || '').toString().trim()
      : (inv.getRange(fila, 18).getValue() || '').toString().trim();

    // ── Email a los novios ──
    const p1 = data.nombre_completo1 || primerNombre(data.nombre) || 'Persona 1';
    const nombres = (data.tipo_web == 2 || data.nombre_completo2)
      ? (data.nombre_completo1 || primerNombre(data.nombre)) +
        (data.nombre_completo2 ? ' y ' + data.nombre_completo2 : '')
      : p1;

    const persona = (nombre, asiste, restr) =>
      '— ' + (nombre || '—') + ' —\n' +
      'Asiste: ' + (asiste === 'si' ? '✅ Sí' : asiste === 'no' ? '❌ No' : '—') + '\n' +
      'Restricción: ' + (restr || 'Ninguna') + '\n';

    let cuerpo = 'Nueva confirmación de boda\n\n';
    cuerpo += 'Usuario: ' + (data.usuario || '—') + '\n\n';
    cuerpo += persona(data.nombre_completo1 || primerNombre(data.nombre), data.asiste1, data.restriccion1);
    if ((data.nombre_completo2 || '').trim() || data.asiste2) {
      cuerpo += '\n' + persona(data.nombre_completo2, data.asiste2, data.restriccion2);
    }
    cuerpo += '\nContacto:\n';
    cuerpo += 'Email: ' + (data.email || '—') + '\n';
    cuerpo += 'Teléfono: ' + (data.telefono || '—') + '\n\n';
    cuerpo += 'Logística:\n';
    cuerpo += 'Llegada: ' + (data.dia_llegada || '—') + '\n';
    cuerpo += 'Salida: ' + (data.dia_salida || '—') + '\n';
    cuerpo += 'Cómo viajan: ' + (data.transporte || '—') +
              (data.ciudad_llegada ? ' (ciudad: ' + data.ciudad_llegada + ')' : '') + '\n';
    cuerpo += 'Alojamiento: ' + (alojEmail || '—') + '\n\n';
    cuerpo += 'Niños:\n' + (data.ninos_resumen || '—') + '\n';
    cuerpo += 'Menús infantiles: ' + (data.menu_infantil || 0) + '\n\n';
    cuerpo += 'Canciones:\n';
    cuerpo += '1. ' + (data.cancion1 || '—') + '\n';
    cuerpo += '2. ' + (data.cancion2 || '—') + '\n\n';
    cuerpo += 'Total confirmados: ' + (data.total_confirmados || 0) + '\n\n';
    cuerpo += 'Mensaje:\n' + (data.mensaje || '—');

    GmailApp.sendEmail(NOVIOS_MAIL, '✅ Confirmación: ' + nombres, cuerpo);

    // ── Save the Date al invitado (si alguien asiste) ──
    const alguienAsiste = data.asiste1 === 'si' || data.asiste2 === 'si';
    if (data.email && alguienAsiste) {
      const restricciones = [data.restriccion1, data.restriccion2]
        .filter(r => r && r.trim()).join(' · ') || 'Ninguna';
      const p1n = primerNombre(data.nombre_completo1 || data.nombre);
      const p2n = data.tipo_web == 2 ? primerNombre(data.nombre_completo2 || data.nombre_pareja) : '';
      try {
        const props = PropertiesService.getScriptProperties();
        props.setProperty('email_q_' + Date.now(), JSON.stringify({
          email:         data.email,
          nombre:        p1n,
          nombre_pareja: p2n,
          total:         data.total_confirmados,
          alojamiento:   alojEmail || '—',
          restricciones: restricciones,
          genero:        genero,
          tipo:          parseInt(data.tipo_web) || 1,
        }));
        ScriptApp.newTrigger('processPendingEmails').timeBased().after(60000).create();
      } catch(triggerErr) {
        sendSaveTheDate(data.email, p1n, p2n, data.total_confirmados,
                        alojEmail, restricciones, genero, parseInt(data.tipo_web) || 1);
      }
    }

    return jsonResponse({ success: true });

  } catch(err) {
    console.error('handleRsvp error: ' + err.toString());
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ── Emails asíncronos ─────────────────────────────────────────

function processPendingEmails() {
  const props = PropertiesService.getScriptProperties();
  const all   = props.getProperties();
  Object.keys(all).forEach(key => {
    if (!key.startsWith('email_q_')) return;
    try {
      const d = JSON.parse(all[key]);
      sendSaveTheDate(d.email, d.nombre, d.nombre_pareja, d.total,
                      d.alojamiento, d.restricciones, d.genero, d.tipo);
      props.deleteProperty(key);
    } catch(err) {
      console.error('Error email pendiente:', key, err.toString());
    }
  });
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'processPendingEmails')
      ScriptApp.deleteTrigger(t);
  });
}

// ── Save the Date ─────────────────────────────────────────────

function sendSaveTheDate(email, nombre, nombrePareja, total, alojamiento, restricciones, genero, tipo) {
  const calUrl =
    'https://calendar.google.com/calendar/render?action=TEMPLATE' +
    '&text=Boda+Roc%C3%ADo+%26+David' +
    '&dates=20261206/20261207' +
    '&details=¡Nos+casamos!+Reserva+este+día+tan+especial.' +
    '&location=Hotel+Los+Escullos,+Cabo+de+Gata,+Almería' +
    '&sf=true&output=xml';

  const esPareja     = parseInt(tipo) === 2 && !!nombrePareja;
  const esFemenino   = parseInt(genero) === 1;
  const destinatario = esPareja ? `${nombre} y ${nombrePareja}` : nombre;
  const saludo       = esPareja
    ? `Queridos ${destinatario}`
    : `${esFemenino ? 'Querida' : 'Querido'} ${nombre}`;

  const html = `
<!DOCTYPE html>
<html lang="es">
<body style="margin:0;padding:0;background:#faf6f0;font-family:'Georgia',serif;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#faf6f0;padding:40px 20px;">
    <tr><td align="center">
      <table width="560" cellpadding="0" cellspacing="0"
             style="background:#ffffff;border-radius:4px;overflow:hidden;box-shadow:0 2px 20px rgba(0,0,0,0.07);">
        <tr>
          <td align="center" style="background:#7FA9CB;padding:50px 40px 40px;">
            <p style="margin:0 0 12px;color:#ffffff;font-size:12px;letter-spacing:4px;text-transform:uppercase;">Save the Date</p>
            <h1 style="margin:0;color:#ffffff;font-size:36px;font-weight:400;letter-spacing:2px;">Rocío &amp; David</h1>
            <p style="margin:16px 0 0;color:#ffffff;font-size:14px;letter-spacing:3px;">6 · XII · 2026</p>
          </td>
        </tr>
        <tr>
          <td style="padding:44px 48px 36px;">
            <p style="margin:0 0 24px;font-size:17px;color:#3a2e28;line-height:1.7;">${saludo},</p>
            <p style="margin:0 0 32px;font-size:15px;color:#5a4a42;line-height:1.8;">${esPareja
              ? '¡Gracias por confirmar vuestra asistencia! Nos hace muchísima ilusión celebrarlo con vosotros.'
              : '¡Gracias por confirmar tu asistencia! Nos hace muchísima ilusión celebrarlo contigo.'}</p>
            <table width="100%" cellpadding="0" cellspacing="0" style="background:#faf6f0;border-radius:4px;margin-bottom:32px;">
              <tr><td style="padding:24px 28px;">
                <p style="margin:0 0 12px;font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#b8956a;">Detalles</p>
                <p style="margin:0 0 8px;font-size:14px;color:#3a2e28;"><span style="color:#b8956a;">&#9679;</span>&nbsp; Puente de diciembre</p>
                <p style="margin:0 0 8px;font-size:14px;color:#3a2e28;"><span style="color:#b8956a;">&#9679;</span>&nbsp; Ceremonia y banquete · 6 de diciembre</p>
                <p style="margin:0 0 8px;font-size:14px;color:#3a2e28;"><span style="color:#b8956a;">&#9679;</span>&nbsp; Ceremonia · Iglesia de Rodalquilar</p>
                <p style="margin:0;font-size:14px;color:#3a2e28;"><span style="color:#b8956a;">&#9679;</span>&nbsp; Banquete · Hotel Los Escullos, Cabo de Gata, Almería</p>
              </td></tr>
            </table>
            <p style="margin:0 0 36px;font-size:15px;color:#5a4a42;line-height:1.8;">${esPareja
              ? 'Tenemos todos los detalles de vuestra confirmación guardados. ¡Muchísimas gracias por confirmar!'
              : 'Tenemos todos los detalles de tu confirmación guardados. ¡Muchísimas gracias por confirmar!'}</p>
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr><td align="center">
                <a href="${calUrl}" style="display:inline-block;background:#152180;color:#ffffff;text-decoration:none;font-size:12px;letter-spacing:2px;text-transform:uppercase;padding:15px 28px;border-radius:2px;max-width:100%;box-sizing:border-box;">
                  Añadir al calendario
                </a>
              </td></tr>
            </table>
          </td>
        </tr>
        <tr>
          <td align="center" style="padding:28px 40px;border-top:1px solid #ede8e0;">
            <p style="margin:0;font-size:12px;color:#b8a898;letter-spacing:1px;">Con mucho cariño · Rocío &amp; David</p>
          </td>
        </tr>
      </table>
    </td></tr>
  </table>
</body>
</html>`;

  const ics = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Rocio y David//Boda//ES',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    'UID:boda-rocio-david-20261206@rocioydavid.com',
    'DTSTAMP:20260101T000000Z',
    'DTSTART;VALUE=DATE:20261206',
    'DTEND;VALUE=DATE:20261207',
    'SUMMARY:Boda Rocío & David',
    'LOCATION:Iglesia de Rodalquilar / Hotel Los Escullos, Cabo de Gata, Almería',
    'DESCRIPTION:¡Nos casamos! Ceremonia en la Iglesia de Rodalquilar y banquete en el Hotel Los Escullos.',
    'TRANSP:TRANSPARENT',
    'BEGIN:VALARM',
    'ACTION:DISPLAY',
    'DESCRIPTION:La boda de Rocío y David es en 2 días',
    'TRIGGER:-P2D',
    'END:VALARM',
    'END:VEVENT',
    'END:VCALENDAR',
  ].join('\r\n');
  const icsBlob = Utilities.newBlob(ics, 'text/calendar', 'Boda-Rocio-David.ics');

  GmailApp.sendEmail(email, 'Save the Date · Rocío y David · 6 Dic 2026', '',
                     { htmlBody: html, name: 'Rocío y David', attachments: [icsBlob] });
}

// ── Guardar alojamiento propio (pestaña Alojamiento) ──────────

function handleSaveAloj(data) {
  try {
    const fila = parseInt(data.fila_excel) || 0;
    if (fila < 17 || fila > 156) throw new Error('Fila fuera de rango: ' + fila);

    const ss  = SpreadsheetApp.openById(SHEET_ID);
    const inv = ss.getSheetByName('Invitados');
    if (!inv) throw new Error('Hoja Invitados no encontrada');

    ensureColumns(inv, TOTAL_COLS);

    // No tocamos R si tiene hotel asignado (L = TRUE): ese caso no ve el recuadro.
    const lVal = inv.getRange(fila, 12).getValue();
    const yaAsignado = lVal === true || lVal === 'TRUE' || lVal === 'true' || lVal === 1;
    if (!yaAsignado) {
      inv.getRange(fila, 18).setValue((data.aloj || '').toString().trim());
    }

    return jsonResponse({ success: true });
  } catch(err) {
    console.error('handleSaveAloj error: ' + err.toString());
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ── Helper JSON ───────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
