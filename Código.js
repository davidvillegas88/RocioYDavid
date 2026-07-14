// ════════════════════════════════════════════════════════════════
// BODA ROCÍO & DAVID — Apps Script v3.3
// ════════════════════════════════════════════════════════════════
//
// CAMBIOS v3.3:
//   · FIX: el mensaje del invitado ya NO se escribe en la columna L
//     (que contiene el alojamiento asignado). Ahora va a la columna T.
//   · Nueva columna P = género (0 masculino / 1 femenino) → saludo
//     "Querido" / "Querida" en el email cuando web = 1 (singular).
//     Si web = 2 el saludo va en plural ("Queridos").
//   · El login lee el rango E:T (16 col) y devuelve género y mensaje previo.
// ════════════════════════════════════════════════════════════════

const SHEET_ID = '14kSEOScPo3WSUk9AitH2BZaieUj-o_Z8';
const NOVIOS_MAIL = 'rocioetdavid@gmail.com';

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
  if (action === 'login') return handleLogin(data.u, data.p);
  if (action === 'rsvp')  return handleRsvp(data);
  return jsonResponse({ ok: false, error: 'unknown_action' });
}

// ── Punto de entrada GET ──────────────────────────────────────

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  const action = (params.action || '').toLowerCase();

  if (action === 'login') {
    return handleLogin(params.u, params.p);
  }

  if (action === 'ping') {
    try {
      const ss  = SpreadsheetApp.openById(SHEET_ID);
      const inv = ss.getSheetByName('Invitados');
      return jsonResponse({
        ok: true, status: 'online', version: '3.3',
        sheet_found: !!inv,
      });
    } catch(err) {
      return jsonResponse({ ok: false, status: 'sheet_error', error: err.toString() });
    }
  }

  return jsonResponse({ ok: true, status: 'online', version: '3.1' });
}

// ── LOGIN ─────────────────────────────────────────────────────
//
// Columnas del rango E:P (índice base 0):
//   [0]  E  Nombre
//   [1]  F  Acompañante  → vacío=0 | 1(num)=acomp opcional | texto=web2
//   [2]  G  Hijos
//   [3]  H  Confirmación
//   [4]  I  Menú infantil
//   [5]  J  Alergias
//   [6]  K  Reserva Hotel
//   [7]  L  Alojamiento asignado  ← NO ESCRIBIR AQUÍ desde el RSVP
//   [8]  M  Web (1 o 2)
//   [9]  N  Usuario
//  [10]  O  Password
//  [11]  P  Género: 0 = masculino | 1 = femenino
//  [12]  Q  (libre)
//  [13]  R  (libre)
//  [14]  S  (libre)
//  [15]  T  Mensaje del invitado (escrito por el RSVP)
//
// Bloques:
//   Rocío: filas 17-84  → getRange(17, 5, 68, 16)
//   David: filas 88-156 → getRange(88, 5, 69, 16)

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

    const bloques = [
      { rango: inv.getRange(17, 5, 68, 16), offset: 17 },
      { rango: inv.getRange(88, 5, 69, 16), offset: 88 },
    ];

    const usuarioNorm = normalize(usuario);
    const passwordStr = (password || '').toString().trim();

    console.log('Login attempt: ' + usuarioNorm);

    for (const bloque of bloques) {
      const valores = bloque.rango.getValues();

      for (let i = 0; i < valores.length; i++) {
        const fila = valores[i];

        const userCol = normalize(fila[9]);
        if (!userCol) continue;
        if (userCol !== usuarioNorm) continue;

        const pwCol = (fila[10] || '').toString().trim();
        if (pwCol !== passwordStr) {
          console.log('Password incorrecto para: ' + userCol);
          return jsonResponse({ ok: false, error: 'wrong_password' });
        }

        const filaExcel = bloque.offset + i;
        const nombre    = (fila[0] || '').toString().trim();
        const acompRaw  = fila[1];
        const acomp     = acompRaw !== null && acompRaw !== undefined
                          ? String(acompRaw).trim() : '';
        const hijos     = fila[2];
        const tipoWeb   = (() => { const v = parseInt(fila[8]); return isNaN(v) ? 1 : v; })();
        const confirmado = (() => {
          const v = fila[3];
          if (v === null || v === '' || v === false || v === undefined) return 0;
          const n = parseInt(v); return isNaN(n) ? (v ? 1 : 0) : n;
        })();
        const hotelReserv = !!(fila[6]);
        const alojamiento = (fila[7] || '').toString().trim();   // columna L
        const genero      = parseInt(fila[11]) === 1 ? 1 : 0;    // columna P
        const mensajePrev = (fila[15] || '').toString().trim();  // columna T
        const tieneHijos  = hijos !== null && hijos !== '' &&
                            hijos !== false && hijos !== 0 && hijos !== undefined;

        // Casos de acompañante:
        // 0 = sin acompañante (F vacío)
        // 1 = acompañante opcional (F = 1 numérico)
        // 2 = pareja dedicada (web=2, F tiene nombre)
        const tieneAcomp = (() => {
          if (tipoWeb === 2) return 2;
          if (acomp === '1' || acompRaw === 1 || acompRaw === 1.0) return 1;
          if (acomp !== '') return 1;
          return 0;
        })();

        console.log('Login OK: ' + nombre + ' (fila ' + filaExcel + ', tipo=' + tipoWeb + ', acomp=' + tieneAcomp + ')');

        return jsonResponse({
          ok:              true,
          usuario:         userCol,
          nombre:          nombre,
          nombre_pareja:   tipoWeb === 2 ? primerNombre(acomp) : '',
          confirmado:      confirmado > 0 ? 1 : 0,
          tipo_web:        tipoWeb,
          tiene_acomp:     tieneAcomp,
          tiene_hijos:     tieneHijos,
          hotel_reservado: hotelReserv,
          alojamiento:     alojamiento,
          genero:          genero,
          mensaje:         mensajePrev,
          fila_excel:      filaExcel,
        });
      }
    }

    console.log('Usuario no encontrado: ' + usuarioNorm);
    return jsonResponse({ ok: false });

  } catch(err) {
    console.error('handleLogin error: ' + err.toString());
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ── Test manual (ejecutar desde el editor para verificar) ─────

function testLogin() {
  const resultado = handleLogin('isabel.caballero', 'ess4di');
  console.log('Resultado testLogin:', resultado.getContent());
}

function testLoginPareja() {
  const resultado = handleLogin('charo', 'saiebz');
  console.log('Resultado testLoginPareja:', resultado.getContent());
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

    // Género: se lee de la hoja (P = col 16), no del cliente.
    const genero = parseInt(inv.getRange(fila, 16).getValue()) === 1 ? 1 : 0;

    // H=8 confirmados | I=9 menús | J=10 alergias | T=20 mensaje
    // OJO: la columna L (12) es el ALOJAMIENTO ASIGNADO. No se toca.
    // OJO: la columna P (16) es el GÉNERO. Tampoco se toca.
    inv.getRange(fila, 8).setValue(parseInt(data.total_confirmados) || 0);
    inv.getRange(fila, 9).setValue(parseInt(data.menu_infantil)     || 0);
    inv.getRange(fila, 10).setValue(data.alergias || '');
    inv.getRange(fila, 20).setValue(data.mensaje  || '');

    const tipo = parseInt(data.tipo_web) || 1;
    const p1   = primerNombre(data.nombre);
    const p2   = primerNombre(data.nombre_pareja || '');
    const nombres = tipo === 2 && p2 ? `${p1} y ${p2}` : p1;

    const as1txt  = data.asistencia1 === 'si' ? '✅ Sí' : '❌ No';
    const as2txt  = tipo === 2
      ? `\nAsistencia (${p2}): ${data.asistencia2 === 'si' ? '✅ Sí' : '❌ No'}` : '';
    const acompTxt = tipo === 1 && data.con_acompanante === 'si'
      ? `\nAcompañante: ${data.nombre_acompanante || '—'}` : '';

    const hijosLineas = (data.hijos_nombres || []).length
      ? data.hijos_nombres.map(h =>
          `  · ${h.nombre}${h.menu_infantil ? ' 🍽️ menú infantil' : ''}`
        ).join('\n')
      : '—';

    GmailApp.sendEmail(
      NOVIOS_MAIL,
      `✅ Confirmación: ${nombres}`,
      `Nueva confirmación de boda\n\n` +
      `Usuario: ${data.usuario || '—'}\n` +
      `Invitado/s: ${nombres}\n` +
      `Email: ${data.email}\n` +
      `Teléfono: ${data.telefono || '—'}\n\n` +
      `Asistencia (${p1}): ${as1txt}` + as2txt + acompTxt + '\n' +
      `Total confirmados: ${data.total_confirmados}\n\n` +
      `Niños:\n${hijosLineas}\n` +
      `Menús infantiles: ${data.menu_infantil || 0}\n` +
      `Alergias: ${data.alergias || 'Ninguna'}\n\n` +
      `Mensaje:\n${data.mensaje || '—'}`
    );

    const alguienAsiste = data.asistencia1 === 'si' || data.asistencia2 === 'si';
    if (data.email && alguienAsiste) {
      try {
        const props = PropertiesService.getScriptProperties();
        props.setProperty('email_q_' + Date.now(), JSON.stringify({
          email:         data.email,
          nombre:        p1,
          nombre_pareja: tipo === 2 ? p2 : '',
          total:         data.total_confirmados,
          alojamiento:   data.alojamiento || '—',
          alergias:      data.alergias    || 'Ninguna',
          genero:        genero,
          tipo:          tipo,
        }));
        ScriptApp.newTrigger('processPendingEmails').timeBased().after(60000).create();
      } catch(triggerErr) {
        sendSaveTheDate(data.email, p1, tipo === 2 ? p2 : '',
                        data.total_confirmados, data.alojamiento, data.alergias,
                        genero, tipo);
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
      sendSaveTheDate(d.email, d.nombre, d.nombre_pareja,
                      d.total, d.alojamiento, d.alergias,
                      d.genero, d.tipo);
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

function sendSaveTheDate(email, nombre, nombrePareja, total, alojamiento, alergias, genero, tipo) {
  const calUrl =
    'https://calendar.google.com/calendar/render?action=TEMPLATE' +
    '&text=Boda+Roc%C3%ADo+%26+David' +
    '&dates=20261206T110000Z/20261206T230000Z' +
    '&details=¡Nos+casamos!+Reserva+este+día+tan+especial.' +
    '&location=Castillo+de+Los+Escullos,+Cabo+de+Gata,+Almería' +
    '&sf=true&output=xml';

  // Saludo:
  //   web = 2 (pareja)  → plural  → "Queridos Ana y Luis"
  //   web = 1 (singular)→ género  → "Querida Ana" / "Querido Luis"
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
          <td align="center" style="background:#3a2e28;padding:50px 40px 40px;">
            <p style="margin:0 0 12px;color:#b8956a;font-size:12px;letter-spacing:4px;text-transform:uppercase;">Save the Date</p>
            <h1 style="margin:0;color:#faf6f0;font-size:36px;font-weight:400;letter-spacing:2px;">Rocío &amp; David</h1>
            <p style="margin:16px 0 0;color:#b8956a;font-size:14px;letter-spacing:3px;">6 · XII · 2026</p>
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
                <p style="margin:0 0 8px;font-size:14px;color:#3a2e28;"><span style="color:#b8956a;">&#9679;</span>&nbsp; 6 de Diciembre de 2026</p>
                <p style="margin:0;font-size:14px;color:#3a2e28;"><span style="color:#b8956a;">&#9679;</span>&nbsp; Castillo de Los Escullos, Almería</p>
              </td></tr>
            </table>
            <p style="margin:0 0 14px;font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#b8956a;">${esPareja ? 'Vuestra confirmación' : 'Tu confirmación'}</p>
            <table width="100%" cellpadding="0" cellspacing="0" style="border-top:1px solid #ede8e0;margin-bottom:36px;">
              <tr>
                <td style="padding:10px 0;border-bottom:1px solid #ede8e0;font-size:13px;color:#8a7a72;width:140px;">Asistentes</td>
                <td style="padding:10px 0;border-bottom:1px solid #ede8e0;font-size:13px;color:#3a2e28;">${total}</td>
              </tr>
              <tr>
                <td style="padding:10px 0;border-bottom:1px solid #ede8e0;font-size:13px;color:#8a7a72;">Alojamiento</td>
                <td style="padding:10px 0;border-bottom:1px solid #ede8e0;font-size:13px;color:#3a2e28;">${alojamiento || '—'}</td>
              </tr>
              <tr>
                <td style="padding:10px 0;font-size:13px;color:#8a7a72;">Alergias</td>
                <td style="padding:10px 0;font-size:13px;color:#3a2e28;">${alergias || 'Ninguna'}</td>
              </tr>
            </table>
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr><td align="center">
                <a href="${calUrl}" style="display:inline-block;background:#3a2e28;color:#faf6f0;text-decoration:none;font-size:12px;letter-spacing:3px;text-transform:uppercase;padding:16px 36px;border-radius:2px;">
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

  GmailApp.sendEmail(email, 'Save the Date · Rocío y David · 6 Dic 2026', '',
                     { htmlBody: html, name: 'Rocío y David' });
}

// ── Helper JSON ───────────────────────────────────────────────

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}