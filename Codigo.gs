/**
 * RONDA DE CAFÉS · servidor en Google Sheets (con inicio de sesión de Google)
 *
 * Cada persona entra con su cuenta de Google. La primera vez elige su nombre y queda
 * dada de alta en la pestaña Personas. Cada ronda y cada pago guardan quién lo apuntó.
 *
 * Instalación (una sola vez):
 *  1. Crea una hoja de cálculo nueva en Google Drive.
 *  2. Extensiones > Apps Script. Borra lo que haya en Código.gs y pega este archivo.
 *  3. Añade un archivo HTML (botón + > HTML) llamado exactamente "Index" y pega Index.html.
 *  4. Elige la función "configurar" arriba y pulsa Ejecutar. Acepta los permisos.
 *  5. Implementar > Nueva implementación > Tipo: Aplicación web. Elige UNA de estas opciones:
 *
 *     A) Si en la empresa usáis Google Workspace (correo de empresa con Gmail) y has creado
 *        la hoja con esa cuenta:
 *          Ejecutar como: Yo
 *          Quién tiene acceso: Cualquier usuario de <tu empresa>
 *
 *     B) Si cada uno usa su Gmail personal:
 *          Ejecutar como: Usuario que accede a la aplicación web
 *          Quién tiene acceso: Cualquier usuario con una cuenta de Google
 *        y comparte la hoja de cálculo con cada compañero como Editor.
 *        La primera vez, Google le pedirá permiso a cada uno (pulsar Configuración avanzada > Ir a...).
 *
 *     OJO: con "Quién tiene acceso: Cualquier usuario" (sin "con una cuenta de Google") Google
 *     no pide iniciar sesión. Y con "Ejecutar como: Yo" y cuentas de Gmail personales, Google
 *     no dice quién entra. En los dos casos la app muestra un aviso con los pasos para arreglarlo.
 *
 *  6. Copia la URL que termina en /exec y pásala al grupo.
 *
 * Si cambias el código: Implementar > Gestionar implementaciones > lápiz > Versión: nueva versión.
 * Si ya tenías la versión anterior instalada, las pestañas se actualizan solas al abrir la app.
 *
 * VERSIÓN EN GITHUB (la web vive en GitHub Pages y este archivo hace de servidor):
 *  - Pega tu ID de cliente de Google en GOOGLE_CLIENT_ID (justo debajo).
 *  - Crea OTRA implementación de tipo Aplicación web con:
 *      Ejecutar como: Yo
 *      Quién tiene acceso: Cualquier usuario
 *    Aquí "Cualquier usuario" es correcto: quién entra lo controla el botón de Google
 *    y la pestaña Permitidos, que se comprueban en cada petición.
 *  - En la pestaña Permitidos escribe un correo por fila, o @tuempresa.com para todo un dominio.
 *    Quien ejecutó "configurar" (el dueño) siempre tiene acceso.
 */

// ID de cliente de Google para la versión en GitHub (Google Cloud > Clientes). No es secreto.
var GOOGLE_CLIENT_ID = 'PEGA_AQUI_TU_ID_DE_CLIENTE.apps.googleusercontent.com';

var HOJAS = {
  Rondas: ['ID', 'Fecha', 'Pagó', 'Persona', 'Producto', 'Precio', 'Apuntado por'],
  Pagos: ['ID', 'Fecha', 'De', 'Para', 'Importe', 'Apuntado por'],
  Personas: ['Nombre', 'Email', 'Activo'],
  Productos: ['Producto', 'Precio'],
  Permitidos: ['Correo o dominio']
};
var ESQUEMA = '2';

var PRODUCTOS_INICIALES = [
  ['Café solo', 1.30],
  ['Cortado', 1.40],
  ['Café con leche', 1.60],
  ['Americano', 1.50],
  ['Capuchino', 2.00],
  ['Descafeinado', 1.60],
  ['Té o infusión', 1.50],
  ['Cola Cao', 1.80]
];

var FORMATO_EUR = '#,##0.00 "€"';
var FORMATO_FECHA = 'dd/mm/yyyy hh:mm';
var SIN_CORREO = 'No puedo saber con qué cuenta de Google has entrado. Quien instaló la app tiene que revisar cómo está implementada (instrucciones al principio de Codigo.gs).';
var SIN_ALTA = 'Antes de apuntar nada, elige tu nombre. Recarga la página.';
var SIN_PERMISO = 'Tu cuenta no tiene acceso a esta app. Pide a quien la gestiona que te añada a la pestaña Permitidos.';

// Solo durante una petición de la versión GitHub: el correo sale del token de Google ya verificado.
var MODO_API_ = false;
var CORREO_API_ = '';

/* ---------- Web ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Ronda de cafés')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Ejecútala una vez desde el editor: crea las pestañas y pide los permisos. */
function configurar() {
  var ss = libro_();
  PropertiesService.getScriptProperties().setProperty('DUENO', correo_(Session.getEffectiveUser().getEmail()));
  asegurar_(ss);
  ['Hoja 1', 'Hoja1', 'Sheet1', 'Full 1'].forEach(function (nombre) {
    var h = ss.getSheetByName(nombre);
    if (h && h.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(h);
  });
  SpreadsheetApp.flush();
  Logger.log('Listo. Ahora: Implementar > Nueva implementación > Aplicación web.');
}

/* ---------- Lectura ---------- */

function cargarDatos() {
  var ss = libro_();
  asegurar_(ss);

  var filasPersonas = filas_(ss, 'Personas');
  var me = usuario_(ss, filasPersonas);
  if (!me.email || me.blocked) {
    return { me: me, people: [], products: [], rounds: [], payments: [], sheetUrl: '', appUrl: urlApp_() };
  }
  var nombrePorCorreo = {};
  filasPersonas.forEach(function (r) {
    var correo = correo_(r[1]);
    if (correo && texto_(r[0]) && !nombrePorCorreo[correo]) nombrePorCorreo[correo] = texto_(r[0]);
  });
  // Los correos no salen del servidor: a la web solo le llega el nombre de quien apuntó.
  function autor(v) {
    var correo = correo_(v);
    return correo ? (nombrePorCorreo[correo] || correo.split('@')[0]) : '';
  }

  var people = filasPersonas
    .filter(function (r) { return texto_(r[0]); })
    .map(function (r) { return { name: texto_(r[0]), active: activo_(r[2]), claimed: !!correo_(r[1]) }; });

  var products = filas_(ss, 'Productos')
    .filter(function (r) { return texto_(r[0]); })
    .map(function (r) { return { name: texto_(r[0]), price: numero_(r[1]) }; });

  var porId = {}, orden = [];
  filas_(ss, 'Rondas').forEach(function (r) {
    var id = texto_(r[0]);
    if (!id || !texto_(r[3])) return;
    if (!porId[id]) {
      porId[id] = { id: id, date: fecha_(r[1]), payer: texto_(r[2]), by: autor(r[6]), items: [] };
      orden.push(id);
    }
    porId[id].items.push({ person: texto_(r[3]), product: texto_(r[4]), price: numero_(r[5]) });
  });

  var payments = filas_(ss, 'Pagos')
    .filter(function (r) { return texto_(r[0]); })
    .map(function (r) {
      return { id: texto_(r[0]), date: fecha_(r[1]), from: texto_(r[2]), to: texto_(r[3]), amount: numero_(r[4]), by: autor(r[5]) };
    });

  return {
    me: me,
    people: people,
    products: products,
    rounds: orden.map(function (id) { return porId[id]; }),
    payments: payments,
    sheetUrl: ss.getUrl(),
    appUrl: urlApp_()
  };
}

/** Dirección pública de la app (la que acaba en /exec), para el enlace «Cambiar de cuenta». */
function urlApp_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

/* ---------- Alta y nombre ---------- */

/** Da de alta a quien ha entrado. Si el nombre ya estaba en la lista sin cuenta, lo reclama. */
function registrarme(nombre) {
  nombre = texto_(nombre);
  if (!nombre) throw new Error('Escribe tu nombre');
  if (nombre.length > 40) throw new Error('El nombre es demasiado largo');
  return conBloqueo_(function (ss, yo) {
    if (!yo.email) throw new Error(SIN_CORREO);
    if (yo.name) return; // ya estaba dado de alta
    var h = ss.getSheetByName('Personas');
    var filas = filas_(ss, 'Personas');
    for (var i = 0; i < filas.length; i++) {
      if (igual_(filas[i][0], nombre)) {
        if (correo_(filas[i][1])) throw new Error('«' + texto_(filas[i][0]) + '» ya está unido a otra cuenta. Elige otro nombre.');
        h.getRange(i + 2, 2, 1, 2).setValues([[yo.email, true]]);
        return;
      }
    }
    var fila = ultimaFilaConNombre_(h) + 1;
    h.getRange(fila, 3).insertCheckboxes();
    h.getRange(fila, 1, 1, 3).setValues([[nombre, yo.email, true]]);
  }, { sinAlta: true });
}

/** Cambia tu nombre en la lista y en todo el historial. */
function cambiarMiNombre(nuevo) {
  nuevo = texto_(nuevo);
  if (!nuevo) throw new Error('Escribe tu nombre');
  if (nuevo.length > 40) throw new Error('El nombre es demasiado largo');
  return conBloqueo_(function (ss, yo) {
    var viejo = yo.name;
    if (viejo === nuevo) return;
    var filas = filas_(ss, 'Personas'), mia = -1;
    for (var i = 0; i < filas.length; i++) {
      if (correo_(filas[i][1]) === yo.email && mia < 0) mia = i;
      else if (igual_(filas[i][0], nuevo)) throw new Error('Ya hay alguien que se llama «' + nuevo + '»');
    }
    ss.getSheetByName('Personas').getRange(mia + 2, 1).setValue(nuevo);
    renombrar_(ss.getSheetByName('Rondas'), 3, viejo, nuevo); // columnas Pagó y Persona
    renombrar_(ss.getSheetByName('Pagos'), 3, viejo, nuevo);  // columnas De y Para
  });
}

/* ---------- Escritura (todas guardan quién la hace) ---------- */

function guardarRonda(ronda) {
  var pagador = texto_(ronda && ronda.payer);
  var items = (ronda && ronda.items) || [];
  if (!pagador || !items.length) throw new Error('La ronda necesita quién paga y al menos un café');
  return conBloqueo_(function (ss, yo) {
    var id = nuevoId_(), ahora = new Date();
    var filas = items.map(function (it) {
      return [id, ahora, pagador, texto_(it.person), texto_(it.product), redondear_(numero_(it.price)), yo.email];
    });
    var h = ss.getSheetByName('Rondas');
    h.getRange(h.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
    return id;
  });
}

function borrarRonda(id) {
  return conBloqueo_(function (ss) { borrarFilas_(ss.getSheetByName('Rondas'), texto_(id)); });
}

function guardarPago(pago) {
  var de = texto_(pago && pago.from), para = texto_(pago && pago.to), importe = redondear_(numero_(pago && pago.amount));
  if (!de || !para || !(importe > 0)) throw new Error('Pago incompleto');
  return conBloqueo_(function (ss, yo) {
    var id = nuevoId_();
    var h = ss.getSheetByName('Pagos');
    h.getRange(h.getLastRow() + 1, 1, 1, 6).setValues([[id, new Date(), de, para, importe, yo.email]]);
    return id;
  });
}

function borrarPago(id) {
  return conBloqueo_(function (ss) { borrarFilas_(ss.getSheetByName('Pagos'), texto_(id)); });
}

/** Guarda la lista de personas (nombre y si siguen bajando). Los correos ya unidos se conservan. */
function guardarPersonas(lista) {
  return conBloqueo_(function (ss) {
    var correoPorNombre = {};
    filas_(ss, 'Personas').forEach(function (r) {
      var k = texto_(r[0]).toLowerCase();
      if (k && correo_(r[1])) correoPorNombre[k] = correo_(r[1]);
    });
    var filas = (lista || [])
      .filter(function (p) { return texto_(p && p.name); })
      .map(function (p) {
        var nombre = texto_(p.name);
        return [nombre, correoPorNombre[nombre.toLowerCase()] || '', p.active !== false];
      });
    var h = ss.getSheetByName('Personas');
    var viejas = Math.max(h.getLastRow() - 1, 1);
    h.getRange(2, 1, viejas, 3).clearDataValidations().clearContent();
    if (filas.length) {
      h.getRange(2, 3, filas.length, 1).insertCheckboxes(); // pone todas a FALSE, por eso va antes
      h.getRange(2, 1, filas.length, 3).setValues(filas);
    }
  });
}

function guardarProductos(lista) {
  var filas = (lista || [])
    .filter(function (p) { return texto_(p && p.name); })
    .map(function (p) { return [texto_(p.name), redondear_(numero_(p.price))]; });
  return conBloqueo_(function (ss) {
    var h = ss.getSheetByName('Productos');
    var viejas = Math.max(h.getLastRow() - 1, 1);
    h.getRange(2, 1, viejas, 2).clearContent();
    if (filas.length) h.getRange(2, 1, filas.length, 2).setValues(filas);
  });
}

/* ---------- Versión GitHub: la web llama aquí con fetch ---------- */

/** Recibe {action, args, token} y responde {ok, data} o {ok:false, error, code}. */
function doPost(e) {
  var acciones = {
    cargarDatos: cargarDatos, registrarme: registrarme, cambiarMiNombre: cambiarMiNombre,
    guardarRonda: guardarRonda, borrarRonda: borrarRonda, guardarPago: guardarPago, borrarPago: borrarPago,
    guardarPersonas: guardarPersonas, guardarProductos: guardarProductos
  };
  var salida;
  try {
    var peticion = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = acciones[peticion.action];
    if (!fn) throw new Error('Acción desconocida');
    MODO_API_ = true;
    CORREO_API_ = verificarToken_(peticion.token);
    salida = { ok: true, data: fn.apply(null, Array.isArray(peticion.args) ? peticion.args : []) };
  } catch (err) {
    salida = { ok: false, error: String((err && err.message) || err), code: (err && err.code) || '' };
  }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

/** Comprueba con Google que el token es auténtico, de esta app y no ha caducado. Devuelve el correo. */
function verificarToken_(token) {
  if (!token) throw conCodigo_('Entra con tu cuenta de Google.', 'login');
  if (GOOGLE_CLIENT_ID.indexOf('PEGA_AQUI') === 0) throw new Error('Falta poner el ID de cliente de Google en Codigo.gs (GOOGLE_CLIENT_ID).');
  var cache = CacheService.getScriptCache();
  var clave = 'tk_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token))).slice(0, 43);
  var guardado = cache.get(clave);
  if (guardado) return guardado;
  var res = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw conCodigo_('Tu sesión ha caducado. Vuelve a entrar.', 'login');
  var info = JSON.parse(res.getContentText());
  var ahora = Date.now() / 1000;
  if (info.aud !== GOOGLE_CLIENT_ID) throw conCodigo_('Ese inicio de sesión no es de esta app.', 'login');
  if (['accounts.google.com', 'https://accounts.google.com'].indexOf(info.iss) < 0) throw conCodigo_('Inicio de sesión no válido.', 'login');
  if (String(info.email_verified) !== 'true') throw conCodigo_('Tu correo de Google no está verificado.', 'login');
  if (!(Number(info.exp) > ahora)) throw conCodigo_('Tu sesión ha caducado. Vuelve a entrar.', 'login');
  var correo = correo_(info.email);
  cache.put(clave, correo, Math.max(30, Math.min(21600, Math.floor(Number(info.exp) - ahora))));
  return correo;
}

function conCodigo_(mensaje, codigo) { var e = new Error(mensaje); e.code = codigo; return e; }

/* ---------- Ayudantes privados (acaban en _ y no se pueden llamar desde la web) ---------- */

function correoActual_() {
  return MODO_API_ ? correo_(CORREO_API_) : correo_(Session.getActiveUser().getEmail());
}

/** El dueño siempre entra. Si la pestaña Permitidos tiene filas, solo entran esos correos o dominios. */
function permitido_(ss, email) {
  var dueno = PropertiesService.getScriptProperties().getProperty('DUENO');
  if (dueno && email === correo_(dueno)) return true;
  var lista = filas_(ss, 'Permitidos').map(function (r) { return correo_(r[0]); }).filter(function (x) { return x; });
  if (!lista.length) return !MODO_API_;
  return lista.some(function (x) {
    return x === email || (x.charAt(0) === '@' && email.length > x.length && email.slice(-x.length) === x);
  });
}

function usuario_(ss, filasPersonas) {
  var email = correoActual_();
  var name = null;
  if (email && !permitido_(ss, email)) return { email: email, name: null, blocked: true };
  if (email) {
    (filasPersonas || filas_(ss, 'Personas')).some(function (r) {
      if (correo_(r[1]) === email && texto_(r[0])) { name = texto_(r[0]); return true; }
      return false;
    });
  }
  return { email: email, name: name };
}

function libro_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('HOJA_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abre Apps Script desde la hoja (Extensiones > Apps Script) y ejecuta "configurar"');
  props.setProperty('HOJA_ID', ss.getId());
  return ss;
}

function asegurar_(ss) {
  Object.keys(HOJAS).forEach(function (nombre) {
    if (ss.getSheetByName(nombre)) return;
    var h = ss.insertSheet(nombre);
    var cab = HOJAS[nombre];
    h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    h.setFrozenRows(1);
    if (nombre === 'Rondas') {
      h.getRange('B:B').setNumberFormat(FORMATO_FECHA);
      h.getRange('F:F').setNumberFormat(FORMATO_EUR);
    } else if (nombre === 'Pagos') {
      h.getRange('B:B').setNumberFormat(FORMATO_FECHA);
      h.getRange('E:E').setNumberFormat(FORMATO_EUR);
    } else if (nombre === 'Productos') {
      h.getRange('B:B').setNumberFormat(FORMATO_EUR);
      h.getRange(2, 1, PRODUCTOS_INICIALES.length, 2).setValues(PRODUCTOS_INICIALES);
    } else if (nombre === 'Permitidos') {
      h.getRange('C1').setValue('Un correo por fila (ana@gmail.com) o un dominio entero (@tuempresa.com). ' +
        'Vacía: en la versión de Apps Script entra quien tenga acceso a la hoja; en la de GitHub, solo el dueño.');
      h.getRange('C1').setFontColor('#666666');
    }
  });
  var props = PropertiesService.getScriptProperties();
  if (MODO_API_ && !props.getProperty('DUENO')) {
    // En la versión GitHub el servidor se ejecuta como el dueño.
    props.setProperty('DUENO', correo_(Session.getEffectiveUser().getEmail()));
  }
  if (!ss.getSheetByName('Saldos')) crearSaldos_(ss);
  migrar_(ss);
}

/** Pone al día una hoja creada con la versión anterior (sin correos ni autor). */
function migrar_(ss) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('ESQUEMA') === ESQUEMA) return;
  var p = ss.getSheetByName('Personas');
  if (texto_(p.getRange(1, 2).getValue()) === 'Activo') {
    p.insertColumnAfter(1);
    p.getRange(1, 2).setValue('Email').setFontWeight('bold');
  }
  ['Rondas', 'Pagos'].forEach(function (nombre) {
    var h = ss.getSheetByName(nombre), col = HOJAS[nombre].length;
    if (!texto_(h.getRange(1, col).getValue())) h.getRange(1, col).setValue('Apuntado por').setFontWeight('bold');
  });
  props.setProperty('ESQUEMA', ESQUEMA);
}

/** Pestaña con el saldo de cada persona, calculado con fórmulas a partir de Rondas y Pagos. */
function crearSaldos_(ss) {
  var h = ss.insertSheet('Saldos', 0);
  h.getRange('A1:C1').setValues([['Persona', 'Saldo', 'Estado']]).setFontWeight('bold');
  h.setFrozenRows(1);
  h.getRange('A2').setFormula('=FILTER(Personas!A2:A, Personas!A2:A<>"")');
  h.getRange('B2').setFormula(
    '=MAP(A2:A, LAMBDA(p, IF(p="", "", ROUND(' +
    'SUMIFS(Rondas!F:F, Rondas!C:C, p) - SUMIFS(Rondas!F:F, Rondas!D:D, p) + ' +
    'SUMIFS(Pagos!E:E, Pagos!C:C, p) - SUMIFS(Pagos!E:E, Pagos!D:D, p), 2))))'
  );
  h.getRange('C2').setFormula('=MAP(B2:B, LAMBDA(s, IF(s="", "", IF(s>0, "Le deben", IF(s<0, "Debe", "En paz")))))');
  h.getRange('B:B').setNumberFormat(FORMATO_EUR);
  h.getRange('E1').setValue('Saldo positivo: ha adelantado más de lo que ha tomado. Negativo: debe dinero.');
  h.getRange('E1').setFontColor('#666666');
  h.setColumnWidth(1, 160);
  h.setColumnWidth(2, 110);
  h.setColumnWidth(3, 110);
}

function conBloqueo_(fn, opciones) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = libro_();
    asegurar_(ss);
    var yo = usuario_(ss);
    if (!yo.email) throw new Error(SIN_CORREO);
    if (yo.blocked) throw new Error(SIN_PERMISO);
    if (!yo.name && !(opciones && opciones.sinAlta)) throw new Error(SIN_ALTA);
    var id = fn(ss, yo);
    SpreadsheetApp.flush();
    var datos = cargarDatos();
    if (id) datos.lastId = id;
    return datos;
  } finally {
    lock.releaseLock();
  }
}

function filas_(ss, nombre) {
  var h = ss.getSheetByName(nombre);
  var n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2, 1, n, HOJAS[nombre].length).getValues();
}

function ultimaFilaConNombre_(h) {
  var n = h.getLastRow();
  if (n < 2) return 1;
  var col = h.getRange(2, 1, n - 1, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) if (texto_(col[i][0])) return i + 2;
  return 1;
}

function borrarFilas_(h, id) {
  if (!id) return;
  var n = h.getLastRow() - 1;
  if (n < 1) return;
  var ids = h.getRange(2, 1, n, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (texto_(ids[i][0]) === id) h.deleteRow(i + 2);
  }
}

/** Cambia un nombre por otro en dos columnas seguidas a partir de primeraCol. */
function renombrar_(h, primeraCol, viejo, nuevo) {
  var n = h.getLastRow() - 1;
  if (n < 1 || !viejo) return;
  var rango = h.getRange(2, primeraCol, n, 2);
  var valores = rango.getValues(), cambiado = false;
  valores.forEach(function (fila) {
    for (var k = 0; k < fila.length; k++) if (texto_(fila[k]) === viejo) { fila[k] = nuevo; cambiado = true; }
  });
  if (cambiado) rango.setValues(valores);
}

function texto_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }

function correo_(v) { return texto_(v).toLowerCase(); }

function igual_(a, b) { return texto_(a).toLowerCase() === texto_(b).toLowerCase(); }

function numero_(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v == null ? '' : v).replace(/[€\s]/g, '').replace(',', '.'));
  return isFinite(n) ? n : 0;
}

function redondear_(n) { return Math.round(n * 100) / 100; }

// Las fechas no viajan a la web como Date: se mandan como texto ISO.
function fecha_(v) {
  if (v instanceof Date && !isNaN(v)) return v.toISOString();
  var d = new Date(v);
  return isNaN(d) ? new Date(0).toISOString() : d.toISOString();
}

function activo_(v) {
  if (v === false) return false;
  return !/^(no|false|falso|0)$/i.test(texto_(v));
}

function nuevoId_() { return Utilities.getUuid().slice(0, 8); }
