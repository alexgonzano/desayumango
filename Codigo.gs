/**
 * DESAYUMANGO · servidor en Google Sheets (con inicio de sesión de Google)
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
 *  - GOOGLE_CLIENT_ID (justo debajo) es el ID de cliente de Google (Google Cloud > Clientes).
 *    Es el mismo que va en config.js.
 *  - Crea OTRA implementación de tipo Aplicación web con:
 *      Ejecutar como: Yo
 *      Quién tiene acceso: Cualquier usuario
 *    Aquí "Cualquier usuario" es correcto: quién entra lo controla el botón de Google
 *    y la pestaña Permitidos, que se comprueban en cada petición.
 *  - Quien no tiene acceso lo pide desde la app y cualquiera del equipo lo acepta (pestaña Solicitudes).
 *    Al aceptarlo se apunta solo en Permitidos; también puedes escribir ahí un correo por fila,
 *    o @tuempresa.com para todo un dominio. Quien ejecutó "configurar" (el dueño) siempre tiene acceso,
 *    y desde la app puede dar y quitar acceso sin solicitud.
 *  - Al entrar con Google, la app recibe una sesión propia de DIAS_SESION días que se renueva sola.
 *    Para cerrar todas las sesiones a la vez, borra SECRETO_SESION en Propiedades del script.
 *  - Avisos en el móvil (opcional): añade el archivo Avisos.gs y sigue los pasos de su cabecera.
 */

// ID de cliente de Google para la versión en GitHub. No es secreto: va en todas las páginas que abre la app.
var GOOGLE_CLIENT_ID = '141212573757-g76sf6huqj4bj1dannkai11tpiuuflia.apps.googleusercontent.com';

var HOJAS = {
  Rondas: ['ID', 'Fecha', 'Pagó', 'Persona', 'Producto', 'Precio', 'Apuntado por'],
  Pagos: ['ID', 'Fecha', 'De', 'Para', 'Importe', 'Apuntado por'],
  Personas: ['Nombre', 'Email', 'Activo'],
  Productos: ['Producto', 'Precio'],
  Permitidos: ['Correo o dominio'],
  Solicitudes: ['Correo', 'Nombre', 'Fecha'],
  Avisos: ['Correo', 'Dirección', 'p256dh', 'auth', 'Fecha']
};
var MAX_SOLICITUDES = 30;
var DIAS_SESION = 30; // la sesión de la app dura esto y se renueva sola si se usa
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
var SIN_PERMISO = 'Tu cuenta no tiene acceso a esta app. Pide acceso desde la app y espera a que alguien del equipo lo acepte.';

// Solo durante una petición de la versión GitHub: el correo sale del token de Google ya verificado.
var MODO_API_ = false;
var CORREO_API_ = '';

/* ---------- Web ---------- */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DesayuMango')
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
  var solicitudes = filas_(ss, 'Solicitudes')
    .filter(function (r) { return correo_(r[0]); })
    .map(function (r) { return { email: correo_(r[0]), name: texto_(r[1]), date: fecha_(r[2]) }; });
  if (!me.email || me.blocked) {
    // A quien no tiene acceso solo se le dice si ya pidió entrar.
    var mia = solicitudes.filter(function (x) { return x.email === me.email; })[0];
    if (me.email) { me.requested = !!mia; me.requestedName = mia ? mia.name : ''; }
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
    appUrl: urlApp_(),
    pushKey: clavePublicaAvisos_(),
    access: {
      requests: solicitudes,
      // La lista completa de quién tiene acceso solo la ve el dueño, que es quien puede quitarlo.
      list: me.owner ? filas_(ss, 'Permitidos').map(function (r) { return correo_(r[0]); }).filter(function (x) { return x; })
        .map(function (e) { return { entry: e, name: nombrePorCorreo[e] || '' }; }) : null
    }
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
  return conBloqueo_(function (ss, yo, avisar) {
    var id = nuevoId_(), ahora = new Date();
    var filas = items.map(function (it) {
      return [id, ahora, pagador, texto_(it.person), texto_(it.product), redondear_(numero_(it.price)), yo.email];
    });
    var h = ss.getSheetByName('Rondas');
    h.getRange(h.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
    avisar('ronda-' + id, textosRonda_(yo.name, filas, false));
    return id;
  });
}

function borrarRonda(id) {
  return conBloqueo_(function (ss, yo, avisar) {
    var rid = texto_(id);
    var filas = filas_(ss, 'Rondas').filter(function (r) { return rid && texto_(r[0]) === rid && texto_(r[3]); });
    borrarFilas_(ss.getSheetByName('Rondas'), rid);
    if (filas.length) avisar('ronda-' + rid, textosRonda_(yo.name, filas, true));
  });
}

function guardarPago(pago) {
  var de = texto_(pago && pago.from), para = texto_(pago && pago.to), importe = redondear_(numero_(pago && pago.amount));
  if (!de || !para || !(importe > 0)) throw new Error('Pago incompleto');
  return conBloqueo_(function (ss, yo, avisar) {
    var id = nuevoId_();
    var fila = [id, new Date(), de, para, importe, yo.email];
    var h = ss.getSheetByName('Pagos');
    h.getRange(h.getLastRow() + 1, 1, 1, 6).setValues([fila]);
    avisar('pago-' + id, textosPago_(yo.name, fila, false));
    return id;
  });
}

function borrarPago(id) {
  return conBloqueo_(function (ss, yo, avisar) {
    var pid = texto_(id);
    var fila = filas_(ss, 'Pagos').filter(function (r) { return pid && texto_(r[0]) === pid; })[0];
    borrarFilas_(ss.getSheetByName('Pagos'), pid);
    if (fila) avisar('pago-' + pid, textosPago_(yo.name, fila, true));
  });
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

/* ---------- Acceso: pedir, aceptar, rechazar y quitar ---------- */

/** Quien no tiene acceso pide entrar. Queda pendiente hasta que alguien del equipo la acepte. */
function solicitarAcceso(nombre) {
  nombre = texto_(nombre);
  if (!nombre) throw new Error('Escribe tu nombre');
  if (nombre.length > 40) throw new Error('El nombre es demasiado largo');
  return conBloqueo_(function (ss, yo) {
    if (!yo.blocked) return; // ya tiene acceso
    var h = ss.getSheetByName('Solicitudes');
    var filas = filas_(ss, 'Solicitudes');
    for (var i = 0; i < filas.length; i++) {
      if (correo_(filas[i][0]) === yo.email) { h.getRange(i + 2, 2, 1, 2).setValues([[nombre, new Date()]]); return; }
    }
    var pendientes = filas.filter(function (r) { return correo_(r[0]); }).length;
    if (pendientes >= MAX_SOLICITUDES) throw new Error('Hay demasiadas solicitudes pendientes. Avisa a alguien del equipo.');
    h.getRange(ultimaFilaConNombre_(h) + 1, 1, 1, 3).setValues([[yo.email, nombre, new Date()]]);
  }, { sinAlta: true, bloqueado: true });
}

/** Da acceso a un correo. Una solicitud la puede aceptar cualquiera del equipo; dar acceso sin solicitud, solo el dueño. */
function aprobarAcceso(correo) {
  correo = correo_(correo);
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo) && !/^@[^@\s]+\.[^@\s]+$/.test(correo)) throw new Error('Escribe un correo válido');
  return conBloqueo_(function (ss, yo) {
    var filas = filas_(ss, 'Solicitudes'), nombre = '', pedida = false;
    filas.forEach(function (r) { if (correo_(r[0]) === correo) { pedida = true; nombre = texto_(r[1]); } });
    if (!pedida && !yo.owner) throw new Error('Solo el dueño puede dar acceso a alguien que no lo ha pedido.');
    if (correo === '@gmail.com' || correo === '@googlemail.com') throw new Error('Eso daría acceso a cualquier persona con Gmail. Añade los correos uno a uno.');
    var p = ss.getSheetByName('Permitidos');
    var lista = filas_(ss, 'Permitidos').map(function (r) { return correo_(r[0]); });
    if (lista.indexOf(correo) < 0) p.getRange(ultimaFilaConNombre_(p) + 1, 1).setValue(correo);
    if (pedida) borrarFilasCorreo_(ss.getSheetByName('Solicitudes'), correo);
    if (nombre && correo.charAt(0) !== '@') altaConNombre_(ss, correo, nombre);
  });
}

/** Descarta una solicitud. Puede hacerlo cualquiera del equipo. */
function rechazarAcceso(correo) {
  correo = correo_(correo);
  return conBloqueo_(function (ss) { borrarFilasCorreo_(ss.getSheetByName('Solicitudes'), correo); });
}

/** Quita el acceso a un correo o dominio. Solo el dueño. */
function quitarAcceso(entrada) {
  entrada = correo_(entrada);
  return conBloqueo_(function (ss, yo) {
    if (!yo.owner) throw new Error('Solo el dueño puede quitar el acceso.');
    if (esDueno_(entrada)) throw new Error('El dueño siempre tiene acceso.');
    borrarFilasCorreo_(ss.getSheetByName('Permitidos'), entrada);
  });
}

/* ---------- Avisos en el móvil: a quién y qué se le dice (el cifrado está en Avisos.gs) ---------- */

/** Guarda este dispositivo para mandarle avisos. La web lo manda cada vez que se abre. */
function guardarAviso(sub) {
  var endpoint = texto_(sub && sub.endpoint), p256dh = texto_(sub && sub.keys && sub.keys.p256dh), auth = texto_(sub && sub.keys && sub.keys.auth);
  // Solo direcciones de los servicios de avisos de los navegadores: este servidor no manda nada a ningún otro sitio.
  if (!/^https:\/\/(fcm\.googleapis\.com|updates\.push\.services\.mozilla\.com|web\.push\.apple\.com|[a-z0-9-]+\.notify\.windows\.com)\/\S+$/.test(endpoint) ||
      endpoint.length > 1000 || !/^[\w-]{86,88}=*$/.test(p256dh) || !/^[\w-]{22,24}=*$/.test(auth)) {
    throw new Error('Este navegador ha dado unos datos de avisos que no son válidos');
  }
  return conBloqueo_(function (ss, yo) {
    var h = ss.getSheetByName('Avisos'), filas = filas_(ss, 'Avisos');
    for (var i = 0; i < filas.length; i++) {
      if (texto_(filas[i][1]) !== endpoint) continue;
      // Ya estaba. Si ahora entra otra cuenta desde el mismo navegador, los avisos pasan a ser suyos.
      if (correo_(filas[i][0]) !== yo.email || texto_(filas[i][2]) !== p256dh || texto_(filas[i][3]) !== auth) {
        h.getRange(i + 2, 1, 1, 4).setValues([[yo.email, endpoint, p256dh, auth]]);
      }
      return;
    }
    h.getRange(h.getLastRow() + 1, 1, 1, 5).setValues([[yo.email, endpoint, p256dh, auth, new Date()]]);
  }, { sinAlta: true, sinDatos: true });
}

/** Textos de una ronda (filas como las de la pestaña Rondas) para quien pagó y para cada uno que tomó algo. */
function textosRonda_(quien, filas, borrada) {
  var pagador = texto_(filas[0][2]), total = 0, deOtros = 0, porPersona = {}, textos = {};
  filas.forEach(function (r) {
    var persona = texto_(r[3]), centimos = Math.round(numero_(r[5]) * 100);
    total += centimos;
    if (persona === pagador) return;
    deOtros += centimos;
    var p = porPersona[persona] = porPersona[persona] || { productos: [], centimos: 0 };
    p.productos.push(texto_(r[4]));
    p.centimos += centimos;
  });
  Object.keys(porPersona).forEach(function (persona) {
    var p = porPersona[persona];
    textos[persona] = borrada
      ? { title: 'Ronda anulada', body: quien + ' ha borrado la ronda que pagó ' + pagador + '. Ya no le debes ' + euros_(p.centimos) + '.' }
      : { title: '☕ ' + pagador + ' ha pagado la ronda', body: lista_(p.productos) + ': ' + euros_(p.centimos) + '. Se lo debes a ' + pagador + '.' };
  });
  var n = filas.length;
  textos[pagador] = borrada
    ? { title: 'Ronda anulada', body: quien + ' ha borrado tu ronda de ' + euros_(total) + '.' }
    : { title: quien + ' ha apuntado tu ronda',
        body: 'Pagaste ' + n + (n === 1 ? ' café' : ' cafés') + ' (' + euros_(total) + ').' + (deOtros ? ' Te deben ' + euros_(deOtros) + '.' : '') };
  return textos;
}

/** Textos de un pago (fila como las de la pestaña Pagos) para quien devolvió el dinero y quien lo recibió. */
function textosPago_(quien, fila, borrado) {
  var de = texto_(fila[2]), para = texto_(fila[3]), importe = euros_(Math.round(numero_(fila[4]) * 100)), textos = {};
  textos[de] = borrado
    ? { title: 'Pago anulado', body: quien + ' ha borrado tu pago de ' + importe + ' a ' + para + '.' }
    : { title: quien + ' ha apuntado tu pago', body: 'Le devolviste ' + importe + ' a ' + para + '.' };
  textos[para] = borrado
    ? { title: 'Pago anulado', body: quien + ' ha borrado el pago de ' + importe + ' que te hizo ' + de + '.' }
    : { title: de + ' te ha devuelto ' + importe, body: quien === de ? 'Ya está apuntado en las cuentas.' : 'Lo ha apuntado ' + quien + '.' };
  return textos;
}

/** Un aviso por dispositivo para cada persona de «textos» que tiene cuenta, menos quien lo ha hecho. */
function destinatarios_(ss, yo, tag, textos) {
  var correoDe = {}, avisos = [];
  filas_(ss, 'Personas').forEach(function (r) { if (texto_(r[0]) && correo_(r[1])) correoDe[texto_(r[0]).toLowerCase()] = correo_(r[1]); });
  var dispositivos = filas_(ss, 'Avisos');
  Object.keys(textos).forEach(function (nombre) {
    var correo = correoDe[nombre.toLowerCase()];
    if (!correo || correo === yo.email) return;
    dispositivos.forEach(function (r) {
      if (correo_(r[0]) !== correo) return;
      avisos.push({ endpoint: texto_(r[1]), p256dh: texto_(r[2]), auth: texto_(r[3]), mensaje: { title: textos[nombre].title, body: textos[nombre].body, tag: tag } });
    });
  });
  return avisos;
}

/** Manda los avisos. Si falla, no pasa nada: lo apuntado ya está guardado. */
function enviarAvisos_(avisos) {
  if (!avisos.length || !clavePublicaAvisos_()) return;
  try {
    var respuestas = UrlFetchApp.fetchAll(avisos.map(function (a) { return peticionAviso_(a, a.mensaje); }));
    var caducados = [];
    respuestas.forEach(function (r, i) {
      var c = r.getResponseCode();
      if (c === 404 || c === 410) caducados.push(avisos[i].endpoint); // app desinstalada o avisos desactivados
      else if (c >= 300) console.warn('Aviso no entregado (' + c + '): ' + r.getContentText().slice(0, 300));
    });
    if (caducados.length) borrarDispositivos_(caducados);
  } catch (err) {
    console.error('No se han podido mandar los avisos: ' + ((err && err.stack) || err));
  }
}

/** Quita de la pestaña Avisos los dispositivos que el servicio de avisos ya no reconoce. */
function borrarDispositivos_(endpoints) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var h = libro_().getSheetByName('Avisos'), n = h.getLastRow() - 1;
    if (n < 1) return;
    var col = h.getRange(2, 2, n, 1).getValues();
    for (var i = col.length - 1; i >= 0; i--) if (endpoints.indexOf(texto_(col[i][0])) >= 0) h.deleteRow(i + 2);
  } finally {
    lock.releaseLock();
  }
}

/** «2× Cortado y Cola Cao» */
function lista_(productos) {
  var cuenta = {}, orden = [];
  productos.forEach(function (p) { if (!cuenta[p]) orden.push(p); cuenta[p] = (cuenta[p] || 0) + 1; });
  var partes = orden.map(function (p) { return (cuenta[p] > 1 ? cuenta[p] + '× ' : '') + p; });
  return partes.length > 1 ? partes.slice(0, -1).join(', ') + ' y ' + partes[partes.length - 1] : partes[0];
}

function euros_(centimos) { return (centimos / 100).toFixed(2).replace('.', ',') + ' €'; }

/* ---------- Versión GitHub: la web llama aquí con fetch ---------- */

/** Recibe {action, args, token} y responde {ok, data} o {ok:false, error, code}. */
function doPost(e) {
  var acciones = {
    cargarDatos: cargarDatos, registrarme: registrarme, cambiarMiNombre: cambiarMiNombre,
    guardarRonda: guardarRonda, borrarRonda: borrarRonda, guardarPago: guardarPago, borrarPago: borrarPago,
    guardarPersonas: guardarPersonas, guardarProductos: guardarProductos,
    solicitarAcceso: solicitarAcceso, aprobarAcceso: aprobarAcceso, rechazarAcceso: rechazarAcceso, quitarAcceso: quitarAcceso,
    guardarAviso: guardarAviso
  };
  var salida;
  try {
    var peticion = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    MODO_API_ = true;
    if (peticion.action === 'iniciarSesion') {
      // Entrar: se comprueba el token de Google una vez y se da una sesión propia de la app.
      CORREO_API_ = verificarToken_(peticion.token);
      salida = { ok: true, data: cargarDatos(), session: crearSesion_(CORREO_API_) };
    } else {
      var fn = acciones[peticion.action];
      if (!fn) throw new Error('Acción desconocida');
      var renovar = false;
      if (peticion.session) {
        var ses = leerSesion_(peticion.session);
        CORREO_API_ = correo_(ses.e);
        renovar = (Date.now() / 1000 - ses.i) > 86400; // como mucho una vez al día
      } else {
        CORREO_API_ = verificarToken_(peticion.token); // webs aún sin actualizar
      }
      salida = { ok: true, data: fn.apply(null, Array.isArray(peticion.args) ? peticion.args : []) };
      if (renovar) salida.session = crearSesion_(CORREO_API_);
    }
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

/* Sesión de la app: «v1.<datos>.<firma>», firmada con un secreto que solo conoce este servidor.
   Quién tiene acceso se sigue comprobando en cada petición (pestaña Permitidos). */
function crearSesion_(correo) {
  var ahora = Math.floor(Date.now() / 1000);
  var datos = sinRelleno_(Utilities.base64EncodeWebSafe(JSON.stringify({ e: correo, i: ahora, x: ahora + DIAS_SESION * 86400 })));
  return 'v1.' + datos + '.' + firma_('v1.' + datos);
}

function leerSesion_(sesion) {
  var partes = String(sesion || '').split('.');
  if (partes.length !== 3 || partes[0] !== 'v1' || firma_('v1.' + partes[1]) !== partes[2]) {
    throw conCodigo_('Vuelve a entrar con tu cuenta de Google.', 'login');
  }
  var datos = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(conRelleno_(partes[1]))).getDataAsString());
  if (!datos.e || !(Number(datos.x) > Date.now() / 1000)) throw conCodigo_('Tu sesión ha caducado. Vuelve a entrar.', 'login');
  return datos;
}

function firma_(texto) {
  return sinRelleno_(Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(texto, secretoSesion_())));
}

/** Para cerrar todas las sesiones abiertas a la vez, borra SECRETO_SESION en Propiedades del script. */
function secretoSesion_() {
  var props = PropertiesService.getScriptProperties();
  var secreto = props.getProperty('SECRETO_SESION');
  if (!secreto) {
    secreto = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty('SECRETO_SESION', secreto);
  }
  return secreto;
}

function sinRelleno_(b) { return String(b).replace(/=+$/, ''); }
function conRelleno_(b) { b = String(b); while (b.length % 4) b += '='; return b; }

function conCodigo_(mensaje, codigo) { var e = new Error(mensaje); e.code = codigo; return e; }

/* ---------- Ayudantes privados (acaban en _ y no se pueden llamar desde la web) ---------- */

function correoActual_() {
  return MODO_API_ ? correo_(CORREO_API_) : correo_(Session.getActiveUser().getEmail());
}

function esDueno_(email) {
  var dueno = PropertiesService.getScriptProperties().getProperty('DUENO');
  return !!(email && dueno && correo_(email) === correo_(dueno));
}

/** Da de alta a quien acaba de recibir acceso, con el nombre que pidió, si está libre o es de alguien sin cuenta. */
function altaConNombre_(ss, correo, nombre) {
  var h = ss.getSheetByName('Personas');
  var filas = filas_(ss, 'Personas');
  for (var i = 0; i < filas.length; i++) if (correo_(filas[i][1]) === correo) return;
  for (var j = 0; j < filas.length; j++) {
    if (igual_(filas[j][0], nombre)) {
      if (correo_(filas[j][1])) return; // ese nombre es de otra cuenta: lo elegirá al entrar
      h.getRange(j + 2, 2, 1, 2).setValues([[correo, true]]);
      return;
    }
  }
  var fila = ultimaFilaConNombre_(h) + 1;
  h.getRange(fila, 3).insertCheckboxes();
  h.getRange(fila, 1, 1, 3).setValues([[nombre, correo, true]]);
}

/** Borra las filas cuya columna A es ese correo (sin distinguir mayúsculas). */
function borrarFilasCorreo_(h, correo) {
  var n = h.getLastRow() - 1;
  if (n < 1 || !correo) return;
  var col = h.getRange(2, 1, n, 1).getValues();
  for (var i = col.length - 1; i >= 0; i--) if (correo_(col[i][0]) === correo) h.deleteRow(i + 2);
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
  return { email: email, name: name, owner: esDueno_(email) };
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
      h.getRange('C1').setValue('Correos con acceso. Se rellena sola al aceptar solicitudes en la app; también puedes escribir ' +
        'un correo por fila o un dominio entero (@tuempresa.com). Vacía: en la versión de Apps Script entra quien tenga acceso a la hoja; en la de GitHub, solo el dueño.');
      h.getRange('C1').setFontColor('#666666');
    } else if (nombre === 'Avisos') {
      h.getRange('B:D').setNumberFormat('@'); // texto tal cual, que la hoja no lo tome por números o fechas
      h.getRange('E:E').setNumberFormat(FORMATO_FECHA);
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

/** Hace un cambio con la hoja bloqueada. fn(ss, yo, avisar) puede pedir avisos con avisar(tag, textos por nombre). */
function conBloqueo_(fn, opciones) {
  var lock = LockService.getScriptLock(), avisos = [], datos;
  lock.waitLock(20000);
  try {
    var ss = libro_();
    asegurar_(ss);
    var yo = usuario_(ss);
    if (!yo.email) throw new Error(SIN_CORREO);
    if (yo.blocked && !(opciones && opciones.bloqueado)) throw new Error(SIN_PERMISO);
    if (!yo.name && !(opciones && opciones.sinAlta)) throw new Error(SIN_ALTA);
    var id = fn(ss, yo, function (tag, textos) { avisos = avisos.concat(destinatarios_(ss, yo, tag, textos)); });
    SpreadsheetApp.flush();
    if (opciones && opciones.sinDatos) return true;
    datos = cargarDatos();
    if (id) datos.lastId = id;
  } finally {
    lock.releaseLock();
  }
  enviarAvisos_(avisos); // ya sin el bloqueo, para no hacer esperar a los demás
  return datos;
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
