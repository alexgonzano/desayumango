// Comprobación de los avisos, sin dependencias: node avisos.test.js
// Carga Codigo.gs y Avisos.gs como Apps Script (sin crypto ni TextEncoder, con Utilities simulado),
// cifra un aviso y lo descifra aparte con el crypto de Node (RFC 8291), verifica la firma VAPID
// y revisa los textos y a quién le llegan.
const assert = require('assert'), crypto = require('crypto'), fs = require('fs'), path = require('path'), vm = require('vm');

const props = { DUENO: 'dueno@ejemplo.com' };
const conSigno = (buf) => Array.from(buf, (b) => (b > 127 ? b - 256 : b));
const sinSigno = (arr) => Buffer.from(Array.from(arr, (b) => b & 255));
const ctx = vm.createContext({
  console,
  Logger: { log() {} },
  PropertiesService: { getScriptProperties: () => ({
    getProperty: (k) => (k in props ? props[k] : null),
    setProperties: (o) => { for (const k in o) props[k] = String(o[k]); }
  }) },
  Utilities: {
    DigestAlgorithm: { SHA_256: 'sha256' },
    computeDigest: (alg, s) => conSigno(crypto.createHash(alg).update(String(s), 'utf8').digest()),
    getUuid: () => crypto.randomUUID(),
    newBlob: (s) => ({ getBytes: () => conSigno(Buffer.from(String(s), 'utf8')) }),
    base64EncodeWebSafe: (bytes) => sinSigno(bytes).toString('base64').replace(/\+/g, '-').replace(/\//g, '_'), // con relleno, como Apps Script
    base64DecodeWebSafe: (s) => conSigno(Buffer.from(String(s), 'base64url'))
  }
});
for (const f of ['Codigo.gs', 'Avisos.gs']) vm.runInContext(fs.readFileSync(path.join(__dirname, f), 'utf8'), ctx, { filename: f });
const gs = (codigo, vars) => { Object.assign(ctx, vars); return vm.runInContext(codigo, ctx); };

// Claves del servidor
gs('crearClavesAvisos()');
assert.match(props.AVISOS_PUBLICA, /^[\w-]{87}$/);
const publica = props.AVISOS_PUBLICA;
gs('crearClavesAvisos()');
assert.equal(props.AVISOS_PUBLICA, publica, 'no se cambian las claves si ya existen');

// Un navegador de prueba
const navegador = crypto.createECDH('prime256v1'); navegador.generateKeys();
const auth = crypto.randomBytes(16);
const dispositivo = { endpoint: 'https://fcm.googleapis.com/fcm/send/abc:APA91b', p256dh: navegador.getPublicKey().toString('base64url'), auth: auth.toString('base64url') };
const mensaje = { title: '☕ Ana ha pagado la ronda', body: 'Cortado: 1,40 €. Se lo debes a Ana.', tag: 'ronda-1' };
const pet = gs('peticionAviso_(dispositivo, mensaje)', { dispositivo, mensaje });
assert.equal(pet.url, dispositivo.endpoint);
assert.equal(pet.headers['Content-Encoding'], 'aes128gcm');

// Descifrar como lo haría el navegador
const cuerpo = sinSigno(pet.payload);
const sal = cuerpo.subarray(0, 16), idlen = cuerpo[20], efimera = cuerpo.subarray(21, 21 + idlen), cifrado = cuerpo.subarray(21 + idlen);
assert.equal(cuerpo.readUInt32BE(16), 4096);
assert.equal(idlen, 65);
const ikm = crypto.hkdfSync('sha256', navegador.computeSecret(efimera), auth, Buffer.concat([Buffer.from('WebPush: info\0'), navegador.getPublicKey(), efimera]), 32);
const cek = crypto.hkdfSync('sha256', ikm, sal, Buffer.from('Content-Encoding: aes128gcm\0'), 16);
const nonce = crypto.hkdfSync('sha256', ikm, sal, Buffer.from('Content-Encoding: nonce\0'), 12);
const d = crypto.createDecipheriv('aes-128-gcm', Buffer.from(cek), Buffer.from(nonce));
d.setAuthTag(cifrado.subarray(-16));
const claro = Buffer.concat([d.update(cifrado.subarray(0, -16)), d.final()]);
assert.equal(claro[claro.length - 1], 2);
assert.deepEqual(JSON.parse(claro.subarray(0, -1).toString('utf8')), mensaje);

// Firma VAPID
const [, jwt, k] = pet.headers.Authorization.match(/^vapid t=([^,]+), k=(.+)$/);
assert.equal(k, publica);
const [cab, datos, firma] = jwt.split('.');
const reclamos = JSON.parse(Buffer.from(datos, 'base64url'));
assert.equal(reclamos.aud, 'https://fcm.googleapis.com');
assert.equal(reclamos.sub, 'mailto:dueno@ejemplo.com');
assert(reclamos.exp > Date.now() / 1000 && reclamos.exp <= Date.now() / 1000 + 24 * 3600);
const pub = Buffer.from(publica, 'base64url');
const clave = crypto.createPublicKey({ key: { kty: 'EC', crv: 'P-256', x: pub.subarray(1, 33).toString('base64url'), y: pub.subarray(33).toString('base64url') }, format: 'jwk' });
assert(crypto.verify('sha256', Buffer.from(cab + '.' + datos), { key: clave, dsaEncoding: 'ieee-p1363' }, Buffer.from(firma, 'base64url')), 'firma VAPID');

// Textos de una ronda: paga Ana, la apunta Marta
const filas = [
  ['r1', 0, 'Ana', 'Ana', 'Cortado', 1.4, 'marta@x'],
  ['r1', 0, 'Ana', 'Javi', 'Café con leche', 1.6, 'marta@x'],
  ['r1', 0, 'Ana', 'Lucía', 'Cortado', 1.4, 'marta@x'],
  ['r1', 0, 'Ana', 'Lucía', 'Cortado', 1.4, 'marta@x'],
  ['r1', 0, 'Ana', 'Lucía', 'Cola Cao', 1.8, 'marta@x']
];
const t = gs("textosRonda_('Marta', filas, false)", { filas });
assert.deepEqual({ ...t.Javi }, { title: '☕ Ana ha pagado la ronda', body: 'Café con leche: 1,60 €. Se lo debes a Ana.' });
assert.equal(t['Lucía'].body, '2× Cortado y Cola Cao: 4,60 €. Se lo debes a Ana.');
assert.deepEqual({ ...t.Ana }, { title: 'Marta ha apuntado tu ronda', body: 'Pagaste 5 cafés (7,60 €). Te deben 6,20 €.' });
const tb = gs("textosRonda_('Javi', filas, true)");
assert.equal(tb['Lucía'].body, 'Javi ha borrado la ronda que pagó Ana. Ya no le debes 4,60 €.');
assert.deepEqual({ ...tb.Ana }, { title: 'Ronda anulada', body: 'Javi ha borrado tu ronda de 7,60 €.' });

// Textos de un pago: Javi devuelve 3,20 € a Ana
const pago = ['p1', 0, 'Javi', 'Ana', 3.2, 'javi@x'];
const tp = gs("textosPago_('Javi', pago, false)", { pago });
assert.deepEqual({ ...tp.Ana }, { title: 'Javi te ha devuelto 3,20 €', body: 'Ya está apuntado en las cuentas.' });
const tp2 = gs("textosPago_('Marta', pago, false)");
assert.equal(tp2.Ana.body, 'Lo ha apuntado Marta.');
assert.deepEqual({ ...tp2.Javi }, { title: 'Marta ha apuntado tu pago', body: 'Le devolviste 3,20 € a Ana.' });
assert.equal(gs("textosPago_('Ana', pago, true)").Javi.body, 'Ana ha borrado tu pago de 3,20 € a Ana.');

// A quién le llega: a cada dispositivo de quien tiene cuenta, nunca a quien lo apunta
const hoja = (f) => ({ getLastRow: () => f.length + 1, getRange: (r, c, n, m) => ({ getValues: () => f.slice(r - 2, r - 2 + n).map((x) => x.slice(c - 1, c - 1 + m)) }) });
const ss = { getSheetByName: (n) => hoja({
  Personas: [['Ana', 'ana@x', true], ['Javi', 'javi@x', true], ['Lucía', '', true], ['Marta', 'marta@x', true]],
  Avisos: [['ana@x', 'https://fcm.googleapis.com/a', 'k', 'a', 0], ['ana@x', 'https://web.push.apple.com/b', 'k', 'a', 0],
    ['javi@x', 'https://fcm.googleapis.com/j', 'k', 'a', 0], ['marta@x', 'https://fcm.googleapis.com/m', 'k', 'a', 0]]
}[n]) };
const av = gs("destinatarios_(ss, { email: 'marta@x', name: 'Marta' }, 'ronda-r1', t)", { ss, t });
assert.deepEqual(av.map((a) => a.endpoint).sort(), ['https://fcm.googleapis.com/a', 'https://fcm.googleapis.com/j', 'https://web.push.apple.com/b']);
assert(av.every((a) => a.mensaje.tag === 'ronda-r1'));

console.log('avisos: todo bien');
