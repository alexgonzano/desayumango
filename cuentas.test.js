// Comprobación de las cuentas, sin dependencias: node cuentas.test.js
// Carga el script de index.html sin navegador (document de mentira) y prueba lo que expone en window.__cafes:
// saldos, pagos para quedar a cero y a quién le toca pagar. Todo en céntimos.
const assert = require('assert'), fs = require('fs'), path = require('path'), vm = require('vm');

const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
// Un objeto que acepta cualquier cosa: la app arranca sin pantalla y se queda esperando config.js.
const nada = new Proxy(function () {}, { get: (t, k) => (k === Symbol.toPrimitive ? () => '' : nada), apply: () => nada, set: () => true });
const window = { addEventListener() {} };
vm.runInContext(html.match(/<script>([\s\S]*?)<\/script>/)[1], vm.createContext({ window, document: nada, navigator: {}, localStorage: nada }));
const { balances, settle, suggestPayer } = window.__cafes;

const saldos = (data) => Object.fromEntries(balances(data));
const personas = ['Ana', 'Javi', 'Lucía', 'Marta'].map((name) => ({ name }));
const ronda = (date, payer, items) => ({ date, payer, items: items.map(([person, product, price]) => ({ person, product, price })) });
// Siempre: los saldos suman cero y, con los pagos que propone la app, todos quedan a cero.
function cuadra(data) {
  const b = balances(data);
  assert.equal([...b.values()].reduce((s, v) => s + v, 0), 0, 'los saldos suman cero');
  settle(b).forEach((t) => { assert(t.amount > 0); b.set(t.from, b.get(t.from) + t.amount); b.set(t.to, b.get(t.to) - t.amount); });
  assert([...b.values()].every((v) => v === 0), 'con los pagos propuestos todos quedan a cero');
}

// Paga Ana: su cortado, el café con leche de Javi y lo de Lucía (2 cortados y un Cola Cao)
const r1 = ronda('2026-11-02T09:15:00.000Z', 'Ana', [
  ['Ana', 'Cortado', 1.4], ['Javi', 'Café con leche', 1.6], ['Lucía', 'Cortado', 1.4], ['Lucía', 'Cortado', 1.4], ['Lucía', 'Cola Cao', 1.8]
]);
const d1 = { people: personas, rounds: [r1], payments: [] };
assert.deepEqual(saldos(d1), { Ana: 620, Javi: -160, 'Lucía': -460, Marta: 0 });
assert.deepEqual(settle(balances(d1)).map((t) => ({ ...t })), [{ from: 'Lucía', to: 'Ana', amount: 460 }, { from: 'Javi', to: 'Ana', amount: 160 }]);
assert.equal(suggestPayer(d1, ['Ana', 'Javi', 'Lucía', 'Marta']), 'Lucía', 'le toca a quien más debe');
cuadra(d1);

// Javi le devuelve 1,60 € a Ana: Javi queda en paz
const d2 = { ...d1, payments: [{ from: 'Javi', to: 'Ana', amount: 1.6 }] };
assert.deepEqual(saldos(d2), { Ana: 460, Javi: 0, 'Lucía': -460, Marta: 0 });
cuadra(d2);

// Lucía también paga: todos en paz, así que le toca a quien no ha pagado nunca (y si empatan, por nombre)
const d3 = { ...d1, payments: [{ from: 'Javi', to: 'Ana', amount: 1.6 }, { from: 'Lucía', to: 'Ana', amount: 4.6 }] };
assert.deepEqual(settle(balances(d3)), []);
assert.equal(suggestPayer(d3, ['Ana', 'Javi', 'Lucía', 'Marta']), 'Javi');
// Si todos han pagado alguna vez, le toca a quien hace más que no paga
const d4 = { people: personas, payments: [], rounds: [
  ronda('2026-11-02T09:00:00.000Z', 'Ana', [['Ana', 'Cortado', 1.4]]),
  ronda('2026-11-03T09:00:00.000Z', 'Javi', [['Javi', 'Cortado', 1.4]])
] };
assert.equal(suggestPayer(d4, ['Javi', 'Ana']), 'Ana');

// Céntimos exactos, aunque los precios lleguen con decimales de coma flotante
const d5 = { people: personas, payments: [], rounds: [ronda('2026-11-04T09:00:00.000Z', 'Marta', [['Javi', 'Té', 1.1], ['Javi', 'Té', 1.1], ['Javi', 'Té', 1.1], ['Lucía', 'Bollo', 1.15]])] };
assert.deepEqual(saldos(d5), { Ana: 0, Javi: -330, 'Lucía': -115, Marta: 445 });
cuadra(d5);

// Quien ya no está en la lista de personas sigue contando en las cuentas
const d6 = { people: personas, payments: [], rounds: [ronda('2026-11-04T09:00:00.000Z', 'Ana', [['Pablo', 'Capuchino', 2]])] };
assert.equal(saldos(d6).Pablo, -200);
cuadra(d6);

// Varios deben y varios cobran: primero la deuda más grande con el que más cobra
const varios = new Map([['A', 500], ['B', 300], ['C', -400], ['D', -400]]);
assert.deepEqual(settle(varios).map((t) => ({ ...t })), [
  { from: 'C', to: 'A', amount: 400 }, { from: 'D', to: 'A', amount: 100 }, { from: 'D', to: 'B', amount: 300 }
]);

console.log('cuentas: todo bien');
