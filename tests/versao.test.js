'use strict';
// A versão mostrada em Ajustes vem do ?v= do index.html; ele e o cache do sw.js têm de subir juntos (10/11/2026).
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const ler = f => fs.readFileSync(path.join(__dirname, '..', f), 'utf8');

test('todos os ?v= do index.html são iguais e batem com o cache do sw.js', () => {
  const vs = [...ler('index.html').matchAll(/\?v=(\d+)/g)].map(m => m[1]);
  assert.equal(vs.length, 5, 'css + 4 scripts');
  assert.equal(new Set(vs).size, 1, 'versões diferentes no index.html: ' + vs.join(', '));
  const sw = ler('sw.js').match(/ponto-vic-v(\d+)/);
  assert.ok(sw, 'cache do sw.js sem número');
  assert.equal(sw[1], vs[0], 'sw.js (v' + sw[1] + ') diferente do index.html (v' + vs[0] + ')');
});

test('a tela Ajustes tem onde mostrar a versão', () => {
  assert.match(ler('index.html'), /<p id="versao"/);
  assert.match(ler('js/app.js'), /\$\('versao'\)\.textContent = 'Versão '/);
});
