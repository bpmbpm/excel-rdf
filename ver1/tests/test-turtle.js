/*
 * test-turtle.js — автоматические проверки конвертеров (запуск в node только для тестирования).
 *
 * Сам продукт работает в браузере. Тесты используют чистые функции модулей
 * (turtle.js, megatype.js, workbook.js) и НЕ требуют библиотеки XLSX:
 * проверяется раскладка по листам (AOA) и обратная сборка триплетов.
 *
 * Запуск:  node ver1/tests/test-turtle.js
 */
'use strict';

var assert = require('assert');
var fs = require('fs');
var path = require('path');

var T = require('../js/turtle.js');
var M = require('../js/megatype.js');
var W = require('../js/workbook.js');

var passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  - ' + name); }
  catch (e) { console.error('  FAIL- ' + name + '\n        ' + e.message); process.exitCode = 1; }
}

// Множество триплетов как набор строк «s|p|o» для сравнения без учёта порядка.
function tripleSet(triples) {
  return new Set(triples.map(function (t) { return t.s + '|' + t.p + '|' + t.o; }));
}
function assertSameTriples(a, b, msg) {
  var sa = tripleSet(a), sb = tripleSet(b);
  var missing = [];
  sa.forEach(function (x) { if (!sb.has(x)) { missing.push('-' + x); } });
  sb.forEach(function (x) { if (!sa.has(x)) { missing.push('+' + x); } });
  assert.strictEqual(missing.length, 0, (msg || 'triples differ') + '\n        ' + missing.join('\n        '));
}

console.log('Парсер Turtle:');

test('разбирает префиксы и базовые триплеты', function () {
  var m = T.parseTurtle('@prefix : <http://ex#> .\n:A :p :B .');
  assert.strictEqual(m.prefixes[''], 'http://ex#');
  assert.strictEqual(m.triples.length, 1);
  assert.deepStrictEqual(m.triples[0], { s: ':A', p: ':p', o: ':B' });
});

test('разбирает ; (несколько предикатов) и , (несколько объектов)', function () {
  var m = T.parseTurtle(':A :p :B , :C ; :q :D .');
  assert.strictEqual(m.triples.length, 3);
  assertSameTriples(m.triples, [
    { s: ':A', p: ':p', o: ':B' },
    { s: ':A', p: ':p', o: ':C' },
    { s: ':A', p: ':q', o: ':D' }
  ]);
});

test('обрабатывает литералы с пробелами и комментарии', function () {
  var m = T.parseTurtle(':A rdfs:label "Привет, мир" . # комментарий с : и ;\n:A :n 5 .');
  assert.strictEqual(m.triples.length, 2);
  assert.strictEqual(m.triples[0].o, '"Привет, мир"');
});

test('ключевое слово a сохраняется как предикат', function () {
  var m = T.parseTurtle(':A a :Class .');
  assert.strictEqual(m.triples[0].p, 'a');
  assert.strictEqual(M.isMegatypePredicate('a'), false);
});

test('localName и expandTerm', function () {
  assert.strictEqual(T.localName(':МегаТип'), 'МегаТип');
  assert.strictEqual(T.localName('ex:Foo'), 'Foo');
  assert.strictEqual(T.localName('<http://ex#Bar>'), 'Bar');
  assert.strictEqual(T.expandTerm(':A', { '': 'http://ex#' }), 'http://ex#A');
});

console.log('Модель МегаТипов:');

test('группирует субъекты по МегаТипу', function () {
  var m = T.parseTurtle('@prefix : <http://ex#> .\n' +
    ':A :МегаТип :Процесс ; :p :X .\n' +
    ':B :МегаТип :Процесс ; :q :Y .\n' +
    ':R :МегаТип :Роль ; :p :Z .');
  var mts = M.buildMegatypes(m);
  assert.strictEqual(mts.length, 2);
  var proc = mts.filter(function (x) { return x.megatype === 'Процесс'; })[0];
  assert.ok(proc);
  assert.strictEqual(proc.rows.length, 2);
  assert.ok(proc.predicates.indexOf(':p') >= 0);
  assert.ok(proc.predicates.indexOf(':q') >= 0);
  // предикат :МегаТип не должен попасть в столбцы
  assert.strictEqual(proc.predicates.indexOf(':МегаТип'), -1);
});

console.log('Книга Excel (раскладка и обратная сборка):');

function roundtripFiles() {
  var dir = path.join(__dirname, '..', 'examples');
  return fs.readdirSync(dir).filter(function (f) { return /\.ttl$/.test(f); })
    .map(function (f) { return path.join(dir, f); });
}

roundtripFiles().forEach(function (file) {
  var name = path.basename(file);
  test('round-trip через листы: ' + name, function () {
    var text = fs.readFileSync(file, 'utf8');
    var model = T.parseTurtle(text);
    model.rawText = text;
    var book = W.modelToSheets(model);     // модель -> листы (AOA)
    var back = W.sheetsToModel(book);      // листы -> модель
    assertSameTriples(model.triples, back.triples, 'round-trip потерял триплеты для ' + name);
    // префиксы сохранены
    Object.keys(model.prefixes).forEach(function (p) {
      assert.strictEqual(back.prefixes[p], model.prefixes[p], 'префикс ' + p + ' потерян');
    });
  });
});

test('книга содержит обязательные служебные листы', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'example1.ttl'), 'utf8');
  var model = T.parseTurtle(text); model.rawText = text;
  var book = W.modelToSheets(model);
  ['main', 'Префиксы', 'Turtle исходный', 'Триплеты простые', 'Триплеты компактные']
    .forEach(function (s) { assert.ok(book.sheets[s], 'нет листа ' + s); });
  // лист main перечисляет остальные листы
  var mainRows = book.sheets['main'].slice(1).map(function (r) { return r[0]; });
  assert.ok(mainRows.indexOf('Префиксы') >= 0);
  // есть лист МегаТипа «Процесс» и «Роль»
  assert.ok(book.sheets['Процесс'], 'нет листа Процесс');
  assert.ok(book.sheets['Роль'], 'нет листа Роль');
});

test('лист «Прочие триплеты» для субъектов без МегаТипа (example3)', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'example3.ttl'), 'utf8');
  var model = T.parseTurtle(text); model.rawText = text;
  var book = W.modelToSheets(model);
  assert.ok(book.sheets['Прочие триплеты'], 'ожидался лист Прочие триплеты');
});

test('сериализация компактная и простая разбираются обратно одинаково', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'example2.ttl'), 'utf8');
  var model = T.parseTurtle(text);
  var compact = T.parseTurtle(T.serializeTurtle(model));
  var simple = T.parseTurtle(T.serializeTurtleSimple(model));
  assertSameTriples(model.triples, compact.triples, 'компактная форма');
  assertSameTriples(model.triples, simple.triples, 'простая форма');
});

console.log('\nПройдено проверок: ' + passed);
