/*
 * test-ver3.js — автоматические проверки редакции ver3 (запуск в node только для тестирования).
 *
 * Сам продукт работает в браузере. Тесты используют чистые функции модулей и библиотеку
 * XLSX (SheetJS) только для проверки целостности готового .xlsx (закрепление строки,
 * автофильтр). Библиотека xlsx ставится локально и в репозиторий не входит (см. .gitignore).
 *
 * Запуск:  node ver3/tests/test-ver3.js
 *
 * Проверяется:
 *   • round-trip Turtle  → листы → Turtle   (сохранение триплетов и префиксов);
 *   • round-trip TriG    → листы МегаТипов → TriG (сохранение квадров и графов, требования 1–3);
 *   • листы МегаТипов TriG: столбцы «Субъект | TriG | предикаты…» (требование 2);
 *   • общий лист «TriG» со свойствами графов и реестр «МегаТипы» (требование 2);
 *   • текстовые листы простой и компактной форм TriG (требование 1);
 *   • второй (простой) пример TriG examples/trig-simple.trig (требование 4);
 *   • закрепление первой строки + автофильтр на каждом листе .xlsx;
 *   • разбор примера TriG VAD с именованными графами и МегаТипами (требование 3).
 */
'use strict';

var assert = require('assert');
var fs = require('fs');
var path = require('path');

var T = require('../js/turtle.js');
var M = require('../js/megatype.js');
var W = require('../js/workbook.js');
var TG = require('../js/trig.js');
var MT = require('../js/trig-megatype.js');
var TW = require('../js/trig-workbook.js');
var E = require('../js/xlsx-extras.js');

// XLSX нужен только для проверки бинарного .xlsx; если не установлен — эти тесты пропускаем.
var XLSX = null;
try { XLSX = require('xlsx'); } catch (e) { /* нет библиотеки — часть тестов будет пропущена */ }

var passed = 0, skipped = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  - ' + name); }
  catch (e) { console.error('  FAIL- ' + name + '\n        ' + e.message); process.exitCode = 1; }
}
function skip(name, reason) { skipped++; console.log('  skip- ' + name + ' (' + reason + ')'); }

// Сравнение множеств триплетов без учёта порядка.
function tripleSet(triples) {
  return new Set(triples.map(function (t) { return t.s + '|' + t.p + '|' + t.o; }));
}
function assertSameTriples(a, b, msg) {
  var sa = tripleSet(a), sb = tripleSet(b);
  var diff = [];
  sa.forEach(function (x) { if (!sb.has(x)) { diff.push('-' + x); } });
  sb.forEach(function (x) { if (!sa.has(x)) { diff.push('+' + x); } });
  assert.strictEqual(diff.length, 0, (msg || 'triples differ') + '\n        ' + diff.join('\n        '));
}
// Сравнение множеств квадров без учёта порядка.
function quadSet(quads) {
  return new Set(quads.map(function (q) { return (q.g == null ? '' : q.g) + '|' + q.s + '|' + q.p + '|' + q.o; }));
}
function assertSameQuads(a, b, msg) {
  var sa = quadSet(a), sb = quadSet(b);
  var diff = [];
  sa.forEach(function (x) { if (!sb.has(x)) { diff.push('-' + x); } });
  sb.forEach(function (x) { if (!sa.has(x)) { diff.push('+' + x); } });
  assert.strictEqual(diff.length, 0, (msg || 'quads differ') + '\n        ' + diff.join('\n        '));
}

console.log('Turtle round-trip (через листы):');

function ttlFiles() {
  var dir = path.join(__dirname, '..', 'examples');
  return fs.readdirSync(dir).filter(function (f) { return /\.ttl$/.test(f); })
    .map(function (f) { return path.join(dir, f); });
}

ttlFiles().forEach(function (file) {
  var name = path.basename(file);
  test('round-trip: ' + name, function () {
    var text = fs.readFileSync(file, 'utf8');
    var model = T.parseTurtle(text); model.rawText = text;
    var book = W.modelToSheets(model);
    var back = W.sheetsToModel(book);
    assertSameTriples(model.triples, back.triples, 'round-trip потерял триплеты: ' + name);
    Object.keys(model.prefixes).forEach(function (p) {
      assert.strictEqual(back.prefixes[p], model.prefixes[p], 'префикс ' + p + ' потерян');
    });
  });
});

console.log('Лист main: столбец «Предикаты» для листов МегаТипов (Turtle):');

test('main перечисляет предикаты листа МегаТипа', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'example4.ttl'), 'utf8');
  var model = T.parseTurtle(text); model.rawText = text;
  var book = W.modelToSheets(model);
  var main = book.sheets['main'];
  // Заголовок содержит третий столбец «Предикаты ...».
  assert.ok(/Предикаты/.test(String(main[0][2])), 'нет столбца Предикаты в заголовке main');
  // Находим строку листа «Процесс» и проверяем, что в столбце предикатов перечислены поля.
  var procRow = main.slice(1).filter(function (r) { return r[0] === 'Процесс'; })[0];
  assert.ok(procRow, 'нет строки листа Процесс на main');
  var preds = String(procRow[2]);
  ['rdfs:label', ':владелец', ':использует', ':приоритет'].forEach(function (p) {
    assert.ok(preds.indexOf(p) >= 0, 'предикат ' + p + ' не указан в столбце Предикаты, получено: ' + preds);
  });
  // Для служебного листа «Префиксы» столбец предикатов пуст.
  var prefRow = main.slice(1).filter(function (r) { return r[0] === 'Префиксы'; })[0];
  assert.strictEqual(String(prefRow[2] || ''), '', 'у служебного листа не должно быть предикатов');
});

console.log('TriG round-trip через листы МегаТипов (требования 1–3):');

// Помощник: список .trig примеров.
function trigFiles() {
  var dir = path.join(__dirname, '..', 'examples');
  return fs.readdirSync(dir).filter(function (f) { return /\.trig$/.test(f); })
    .map(function (f) { return path.join(dir, f); });
}

trigFiles().forEach(function (file) {
  var name = path.basename(file);
  test('round-trip TriG через листы: ' + name, function () {
    var text = fs.readFileSync(file, 'utf8');
    var model = TG.parseTrig(text); model.rawText = text;
    var book = TW.modelToSheets(model);
    var back = TW.sheetsToModel(book);
    assertSameQuads(model.quads, back.quads, 'round-trip TriG потерял квадры: ' + name);
    Object.keys(model.prefixes).forEach(function (p) {
      assert.strictEqual(back.prefixes[p], model.prefixes[p], 'префикс ' + p + ' потерян');
    });
  });
});

test('round-trip TriG через сериализацию компактную (parse → serialize → parse)', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text);
  var reparsed = TG.parseTrig(TG.serializeTrig(model));
  assertSameQuads(model.quads, reparsed.quads, 'компактная сериализация TriG потеряла квадры');
});

test('round-trip TriG через сериализацию простую (parse → serializeSimple → parse)', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text);
  var reparsed = TG.parseTrig(TG.serializeTrigSimple(model));
  assertSameQuads(model.quads, reparsed.quads, 'простая сериализация TriG потеряла квадры');
});

console.log('Листы анализа TriG строятся вокруг МегаТипов, а не графов (требование 2):');

test('trig-vad.trig: МегаТипы по rdf:type (ObjectTree, TypeProcess, TypeExecutor, VADProcessDia, ExecutorGroup)', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text);
  var mts = MT.buildTrigMegatypes(model);
  var locals = mts.map(function (m) { return m.megatype; });
  ['ObjectTree', 'TypeProcess', 'TypeExecutor', 'VADProcessDia', 'ExecutorGroup'].forEach(function (m) {
    assert.ok(locals.indexOf(m) >= 0, 'не найден лист МегаТипа ' + m + ', получено: ' + locals.join(', '));
  });
});

test('лист МегаТипа имеет столбцы «Субъект | TriG | предикаты…», строка = (субъект, граф)', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text); model.rawText = text;
  var book = TW.modelToSheets(model);
  // Лист МегаТипа ObjectTree существует и имеет нужные первые два столбца.
  var sheet = book.sheets['ObjectTree'];
  assert.ok(sheet, 'нет листа МегаТипа ObjectTree');
  assert.strictEqual(String(sheet[0][0]), TW.SUBJECT_HEADER, 'первый столбец листа МегаТипа должен быть «Субъект»');
  assert.strictEqual(String(sheet[0][1]), TW.TRIG_HEADER, 'второй столбец листа МегаТипа должен быть «TriG»');
  // В столбце «TriG» строки указан именованный граф (терм графа).
  var dataRow = sheet[1];
  assert.ok(dataRow && String(dataRow[1]).indexOf(':') >= 0, 'в столбце TriG должен быть указан граф, получено: ' + (dataRow && dataRow[1]));
});

test('общий лист «TriG» (сводка по графам) присутствует и перечисляет графы', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text); model.rawText = text;
  var book = TW.modelToSheets(model);
  var summary = book.sheets[TW.SHEET.SUMMARY];
  assert.ok(summary, 'нет общего листа «TriG»');
  assert.ok(/Граф/.test(String(summary[0][0])), 'заголовок общего листа TriG должен начинаться со столбца «Граф»');
  var graphsListed = summary.slice(1).map(function (r) { return String(r[0]); });
  ['vad:root', 'vad:ptree', 'vad:rtree'].forEach(function (g) {
    assert.ok(graphsListed.indexOf(g) >= 0, 'граф ' + g + ' не указан в общем листе TriG, получено: ' + graphsListed.join(', '));
  });
});

test('реестр «МегаТипы» перечисляет листы МегаТипов и их термы', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text); model.rawText = text;
  var book = TW.modelToSheets(model);
  var reg = book.sheets[TW.SHEET.MEGATYPES];
  assert.ok(reg, 'нет листа реестра «МегаТипы»');
  var names = reg.slice(1).map(function (r) { return String(r[0]); });
  ['ObjectTree', 'TypeProcess', 'TypeExecutor', 'VADProcessDia', 'ExecutorGroup'].forEach(function (m) {
    assert.ok(names.indexOf(m) >= 0, 'реестр МегаТипов не содержит лист ' + m);
  });
  // Терм rdf:type записан во втором столбце.
  var objRow = reg.slice(1).filter(function (r) { return r[0] === 'ObjectTree'; })[0];
  assert.strictEqual(String(objRow[1]), 'vad:ObjectTree', 'терм МегаТипа ObjectTree должен быть vad:ObjectTree');
});

test('текстовые листы простой и компактной форм TriG присутствуют (требование 1)', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text); model.rawText = text;
  var book = TW.modelToSheets(model);
  assert.ok(book.sheets[TW.SHEET.SIMPLE], 'нет листа «TriG с Триплеты простые»');
  assert.ok(book.sheets[TW.SHEET.COMPACT], 'нет листа «TriG с Триплеты компактные»');
  // Простая форма: текст листа должен заново разбираться в те же квадры.
  var simpleText = book.sheets[TW.SHEET.SIMPLE].slice(1).map(function (r) { return r[0]; }).join('\n');
  var simpleModel = TG.parseTrig(simpleText);
  assertSameQuads(model.quads, simpleModel.quads, 'текст простой формы не воспроизводит квадры');
  var compactText = book.sheets[TW.SHEET.COMPACT].slice(1).map(function (r) { return r[0]; }).join('\n');
  var compactModel = TG.parseTrig(compactText);
  assertSameQuads(model.quads, compactModel.quads, 'текст компактной формы не воспроизводит квадры');
});

test('rdf:type сохранён как обычный столбец (round-trip не выдумывает квадров)', function () {
  // Один субъект в двух графах с разными rdf:type — должны сохраниться оба, без дублей.
  var src = '@prefix rdf: <http://www.w3.org/1999/02/22-rdf-syntax-ns#> .\n' +
    '@prefix : <http://ex#> .\n' +
    ':g1 { :s rdf:type :A ; :p :x . }\n' +
    ':g2 { :s rdf:type :B ; :p :y . }\n';
  var model = TG.parseTrig(src);
  var book = TW.modelToSheets(model);
  var back = TW.sheetsToModel(book);
  assertSameQuads(model.quads, back.quads, 'round-trip потерял или выдумал квадры при двух графах одного субъекта');
});

console.log('Второй (простой) пример TriG — требование 4:');

test('examples/trig-simple.trig существует и round-trip через листы сохраняет квадры', function () {
  var file = path.join(__dirname, '..', 'examples', 'trig-simple.trig');
  assert.ok(fs.existsSync(file), 'нет файла второго примера trig-simple.trig');
  var text = fs.readFileSync(file, 'utf8');
  var model = TG.parseTrig(text); model.rawText = text;
  assert.ok(model.quads.length > 0, 'второй пример не содержит квадров');
  var book = TW.modelToSheets(model);
  var back = TW.sheetsToModel(book);
  assertSameQuads(model.quads, back.quads, 'round-trip второго примера потерял квадры');
  // МегаТипы простого примера: Product, Customer, Order.
  var mts = MT.buildTrigMegatypes(model).map(function (m) { return m.megatype; });
  ['Product', 'Customer', 'Order'].forEach(function (m) {
    assert.ok(mts.indexOf(m) >= 0, 'в простом примере нет МегаТипа ' + m + ', получено: ' + mts.join(', '));
  });
});

console.log('TriG: базовые свойства парсера:');

test('TriG: несколько объектов на предикат (:next :b , :c)', function () {
  var m = TG.parseTrig('@prefix : <http://ex#> .\n:g { :a :next :b , :c . }');
  assertSameQuads(m.quads, [
    { g: ':g', s: ':a', p: ':next', o: ':b' },
    { g: ':g', s: ':a', p: ':next', o: ':c' }
  ]);
});

test('TriG: граф по умолчанию (триплеты вне скобок)', function () {
  var m = TG.parseTrig('@prefix : <http://ex#> .\n:a :p :b .\n:g { :c :q :d . }');
  assertSameQuads(m.quads, [
    { g: null, s: ':a', p: ':p', o: ':b' },
    { g: ':g', s: ':c', p: ':q', o: ':d' }
  ]);
});

console.log('Бинарный .xlsx: закрепление строки и автофильтр:');

if (!XLSX) {
  skip('freeze + autofilter на всех листах', 'библиотека xlsx не установлена');
} else {
  test('freeze panes и autoFilter присутствуют в каждом листе (Turtle)', function () {
    var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'example4.ttl'), 'utf8');
    var model = T.parseTurtle(text); model.rawText = text;
    var bytes = W.modelToXlsxBytes(model, XLSX);
    var wb = XLSX.read(bytes, { type: 'array' });
    assert.ok(wb.SheetNames.length > 0, 'книга не открылась');
    var entries = E.parseZip(bytes).filter(function (e) { return /xl\/worksheets\/sheet\d+\.xml$/.test(e.name); });
    assert.ok(entries.length > 0, 'нет листов worksheet в архиве');
    var dec = new TextDecoder('utf-8');
    entries.forEach(function (e) {
      var xml = dec.decode(e.data);
      assert.ok(xml.indexOf('<pane ') >= 0, 'нет закрепления (<pane>) в ' + e.name);
      assert.ok(xml.indexOf('autoFilter') >= 0, 'нет автофильтра (<autoFilter>) в ' + e.name);
    });
  });

  test('TriG .xlsx: листы МегаТипов с закреплением и автофильтром', function () {
    var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
    var model = TG.parseTrig(text); model.rawText = text;
    var bytes = TW.modelToXlsxBytes(model, XLSX);
    var wb = XLSX.read(bytes, { type: 'array' });
    // В ver3 листы строятся по МегаТипам, а не по графам.
    assert.ok(wb.SheetNames.indexOf('ObjectTree') >= 0, 'нет листа МегаТипа ObjectTree');
    assert.ok(wb.SheetNames.indexOf(TW.SHEET.SUMMARY) >= 0, 'нет общего листа TriG');
    assert.ok(wb.SheetNames.indexOf(TW.SHEET.MEGATYPES) >= 0, 'нет листа реестра МегаТипов');
    var entries = E.parseZip(bytes).filter(function (e) { return /xl\/worksheets\/sheet\d+\.xml$/.test(e.name); });
    var dec = new TextDecoder('utf-8');
    entries.forEach(function (e) {
      var xml = dec.decode(e.data);
      assert.ok(xml.indexOf('<pane ') >= 0, 'нет закрепления в ' + e.name);
      assert.ok(xml.indexOf('autoFilter') >= 0, 'нет автофильтра в ' + e.name);
    });
  });

  test('round-trip через настоящий .xlsx (Turtle → xlsx → Turtle)', function () {
    var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'example2.ttl'), 'utf8');
    var model = T.parseTurtle(text); model.rawText = text;
    var bytes = W.modelToXlsxBytes(model, XLSX);
    var wb = XLSX.read(bytes, { type: 'array' });
    var back = W.workbookToModel(wb, XLSX);
    assertSameTriples(model.triples, back.triples, 'round-trip через .xlsx потерял триплеты');
  });

  test('round-trip через настоящий .xlsx (TriG → xlsx → TriG)', function () {
    var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
    var model = TG.parseTrig(text); model.rawText = text;
    var bytes = TW.modelToXlsxBytes(model, XLSX);
    var wb = XLSX.read(bytes, { type: 'array' });
    var back = TW.workbookToModel(wb, XLSX);
    assertSameQuads(model.quads, back.quads, 'round-trip TriG через .xlsx потерял квадры');
  });
}

console.log('\nПройдено проверок: ' + passed + (skipped ? (', пропущено: ' + skipped) : ''));
