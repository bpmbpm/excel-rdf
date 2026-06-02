/*
 * test-ver2.js — автоматические проверки редакции ver2 (запуск в node только для тестирования).
 *
 * Сам продукт работает в браузере. Тесты используют чистые функции модулей и библиотеку
 * XLSX (SheetJS) только для проверки целостности готового .xlsx (закрепление строки,
 * автофильтр). Библиотека xlsx ставится локально и в репозиторий не входит (см. .gitignore).
 *
 * Запуск:  node ver2/tests/test-ver2.js
 *
 * Проверяется:
 *   • round-trip Turtle  → листы → Turtle   (сохранение триплетов и префиксов);
 *   • round-trip TriG    → листы → TriG      (сохранение квадров и графов);
 *   • столбец «Предикаты» на листе main для листов МегаТипов (требование 3);
 *   • закрепление первой строки + автофильтр на каждом листе .xlsx (требование 1);
 *   • разбор примера TriG VADv8 с именованными графами (требование 4).
 */
'use strict';

var assert = require('assert');
var fs = require('fs');
var path = require('path');

var T = require('../js/turtle.js');
var M = require('../js/megatype.js');
var W = require('../js/workbook.js');
var TG = require('../js/trig.js');
var G = require('../js/graph.js');
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

console.log('Лист main: столбец «Предикаты» для листов МегаТипов (требование 3):');

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

console.log('TriG round-trip (через листы):');

test('round-trip TriG: trig-vad.trig', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text); model.rawText = text;
  var book = TW.modelToSheets(model);
  var back = TW.sheetsToModel(book);
  assertSameQuads(model.quads, back.quads, 'round-trip TriG потерял квадры');
  Object.keys(model.prefixes).forEach(function (p) {
    assert.strictEqual(back.prefixes[p], model.prefixes[p], 'префикс ' + p + ' потерян');
  });
});

test('round-trip TriG через сериализацию (parse → serialize → parse)', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text);
  var reparsed = TG.parseTrig(TG.serializeTrig(model));
  assertSameQuads(model.quads, reparsed.quads, 'сериализация TriG потеряла квадры');
});

test('TriG: именованные графы распознаются', function () {
  var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
  var model = TG.parseTrig(text);
  var graphTerms = {};
  model.quads.forEach(function (q) { if (q.g != null) { graphTerms[q.g] = true; } });
  ['vad:root', 'vad:ptree', 'vad:rtree', 'vad:t_p1', 'vad:t_p1_1', 'vad:t_p2'].forEach(function (g) {
    assert.ok(graphTerms[g], 'не найден именованный граф ' + g);
  });
  // Группировка по графам даёт отдельную таблицу на каждый граф.
  var graphs = G.buildGraphs(model);
  assert.ok(graphs.length >= 6, 'ожидалось не менее 6 графов, получено ' + graphs.length);
});

test('TriG: несколько объектов на предикат (vad:hasNext vad:p2_2, vad:p1_1)', function () {
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

console.log('Бинарный .xlsx: закрепление строки и автофильтр (требование 1):');

if (!XLSX) {
  skip('freeze + autofilter на всех листах', 'библиотека xlsx не установлена');
} else {
  test('freeze panes и autoFilter присутствуют в каждом листе', function () {
    var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'example4.ttl'), 'utf8');
    var model = T.parseTurtle(text); model.rawText = text;
    var bytes = W.modelToXlsxBytes(model, XLSX);
    // Книгу можно открыть обратно — ZIP целостен.
    var wb = XLSX.read(bytes, { type: 'array' });
    assert.ok(wb.SheetNames.length > 0, 'книга не открылась');
    // В каждом листе worksheet есть <pane> и <autoFilter>.
    var entries = E.parseZip(bytes).filter(function (e) { return /xl\/worksheets\/sheet\d+\.xml$/.test(e.name); });
    assert.ok(entries.length > 0, 'нет листов worksheet в архиве');
    var dec = new TextDecoder('utf-8');
    entries.forEach(function (e) {
      var xml = dec.decode(e.data);
      assert.ok(xml.indexOf('<pane ') >= 0, 'нет закрепления (<pane>) в ' + e.name);
      assert.ok(xml.indexOf('autoFilter') >= 0, 'нет автофильтра (<autoFilter>) в ' + e.name);
    });
  });

  test('TriG .xlsx также с закреплением и автофильтром', function () {
    var text = fs.readFileSync(path.join(__dirname, '..', 'examples', 'trig-vad.trig'), 'utf8');
    var model = TG.parseTrig(text); model.rawText = text;
    var bytes = TW.modelToXlsxBytes(model, XLSX);
    var wb = XLSX.read(bytes, { type: 'array' });
    assert.ok(wb.SheetNames.indexOf('ptree') >= 0, 'нет листа графа ptree');
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
}

console.log('\nПройдено проверок: ' + passed + (skipped ? (', пропущено: ' + skipped) : ''));
