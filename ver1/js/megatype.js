/*
 * megatype.js — построение модели «МегаТип» из разобранного Turtle.
 *
 * Концепция: объект (субъект) относится к некоторому МегаТипу через предикат «:МегаТип».
 * Например:  :Процесс_А :МегаТип :Процесс .
 * Тогда субъект :Процесс_А попадает на лист «Процесс», а столбцы этого листа —
 * это все предикаты, которые встречаются у субъектов данного МегаТипа.
 * На пересечении субъекта (строка) и предиката (столбец) стоит значение объекта.
 *
 * Предикат, обозначающий МегаТип, определяется по локальному имени «МегаТип»
 * (часть после префикса), поэтому работает и «:МегаТип», и «ex:МегаТип».
 *
 * Модуль не зависит от node.js; экспортируется и в window, и в module.exports (для тестов).
 */
(function (global) {
  'use strict';

  var T = (typeof module !== 'undefined' && module.exports) ? require('./turtle.js') : global.TurtleRDF;

  // Имя предиката, помечающего МегаТип (локальная часть).
  var MEGATYPE_LOCAL = 'МегаТип';

  function isMegatypePredicate(p) {
    return T.localName(p) === MEGATYPE_LOCAL;
  }

  /*
   * Построить модель листов МегаТипов.
   * Возвращает массив объектов:
   *   { megatype, megatypeTerm, predicates: [...], rows: [ { subject, cells: { pred: [objs] } } ] }
   * megatype       — локальное имя МегаТипа (для имени листа), напр. «Процесс»
   * megatypeTerm   — терм объекта МегаТипа, как записан, напр. «:Процесс»
   * predicates     — упорядоченный список предикатов-столбцов (без самого «:МегаТип»)
   * rows           — строки: субъект и значения по предикатам (список объектов на каждый предикат)
   */
  function buildMegatypes(model) {
    var triples = model.triples || [];

    // 1) Для каждого субъекта определяем его МегаТип(ы).
    var subjectMegatype = {}; // subject -> { term, local }
    triples.forEach(function (t) {
      if (isMegatypePredicate(t.p)) {
        subjectMegatype[t.s] = { term: t.o, local: T.localName(t.o) };
      }
    });

    // 2) Группируем субъекты по МегаТипу и собираем предикаты/значения.
    var groups = {}; // local -> group
    var order = [];

    triples.forEach(function (t) {
      var mt = subjectMegatype[t.s];
      if (!mt) { return; } // субъект без МегаТипа на отдельные листы не попадает
      var key = mt.local;
      if (!groups[key]) {
        groups[key] = {
          megatype: key,
          megatypeTerm: mt.term,
          predicates: [],
          predicateSet: {},
          rowsOrder: [],
          rowsMap: {}
        };
        order.push(key);
      }
      var g = groups[key];

      // Строка субъекта.
      if (!g.rowsMap[t.s]) {
        g.rowsMap[t.s] = { subject: t.s, cells: {} };
        g.rowsOrder.push(t.s);
      }
      var row = g.rowsMap[t.s];

      // Сам предикат МегаТипа не выносим в отдельный столбец (он определяет лист),
      // но прочие предикаты становятся столбцами.
      if (!isMegatypePredicate(t.p)) {
        if (!g.predicateSet[t.p]) { g.predicateSet[t.p] = true; g.predicates.push(t.p); }
        if (!row.cells[t.p]) { row.cells[t.p] = []; }
        row.cells[t.p].push(t.o);
      }
    });

    // 3) Формируем итоговый массив.
    return order.map(function (key) {
      var g = groups[key];
      return {
        megatype: g.megatype,
        megatypeTerm: g.megatypeTerm,
        predicates: g.predicates,
        rows: g.rowsOrder.map(function (s) { return g.rowsMap[s]; })
      };
    });
  }

  /*
   * Преобразовать лист МегаТипа (заголовок + строки таблицы) обратно в триплеты.
   * sheet: { megatype, megatypeTerm, predicates, rows }
   * Возвращает массив триплетов { s, p, o }, включая утверждение «:Субъект :МегаТип :Тип».
   */
  function megatypeSheetToTriples(sheet, megatypePredicate) {
    var mtPred = megatypePredicate || ':МегаТип';
    var triples = [];
    sheet.rows.forEach(function (row) {
      // Утверждение о принадлежности МегаТипу.
      triples.push({ s: row.subject, p: mtPred, o: sheet.megatypeTerm });
      sheet.predicates.forEach(function (p) {
        var objs = row.cells[p];
        if (!objs) { return; }
        objs.forEach(function (o) {
          if (o === '' || o == null) { return; }
          triples.push({ s: row.subject, p: p, o: o });
        });
      });
    });
    return triples;
  }

  var api = {
    MEGATYPE_LOCAL: MEGATYPE_LOCAL,
    isMegatypePredicate: isMegatypePredicate,
    buildMegatypes: buildMegatypes,
    megatypeSheetToTriples: megatypeSheetToTriples
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  }
  global.MegatypeModel = api;
})(typeof window !== 'undefined' ? window : this);
