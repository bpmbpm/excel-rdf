/*
 * workbook.js — преобразование модели Turtle в книгу Excel и обратно.
 *
 * Использует библиотеку SheetJS (глобальный объект XLSX), подключаемую в браузере через CDN.
 * Чистая логика (раскладка данных по листам и сборка обратно) вынесена в функции,
 * которые можно тестировать в node, передав совместимый объект XLSX (или работая со структурами AOA).
 *
 * Состав книги (концепция проекта, см. doc/):
 *   1. main                  — перечень листов книги и их назначение
 *   2. Префиксы              — префиксы, используемые в исходном Turtle
 *   3. Turtle исходный       — исходный текст Turtle RDF
 *   4. Триплеты простые      — по одному элементарному триплету на строку (без ; и ,)
 *   5. Триплеты компактные   — Turtle с ; и , (сокращённое число утверждений)
 *   6+. Лист на каждый МегаТип — субъекты данного МегаТипа, столбцы = предикаты
 *   N.  Прочие триплеты      — триплеты субъектов, не имеющих МегаТипа
 */
(function (global) {
  'use strict';

  var isNode = (typeof module !== 'undefined' && module.exports);
  var T = isNode ? require('./turtle.js') : global.TurtleRDF;
  var M = isNode ? require('./megatype.js') : global.MegatypeModel;

  // Зарезервированные имена служебных листов.
  var SHEET = {
    MAIN: 'main',
    PREFIXES: 'Префиксы',
    TURTLE: 'Turtle исходный',
    SIMPLE: 'Триплеты простые',
    COMPACT: 'Триплеты компактные',
    OTHER: 'Прочие триплеты'
  };
  var RESERVED = {};
  Object.keys(SHEET).forEach(function (k) { RESERVED[SHEET[k]] = true; });

  var SUBJECT_HEADER = 'Субъект';
  var MULTI_SEP = ' , '; // разделитель нескольких объектов в одной ячейке

  // Очистка имени листа Excel: запрещены : \\ / ? * [ ] и длина <= 31 символ.
  function sanitizeSheetName(name, used) {
    var clean = String(name).replace(/[:\\\/?*\[\]]/g, '_').slice(0, 31) || 'Лист';
    var base = clean;
    var i = 1;
    while (used[clean]) {
      var suffix = '_' + (i++);
      clean = base.slice(0, 31 - suffix.length) + suffix;
    }
    used[clean] = true;
    return clean;
  }

  /*
   * Построить набор листов в виде «массивов массивов» (AOA) из модели Turtle.
   * Возвращает { order: [имена], sheets: { имя: aoa }, meta: {...} }.
   * Эта функция не зависит от XLSX и легко тестируется.
   */
  function modelToSheets(model) {
    var prefixes = model.prefixes || {};
    var triples = model.triples || [];
    var megatypes = M.buildMegatypes(model);

    var order = [];
    var sheets = {};
    var used = {};

    function addSheet(name, aoa) {
      var safe = sanitizeSheetName(name, used);
      order.push(safe);
      sheets[safe] = aoa;
      return safe;
    }

    // 6+. Листы МегаТипов (сначала вычисляем имена, чтобы перечислить их в main).
    var megaSheetInfo = megatypes.map(function (mt) {
      // Заголовок: Субъект | :МегаТип | предикаты...
      var header = [SUBJECT_HEADER, mtPredicateTerm(triples)].concat(mt.predicates);
      var aoa = [header];
      mt.rows.forEach(function (row) {
        var line = [row.subject, mt.megatypeTerm];
        mt.predicates.forEach(function (p) {
          var objs = row.cells[p] || [];
          line.push(objs.join(MULTI_SEP));
        });
        aoa.push(line);
      });
      return { megatype: mt.megatype, megatypeTerm: mt.megatypeTerm, aoa: aoa, name: null };
    });

    // 1. main — перечень листов (строки добавим ниже, когда узнаем все имена).
    var mainAoa = [['Лист', 'Назначение']];
    addSheet(SHEET.MAIN, mainAoa);

    // 2. Префиксы.
    var prefAoa = [['Префикс', 'Пространство имён (IRI)']];
    Object.keys(prefixes).forEach(function (p) {
      prefAoa.push([p === '' ? '' : p, prefixes[p]]);
    });
    addSheet(SHEET.PREFIXES, prefAoa);

    // 3. Turtle исходный.
    var rawText = model.rawText != null ? model.rawText : T.serializeTurtle(model);
    var turtleAoa = rawText.split('\n').map(function (l) { return [l]; });
    turtleAoa.unshift(['Исходный Turtle RDF']);
    addSheet(SHEET.TURTLE, turtleAoa);

    // 4. Триплеты простые.
    var simpleAoa = [[SUBJECT_HEADER, 'Предикат', 'Объект']];
    triples.forEach(function (t) { simpleAoa.push([t.s, t.p, t.o]); });
    addSheet(SHEET.SIMPLE, simpleAoa);

    // 5. Триплеты компактные.
    var compactText = T.serializeTurtle(model);
    var compactAoa = compactText.split('\n').map(function (l) { return [l]; });
    compactAoa.unshift(['Turtle с ; и , (сокращённая форма)']);
    addSheet(SHEET.COMPACT, compactAoa);

    // 6+. Добавляем листы МегаТипов.
    megaSheetInfo.forEach(function (info) {
      info.name = addSheet(info.megatype, info.aoa);
    });

    // N. Прочие триплеты (субъекты без МегаТипа).
    var subjectsWithMega = {};
    triples.forEach(function (t) {
      if (M.isMegatypePredicate(t.p)) { subjectsWithMega[t.s] = true; }
    });
    var otherAoa = [[SUBJECT_HEADER, 'Предикат', 'Объект']];
    var hasOther = false;
    triples.forEach(function (t) {
      if (!subjectsWithMega[t.s]) { otherAoa.push([t.s, t.p, t.o]); hasOther = true; }
    });
    if (hasOther) { addSheet(SHEET.OTHER, otherAoa); }

    // Заполняем main описанием листов.
    var descriptions = {};
    descriptions[SHEET.MAIN] = 'Перечень листов книги и их назначение';
    descriptions[SHEET.PREFIXES] = 'Префиксы, используемые в исходном Turtle RDF';
    descriptions[SHEET.TURTLE] = 'Исходный текст Turtle RDF';
    descriptions[SHEET.SIMPLE] = 'Триплеты по одному на строку (без ; и ,)';
    descriptions[SHEET.COMPACT] = 'Turtle с ; и , (сокращённое число утверждений)';
    descriptions[SHEET.OTHER] = 'Триплеты субъектов без МегаТипа';
    megaSheetInfo.forEach(function (info) {
      descriptions[info.name] = 'МегаТип «' + info.megatype + '» (' + info.megatypeTerm + '): субъекты и их предикаты';
    });
    order.forEach(function (name) {
      if (name === SHEET.MAIN) { return; }
      mainAoa.push([name, descriptions[name] || '']);
    });

    return { order: order, sheets: sheets, meta: { megatypes: megaSheetInfo } };
  }

  // Определить терм предиката МегаТипа, как он записан в исходнике (по умолчанию «:МегаТип»).
  function mtPredicateTerm(triples) {
    for (var i = 0; i < triples.length; i++) {
      if (M.isMegatypePredicate(triples[i].p)) { return triples[i].p; }
    }
    return ':МегаТип';
  }

  // Построить книгу XLSX (SheetJS) из модели Turtle.
  function modelToWorkbook(model, XLSXlib) {
    var X = XLSXlib || global.XLSX;
    var built = modelToSheets(model);
    var wb = X.utils.book_new();
    built.order.forEach(function (name) {
      var ws = X.utils.aoa_to_sheet(built.sheets[name]);
      X.utils.book_append_sheet(wb, ws, name);
    });
    return wb;
  }

  /*
   * Прочитать книгу (в виде { order, sheets } из AOA) обратно в модель Turtle.
   * Источник истины для экспорта — листы МегаТипов и лист «Прочие триплеты».
   * Это позволяет пользователю редактировать удобную форму и выгружать корректный Turtle.
   */
  function sheetsToModel(book) {
    var sheets = book.sheets;
    var order = book.order || Object.keys(sheets);
    var prefixes = {};
    var triples = [];

    // Префиксы.
    var prefAoa = sheets[SHEET.PREFIXES];
    if (prefAoa) {
      for (var r = 1; r < prefAoa.length; r++) {
        var row = prefAoa[r] || [];
        var pfx = (row[0] == null ? '' : String(row[0])).trim();
        var iri = (row[1] == null ? '' : String(row[1])).trim();
        if (iri) { prefixes[pfx] = iri; }
      }
    }

    function pushCell(subject, predicate, cellVal) {
      if (cellVal == null) { return; }
      var s = String(cellVal).trim();
      if (s === '') { return; }
      // Несколько объектов в одной ячейке разделяются запятой (с пробелами).
      s.split(/\s*,\s*/).forEach(function (obj) {
        var v = obj.trim();
        if (v !== '') { triples.push({ s: subject, p: predicate, o: v }); }
      });
    }

    // Листы МегаТипов: любой лист с заголовком первой ячейки «Субъект»,
    // не входящий в число служебных листов с другой структурой.
    order.forEach(function (name) {
      if (name === SHEET.SIMPLE || name === SHEET.OTHER) { return; } // обработаем отдельно
      if (RESERVED[name]) { return; }
      var aoa = sheets[name];
      if (!aoa || !aoa.length) { return; }
      var header = aoa[0] || [];
      if (String(header[0]).trim() !== SUBJECT_HEADER) { return; }
      for (var i = 1; i < aoa.length; i++) {
        var row = aoa[i] || [];
        var subject = row[0] == null ? '' : String(row[0]).trim();
        if (subject === '') { continue; }
        for (var c = 1; c < header.length; c++) {
          var pred = header[c] == null ? '' : String(header[c]).trim();
          if (pred === '') { continue; }
          pushCell(subject, pred, row[c]);
        }
      }
    });

    // Прочие триплеты (Субъект | Предикат | Объект).
    var otherAoa = sheets[SHEET.OTHER];
    if (otherAoa) {
      for (var k = 1; k < otherAoa.length; k++) {
        var orow = otherAoa[k] || [];
        var os = orow[0] == null ? '' : String(orow[0]).trim();
        var op = orow[1] == null ? '' : String(orow[1]).trim();
        var oo = orow[2] == null ? '' : String(orow[2]).trim();
        if (os && op && oo) { triples.push({ s: os, p: op, o: oo }); }
      }
    }

    // Если листов МегаТипов не оказалось, используем лист «Триплеты простые».
    if (triples.length === 0 && sheets[SHEET.SIMPLE]) {
      var simpleAoa = sheets[SHEET.SIMPLE];
      for (var j = 1; j < simpleAoa.length; j++) {
        var srow = simpleAoa[j] || [];
        var ss = srow[0] == null ? '' : String(srow[0]).trim();
        var sp = srow[1] == null ? '' : String(srow[1]).trim();
        var so = srow[2] == null ? '' : String(srow[2]).trim();
        if (ss && sp && so) { triples.push({ s: ss, p: sp, o: so }); }
      }
    }

    return { prefixes: prefixes, triples: triples, base: null };
  }

  // Прочитать книгу XLSX (SheetJS) в структуру { order, sheets: AOA }.
  function workbookToBook(wb, XLSXlib) {
    var X = XLSXlib || global.XLSX;
    var order = wb.SheetNames.slice();
    var sheets = {};
    order.forEach(function (name) {
      sheets[name] = X.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, defval: '' });
    });
    return { order: order, sheets: sheets };
  }

  // Полное преобразование книги XLSX в модель Turtle.
  function workbookToModel(wb, XLSXlib) {
    return sheetsToModel(workbookToBook(wb, XLSXlib));
  }

  var api = {
    SHEET: SHEET,
    SUBJECT_HEADER: SUBJECT_HEADER,
    MULTI_SEP: MULTI_SEP,
    sanitizeSheetName: sanitizeSheetName,
    modelToSheets: modelToSheets,
    modelToWorkbook: modelToWorkbook,
    sheetsToModel: sheetsToModel,
    workbookToBook: workbookToBook,
    workbookToModel: workbookToModel
  };

  if (isNode) { module.exports = api; }
  global.WorkbookModel = api;
})(typeof window !== 'undefined' ? window : this);
