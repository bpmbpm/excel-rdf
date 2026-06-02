/*
 * trig-workbook.js — преобразование модели TriG в книгу Excel и обратно (редакция ver3).
 *
 * Отличие ver3 от ver2. В ver2 листы TriG строились вокруг ГРАФОВ (один лист = один граф).
 * В ver3 (см. issue #5) листы анализа утверждений строятся вокруг МегаТипов — так же, как в
 * книге Turtle. Для TriG роль предиката «МегаТип» играет rdf:type: значение rdf:type субъекта
 * задаёт лист, на который попадает субъект. Принадлежность утверждения к именованному графу
 * (TriG) фиксируется ОТДЕЛЬНЫМ столбцом «TriG» на листе МегаТипа.
 *
 * Состав книги TriG (ver3):
 *   1. main                       — перечень листов книги и их назначение
 *   2. Префиксы                   — префиксы исходного TriG
 *   3. TriG исходный              — исходный текст TriG
 *   4. TriG с Триплеты простые    — TriG, где внутри графов по одному триплету на строку (без ; и ,)
 *   5. TriG с Триплеты компактные — TriG в компактной форме (с ; и ,)
 *   6. Квадры простые             — по одной четвёрке (граф, субъект, предикат, объект) на строку
 *   7. TriG                       — общий лист: свойства каждого графа TriG (требование 2)
 *   8. МегаТипы                   — реестр листов МегаТипов (используется при обратной сборке)
 *   9+. Лист на каждый МегаТип     — столбцы: Субъект | TriG | <предикаты>; строка = (субъект, граф)
 *
 * Источник истины при обратной сборке (Excel → TriG) — листы МегаТипов, перечисленные на листе
 * «МегаТипы». Предикат rdf:type сохраняется как обычный столбец, поэтому round-trip обратим.
 *
 * Чистая логика (раскладка/сборка) не зависит от XLSX и тестируется в node.
 */
(function (global) {
  'use strict';

  var isNode = (typeof module !== 'undefined' && module.exports);
  var T = isNode ? require('./turtle.js') : global.TurtleRDF;
  var TG = isNode ? require('./trig.js') : global.TrigRDF;
  var MT = isNode ? require('./trig-megatype.js') : global.TrigMegatypeModel;
  var E = isNode ? require('./xlsx-extras.js') : global.XlsxExtras;

  var SHEET = {
    MAIN: 'main',
    PREFIXES: 'Префиксы',
    TRIG: 'TriG исходный',
    SIMPLE: 'TriG с Триплеты простые',
    COMPACT: 'TriG с Триплеты компактные',
    QUADS: 'Квадры простые',
    SUMMARY: 'TriG',
    MEGATYPES: 'МегаТипы'
  };
  var RESERVED = {};
  Object.keys(SHEET).forEach(function (k) { RESERVED[SHEET[k]] = true; });

  var SUBJECT_HEADER = 'Субъект';
  var TRIG_HEADER = 'TriG';
  var MULTI_SEP = ' , ';
  var PRED_SEP = ' , ';

  function sanitizeSheetName(name, used) {
    var clean = String(name).replace(/[:\\\/?*\[\]]/g, '_').slice(0, 31) || 'МегаТип';
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
   * Построить набор листов (AOA) из модели TriG { prefixes, base, quads }.
   * Возвращает { order, sheets, meta: { megatypes } }.
   */
  function modelToSheets(model) {
    var prefixes = model.prefixes || {};
    var quads = model.quads || [];
    var megatypes = MT.buildTrigMegatypes(model);
    var summary = MT.buildTrigSummary(model);

    var order = [];
    var sheets = {};
    var used = {};

    function addSheet(name, aoa) {
      var safe = sanitizeSheetName(name, used);
      order.push(safe);
      sheets[safe] = aoa;
      return safe;
    }

    // 9+. Листы МегаТипов (сначала вычисляем имена, чтобы перечислить их в main и МегаТипы).
    var mtSheetInfo = megatypes.map(function (mt) {
      var header = [SUBJECT_HEADER, TRIG_HEADER].concat(mt.predicates);
      var aoa = [header];
      mt.rows.forEach(function (row) {
        var line = [row.subject, row.graph == null ? '' : row.graph];
        mt.predicates.forEach(function (p) {
          var objs = row.cells[p] || [];
          line.push(objs.join(MULTI_SEP));
        });
        aoa.push(line);
      });
      return {
        megatype: mt.megatype,
        megatypeTerm: mt.megatypeTerm,
        predicates: mt.predicates,
        aoa: aoa,
        name: null
      };
    });

    // 1. main.
    var mainAoa = [['Лист', 'Назначение', 'Предикаты (для листов МегаТипов)']];
    addSheet(SHEET.MAIN, mainAoa);

    // 2. Префиксы.
    var prefAoa = [['Префикс', 'Пространство имён (IRI)']];
    Object.keys(prefixes).forEach(function (p) {
      prefAoa.push([p === '' ? '' : p, prefixes[p]]);
    });
    addSheet(SHEET.PREFIXES, prefAoa);

    // 3. TriG исходный.
    var rawText = model.rawText != null ? model.rawText : TG.serializeTrig(model);
    var trigAoa = rawText.split('\n').map(function (l) { return [l]; });
    trigAoa.unshift(['Исходный TriG']);
    addSheet(SHEET.TRIG, trigAoa);

    // 4. TriG с Триплеты простые (внутри графов — по одному триплету на строку).
    var simpleText = TG.serializeTrigSimple(model);
    var simpleAoa = simpleText.split('\n').map(function (l) { return [l]; });
    simpleAoa.unshift(['TriG: простые триплеты (по одному на строку, без ; и ,)']);
    addSheet(SHEET.SIMPLE, simpleAoa);

    // 5. TriG с Триплеты компактные (с ; и ,).
    var compactText = TG.serializeTrig(model);
    var compactAoa = compactText.split('\n').map(function (l) { return [l]; });
    compactAoa.unshift(['TriG: компактная форма (с ; и ,)']);
    addSheet(SHEET.COMPACT, compactAoa);

    // 6. Квадры простые (фоллбэк при обратной сборке).
    var quadsAoa = [['Граф (TriG)', SUBJECT_HEADER, 'Предикат', 'Объект']];
    quads.forEach(function (q) {
      quadsAoa.push([q.g == null ? '' : q.g, q.s, q.p, q.o]);
    });
    addSheet(SHEET.QUADS, quadsAoa);

    // 7. TriG — общий лист: свойства каждого графа.
    var summaryAoa = [['Граф (TriG)', 'Кол-во квадров', 'Субъекты', 'Свойства (предикаты)']];
    summary.forEach(function (s) {
      summaryAoa.push([
        s.graphTerm == null ? '(граф по умолчанию)' : s.graphTerm,
        s.quadCount,
        s.subjects.join(PRED_SEP),
        s.predicates.join(PRED_SEP)
      ]);
    });
    addSheet(SHEET.SUMMARY, summaryAoa);

    // 8. МегаТипы — реестр листов МегаТипов (заполним именами ниже).
    var megatypesAoa = [['Лист', 'МегаТип (терм rdf:type)', 'Назначение']];
    addSheet(SHEET.MEGATYPES, megatypesAoa);

    // 9+. Листы МегаТипов.
    mtSheetInfo.forEach(function (info) {
      info.name = addSheet(info.megatype, info.aoa);
    });

    // Заполняем реестр МегаТипов.
    mtSheetInfo.forEach(function (info) {
      var term = info.megatypeTerm == null ? '' : info.megatypeTerm;
      var purpose = info.megatypeTerm == null
        ? 'Субъекты без явного rdf:type'
        : 'МегаТип ' + info.megatypeTerm + ' (rdf:type)';
      megatypesAoa.push([info.name, term, purpose]);
    });

    // Заполняем main.
    var descriptions = {};
    descriptions[SHEET.MAIN] = 'Перечень листов книги и их назначение';
    descriptions[SHEET.PREFIXES] = 'Префиксы, используемые в исходном TriG';
    descriptions[SHEET.TRIG] = 'Исходный текст TriG';
    descriptions[SHEET.SIMPLE] = 'TriG: внутри графов по одному триплету на строку (без ; и ,)';
    descriptions[SHEET.COMPACT] = 'TriG: компактная форма (с ; и ,)';
    descriptions[SHEET.QUADS] = 'Четвёрки (граф, субъект, предикат, объект) по одной на строку';
    descriptions[SHEET.SUMMARY] = 'Свойства каждого графа TriG (субъекты и предикаты)';
    descriptions[SHEET.MEGATYPES] = 'Реестр листов МегаТипов (rdf:type субъекта)';
    var predsByName = {};
    mtSheetInfo.forEach(function (info) {
      var who = info.megatypeTerm == null
        ? 'Субъекты без явного rdf:type'
        : 'МегаТип «' + info.megatypeTerm + '»: субъекты, графы (TriG) и предикаты';
      descriptions[info.name] = who;
      predsByName[info.name] = info.predicates.join(PRED_SEP);
    });
    order.forEach(function (name) {
      if (name === SHEET.MAIN) { return; }
      mainAoa.push([name, descriptions[name] || '', predsByName[name] || '']);
    });

    return { order: order, sheets: sheets, meta: { megatypes: mtSheetInfo } };
  }

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

  // Собрать готовый .xlsx (Uint8Array) с закреплением первой строки и автофильтром.
  function modelToXlsxBytes(model, XLSXlib) {
    var X = XLSXlib || global.XLSX;
    var wb = modelToWorkbook(model, X);
    return E.writeWorkbookBytes(wb, X);
  }

  /*
   * Обратная сборка: листы (AOA) -> модель TriG { prefixes, base, quads }.
   * Источник истины — листы МегаТипов, перечисленные на листе «МегаТипы».
   */
  function sheetsToModel(book) {
    var sheets = book.sheets;
    var prefixes = {};
    var quads = [];

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

    function readMegatypeSheet(sheetName) {
      var aoa = sheets[sheetName];
      if (!aoa || !aoa.length) { return false; }
      var header = aoa[0] || [];
      if (String(header[0]).trim() !== SUBJECT_HEADER) { return false; }
      if (String(header[1]).trim() !== TRIG_HEADER) { return false; }
      for (var i = 1; i < aoa.length; i++) {
        var arow = aoa[i] || [];
        var subject = arow[0] == null ? '' : String(arow[0]).trim();
        if (subject === '') { continue; }
        var graphCell = arow[1] == null ? '' : String(arow[1]).trim();
        var graphTerm = graphCell === '' ? null : graphCell;
        for (var c = 2; c < header.length; c++) {
          var pred = header[c] == null ? '' : String(header[c]).trim();
          if (pred === '') { continue; }
          // Запятые внутри литералов «"..."» и IRI «<...>» не считаются разделителями.
          T.splitObjects(arow[c]).forEach(function (v) {
            quads.push({ g: graphTerm, s: subject, p: pred, o: v });
          });
        }
      }
      return true;
    }

    // Перечень листов МегаТипов из реестра «МегаТипы».
    var megatypesAoa = sheets[SHEET.MEGATYPES];
    var mapped = false;
    if (megatypesAoa) {
      for (var k = 1; k < megatypesAoa.length; k++) {
        var mrow = megatypesAoa[k] || [];
        var sheetName = mrow[0] == null ? '' : String(mrow[0]).trim();
        if (sheetName === '') { continue; }
        if (readMegatypeSheet(sheetName)) { mapped = true; }
      }
    }

    // Фоллбэк: если реестра нет — пытаемся прочитать любые листы с сигнатурой (Субъект, TriG).
    if (!mapped) {
      Object.keys(sheets).forEach(function (name) {
        if (RESERVED[name]) { return; }
        if (readMegatypeSheet(name)) { mapped = true; }
      });
    }

    // Фоллбэк второго уровня: «Квадры простые».
    if (!mapped || quads.length === 0) {
      var quadsAoa = sheets[SHEET.QUADS];
      if (quadsAoa) {
        for (var j = 1; j < quadsAoa.length; j++) {
          var qrow = quadsAoa[j] || [];
          var qg = qrow[0] == null ? '' : String(qrow[0]).trim();
          var qs = qrow[1] == null ? '' : String(qrow[1]).trim();
          var qp = qrow[2] == null ? '' : String(qrow[2]).trim();
          var qo = qrow[3] == null ? '' : String(qrow[3]).trim();
          if (qs && qp && qo) { quads.push({ g: qg === '' ? null : qg, s: qs, p: qp, o: qo }); }
        }
      }
    }

    return { prefixes: prefixes, base: null, quads: quads };
  }

  function workbookToBook(wb, XLSXlib) {
    var X = XLSXlib || global.XLSX;
    var order = wb.SheetNames.slice();
    var sheets = {};
    order.forEach(function (name) {
      sheets[name] = X.utils.sheet_to_json(wb.Sheets[name], { header: 1, blankrows: false, defval: '' });
    });
    return { order: order, sheets: sheets };
  }

  function workbookToModel(wb, XLSXlib) {
    return sheetsToModel(workbookToBook(wb, XLSXlib));
  }

  var api = {
    SHEET: SHEET,
    SUBJECT_HEADER: SUBJECT_HEADER,
    TRIG_HEADER: TRIG_HEADER,
    MULTI_SEP: MULTI_SEP,
    sanitizeSheetName: sanitizeSheetName,
    modelToSheets: modelToSheets,
    modelToWorkbook: modelToWorkbook,
    modelToXlsxBytes: modelToXlsxBytes,
    sheetsToModel: sheetsToModel,
    workbookToBook: workbookToBook,
    workbookToModel: workbookToModel
  };

  if (isNode) { module.exports = api; }
  global.TrigWorkbookModel = api;
})(typeof window !== 'undefined' ? window : this);
