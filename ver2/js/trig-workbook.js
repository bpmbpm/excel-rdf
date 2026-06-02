/*
 * trig-workbook.js — преобразование модели TriG в книгу Excel и обратно (редакция ver2).
 *
 * TriG отличается от Turtle наличием именованных графов. Поэтому раскладка по листам
 * строится не вокруг МегаТипов, а вокруг ГРАФОВ: один лист = один граф, строки = субъекты
 * графа, столбцы = предикаты (см. js/graph.js). Это даёт удобную таблицу на каждый граф.
 *
 * Состав книги TriG:
 *   1. main            — перечень листов книги, назначение и (для листов-графов) их предикаты
 *   2. Префиксы        — префиксы, используемые в исходном TriG
 *   3. TriG исходный   — исходный текст TriG
 *   4. Квадры простые  — по одной четвёрке (граф, субъект, предикат, объект) на строку
 *   5. Графы           — соответствие: лист ⇄ метка графа (используется при обратной сборке)
 *   6+. Лист на каждый граф — субъекты графа, столбцы = предикаты
 *
 * Источник истины при обратной сборке (Excel → TriG) — листы-графы, перечисленные на листе
 * «Графы». Это позволяет редактировать удобные таблицы и выгружать корректный TriG.
 *
 * Чистая логика (раскладка/сборка) не зависит от XLSX и тестируется в node.
 */
(function (global) {
  'use strict';

  var isNode = (typeof module !== 'undefined' && module.exports);
  var T = isNode ? require('./turtle.js') : global.TurtleRDF;
  var TG = isNode ? require('./trig.js') : global.TrigRDF;
  var G = isNode ? require('./graph.js') : global.GraphModel;
  var E = isNode ? require('./xlsx-extras.js') : global.XlsxExtras;

  var SHEET = {
    MAIN: 'main',
    PREFIXES: 'Префиксы',
    TRIG: 'TriG исходный',
    QUADS: 'Квадры простые',
    GRAPHS: 'Графы'
  };
  var RESERVED = {};
  Object.keys(SHEET).forEach(function (k) { RESERVED[SHEET[k]] = true; });

  var SUBJECT_HEADER = 'Субъект';
  var MULTI_SEP = ' , ';
  var PRED_SEP = ' , ';

  function sanitizeSheetName(name, used) {
    var clean = String(name).replace(/[:\\\/?*\[\]]/g, '_').slice(0, 31) || 'Граф';
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
   * Возвращает { order, sheets, meta: { graphs } }.
   */
  function modelToSheets(model) {
    var prefixes = model.prefixes || {};
    var quads = model.quads || [];
    var graphs = G.buildGraphs(model);

    var order = [];
    var sheets = {};
    var used = {};

    function addSheet(name, aoa) {
      var safe = sanitizeSheetName(name, used);
      order.push(safe);
      sheets[safe] = aoa;
      return safe;
    }

    // 6+. Листы графов (сначала вычисляем имена, чтобы перечислить их в main и Графы).
    var graphSheetInfo = graphs.map(function (gr) {
      var header = [SUBJECT_HEADER].concat(gr.predicates);
      var aoa = [header];
      gr.rows.forEach(function (row) {
        var line = [row.subject];
        gr.predicates.forEach(function (p) {
          var objs = row.cells[p] || [];
          line.push(objs.join(MULTI_SEP));
        });
        aoa.push(line);
      });
      return {
        graphTerm: gr.graphTerm,
        graphLocal: gr.graphLocal,
        predicates: gr.predicates,
        aoa: aoa,
        name: null
      };
    });

    // 1. main.
    var mainAoa = [['Лист', 'Назначение', 'Предикаты (для листов-графов)']];
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

    // 4. Квадры простые.
    var quadsAoa = [['Граф', SUBJECT_HEADER, 'Предикат', 'Объект']];
    quads.forEach(function (q) {
      quadsAoa.push([q.g == null ? '' : q.g, q.s, q.p, q.o]);
    });
    addSheet(SHEET.QUADS, quadsAoa);

    // 5. Графы — соответствие лист ⇄ метка графа (заполним именами ниже).
    var graphsAoa = [['Лист', 'Граф (метка)', 'Назначение']];
    addSheet(SHEET.GRAPHS, graphsAoa);

    // 6+. Листы графов.
    graphSheetInfo.forEach(function (info) {
      info.name = addSheet(info.graphLocal, info.aoa);
    });

    // Заполняем Графы.
    graphSheetInfo.forEach(function (info) {
      var label = info.graphTerm == null ? '' : info.graphTerm;
      var purpose = info.graphTerm == null ? 'Граф по умолчанию (без метки)' : 'Именованный граф ' + info.graphTerm;
      graphsAoa.push([info.name, label, purpose]);
    });

    // Заполняем main.
    var descriptions = {};
    descriptions[SHEET.MAIN] = 'Перечень листов книги и их назначение';
    descriptions[SHEET.PREFIXES] = 'Префиксы, используемые в исходном TriG';
    descriptions[SHEET.TRIG] = 'Исходный текст TriG';
    descriptions[SHEET.QUADS] = 'Четвёрки (граф, субъект, предикат, объект) по одной на строку';
    descriptions[SHEET.GRAPHS] = 'Соответствие листов и меток именованных графов';
    var predsByName = {};
    graphSheetInfo.forEach(function (info) {
      var who = info.graphTerm == null ? 'граф по умолчанию' : 'граф «' + info.graphTerm + '»';
      descriptions[info.name] = 'Именованный ' + who + ': субъекты и их предикаты';
      predsByName[info.name] = info.predicates.join(PRED_SEP);
    });
    order.forEach(function (name) {
      if (name === SHEET.MAIN) { return; }
      mainAoa.push([name, descriptions[name] || '', predsByName[name] || '']);
    });

    return { order: order, sheets: sheets, meta: { graphs: graphSheetInfo } };
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
   * Источник истины — листы-графы, перечисленные на листе «Графы».
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

    function readGraphSheet(sheetName, graphTerm) {
      var aoa = sheets[sheetName];
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
          // Запятые внутри литералов «"..."» и IRI «<...>» не считаются разделителями.
          T.splitObjects(row[c]).forEach(function (v) {
            quads.push({ g: graphTerm, s: subject, p: pred, o: v });
          });
        }
      }
    }

    // Перечень листов-графов из листа «Графы».
    var graphsAoa = sheets[SHEET.GRAPHS];
    var mapped = false;
    if (graphsAoa) {
      for (var k = 1; k < graphsAoa.length; k++) {
        var grow = graphsAoa[k] || [];
        var sheetName = grow[0] == null ? '' : String(grow[0]).trim();
        var label = grow[1] == null ? '' : String(grow[1]).trim();
        if (sheetName === '') { continue; }
        readGraphSheet(sheetName, label === '' ? null : label);
        mapped = true;
      }
    }

    // Фоллбэк: если карты графов нет, читаем «Квадры простые».
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
