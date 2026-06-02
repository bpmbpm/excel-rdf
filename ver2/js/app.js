/*
 * app.js — связывание интерфейса (index.html) с логикой конвертеров (редакция ver2).
 * Только браузерный JavaScript. Использует глобальные объекты:
 *   XLSX                — библиотека SheetJS (CDN)
 *   TurtleRDF           — js/turtle.js
 *   MegatypeModel       — js/megatype.js
 *   XlsxExtras          — js/xlsx-extras.js (закрепление строки + автофильтр)
 *   WorkbookModel       — js/workbook.js
 *   TrigRDF             — js/trig.js
 *   GraphModel          — js/graph.js
 *   TrigWorkbookModel   — js/trig-workbook.js
 */
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };

  // -------- Переключение вкладок --------
  var tabs = document.querySelectorAll('.tab');
  tabs.forEach(function (tab) {
    tab.addEventListener('click', function () {
      tabs.forEach(function (t) { t.classList.remove('active'); });
      document.querySelectorAll('.panel').forEach(function (p) { p.classList.remove('active'); });
      tab.classList.add('active');
      $(tab.getAttribute('data-tab')).classList.add('active');
    });
  });

  function setStatus(el, text, ok) {
    el.textContent = text;
    el.className = 'status ' + (ok === true ? 'ok' : ok === false ? 'err' : '');
  }

  // -------- Скачивание файла из Blob --------
  function downloadBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  // Сохранить готовый .xlsx (Uint8Array) с закреплением строки и автофильтром.
  function downloadXlsx(bytes, filename) {
    var blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    downloadBlob(blob, filename);
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  // Предпросмотр: показываем лист main (Лист | Назначение | Предикаты).
  function renderPreview(containerId, built) {
    var container = $(containerId);
    var main = built.sheets['main'] || [];
    var header = main[0] || ['Лист', 'Назначение', 'Предикаты'];
    var html = '<h3>Листы книги</h3><table><thead><tr>' +
      '<th>' + esc(header[0]) + '</th><th>' + esc(header[1]) + '</th><th>' + esc(header[2] || '') + '</th>' +
      '</tr></thead><tbody>';
    main.slice(1).forEach(function (row) {
      html += '<tr><td>' + esc(row[0]) + '</td><td>' + esc(row[1]) + '</td><td>' + esc(row[2] || '') + '</td></tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
  }

  // ================= Turtle → Excel =================
  var ttlFile = $('ttlFile');
  var ttlInput = $('ttlInput');

  ttlFile.addEventListener('change', function () {
    var file = ttlFile.files[0];
    if (!file) { return; }
    var reader = new FileReader();
    reader.onload = function () { ttlInput.value = reader.result; };
    reader.readAsText(file, 'utf-8');
  });

  $('loadExample').addEventListener('click', function () {
    fetch('examples/example4.ttl')
      .then(function (r) { if (!r.ok) { throw new Error('нет файла примера'); } return r.text(); })
      .then(function (t) { ttlInput.value = t; })
      .catch(function () {
        // Резервный пример, если запуск не на GitHub Pages (file://).
        ttlInput.value = '@prefix : <http://example.org/bpm#> .\n' +
          '@prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#> .\n\n' +
          ':Процесс_А :МегаТип :Процесс ;\n    rdfs:label "Согласование договора" ;\n' +
          '    :владелец :Иванов .\n\n' +
          ':Иванов :МегаТип :Роль ;\n    rdfs:label "Менеджер по договорам" .\n';
      });
  });

  $('convertT2E').addEventListener('click', function () {
    var status = $('t2eStatus');
    var text = ttlInput.value;
    if (!text || !text.trim()) { setStatus(status, 'Введите или загрузите Turtle.', false); return; }
    try {
      var model = TurtleRDF.parseTurtle(text);
      model.rawText = text;
      if (!model.triples.length) { setStatus(status, 'Не найдено ни одного триплета.', false); return; }
      var built = WorkbookModel.modelToSheets(model);
      var bytes = WorkbookModel.modelToXlsxBytes(model, XLSX);
      downloadXlsx(bytes, 'turtle.xlsx');
      setStatus(status, 'Готово: ' + built.order.length + ' листов, ' + model.triples.length + ' триплетов.', true);
      renderPreview('t2ePreview', built);
    } catch (e) {
      setStatus(status, 'Ошибка разбора: ' + e.message, false);
    }
  });

  // ================= Excel → Turtle =================
  var xlsxFile = $('xlsxFile');
  var ttlOutput = $('ttlOutput');

  $('convertE2T').addEventListener('click', function () {
    var status = $('e2tStatus');
    var file = xlsxFile.files[0];
    if (!file) { setStatus(status, 'Выберите файл .xlsx.', false); return; }
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var data = new Uint8Array(reader.result);
        var wb = XLSX.read(data, { type: 'array' });
        var model = WorkbookModel.workbookToModel(wb, XLSX);
        if (!model.triples.length) {
          setStatus(status, 'В книге не найдено триплетов (проверьте листы МегаТипов).', false);
          return;
        }
        ttlOutput.value = TurtleRDF.serializeTurtle(model);
        setStatus(status, 'Готово: ' + model.triples.length + ' триплетов.', true);
      } catch (e) {
        setStatus(status, 'Ошибка чтения книги: ' + e.message, false);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  $('downloadTtl').addEventListener('click', function () {
    if (!ttlOutput.value) { return; }
    downloadBlob(new Blob([ttlOutput.value], { type: 'text/turtle;charset=utf-8' }), 'export.ttl');
  });

  $('copyTtl').addEventListener('click', function () {
    if (!ttlOutput.value) { return; }
    ttlOutput.select();
    try { document.execCommand('copy'); } catch (e) { /* игнорируем */ }
  });

  // ================= TriG → Excel =================
  var trigFile = $('trigFile');
  var trigInput = $('trigInput');

  trigFile.addEventListener('change', function () {
    var file = trigFile.files[0];
    if (!file) { return; }
    var reader = new FileReader();
    reader.onload = function () { trigInput.value = reader.result; };
    reader.readAsText(file, 'utf-8');
  });

  $('loadTrigExample').addEventListener('click', function () {
    fetch('examples/trig-vad.trig')
      .then(function (r) { if (!r.ok) { throw new Error('нет файла примера'); } return r.text(); })
      .then(function (t) { trigInput.value = t; })
      .catch(function () {
        trigInput.value = '@prefix : <http://example.org/vad#> .\n' +
          '@prefix rdfs: <http://www.w3.org/2000/01/rdf-schema#> .\n\n' +
          ':root {\n    :root rdfs:label "Корень" .\n}\n\n' +
          ':ptree {\n    :p1 rdfs:label "Процесс 1" ;\n        :hasParentObj :root .\n}\n';
      });
  });

  $('convertG2E').addEventListener('click', function () {
    var status = $('g2eStatus');
    var text = trigInput.value;
    if (!text || !text.trim()) { setStatus(status, 'Введите или загрузите TriG.', false); return; }
    try {
      var model = TrigRDF.parseTrig(text);
      model.rawText = text;
      if (!model.quads.length) { setStatus(status, 'Не найдено ни одной четвёрки (квадра).', false); return; }
      var built = TrigWorkbookModel.modelToSheets(model);
      var bytes = TrigWorkbookModel.modelToXlsxBytes(model, XLSX);
      downloadXlsx(bytes, 'trig.xlsx');
      setStatus(status, 'Готово: ' + built.order.length + ' листов, ' + model.quads.length + ' квадров.', true);
      renderPreview('g2ePreview', built);
    } catch (e) {
      setStatus(status, 'Ошибка разбора: ' + e.message, false);
    }
  });

  // ================= Excel → TriG =================
  var xlsxTrigFile = $('xlsxTrigFile');
  var trigOutput = $('trigOutput');

  $('convertE2G').addEventListener('click', function () {
    var status = $('e2gStatus');
    var file = xlsxTrigFile.files[0];
    if (!file) { setStatus(status, 'Выберите файл .xlsx.', false); return; }
    var reader = new FileReader();
    reader.onload = function () {
      try {
        var data = new Uint8Array(reader.result);
        var wb = XLSX.read(data, { type: 'array' });
        var model = TrigWorkbookModel.workbookToModel(wb, XLSX);
        if (!model.quads.length) {
          setStatus(status, 'В книге не найдено квадров (проверьте листы графов и лист «Графы»).', false);
          return;
        }
        trigOutput.value = TrigRDF.serializeTrig(model);
        setStatus(status, 'Готово: ' + model.quads.length + ' квадров.', true);
      } catch (e) {
        setStatus(status, 'Ошибка чтения книги: ' + e.message, false);
      }
    };
    reader.readAsArrayBuffer(file);
  });

  $('downloadTrig').addEventListener('click', function () {
    if (!trigOutput.value) { return; }
    downloadBlob(new Blob([trigOutput.value], { type: 'application/trig;charset=utf-8' }), 'export.trig');
  });

  $('copyTrig').addEventListener('click', function () {
    if (!trigOutput.value) { return; }
    trigOutput.select();
    try { document.execCommand('copy'); } catch (e) { /* игнорируем */ }
  });
})();
