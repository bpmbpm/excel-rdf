/*
 * app.js — связывание интерфейса (index.html) с логикой конвертеров.
 * Только браузерный JavaScript. Использует глобальные объекты:
 *   XLSX           — библиотека SheetJS (CDN)
 *   TurtleRDF      — js/turtle.js
 *   MegatypeModel  — js/megatype.js
 *   WorkbookModel  — js/workbook.js
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
    fetch('examples/example1.ttl')
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
      var wb = WorkbookModel.modelToWorkbook(model, XLSX);
      XLSX.writeFile(wb, 'turtle.xlsx');
      setStatus(status, 'Готово: ' + built.order.length + ' листов, ' + model.triples.length + ' триплетов.', true);
      renderPreview(built);
    } catch (e) {
      setStatus(status, 'Ошибка разбора: ' + e.message, false);
    }
  });

  // Предпросмотр: показываем список листов (лист main).
  function renderPreview(built) {
    var container = $('t2ePreview');
    var main = built.sheets['main'] || [];
    var html = '<h3>Листы книги</h3><table><thead><tr><th>Лист</th><th>Назначение</th></tr></thead><tbody>';
    main.slice(1).forEach(function (row) {
      html += '<tr><td>' + esc(row[0]) + '</td><td>' + esc(row[1]) + '</td></tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

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
})();
