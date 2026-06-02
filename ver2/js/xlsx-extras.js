/*
 * xlsx-extras.js — доработка книги Excel: закрепление первой строки (freeze panes)
 * и автофильтр на КАЖДОМ листе книги (требование 1 редакции ver2).
 *
 * Почему отдельный модуль. Библиотека SheetJS (community edition) умеет записывать
 * автофильтр (через ws['!autofilter']), но НЕ умеет записывать закрепление областей
 * (freeze panes) — в XML листа выводится пустой <sheetView/> без элемента <pane>.
 * Поэтому freeze-panes мы добавляем сами: после того как SheetJS собрал .xlsx
 * (ZIP-архив без сжатия, compression:false), мы разбираем архив, в каждый файл
 * xl/worksheets/sheetN.xml вставляем элемент <pane>, и пересобираем архив заново.
 *
 * Модуль работает и в браузере, и в node (для тестов): используются TextEncoder/
 * TextDecoder и собственная реализация CRC32 — внешних зависимостей нет.
 *
 * Экспортируемые функции:
 *   setAutofilters(wb)        — проставить ws['!autofilter'] на каждом листе книги
 *   injectFreezePanes(u8)     — вставить <pane> в каждый лист готового .xlsx (Uint8Array)
 *   writeWorkbookBytes(wb, X) — собрать .xlsx с автофильтром и закреплением (Uint8Array)
 */
(function (global) {
  'use strict';

  var isNode = (typeof module !== 'undefined' && module.exports);

  // --- CRC32 (для пересборки ZIP) ---
  var CRC_TABLE = (function () {
    var t = new Array(256);
    for (var n = 0; n < 256; n++) {
      var c = n;
      for (var k = 0; k < 8; k++) { c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); }
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(buf) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < buf.length; i++) { c = CRC_TABLE[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); }
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  function rd2(b, o) { return b[o] | (b[o + 1] << 8); }
  function rd4(b, o) { return (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0; }
  function strBytes(s) {
    var u = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) { u[i] = s.charCodeAt(i) & 0xFF; }
    return u;
  }

  // Разобрать STORED-ZIP (без сжатия) в список записей [{ name, data }].
  function parseZip(u8) {
    // Поиск End Of Central Directory (сигнатура 0x06054b50).
    var i = u8.length - 22;
    while (i >= 0 && rd4(u8, i) !== 0x06054b50) { i--; }
    if (i < 0) { return []; }
    var cdOff = rd4(u8, i + 16), cnt = rd2(u8, i + 10);
    var entries = [], p = cdOff;
    for (var e = 0; e < cnt; e++) {
      var nameLen = rd2(u8, p + 28), extraLen = rd2(u8, p + 30), commLen = rd2(u8, p + 32);
      var lho = rd4(u8, p + 42);
      var name = '';
      for (var j = 0; j < nameLen; j++) { name += String.fromCharCode(u8[p + 46 + j]); }
      var lnameLen = rd2(u8, lho + 26), lextraLen = rd2(u8, lho + 28);
      var compSize = rd4(u8, p + 20);
      var dataStart = lho + 30 + lnameLen + lextraLen;
      var data = u8.slice(dataStart, dataStart + compSize);
      entries.push({ name: name, data: data });
      p += 46 + nameLen + extraLen + commLen;
    }
    return entries;
  }

  // Пересобрать STORED-ZIP (без сжатия) из списка записей.
  function buildZip(entries) {
    var locals = [], central = [], offset = 0;
    function num2(n) { return [n & 0xFF, (n >> 8) & 0xFF]; }
    function num4(n) { return [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >>> 24) & 0xFF]; }
    for (var k = 0; k < entries.length; k++) {
      var en = entries[k];
      var nameB = strBytes(en.name);
      var crc = crc32(en.data);
      var sz = en.data.length;
      var lh = [].concat(
        [0x50, 0x4b, 0x03, 0x04], num2(20), num2(0), num2(0), num2(0), num2(0),
        num4(crc), num4(sz), num4(sz), num2(nameB.length), num2(0)
      );
      locals.push(new Uint8Array(lh)); locals.push(nameB); locals.push(en.data);
      var cd = [].concat(
        [0x50, 0x4b, 0x01, 0x02], num2(20), num2(20), num2(0), num2(0), num2(0), num2(0),
        num4(crc), num4(sz), num4(sz), num2(nameB.length), num2(0), num2(0), num2(0), num2(0),
        num4(0), num4(offset)
      );
      central.push(new Uint8Array(cd)); central.push(nameB);
      offset += lh.length + nameB.length + sz;
    }
    var cdStart = offset, cdLen = 0;
    for (var c = 0; c < central.length; c++) { cdLen += central[c].length; }
    var eocd = new Uint8Array([].concat(
      [0x50, 0x4b, 0x05, 0x06], num2(0), num2(0), num2(entries.length), num2(entries.length),
      num4(cdLen), num4(cdStart), num2(0)
    ));
    var parts = locals.concat(central, [eocd]);
    var total = 0;
    for (var t = 0; t < parts.length; t++) { total += parts[t].length; }
    var out = new Uint8Array(total), off = 0;
    for (var u = 0; u < parts.length; u++) { out.set(parts[u], off); off += parts[u].length; }
    return out;
  }

  // XML закрепления первой строки и выделения первой ячейки данных.
  var PANE_XML = '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>' +
                 '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>';

  /*
   * Вставить элемент <pane> (freeze panes) в каждый лист готового .xlsx.
   * Принимает Uint8Array (важно: X.write(...,{type:'array'}) возвращает ArrayBuffer,
   * поэтому оборачивайте результат в new Uint8Array(...) перед вызовом).
   * Возвращает Uint8Array пересобранного архива.
   */
  function injectFreezePanes(u8) {
    if (!(u8 instanceof Uint8Array)) { u8 = new Uint8Array(u8); }
    var entries = parseZip(u8);
    if (!entries.length) { return u8; }
    var dec = new TextDecoder('utf-8'), enc = new TextEncoder();
    for (var i = 0; i < entries.length; i++) {
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(entries[i].name)) {
        var xml = dec.decode(entries[i].data);
        if (xml.indexOf('<pane ') >= 0) { continue; } // уже закреплено
        // SheetJS пишет самозакрывающийся <sheetView .../> — раскрываем и вставляем <pane>.
        xml = xml.replace(/<sheetView( [^>]*?)\/>/, '<sheetView$1>' + PANE_XML + '</sheetView>');
        entries[i].data = enc.encode(xml);
      }
    }
    return buildZip(entries);
  }

  /*
   * Проставить автофильтр на первой строке КАЖДОГО листа книги SheetJS.
   * Ref берётся из полного диапазона листа (!ref), чтобы фильтр охватывал данные.
   */
  function setAutofilters(wb) {
    if (!wb || !wb.SheetNames) { return wb; }
    wb.SheetNames.forEach(function (name) {
      var ws = wb.Sheets[name];
      if (ws && ws['!ref']) { ws['!autofilter'] = { ref: ws['!ref'] }; }
    });
    return wb;
  }

  /*
   * Собрать .xlsx из книги SheetJS с автофильтром и закреплением первой строки.
   * Возвращает Uint8Array — готовый для сохранения через Blob.
   * X — объект библиотеки SheetJS (в браузере global.XLSX, в node — require('xlsx')).
   */
  function writeWorkbookBytes(wb, X) {
    X = X || global.XLSX;
    setAutofilters(wb);
    // compression:false обязательно — пересборка ZIP работает только со STORED-записями.
    var buf = X.write(wb, { type: 'array', bookType: 'xlsx', compression: false });
    return injectFreezePanes(new Uint8Array(buf));
  }

  var api = {
    crc32: crc32,
    parseZip: parseZip,
    buildZip: buildZip,
    injectFreezePanes: injectFreezePanes,
    setAutofilters: setAutofilters,
    writeWorkbookBytes: writeWorkbookBytes,
    PANE_XML: PANE_XML
  };

  if (isNode) { module.exports = api; }
  global.XlsxExtras = api;
})(typeof window !== 'undefined' ? window : this);
