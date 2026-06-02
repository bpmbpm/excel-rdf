/*
 * turtle.js — разбор (parse) и сборка (serialize) Turtle RDF в браузерном JavaScript.
 *
 * Модуль не использует node.js и предназначен для запуска в браузере (GitHub Pages).
 * Для удобства автоматического тестирования он также экспортируется через module.exports,
 * если переменная module доступна (например, в node при запуске тестов).
 *
 * Основные функции:
 *   parseTurtle(text)      -> { prefixes, triples, base }
 *   serializeTurtle(model) -> строка Turtle (компактная форма с ; и ,)
 *   localName(term)        -> локальное имя терма (часть после префикса)
 *   expandTerm(term, ...)  -> развёрнутый IRI терма
 *
 * Структура триплета:
 *   { s, p, o }  — где s, p, o это строки-термы в том виде, как они записаны в Turtle
 *                  (префиксное имя «:Процесс_А», полный IRI «<http://...>» или литерал «"текст"»).
 *   Предикат «a» нормализуется в «rdf:type» только при необходимости; по умолчанию сохраняется как есть.
 */
(function (global) {
  'use strict';

  // Удаляем комментарии «# ...», но не трогаем символ # внутри строк «"..."» и IRI «<...>».
  function stripComment(line) {
    var inString = false;
    var stringChar = '';
    var inIri = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i];
      if (inString) {
        if (ch === '\\') { i++; continue; }
        if (ch === stringChar) { inString = false; }
        continue;
      }
      if (inIri) {
        if (ch === '>') { inIri = false; }
        continue;
      }
      if (ch === '"' || ch === "'") { inString = true; stringChar = ch; continue; }
      if (ch === '<') { inIri = true; continue; }
      if (ch === '#') { return line.slice(0, i); }
    }
    return line;
  }

  /*
   * Токенизатор Turtle. Разбивает текст на термы и структурные символы (. ; ,).
   * Корректно обрабатывает:
   *   - IRI в угловых скобках <...>
   *   - префиксные имена pre:local и :local
   *   - литералы "..." и '...' с экранированием, языковыми тегами @lang и типами ^^тип
   *   - ключевое слово a (rdf:type)
   *   - директивы @prefix / @base (и формы PREFIX / BASE в верхнем регистре)
   */
  function tokenize(text) {
    var tokens = [];
    var i = 0;
    var n = text.length;

    function isWhitespace(ch) { return ch === ' ' || ch === '\t' || ch === '\r' || ch === '\n'; }

    while (i < n) {
      var ch = text[i];

      if (isWhitespace(ch)) { i++; continue; }

      // Структурные символы.
      if (ch === '.' || ch === ';' || ch === ',') {
        tokens.push({ type: 'punct', value: ch });
        i++;
        continue;
      }

      // Открытие/закрытие списков и blank-node (минимальная поддержка как отдельных токенов).
      if (ch === '[' || ch === ']' || ch === '(' || ch === ')') {
        tokens.push({ type: 'punct', value: ch });
        i++;
        continue;
      }

      // IRI <...>
      if (ch === '<') {
        var j = text.indexOf('>', i);
        if (j === -1) { throw new Error('Turtle: незакрытый IRI начиная с позиции ' + i); }
        tokens.push({ type: 'iri', value: text.slice(i, j + 1) });
        i = j + 1;
        continue;
      }

      // Директивы @prefix / @base
      if (ch === '@') {
        var m = /^@[A-Za-z]+/.exec(text.slice(i));
        if (m) {
          tokens.push({ type: 'directive', value: m[0] });
          i += m[0].length;
          continue;
        }
      }

      // Литерал-строка "..." или '...', включая многострочные """ ... """
      if (ch === '"' || ch === "'") {
        var triple = text.substr(i, 3);
        var literal;
        if (triple === '"""' || triple === "'''") {
          var endSeq = triple;
          var k = text.indexOf(endSeq, i + 3);
          if (k === -1) { throw new Error('Turtle: незакрытый многострочный литерал'); }
          literal = text.slice(i, k + 3);
          i = k + 3;
        } else {
          var quote = ch;
          var p = i + 1;
          while (p < n) {
            if (text[p] === '\\') { p += 2; continue; }
            if (text[p] === quote) { p++; break; }
            p++;
          }
          literal = text.slice(i, p);
          i = p;
        }
        // Языковой тег @lang или тип данных ^^...
        if (text[i] === '@') {
          var lm = /^@[A-Za-z0-9-]+/.exec(text.slice(i));
          if (lm) { literal += lm[0]; i += lm[0].length; }
        } else if (text.substr(i, 2) === '^^') {
          i += 2;
          if (text[i] === '<') {
            var gt = text.indexOf('>', i);
            literal += '^^' + text.slice(i, gt + 1);
            i = gt + 1;
          } else {
            var pm = /^[^\s.;,]+/.exec(text.slice(i));
            literal += '^^' + pm[0];
            i += pm[0].length;
          }
        }
        tokens.push({ type: 'literal', value: literal });
        continue;
      }

      // Прочее: префиксное имя, ключевое слово a, число, булево, PREFIX/BASE.
      var rest = text.slice(i);
      var wm = /^[^\s.;,()\[\]]+/.exec(rest);
      if (!wm) { i++; continue; }
      var word = wm[0];
      i += word.length;

      if (word === 'a') {
        tokens.push({ type: 'a', value: 'a' });
      } else if (/^(PREFIX|BASE)$/i.test(word)) {
        tokens.push({ type: 'directive', value: '@' + word.toLowerCase() });
      } else {
        tokens.push({ type: 'term', value: word });
      }
    }
    return tokens;
  }

  // Разбор Turtle-текста в модель { prefixes, triples, base }.
  function parseTurtle(text) {
    if (text == null) { text = ''; }
    // Предварительно убираем комментарии построчно (безопасно для строк и IRI).
    var cleaned = text.split('\n').map(stripComment).join('\n');
    var tokens = tokenize(cleaned);

    var prefixes = {};
    var base = null;
    var triples = [];

    var idx = 0;
    function peek() { return tokens[idx]; }
    function next() { return tokens[idx++]; }

    var curSubject = null;
    var curPredicate = null;

    while (idx < tokens.length) {
      var tok = peek();

      // Директивы @prefix / @base.
      if (tok.type === 'directive') {
        next();
        var dname = tok.value.toLowerCase();
        if (dname === '@prefix') {
          var pfxTok = next(); // что-то вида «ex:» или «:»
          var iriTok = next(); // <...>
          var pfx = pfxTok.value.replace(/:$/, '');
          prefixes[pfx] = stripIri(iriTok.value);
          // Поглощаем завершающую точку, если есть.
          if (peek() && peek().type === 'punct' && peek().value === '.') { next(); }
        } else if (dname === '@base') {
          var baseTok = next();
          base = stripIri(baseTok.value);
          if (peek() && peek().type === 'punct' && peek().value === '.') { next(); }
        }
        continue;
      }

      // Завершение утверждения.
      if (tok.type === 'punct' && tok.value === '.') {
        next();
        curSubject = null;
        curPredicate = null;
        continue;
      }

      // Переход к следующему предикату того же субъекта.
      if (tok.type === 'punct' && tok.value === ';') {
        next();
        curPredicate = null;
        continue;
      }

      // Переход к следующему объекту того же предиката.
      if (tok.type === 'punct' && tok.value === ',') {
        next();
        continue;
      }

      // Пропускаем скобки blank-node/коллекций (минимальная поддержка).
      if (tok.type === 'punct') { next(); continue; }

      // Содержательный токен: субъект -> предикат -> объект.
      if (curSubject === null) {
        curSubject = termValue(next());
        continue;
      }
      if (curPredicate === null) {
        var pt = next();
        curPredicate = pt.type === 'a' ? 'a' : termValue(pt);
        continue;
      }
      // Объект.
      var ot = next();
      triples.push({ s: curSubject, p: curPredicate, o: termValue(ot) });
    }

    return { prefixes: prefixes, triples: triples, base: base };
  }

  function termValue(tok) { return tok.value; }
  function stripIri(v) { return v.replace(/^</, '').replace(/>$/, ''); }

  // Локальное имя терма: часть после префикса/решётки/слэша.
  function localName(term) {
    if (term == null) { return ''; }
    if (term === 'a') { return 'type'; }
    if (term[0] === '"' || term[0] === "'") { return term; }
    if (term[0] === '<') {
      var iri = stripIri(term);
      var h = iri.lastIndexOf('#');
      var s = iri.lastIndexOf('/');
      var pos = Math.max(h, s);
      return pos >= 0 ? iri.slice(pos + 1) : iri;
    }
    var colon = term.indexOf(':');
    return colon >= 0 ? term.slice(colon + 1) : term;
  }

  // Префикс терма («ex» для «ex:foo», «» для «:foo»). Для IRI/литерала возвращает null.
  function prefixOf(term) {
    if (term == null || term === 'a') { return null; }
    if (term[0] === '<' || term[0] === '"' || term[0] === "'") { return null; }
    var colon = term.indexOf(':');
    return colon >= 0 ? term.slice(0, colon) : null;
  }

  // Развернуть терм в полный IRI, используя карту префиксов.
  function expandTerm(term, prefixes, base) {
    if (term == null) { return ''; }
    if (term === 'a') { return 'http://www.w3.org/1999/02/22-rdf-syntax-ns#type'; }
    if (term[0] === '"' || term[0] === "'") { return term; }
    if (term[0] === '<') {
      var iri = stripIri(term);
      if (base && !/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(iri)) { return base + iri; }
      return iri;
    }
    var colon = term.indexOf(':');
    if (colon >= 0) {
      var pfx = term.slice(0, colon);
      var local = term.slice(colon + 1);
      if (prefixes && Object.prototype.hasOwnProperty.call(prefixes, pfx)) {
        return prefixes[pfx] + local;
      }
    }
    return term;
  }

  /*
   * Собрать Turtle-текст из модели в компактной форме (с ; и ,).
   * Триплеты группируются по субъекту, затем по предикату.
   * Это «пятый лист» концепции — сокращённое число утверждений.
   */
  function serializeTurtle(model) {
    var prefixes = model.prefixes || {};
    var triples = model.triples || [];
    var lines = [];

    // Директивы префиксов.
    var pfxKeys = Object.keys(prefixes);
    pfxKeys.forEach(function (p) {
      lines.push('@prefix ' + p + ': <' + prefixes[p] + '> .');
    });
    if (model.base) { lines.push('@base <' + model.base + '> .'); }
    if (pfxKeys.length || model.base) { lines.push(''); }

    // Группировка: subject -> predicate -> [objects], с сохранением порядка появления.
    var subjects = [];
    var bySubject = {};
    triples.forEach(function (t) {
      if (!bySubject[t.s]) { bySubject[t.s] = { preds: [], map: {} }; subjects.push(t.s); }
      var entry = bySubject[t.s];
      if (!entry.map[t.p]) { entry.map[t.p] = []; entry.preds.push(t.p); }
      entry.map[t.p].push(t.o);
    });

    subjects.forEach(function (s) {
      var entry = bySubject[s];
      var predParts = entry.preds.map(function (p) {
        var objs = entry.map[p].join(' , ');
        return '    ' + p + ' ' + objs;
      });
      lines.push(s + ' ' + predParts.join(' ;\n').replace(/^ {4}/, '') + ' .');
      lines.push('');
    });

    return lines.join('\n').replace(/\n+$/, '\n');
  }

  /*
   * Собрать Turtle-текст в простой форме — по одному элементарному триплету на строку
   * (без ; и ,). Это «четвёртый лист» концепции.
   */
  function serializeTurtleSimple(model) {
    var prefixes = model.prefixes || {};
    var triples = model.triples || [];
    var lines = [];
    Object.keys(prefixes).forEach(function (p) {
      lines.push('@prefix ' + p + ': <' + prefixes[p] + '> .');
    });
    if (model.base) { lines.push('@base <' + model.base + '> .'); }
    if (lines.length) { lines.push(''); }
    triples.forEach(function (t) {
      lines.push(t.s + ' ' + t.p + ' ' + t.o + ' .');
    });
    return lines.join('\n') + '\n';
  }

  var api = {
    parseTurtle: parseTurtle,
    serializeTurtle: serializeTurtle,
    serializeTurtleSimple: serializeTurtleSimple,
    localName: localName,
    prefixOf: prefixOf,
    expandTerm: expandTerm,
    tokenize: tokenize
  };

  // Экспорт для браузера (window) и для node (module.exports — для тестов).
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  }
  global.TurtleRDF = api;
})(typeof window !== 'undefined' ? window : this);
