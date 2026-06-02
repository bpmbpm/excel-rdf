/*
 * trig.js — разбор (parse) и сборка (serialize) TriG в браузерном JavaScript.
 *
 * TriG — расширение Turtle, добавляющее именованные графы:
 *
 *     graphLabel {
 *         :s :p :o .
 *         ...
 *     }
 *
 * Триплеты вне фигурных скобок относятся к графу по умолчанию (default graph).
 * Модуль повторно использует токенизатор и утилиты из turtle.js и не зависит от node.js;
 * для тестов экспортируется и в window, и в module.exports.
 *
 * Основные функции:
 *   parseTrig(text)      -> { prefixes, base, quads }
 *   serializeTrig(model) -> строка TriG (компактная форма с ; и ,)
 *
 * Структура «квадра» (quad):
 *   { g, s, p, o }
 *   g — терм именованного графа (как записан, напр. «vad:ptree») либо null для графа
 *       по умолчанию; s, p, o — термы триплета (как в turtle.js).
 */
(function (global) {
  'use strict';

  var isNode = (typeof module !== 'undefined' && module.exports);
  var T = isNode ? require('./turtle.js') : global.TurtleRDF;

  function stripComment(line) {
    var inString = false, stringChar = '', inIri = false;
    for (var i = 0; i < line.length; i++) {
      var ch = line[i];
      if (inString) {
        if (ch === '\\') { i++; continue; }
        if (ch === stringChar) { inString = false; }
        continue;
      }
      if (inIri) { if (ch === '>') { inIri = false; } continue; }
      if (ch === '"' || ch === "'") { inString = true; stringChar = ch; continue; }
      if (ch === '<') { inIri = true; continue; }
      if (ch === '#') { return line.slice(0, i); }
    }
    return line;
  }

  function stripIri(v) { return v.replace(/^</, '').replace(/>$/, ''); }

  // Разбор TriG-текста в модель { prefixes, base, quads }.
  function parseTrig(text) {
    if (text == null) { text = ''; }
    var cleaned = text.split('\n').map(stripComment).join('\n');
    var tokens = T.tokenize(cleaned);

    var prefixes = {};
    var base = null;
    var quads = [];

    var idx = 0;
    function peek() { return tokens[idx]; }
    function next() { return tokens[idx++]; }
    // Ближайший следующий значимый токен (с учётом сдвига).
    function lookahead(offset) { return tokens[idx + (offset || 1)]; }

    var curGraph = null;       // текущий именованный граф (null = граф по умолчанию)
    var pendingGraph = null;   // кандидат на метку графа (терм перед «{»)
    var curSubject = null;
    var curPredicate = null;

    while (idx < tokens.length) {
      var tok = peek();

      // Директивы @prefix / @base.
      if (tok.type === 'directive') {
        next();
        var dname = tok.value.toLowerCase();
        if (dname === '@prefix') {
          var pfxTok = next();
          var iriTok = next();
          var pfx = pfxTok.value.replace(/:$/, '');
          prefixes[pfx] = stripIri(iriTok.value);
          if (peek() && peek().type === 'punct' && peek().value === '.') { next(); }
        } else if (dname === '@base') {
          var baseTok = next();
          base = stripIri(baseTok.value);
          if (peek() && peek().type === 'punct' && peek().value === '.') { next(); }
        }
        continue;
      }

      // Начало блока графа «{».
      if (tok.type === 'punct' && tok.value === '{') {
        next();
        curGraph = pendingGraph;   // метка графа (или null для анонимного блока графа по умолчанию)
        pendingGraph = null;
        curSubject = null;
        curPredicate = null;
        continue;
      }

      // Конец блока графа «}».
      if (tok.type === 'punct' && tok.value === '}') {
        next();
        curGraph = null;
        pendingGraph = null;
        curSubject = null;
        curPredicate = null;
        continue;
      }

      // Завершение утверждения.
      if (tok.type === 'punct' && tok.value === '.') {
        next();
        curSubject = null;
        curPredicate = null;
        continue;
      }

      // Следующий предикат того же субъекта.
      if (tok.type === 'punct' && tok.value === ';') { next(); curPredicate = null; continue; }

      // Следующий объект того же предиката.
      if (tok.type === 'punct' && tok.value === ',') { next(); continue; }

      // Прочие скобки (blank-node/коллекции) пропускаем.
      if (tok.type === 'punct') { next(); continue; }

      // Содержательный токен.
      if (curSubject === null) {
        // Необязательное ключевое слово GRAPH перед меткой графа.
        if (tok.type === 'term' && /^GRAPH$/i.test(tok.value)) { next(); continue; }
        // Если следующий значимый токен — «{», то текущий терм это метка графа.
        var nx = lookahead(1);
        if (nx && nx.type === 'punct' && nx.value === '{') {
          pendingGraph = tok.value;
          next();
          continue;
        }
        curSubject = next().value;
        continue;
      }
      if (curPredicate === null) {
        var pt = next();
        curPredicate = pt.type === 'a' ? 'a' : pt.value;
        continue;
      }
      // Объект.
      var ot = next();
      quads.push({ g: curGraph, s: curSubject, p: curPredicate, o: ot.value });
    }

    return { prefixes: prefixes, base: base, quads: quads };
  }

  // Сгруппировать список квадров по предикату/объекту для одного субъекта (компактная форма).
  function serializeSubjectBlock(triples, indent) {
    var preds = [], map = {};
    triples.forEach(function (t) {
      if (!map[t.p]) { map[t.p] = []; preds.push(t.p); }
      map[t.p].push(t.o);
    });
    var parts = preds.map(function (p) {
      return indent + '    ' + p + ' ' + map[p].join(' , ');
    });
    return parts.join(' ;\n');
  }

  // Сериализовать множество триплетов одного графа (компактная форма с ; и ,).
  function serializeGraphBody(triples, indent) {
    var subjects = [], bySubject = {};
    triples.forEach(function (t) {
      if (!bySubject[t.s]) { bySubject[t.s] = []; subjects.push(t.s); }
      bySubject[t.s].push(t);
    });
    var blocks = subjects.map(function (s) {
      var body = serializeSubjectBlock(bySubject[s], indent);
      // Убираем ведущий отступ у первой строки: «subject pred obj ...».
      var firstNl = body.indexOf('\n');
      var head = body.slice(0, firstNl < 0 ? body.length : firstNl).replace(/^\s+/, '');
      var tail = firstNl < 0 ? '' : body.slice(firstNl);
      return indent + s + ' ' + head + tail + ' .';
    });
    return blocks.join('\n');
  }

  /*
   * Собрать TriG-текст из модели { prefixes, base, quads }.
   * Квадры группируются по графу (с сохранением порядка появления), затем
   * по субъекту и предикату. Граф по умолчанию выводится без фигурных скобок.
   */
  function serializeTrig(model) {
    var prefixes = model.prefixes || {};
    var quads = model.quads || [];
    var lines = [];

    Object.keys(prefixes).forEach(function (p) {
      lines.push('@prefix ' + p + ': <' + prefixes[p] + '> .');
    });
    if (model.base) { lines.push('@base <' + model.base + '> .'); }
    if (lines.length) { lines.push(''); }

    // Группировка по графу с сохранением порядка.
    var graphOrder = [], byGraph = {};
    quads.forEach(function (q) {
      var key = q.g == null ? ' default' : q.g;
      if (!byGraph[key]) { byGraph[key] = { term: q.g, triples: [] }; graphOrder.push(key); }
      byGraph[key].triples.push({ s: q.s, p: q.p, o: q.o });
    });

    graphOrder.forEach(function (key) {
      var grp = byGraph[key];
      if (grp.term == null) {
        // Граф по умолчанию — без скобок.
        lines.push(serializeGraphBody(grp.triples, ''));
        lines.push('');
      } else {
        lines.push(grp.term + ' {');
        lines.push(serializeGraphBody(grp.triples, '    '));
        lines.push('}');
        lines.push('');
      }
    });

    return lines.join('\n').replace(/\n+$/, '\n');
  }

  var api = {
    parseTrig: parseTrig,
    serializeTrig: serializeTrig
  };

  if (isNode) { module.exports = api; }
  global.TrigRDF = api;
})(typeof window !== 'undefined' ? window : this);
