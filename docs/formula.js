/*
 * FxTrace – Excel formula tokenizer & parser
 * Works on formulas as returned by Office.js (Range.formulas), which are
 * always en-US syntax: comma argument separators, dot decimals, English
 * function names.
 */
(function (root) {
  "use strict";

  // ---------- Tokenizer ----------
  var NAMECH = "A-Za-z0-9_.\\\\?\\u00A0-\\uFFFF";
  var RE = {
    ws: /\s+/y,
    str: /"(?:[^"]|"")*"/y,
    err: /#(?:NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!|PYTHON!)/iy,
    // optional external book, then quoted or plain sheet (optionally 3D Sheet1:Sheet3), then "!"
    sheet: new RegExp(
      "(?:'((?:\\[[^\\]]+\\])?(?:[^']|'')+)'|((?:\\[[^\\]]+\\])?[" + NAMECH + "]+(?::[" + NAMECH + "]+)?))!",
      "y"
    ),
    area: new RegExp(
      "\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7}(?::\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7})?(?![" + NAMECH + "(\\[!])",
      "y"
    ),
    colRange: new RegExp("\\$?[A-Za-z]{1,3}:\\$?[A-Za-z]{1,3}(?![" + NAMECH + "(\\[!])", "y"),
    rowRange: /\$?[0-9]{1,7}:\$?[0-9]{1,7}(?![0-9.])/y,
    num: /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y,
    bool: new RegExp("(?:TRUE|FALSE)(?![" + NAMECH + "(\\[])", "iy"),
    name: new RegExp("[A-Za-z_\\\\\\u00A0-\\uFFFF][" + NAMECH + "]*", "y"),
    op: /<=|>=|<>|[-+*\/^&=<>%:@]/y,
  };

  function match(re, s, i) {
    re.lastIndex = i;
    var m = re.exec(s);
    return m;
  }

  function scanBrackets(s, i) {
    // s[i] === "[" ; return index after matching "]"
    var depth = 0;
    for (var j = i; j < s.length; j++) {
      var ch = s[j];
      if (ch === "'" && j + 1 < s.length) { j++; continue; } // escape char in structured refs
      if (ch === "[") depth++;
      else if (ch === "]") { depth--; if (depth === 0) return j + 1; }
    }
    throw new Error("Unbalanced [ in structured reference");
  }

  function scanArray(s, i) {
    var inStr = false;
    for (var j = i + 1; j < s.length; j++) {
      var ch = s[j];
      if (ch === '"') inStr = !inStr;
      else if (!inStr && ch === "}") return j + 1;
    }
    throw new Error("Unbalanced { in array constant");
  }

  function isValidCol(letters) {
    letters = letters.toUpperCase();
    var n = 0;
    for (var k = 0; k < letters.length; k++) n = n * 26 + (letters.charCodeAt(k) - 64);
    return n >= 1 && n <= 16384;
  }

  function tokenize(formula) {
    var s = formula;
    var i = s[0] === "=" ? 1 : 0;
    var toks = [];
    var m;
    function push(type, start, end, extra) {
      var t = { type: type, text: s.slice(start, end), start: start, end: end };
      if (extra) for (var k in extra) t[k] = extra[k];
      toks.push(t);
    }
    while (i < s.length) {
      var ch = s[i];
      if ((m = match(RE.ws, s, i))) { i += m[0].length; continue; }
      if (ch === '"') { m = match(RE.str, s, i); if (!m) throw new Error("Unterminated string"); push("str", i, i + m[0].length); i += m[0].length; continue; }
      if (ch === "{") { var e = scanArray(s, i); push("array", i, e); i = e; continue; }
      if (ch === "(") { push("lparen", i, i + 1); i++; continue; }
      if (ch === ")") { push("rparen", i, i + 1); i++; continue; }
      if (ch === ",") { push("comma", i, i + 1); i++; continue; }
      if (ch === ";") { push("comma", i, i + 1); i++; continue; }
      if (ch === "#") {
        if ((m = match(RE.err, s, i))) { push("err", i, i + m[0].length); i += m[0].length; continue; }
        // spill operator after a reference
        var prev = toks[toks.length - 1];
        if (prev && (prev.type === "ref" || prev.type === "name")) {
          prev.spill = true; prev.end = i + 1; prev.text = s.slice(prev.start, i + 1); i++; continue;
        }
        throw new Error("Unexpected # at " + i);
      }

      // Reference with sheet prefix
      var start = i, sheet = null, book = null;
      if ((m = match(RE.sheet, s, i))) {
        var raw = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
        var bm = /^\[([^\]]+)\](.*)$/.exec(raw);
        if (bm) { book = bm[1]; raw = bm[2]; }
        var j = i + m[0].length;
        var m2;
        if ((m2 = match(RE.area, s, j)) && checkArea(m2[0])) {
          push("ref", start, j + m2[0].length, { sheet: raw, book: book, address: m2[0] }); i = j + m2[0].length; continue;
        }
        if ((m2 = match(RE.colRange, s, j)) || (m2 = match(RE.rowRange, s, j))) {
          push("ref", start, j + m2[0].length, { sheet: raw, book: book, address: m2[0] }); i = j + m2[0].length; continue;
        }
        if ((m2 = match(RE.err, s, j))) {
          push("err", start, j + m2[0].length); i = j + m2[0].length; continue;
        }
        if ((m2 = match(RE.name, s, j))) {
          push("name", start, j + m2[0].length, { sheet: raw, book: book, name: m2[0] }); i = j + m2[0].length; continue;
        }
        // not a reference after all – fall through
      }

      if ((m = match(RE.rowRange, s, i))) { push("ref", i, i + m[0].length, { sheet: null, address: m[0] }); i += m[0].length; continue; }
      if (/[0-9.]/.test(ch) && (m = match(RE.num, s, i))) { push("num", i, i + m[0].length); i += m[0].length; continue; }
      if ((m = match(RE.area, s, i)) && checkArea(m[0])) { push("ref", i, i + m[0].length, { sheet: null, address: m[0] }); i += m[0].length; continue; }
      if ((m = match(RE.colRange, s, i)) && checkCols(m[0])) { push("ref", i, i + m[0].length, { sheet: null, address: m[0] }); i += m[0].length; continue; }
      if ((m = match(RE.bool, s, i))) { push("bool", i, i + m[0].length); i += m[0].length; continue; }
      if (ch === "[") { var e2 = scanBrackets(s, i); push("struct", i, e2, { table: null }); i = e2; continue; }
      if ((m = match(RE.name, s, i))) {
        var end = i + m[0].length;
        if (s[end] === "(") { push("func", i, end, { name: m[0] }); i = end; continue; }
        if (s[end] === "[") { var e3 = scanBrackets(s, end); push("struct", i, e3, { table: m[0] }); i = e3; continue; }
        push("name", i, end, { sheet: null, name: m[0] }); i = end; continue;
      }
      if ((m = match(RE.op, s, i))) { push("op", i, i + m[0].length); i += m[0].length; continue; }
      throw new Error("Unexpected character '" + ch + "' at position " + i);
    }
    return toks;
  }

  function checkArea(a) {
    var parts = a.replace(/\$/g, "").split(":");
    for (var k = 0; k < parts.length; k++) {
      var mm = /^([A-Za-z]+)([0-9]+)$/.exec(parts[k]);
      if (!mm || !isValidCol(mm[1]) || +mm[2] < 1 || +mm[2] > 1048576) return false;
    }
    return true;
  }
  function checkCols(a) {
    var parts = a.replace(/\$/g, "").split(":");
    return isValidCol(parts[0]) && isValidCol(parts[1]);
  }

  // ---------- Parser (Pratt) ----------
  var BIN = {
    ":": 8,
    "^": 5, "*": 4, "/": 4, "+": 3, "-": 3, "&": 2,
    "=": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "<>": 1,
  };
  var UNARY_BP = 7, PERCENT_BP = 6;

  function parse(formula) {
    var toks = tokenize(formula);
    var p = 0;
    function peek() { return toks[p]; }
    function next() { return toks[p++]; }
    function expect(type) {
      var t = next();
      if (!t || t.type !== type) throw new Error("Expected " + type + (t ? " but found '" + t.text + "'" : " but reached end"));
      return t;
    }

    function parseExpr(minBp) {
      var t = next();
      if (!t) throw new Error("Unexpected end of formula");
      var left;
      switch (t.type) {
        case "num": case "str": case "bool": case "err": case "array":
          left = { type: t.type, start: t.start, end: t.end, text: t.text }; break;
        case "ref":
          left = { type: "ref", start: t.start, end: t.end, text: t.text, sheet: t.sheet, book: t.book || null, address: t.address, spill: !!t.spill }; break;
        case "name":
          left = { type: "name", start: t.start, end: t.end, text: t.text, sheet: t.sheet, book: t.book || null, name: t.name, spill: !!t.spill }; break;
        case "struct":
          left = { type: "struct", start: t.start, end: t.end, text: t.text, table: t.table }; break;
        case "func":
          left = parseFunc(t); break;
        case "lparen": {
          var items = [parseExpr(0)];
          while (peek() && peek().type === "comma") { next(); items.push(parseExpr(0)); }
          var r = expect("rparen");
          left = items.length === 1
            ? { type: "paren", start: t.start, end: r.end, children: items }
            : { type: "union", start: t.start, end: r.end, children: items };
          break;
        }
        case "op":
          if (t.text === "-" || t.text === "+" || t.text === "@") {
            var operand = parseExpr(UNARY_BP);
            left = { type: "unop", op: t.text, start: t.start, end: operand.end, children: [operand] };
            break;
          }
        // fallthrough
        default:
          throw new Error("Unexpected '" + t.text + "'");
      }

      for (;;) {
        var o = peek();
        // immediate invocation, e.g. LAMBDA(x,x*2)(A1)
        if (o && o.type === "lparen" && (left.type === "func" || left.type === "call")) {
          left = parseFunc({ name: "(call)", start: left.start }, left);
          continue;
        }
        if (!o || o.type !== "op") break;
        if (o.text === "%") {
          if (PERCENT_BP < minBp) break;
          next();
          left = { type: "postfix", op: "%", start: left.start, end: o.end, children: [left] };
          continue;
        }
        var bp = BIN[o.text];
        if (bp === undefined || bp <= minBp) break;
        next();
        // ^ is left-associative in Excel
        var right = parseExpr(bp);
        left = { type: "binop", op: o.text, start: left.start, end: right.end, children: [left, right] };
      }
      return left;
    }

    function parseFunc(t, callee) {
      expect("lparen");
      var args = callee ? [callee] : [];
      var type = callee ? "call" : "func";
      if (peek() && peek().type === "rparen") {
        var r0 = next();
        return { type: type, name: t.name, start: t.start, end: r0.end, children: args };
      }
      for (;;) {
        var nt = peek();
        if (nt && (nt.type === "comma" || nt.type === "rparen")) {
          args.push({ type: "empty", start: nt.start, end: nt.start, text: "" });
        } else {
          args.push(parseExpr(0));
        }
        var sep = next();
        if (!sep) throw new Error("Missing ) for " + t.name);
        if (sep.type === "rparen") return { type: type, name: t.name, start: t.start, end: sep.end, children: args };
        if (sep.type !== "comma") throw new Error("Unexpected '" + sep.text + "' in " + t.name + "()");
      }
    }

    if (!toks.length) throw new Error("Empty formula");
    var ast = parseExpr(0);
    if (p < toks.length) throw new Error("Unexpected '" + toks[p].text + "'");
    return { ast: ast, tokens: toks, formula: formula };
  }

  // ---------- Helpers ----------
  function walk(node, fn, depth) {
    depth = depth || 0;
    fn(node, depth);
    if (node.children) for (var k = 0; k < node.children.length; k++) walk(node.children[k], fn, depth + 1);
  }

  /** Reference-like leaves (refs, names, structured refs) in order of appearance. */
  function collectRefs(ast) {
    // LET / LAMBDA parameter names are local variables, not workbook names
    var locals = {};
    walk(ast, function (n) {
      if (n.type !== "func") return;
      var fn = n.name.replace(/^_xlfn\./i, "").toUpperCase();
      var args = n.children, k;
      if (fn === "LET") { for (k = 0; k < args.length - 1; k += 2) if (args[k].type === "name") locals[localKey(args[k].name)] = true; }
      if (fn === "LAMBDA") { for (k = 0; k < args.length - 1; k++) if (args[k].type === "name") locals[localKey(args[k].name)] = true; }
    });
    var out = [];
    walk(ast, function (n) {
      if (n.type === "name" && !n.sheet && locals[localKey(n.name)]) return;
      if (n.type === "ref" || n.type === "name" || n.type === "struct") out.push(n);
    });
    return out;
  }
  function localKey(nm) { return nm.replace(/^_xlpm\./i, "").toUpperCase(); }

  var DYNAMIC = /^(?:_xlfn\.)?(OFFSET|INDIRECT)$/i;
  function dynamicFunctions(ast) {
    var found = {};
    walk(ast, function (n) { if (n.type === "func" && DYNAMIC.test(n.name)) found[n.name.replace(/^_xlfn\./i, "").toUpperCase()] = true; });
    return Object.keys(found);
  }

  function quoteSheet(name) { return "'" + String(name).replace(/'/g, "''") + "'"; }

  /** Source text of a node, with every unqualified A1 reference prefixed by defaultSheet. */
  function qualifiedText(parsed, node, defaultSheet, sheetScopedNames) {
    var s = parsed.formula, out = "", pos = node.start;
    for (var k = 0; k < parsed.tokens.length; k++) {
      var t = parsed.tokens[k];
      if (t.start < node.start || t.end > node.end) continue;
      if ((t.type === "ref" && !t.sheet) ||
          (t.type === "name" && !t.sheet && sheetScopedNames && sheetScopedNames[t.name.toUpperCase()])) {
        out += s.slice(pos, t.start) + quoteSheet(defaultSheet) + "!" + t.text;
        pos = t.end;
      }
    }
    return out + s.slice(pos, node.end);
  }

  function nodeText(parsed, node) { return parsed.formula.slice(node.start, node.end); }

  // ---------- A1 address maths ----------
  function colToNum(c) { c = c.toUpperCase(); var n = 0; for (var k = 0; k < c.length; k++) n = n * 26 + (c.charCodeAt(k) - 64); return n; }
  function numToCol(n) { var s = ""; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }

  /** "$A$1:B2" | "A:C" | "3:5" -> {r1,c1,r2,c2} 1-based inclusive */
  function parseArea(addr) {
    var a = addr.replace(/\$/g, "").replace(/#$/, "");
    var parts = a.split(":");
    function one(p) {
      var m;
      if ((m = /^([A-Za-z]+)([0-9]+)$/.exec(p))) return { r: +m[2], c: colToNum(m[1]) };
      if ((m = /^([A-Za-z]+)$/.exec(p))) return { r: null, c: colToNum(m[1]) };
      if ((m = /^([0-9]+)$/.exec(p))) return { r: +m[1], c: null };
      return null;
    }
    var p1 = one(parts[0]), p2 = parts[1] ? one(parts[1]) : p1;
    if (!p1 || !p2) return null;
    var r1 = p1.r == null ? 1 : p1.r, r2 = p2.r == null ? 1048576 : p2.r;
    var c1 = p1.c == null ? 1 : p1.c, c2 = p2.c == null ? 16384 : p2.c;
    return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
  }
  function areaContains(area, r, c) { return r >= area.r1 && r <= area.r2 && c >= area.c1 && c <= area.c2; }
  function areaSize(area) { return (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1); }

  /** Split "Sheet1!A1:B2" / "'My Sheet'!A1" into {sheet, address} */
  function splitAddress(full) {
    var idx = full.lastIndexOf("!");
    if (idx < 0) return { sheet: null, address: full };
    var sh = full.slice(0, idx);
    if (sh[0] === "'" && sh[sh.length - 1] === "'") sh = sh.slice(1, -1).replace(/''/g, "'");
    return { sheet: sh, address: full.slice(idx + 1) };
  }

  /**
   * Structured reference spec. Input: the struct token text, e.g.
   * "Table1[[#This Row],[Sales]]", "[@Sales]", "T[[Jan]:[Mar]]", "T[#All]".
   * Returns {table, thisRow, specials:[...upper], col1, col2}
   */
  function parseStructSpec(text, table) {
    var s = table ? text.slice(table.length) : text;
    s = s.slice(1, -1); // strip outer [ ]
    var res = { table: table || null, thisRow: false, specials: [], col1: null, col2: null };
    function unesc(x) { return x.replace(/'(.)/g, "$1").trim(); }
    function addItem(item) {
      item = item.trim();
      if (!item) return;
      if (item[0] === "#") {
        var sp = item.toUpperCase();
        if (sp === "#THIS ROW") res.thisRow = true; else res.specials.push(sp);
      } else if (res.col1 === null) res.col1 = unesc(item);
      else res.col2 = unesc(item);
    }
    if (s[0] === "@") {
      res.thisRow = true;
      s = s.slice(1).trim();
      if (!s) return res;
    }
    if (s.indexOf("[") < 0) { addItem(s); return res; }
    // top-level bracket items, separated by "," or ":"
    var depth = 0, cur = "";
    for (var k = 0; k < s.length; k++) {
      var ch = s[k];
      if (ch === "'" && k + 1 < s.length) { cur += ch + s[++k]; continue; }
      if (ch === "[") { if (depth++ === 0) { cur = ""; continue; } }
      else if (ch === "]") { if (--depth === 0) { addItem(cur); cur = ""; continue; } }
      if (depth > 0) cur += ch;
    }
    return res;
  }

  /** Split "Sheet1!A1:B2, 'My, Sheet'!C3" style address lists on top-level commas. */
  function splitAddressList(str) {
    var out = [], cur = "", q = false;
    for (var k = 0; k < str.length; k++) {
      var ch = str[k];
      if (ch === "'") q = !q;
      if (ch === "," && !q) { if (cur.trim()) out.push(cur.trim()); cur = ""; continue; }
      cur += ch;
    }
    if (cur.trim()) out.push(cur.trim());
    return out;
  }

  var api = {
    parseStructSpec: parseStructSpec, splitAddressList: splitAddressList,
    tokenize: tokenize, parse: parse, walk: walk, collectRefs: collectRefs,
    dynamicFunctions: dynamicFunctions, qualifiedText: qualifiedText, nodeText: nodeText,
    quoteSheet: quoteSheet, parseArea: parseArea, areaContains: areaContains, areaSize: areaSize,
    splitAddress: splitAddress, colToNum: colToNum, numToCol: numToCol,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FxFormula = api;
})(this);
