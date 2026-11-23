/* FxTrace – precedents / dependents tracer and Formula Explorer for Excel (Mac, Windows, web) */
(function () {
  "use strict";
  var F = window.FxFormula;
  var MAX_EXPAND = 500;      // max cells listed when a range is expanded
  var MAX_SUM_CELLS = 5000;  // max cells read to show a Σ for a range
  var SPLIT_DEP_AREAS = 50;  // dependents areas up to this size are listed cell by cell

  var state = {
    mode: "prec", anchor: null, history: [],
    roots: [], flat: [], selIdx: -1,
    fx: null, fxFlat: [], fxSel: -1,
    follow: true, lastNav: null, gen: 0, depIndex: null,
    api: { prec: false, dep: false, events: false },
  };
  var uid = 0;
  var $ = function (id) { return document.getElementById(id); };

  // ---------------------------------------------------------------- startup
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) { setStatus("FxTrace runs inside Excel.", true); return; }
    var req = Office.context.requirements;
    state.api.dep = req.isSetSupported("ExcelApi", "1.15");
    state.api.events = req.isSetSupported("ExcelApi", "1.9");
    wireUi();
    registerShortcuts();
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, onSelectionChanged);
    if (state.api.events) {
      Excel.run(function (ctx) {
        ctx.workbook.worksheets.onChanged.add(function () { state.depIndex = null; return Promise.resolve(); });
        return ctx.sync();
      }).catch(function () { state.api.events = false; });
    }
    traceActive();
  });

  function wireUi() {
    document.querySelectorAll(".tab").forEach(function (b) {
      b.addEventListener("click", function () { setMode(b.getAttribute("data-mode")); });
    });
    $("btn-trace").addEventListener("click", traceActive);
    $("btn-origin").addEventListener("click", goOrigin);
    $("btn-back").addEventListener("click", goBack);
    $("btn-eval").addEventListener("click", evaluateFx);
    $("chk-follow").addEventListener("change", function (e) { state.follow = e.target.checked; });
    $("main").addEventListener("keydown", onKey);
  }

  // ---------------------------------------------------------------- helpers
  function esc(s) { return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]; }); }
  function norm(sheet, address) { return (sheet + "!" + address).replace(/\$/g, "").toUpperCase(); }
  function setStatus(msg, isErr) { var s = $("status"); s.textContent = msg || ""; s.className = "status" + (isErr ? " err" : ""); }
  function fmtVal(v) {
    if (v === null || v === undefined || v === "") return "";
    if (typeof v === "number") {
      if (!isFinite(v)) return String(v);
      var a = Math.abs(v);
      if (a !== 0 && (a < 0.0001 || a >= 1e15)) return v.toExponential(3);
      if (a >= 1000) return v.toLocaleString("en-US", { maximumFractionDigits: 2 });
      return String(+v.toPrecision(8));
    }
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    return String(v);
  }
  function sheetRef(sheet) {
    return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet) && !/^[A-Za-z]{1,3}[0-9]+$/.test(sheet) ? sheet : F.quoteSheet(sheet);
  }
  function labelFor(sheet, address, baseSheet) {
    return (sheet && sheet !== baseSheet ? sheetRef(sheet) + "!" : "") + address.replace(/\$/g, "");
  }
  function baseSheet() { return state.anchor ? state.anchor.sheet : null; }
  function mkEntry(o) { o.id = ++uid; o.children = null; o.expanded = false; return o; }
  function isNotFound(e) {
    return e && (e.code === "ItemNotFound" || /not\s*found|no cells|couldn.t find/i.test(e.message || ""));
  }

  function getSheetMap(ctx) {
    var wss = ctx.workbook.worksheets; wss.load("items/name");
    return ctx.sync().then(function () {
      var m = {}; wss.items.forEach(function (w) { m[w.name.toUpperCase()] = w.name; }); return m;
    });
  }

  // ---------------------------------------------------------------- tracing entry points
  function traceActive() {
    return Excel.run(function (ctx) {
      var cell = ctx.workbook.getActiveCell();
      cell.load("address"); cell.worksheet.load("name");
      return ctx.sync().then(function () {
        var sp = F.splitAddress(cell.address);
        return setAnchor(cell.worksheet.name, sp.address, true);
      });
    }).catch(showError);
  }

  var selTimer = null;
  function onSelectionChanged() {
    if (!state.follow) return;
    clearTimeout(selTimer);
    selTimer = setTimeout(function () {
      Excel.run(function (ctx) {
        var r = ctx.workbook.getSelectedRange(); r.load("address");
        var cell = ctx.workbook.getActiveCell(); cell.load("address");
        return ctx.sync().then(function () {
          var sp = F.splitAddress(r.address);
          var key = norm(sp.sheet, sp.address);
          if (state.lastNav && key === state.lastNav) return;      // our own navigation
          var cs = F.splitAddress(cell.address);
          if (state.anchor && norm(cs.sheet, cs.address) === norm(state.anchor.sheet, state.anchor.address)) return;
          state.lastNav = null;
          return setAnchor(cs.sheet, cs.address, true);
        });
      }).catch(showError);
    }, 180);
  }

  function setAnchor(sheet, address, pushHistory) {
    if (pushHistory && state.anchor && norm(state.anchor.sheet, state.anchor.address) !== norm(sheet, address)) {
      state.history.push(state.anchor);
      if (state.history.length > 50) state.history.shift();
    }
    state.anchor = { sheet: sheet, address: address.replace(/\$/g, "") };
    $("btn-back").disabled = !state.history.length;
    return Excel.run(function (ctx) {
      var r = ctx.workbook.worksheets.getItem(sheet).getRange(address);
      r.load("formulas,text");
      return ctx.sync().then(function () {
        var f = r.formulas[0][0];
        state.anchor.formula = typeof f === "string" && f[0] === "=" ? f : "";
        state.anchor.text = r.text[0][0];
        renderHeader();
      });
    }).then(refresh).catch(showError);
  }

  function setMode(mode) {
    state.mode = mode;
    document.querySelectorAll(".tab").forEach(function (b) { b.classList.toggle("active", b.getAttribute("data-mode") === mode); });
    $("tree").hidden = mode === "fx";
    $("fx").hidden = mode !== "fx";
    refresh();
  }

  function refresh() {
    if (!state.anchor) return;
    var gen = ++state.gen, a = state.anchor;
    setStatus("");
    $("detail").innerHTML = "";
    if (state.mode === "fx") return buildFx(gen);
    $("tree").innerHTML = '<div class="empty">Loading…</div>';
    var p = state.mode === "prec" ? getPrecedents(a.sheet, a.address) : getDependents(a.sheet, a.address);
    return p.then(function (res) {
      if (gen !== state.gen) return;
      res.entries.forEach(function (e) { e.depth = 0; e.parent = null; });
      state.roots = res.entries;
      state.selIdx = res.entries.length ? 0 : -1;
      renderTree();
      if (!res.entries.length) $("tree").innerHTML = '<div class="empty">' + esc(res.empty || "Nothing found.") + "</div>";
      setStatus(res.note || "");
    }).catch(showError);
  }

  // ---------------------------------------------------------------- precedents
  function getPrecedents(sheet, address) {
    return Excel.run(function (ctx) {
      var r = ctx.workbook.worksheets.getItem(sheet).getRange(address);
      r.load("formulas");
      return ctx.sync().then(function () {
        var f = r.formulas[0][0];
        if (typeof f !== "string" || f[0] !== "=") return { entries: [], empty: "No formula – this cell is an input / constant." };
        var parsed;
        try { parsed = F.parse(f); } catch (e) { return { entries: [], empty: "Could not read this formula: " + e.message }; }
        var dyn = F.dynamicFunctions(parsed.ast);
        return resolveRefs(ctx, F.collectRefs(parsed.ast), sheet, address).then(function (entries) {
          return loadInfo(ctx, entries).then(function () {
            return {
              entries: entries,
              empty: "The formula has no cell references.",
              note: dyn.length ? "Uses " + dyn.join(" / ") + ": the range it points to is decided at runtime and is not listed." : "",
            };
          });
        });
      });
    });
  }

  /** Turn ref / name / structured-ref AST nodes into navigable entries. */
  function resolveRefs(ctx, refNodes, baseSheet, baseAddress) {
    return getSheetMap(ctx).then(function (sheets) {
      var entries = [], pend = [], seen = {};
      refNodes.forEach(function (n) {
        var key = n.text.replace(/\$/g, "").toUpperCase();
        if (seen[key]) return;
        seen[key] = true;
        if (n.type === "ref") {
          if (n.book) { entries.push(mkEntry({ label: n.text, kind: "ext", note: "External workbook", node: n })); return; }
          var sh = n.sheet || baseSheet, note = "";
          if (!sheets[sh.toUpperCase()] && sh.indexOf(":") > 0) { note = "3D reference – showing first sheet"; sh = sh.split(":")[0]; }
          var real = sheets[sh.toUpperCase()];
          if (!real) { entries.push(mkEntry({ label: n.text, kind: "ref", note: "Sheet not found", node: n })); return; }
          if (n.spill) note = "Spill range – showing its anchor cell";
          entries.push(mkEntry({ label: n.text, kind: "ref", sheet: real, address: n.address.replace(/#$/, ""), note: note, node: n }));
        } else if (n.type === "name") {
          var e = mkEntry({ label: n.text, kind: "name", node: n });
          entries.push(e);
          var scope = n.sheet ? sheets[n.sheet.toUpperCase()] : baseSheet;
          var p = { e: e, ws: null, wb: null };
          if (scope) { p.ws = ctx.workbook.worksheets.getItem(scope).names.getItemOrNullObject(n.name); p.ws.load("name,type,formula,value"); }
          if (!n.sheet) { p.wb = ctx.workbook.names.getItemOrNullObject(n.name); p.wb.load("name,type,formula,value"); }
          pend.push(p);
        } else if (n.type === "struct") {
          var e2 = mkEntry({ label: n.text, kind: "struct", node: n });
          entries.push(e2);
          var spec = F.parseStructSpec(n.text, n.table);
          var tbl = spec.table
            ? ctx.workbook.tables.getItemOrNullObject(spec.table)
            : ctx.workbook.worksheets.getItem(baseSheet).getRange(baseAddress).getTables(false).getFirstOrNullObject();
          tbl.load("name,showHeaders,showTotals");
          pend.push({ e: e2, tbl: tbl, spec: spec, struct: true });
        }
      });
      if (!pend.length) return entries;
      return ctx.sync().then(function () {
        var structs = [];
        pend.forEach(function (p) {
          if (p.struct) {
            if (p.tbl.isNullObject) { p.e.note = "Table not found"; return; }
            p.rng = p.tbl.getRange(); p.rng.load("address");
            p.cols = p.tbl.columns; p.cols.load("items/name");
            structs.push(p);
            return;
          }
          var it = p.ws && !p.ws.isNullObject ? p.ws : (p.wb && !p.wb.isNullObject ? p.wb : null);
          if (!it) { p.e.note = "Name not found"; return; }
          p.e.scope = it === p.ws ? "sheet" : "workbook";
          p.e.nameFormula = it.formula;
          try {
            var pa = F.parse(it.formula).ast;
            if (pa.type === "ref" && !pa.book) {
              var s2 = sheets[(pa.sheet || baseSheet).toUpperCase()];
              if (s2) { p.e.sheet = s2; p.e.address = pa.address.replace(/#$/, ""); return; }
            }
          } catch (ignore) { /* not a plain reference */ }
          p.e.text = it.type === "Range" ? "" : fmtVal(it.value);
          var body = it.formula.replace(/^=/, "");
          p.e.info = body === p.e.text ? "(constant)" : "= " + body;
        });
        if (!structs.length) return entries;
        return ctx.sync().then(function () {
          var base = F.parseArea(baseAddress);
          structs.forEach(function (p) {
            var sp = F.splitAddress(p.rng.address), area = F.parseArea(sp.address), t = p.tbl, s = p.spec;
            var hdr = t.showHeaders ? area.r1 : null, tot = t.showTotals ? area.r2 : null;
            var d1 = area.r1 + (t.showHeaders ? 1 : 0), d2 = area.r2 - (t.showTotals ? 1 : 0);
            var names = p.cols.items.map(function (c) { return c.name.toUpperCase(); });
            var c1 = area.c1, c2 = area.c2;
            if (s.col1) {
              var i1 = names.indexOf(s.col1.toUpperCase());
              if (i1 < 0) { p.e.note = "Column '" + s.col1 + "' not found"; return; }
              c1 = c2 = area.c1 + i1;
              if (s.col2) {
                var i2 = names.indexOf(s.col2.toUpperCase());
                if (i2 >= 0) { c2 = area.c1 + i2; if (c2 < c1) { var tmp = c1; c1 = c2; c2 = tmp; } }
              }
            }
            var r1, r2;
            if (s.thisRow) { r1 = r2 = base.r1; }
            else if (!s.specials.length) { r1 = d1; r2 = d2; }
            else {
              r1 = Infinity; r2 = -Infinity;
              s.specials.forEach(function (x) {
                var a = x === "#ALL" ? [area.r1, area.r2] : x === "#DATA" ? [d1, d2] : x === "#HEADERS" ? [hdr, hdr] : x === "#TOTALS" ? [tot, tot] : null;
                if (a && a[0] != null) { r1 = Math.min(r1, a[0]); r2 = Math.max(r2, a[1]); }
              });
              if (!isFinite(r1)) { p.e.note = "Empty table section"; return; }
            }
            if (r2 < r1) { p.e.note = "Table has no data rows"; return; }
            p.e.sheet = sp.sheet;
            var a1 = F.numToCol(c1) + r1, a2 = F.numToCol(c2) + r2;
            p.e.address = a1 === a2 ? a1 : a1 + ":" + a2;
          });
          return entries;
        });
      });
    });
  }

  /** Load value / formula / Σ for entries that point at cells. */
  function loadInfo(ctx, entries) {
    var jobs = [];
    entries.forEach(function (e) {
      if (!e.sheet || !e.address) { e.expandable = false; return; }
      var area = F.parseArea(e.address);
      if (!area) return;
      e.cells = F.areaSize(area);
      var r = ctx.workbook.worksheets.getItem(e.sheet).getRange(e.address);
      if (e.cells === 1) r.load("formulas,text,values");
      else if (e.cells <= MAX_SUM_CELLS) r.load("values");
      jobs.push({ e: e, r: r });
    });
    if (!jobs.length) return Promise.resolve(entries);
    return ctx.sync().then(function () {
      jobs.forEach(function (j) {
        var e = j.e;
        if (e.cells === 1) {
          var f = j.r.formulas[0][0];
          e.formula = typeof f === "string" && f[0] === "=" ? f : "";
          e.text = j.r.text[0][0];
          e.raw = j.r.values[0][0];
          e.expandable = state.mode === "dep" ? true : !!e.formula;
        } else {
          e.expandable = true;
          if (e.cells <= MAX_SUM_CELLS) {
            var sum = 0, n = 0;
            j.r.values.forEach(function (row) { row.forEach(function (v) { if (typeof v === "number") { sum += v; n++; } }); });
            e.text = n ? "Σ " + fmtVal(sum) : "";
          }
        }
      });
      return entries;
    }).catch(function (err) {
      entries.forEach(function (e) { if (e.sheet && e.text == null) e.note = e.note || "Could not read: " + err.message; });
      return entries;
    });
  }

  // ---------------------------------------------------------------- dependents
  function getDependents(sheet, address) {
    return Excel.run(function (ctx) {
      var viaApi = state.api.dep ? apiDependents(ctx, sheet, address) : Promise.resolve(null);
      return viaApi.then(function (addrs) {
        if (addrs) return { addrs: addrs, scanned: false };
        setStatus("Indexing formulas in the workbook…");
        return scanDependents(ctx, sheet, address).then(function (a) { return { addrs: a, scanned: true }; });
      }).then(function (res) {
        var entries = [];
        res.addrs.forEach(function (full) {
          var sp = F.splitAddress(full), area = F.parseArea(sp.address);
          if (!area) return;
          var n = F.areaSize(area);
          if (n > 1 && n <= SPLIT_DEP_AREAS) {
            for (var r = area.r1; r <= area.r2; r++) for (var c = area.c1; c <= area.c2; c++) {
              var a = F.numToCol(c) + r;
              entries.push(mkEntry({ label: labelFor(sp.sheet, a, baseSheet()), kind: "cell", sheet: sp.sheet, address: a }));
            }
          } else {
            entries.push(mkEntry({ label: labelFor(sp.sheet, sp.address, baseSheet()), kind: "cell", sheet: sp.sheet, address: sp.address.replace(/\$/g, "") }));
          }
        });
        return loadInfo(ctx, entries).then(function () {
          return {
            entries: entries,
            empty: "No formulas depend on this cell.",
            note: res.scanned ? "Dependents found by scanning formulas (structured references and OFFSET/INDIRECT not included)." : "",
          };
        });
      });
    });
  }

  function apiDependents(ctx, sheet, address) {
    var r = ctx.workbook.worksheets.getItem(sheet).getRange(address);
    var d = r.getDirectDependents();
    d.load("addresses");
    return ctx.sync().then(function () {
      var out = [];
      d.addresses.forEach(function (a) { out.push.apply(out, F.splitAddressList(a)); });
      return out;
    }).catch(function (e) {
      if (isNotFound(e)) return [];
      return null; // fall back to scanning
    });
  }

  function buildIndex(ctx) {
    var idx = state.depIndex;
    if (idx && (state.api.events || Date.now() - idx.t < 30000)) return Promise.resolve(idx);
    var wss = ctx.workbook.worksheets; wss.load("items/name");
    var wbNames = ctx.workbook.names; wbNames.load("items/name,formula");
    return ctx.sync().then(function () {
      var used = wss.items.map(function (ws) {
        var u = ws.getUsedRangeOrNullObject(); u.load("address,formulas");
        var nm = ws.names; nm.load("items/name,formula");
        return { ws: ws, u: u, names: nm };
      });
      return ctx.sync().then(function () {
        var sheetMap = {}; wss.items.forEach(function (w) { sheetMap[w.name.toUpperCase()] = w.name; });
        function nameTarget(formula, host) {
          try {
            var a = F.parse(formula).ast;
            if (a.type === "ref" && !a.book) return { sheet: (sheetMap[(a.sheet || host).toUpperCase()] || a.sheet || host).toUpperCase(), area: F.parseArea(a.address) };
          } catch (ignore) { /* not a range */ }
          return null;
        }
        var wbN = {}, wsN = {};
        wbNames.items.forEach(function (n) { var t = nameTarget(n.formula, ""); if (t) wbN[n.name.toUpperCase()] = t; });
        used.forEach(function (x) { x.names.items.forEach(function (n) { var t = nameTarget(n.formula, x.ws.name); if (t) wsN[x.ws.name.toUpperCase() + "|" + n.name.toUpperCase()] = t; }); });
        var cache = {}, cells = [];
        used.forEach(function (x) {
          if (x.u.isNullObject) return;
          var host = x.ws.name, hostU = host.toUpperCase();
          var org = F.parseArea(F.splitAddress(x.u.address).address);
          x.u.formulas.forEach(function (row, ri) {
            row.forEach(function (f, ci) {
              if (typeof f !== "string" || f[0] !== "=") return;
              var refs = cache[f];
              if (!refs) {
                refs = [];
                try {
                  F.collectRefs(F.parse(f).ast).forEach(function (n) {
                    if (n.type === "ref" && !n.book) refs.push({ sheet: n.sheet, area: F.parseArea(n.address) });
                    else if (n.type === "name") refs.push({ name: n.name.toUpperCase(), sheet: n.sheet });
                  });
                } catch (ignore) { /* unparsable */ }
                cache[f] = refs;
              }
              var resolved = [];
              refs.forEach(function (rf) {
                if (rf.name) {
                  var t = wsN[(rf.sheet || host).toUpperCase() + "|" + rf.name] || (!rf.sheet && wbN[rf.name]);
                  if (t && t.area) resolved.push(t);
                } else if (rf.area) {
                  var s = rf.sheet ? rf.sheet.toUpperCase() : hostU;
                  if (s.indexOf(":") > 0 && !sheetMap[s]) {
                    // 3D ref: expand to all sheets between the two names
                    var order = wss.items.map(function (w) { return w.name.toUpperCase(); });
                    var ends = s.split(":"), i1 = order.indexOf(ends[0]), i2 = order.indexOf(ends[1]);
                    for (var k = Math.min(i1, i2); k >= 0 && k <= Math.max(i1, i2); k++) resolved.push({ sheet: order[k], area: rf.area });
                  } else resolved.push({ sheet: s, area: rf.area });
                }
              });
              if (resolved.length) cells.push({ sheet: host, r: org.r1 + ri, c: org.c1 + ci, refs: resolved });
            });
          });
        });
        state.depIndex = { t: Date.now(), cells: cells };
        return state.depIndex;
      });
    });
  }

  function scanDependents(ctx, sheet, address) {
    return buildIndex(ctx).then(function (idx) {
      var tgt = F.parseArea(address), su = sheet.toUpperCase(), out = [];
      idx.cells.forEach(function (cell) {
        for (var k = 0; k < cell.refs.length; k++) {
          var rf = cell.refs[k];
          if (rf.sheet === su && rf.area.r1 <= tgt.r2 && rf.area.r2 >= tgt.r1 && rf.area.c1 <= tgt.c2 && rf.area.c2 >= tgt.c1) {
            out.push(F.quoteSheet(cell.sheet) + "!" + F.numToCol(cell.c) + cell.r);
            break;
          }
        }
      });
      return out;
    });
  }

  // ---------------------------------------------------------------- tree expand / collapse
  function ancestorsInclude(e, sheet, address) {
    var key = norm(sheet, address);
    if (state.anchor && norm(state.anchor.sheet, state.anchor.address) === key) return true;
    for (var p = e; p; p = p.parent) if (p.sheet && norm(p.sheet, p.address) === key) return true;
    return false;
  }

  function expand(e) {
    if (!e.expandable || e.expanded) return Promise.resolve();
    if (e.children) { e.expanded = true; renderTree(); return Promise.resolve(); }
    var p;
    if (e.cells === 1) {
      p = (state.mode === "prec" ? getPrecedents(e.sheet, e.address) : getDependents(e.sheet, e.address))
        .then(function (res) { if (res.note) setStatus(res.note); return res.entries; });
    } else {
      p = expandRange(e);
    }
    return p.then(function (kids) {
      kids.forEach(function (k) {
        k.parent = e; k.depth = e.depth + 1;
        if (k.sheet && k.cells === 1 && ancestorsInclude(e, k.sheet, k.address)) { k.expandable = false; k.note = "Circular – already in this path"; }
      });
      e.children = kids;
      e.expanded = true;
      if (!kids.length) e.expandable = false;
      renderTree();
    }).catch(showError);
  }

  function expandRange(e) {
    return Excel.run(function (ctx) {
      var area = F.parseArea(e.address);
      var rows = area.r2 - area.r1 + 1, cols = area.c2 - area.c1 + 1;
      if (rows * cols > MAX_EXPAND) {
        // keep the first MAX_EXPAND cells, row by row
        rows = Math.max(1, Math.floor(MAX_EXPAND / cols));
        if (cols > MAX_EXPAND) cols = MAX_EXPAND;
        setStatus("Large range – showing the first " + rows * cols + " cells.");
      }
      var r = ctx.workbook.worksheets.getItem(e.sheet).getRangeByIndexes(area.r1 - 1, area.c1 - 1, rows, cols);
      r.load("formulas,text,values");
      return ctx.sync().then(function () {
        var kids = [];
        for (var i = 0; i < rows; i++) for (var j = 0; j < cols; j++) {
          var f = r.formulas[i][j], a = F.numToCol(area.c1 + j) + (area.r1 + i);
          var isF = typeof f === "string" && f[0] === "=";
          if (!isF && r.values[i][j] === "") continue; // skip blanks
          var k = mkEntry({ label: labelFor(e.sheet, a, baseSheet()), kind: "cell", sheet: e.sheet, address: a, cells: 1 });
          k.formula = isF ? f : ""; k.text = r.text[i][j]; k.raw = r.values[i][j];
          k.expandable = state.mode === "dep" ? true : isF;
          kids.push(k);
        }
        return kids;
      });
    });
  }

  function collapse(e) { e.expanded = false; renderTree(); }

  // ---------------------------------------------------------------- tree rendering
  function flatten() {
    var out = [];
    (function rec(list) { list.forEach(function (e) { out.push(e); if (e.expanded && e.children) rec(e.children); }); })(state.roots);
    state.flat = out;
  }

  function renderTree() {
    flatten();
    var html = state.flat.map(function (e, i) {
      var caret = e.expandable ? (e.expanded ? "▾" : "▸") : "";
      if (e.kind === "cell") e.label = labelFor(e.sheet, e.address, baseSheet());
      var lblCls = "lbl mono" + (e.kind === "name" ? " k-name" : e.kind === "struct" ? " k-struct" : e.kind === "ext" ? " k-ext" : "");
      var sub = "";
      if ((e.kind === "name" || e.kind === "struct") && e.sheet) sub = "→ " + labelFor(e.sheet, e.address, baseSheet());
      else if (e.info) sub = e.info;
      var badges = "";
      if (e.formula) badges += '<span class="badge fx">fx</span>';
      if (e.cells > 1) badges += '<span class="badge">' + e.cells.toLocaleString("en-US") + " cells</span>";
      if (e.note) badges += '<span class="badge warn" title="' + esc(e.note) + '">!</span>';
      return '<div class="row' + (i === state.selIdx ? " sel" : "") + '" role="treeitem" data-i="' + i + '" style="padding-left:' + (e.depth * 14 + 4) + 'px">' +
        '<span class="caret" data-caret="1">' + caret + "</span>" +
        '<span class="' + lblCls + '" title="' + esc(e.label) + '">' + esc(e.label) + "</span>" +
        (sub ? '<span class="sub mono">' + esc(sub) + "</span>" : "") + badges +
        '<span class="rv">' + esc(e.text || "") + "</span></div>";
    }).join("");
    var t = $("tree");
    t.innerHTML = html;
    t.querySelectorAll(".row").forEach(function (row) {
      row.addEventListener("click", function (ev) {
        var i = +row.getAttribute("data-i"), e = state.flat[i];
        if (ev.target.getAttribute("data-caret")) { e.expanded ? collapse(e) : expand(e); return; }
        select(i, true);
      });
      row.addEventListener("dblclick", function () { var e = state.flat[+row.getAttribute("data-i")]; e.expanded ? collapse(e) : expand(e); });
    });
    renderDetail();
    var sel = t.querySelector(".row.sel");
    if (sel) sel.scrollIntoView({ block: "nearest" });
  }

  function renderDetail() {
    var d = $("detail");
    if (state.mode === "fx") return;
    var e = state.flat[state.selIdx];
    if (!e) { d.innerHTML = ""; return; }
    var where = e.sheet ? sheetRef(e.sheet) + "!" + e.address : e.label;
    d.innerHTML = '<div class="d-addr mono">' + esc(where) + (e.text ? " &nbsp;=&nbsp; " + esc(e.text) : "") + "</div>" +
      (e.formula ? '<div class="d-f mono">' + esc(e.formula) + "</div>" : "") +
      (e.nameFormula ? '<div class="d-f mono">' + esc(e.label + " " + e.nameFormula) + (e.scope ? "  (" + e.scope + " name)" : "") + "</div>" : "") +
      (e.note ? '<div class="d-f">' + esc(e.note) + "</div>" : "");
  }

  function select(i, nav) {
    if (i < 0 || i >= state.flat.length) return;
    state.selIdx = i;
    $("tree").querySelectorAll(".row").forEach(function (r) { r.classList.toggle("sel", +r.getAttribute("data-i") === i); });
    var sel = $("tree").querySelector(".row.sel"); if (sel) sel.scrollIntoView({ block: "nearest" });
    renderDetail();
    var e = state.flat[i];
    if (nav && e.sheet && e.address) navigate(e.sheet, e.address);
  }

  function navigate(sheet, address) {
    state.lastNav = norm(sheet, address);
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.getItem(sheet);
      ws.activate();
      ws.getRange(address).select();
      return ctx.sync();
    }).catch(showError);
  }

  function goOrigin() { if (state.anchor) return navigate(state.anchor.sheet, state.anchor.address); }
  function goBack() {
    var prev = state.history.pop();
    if (!prev) return;
    setAnchor(prev.sheet, prev.address, false).then(function () { navigate(prev.sheet, prev.address); });
  }

  // ---------------------------------------------------------------- keyboard
  function onKey(ev) {
    var k = ev.key;
    if (state.mode === "fx") return onFxKey(ev);
    var e = state.flat[state.selIdx];
    if (k === "ArrowDown") { select(Math.min(state.selIdx + 1, state.flat.length - 1), true); }
    else if (k === "ArrowUp") { select(Math.max(state.selIdx - 1, 0), true); }
    else if (k === "ArrowRight" && e) {
      if (e.expandable && !e.expanded) expand(e);
      else if (e.expanded && e.children && e.children.length) select(state.selIdx + 1, true);
    }
    else if (k === "ArrowLeft" && e) {
      if (e.expanded) collapse(e);
      else if (e.parent) select(state.flat.indexOf(e.parent), true);
    }
    else if (k === "Enter" && e && e.sheet && e.cells === 1) {
      navigate(e.sheet, e.address); setAnchor(e.sheet, e.address, true);
    }
    else if (k === "Escape") { goOrigin(); }
    else if (k === "Backspace") { goBack(); }
    else return;
    ev.preventDefault();
  }

  // ---------------------------------------------------------------- header
  function renderHeader(markStart, markEnd) {
    var a = state.anchor;
    $("btn-origin").textContent = sheetRef(a.sheet) + "!" + a.address;
    $("anchor-value").textContent = a.text || "";
    var f = a.formula || "(no formula)";
    if (markStart != null && a.formula) {
      $("anchor-formula").innerHTML = esc(f.slice(0, markStart)) + "<mark>" + esc(f.slice(markStart, markEnd)) + "</mark>" + esc(f.slice(markEnd));
    } else $("anchor-formula").textContent = f;
  }

  // ---------------------------------------------------------------- Formula Explorer
  var LEAF = { num: 1, str: 1, bool: 1, err: 1, array: 1, empty: 1, ref: 1, name: 1, struct: 1 };

  function buildFx(gen) {
    var a = state.anchor, box = $("fx-tree");
    state.fx = null; state.fxFlat = []; state.fxSel = -1;
    $("btn-eval").disabled = true;
    if (!a.formula) { box.innerHTML = '<div class="empty">No formula – this cell is an input / constant.</div>'; return; }
    var parsed;
    try { parsed = F.parse(a.formula); } catch (e) { box.innerHTML = '<div class="empty">Could not read this formula: ' + esc(e.message) + "</div>"; return; }

    function mk(node, depth) {
      // skip redundant parentheses
      if (node.type === "paren") return mk(node.children[0], depth);
      var x = { id: ++uid, node: node, depth: depth, expanded: depth < 3, kids: [] };
      var kids = node.children || [];
      // flatten chains of the same operator: a+b+c -> one node with three operands
      var prefixes = [];
      var family = node.type === "binop" ? ({ "+": "+-", "-": "+-", "*": "*/", "/": "*/", "&": "&" })[node.op] : null;
      if (family) {
        // left-associative chain of the same precedence: a+b-c -> operands a, +b, -c
        var flat = [];
        (function rec(n, op) {
          if (n.type === "binop" && family.indexOf(n.op) >= 0) { rec(n.children[0], op); rec(n.children[1], n.op); }
          else { flat.push(n); prefixes.push(op); }
        })(node, "");
        kids = flat;
      }
      x.kids = [];
      kids.forEach(function (k, i) {
        if (k.type === "empty") return;
        var c = mk(k, depth + 1);
        if (prefixes[i]) c.prefix = prefixes[i];
        x.kids.push(c);
      });
      return x;
    }
    var root = mk(parsed.ast, 0);
    state.fx = { parsed: parsed, root: root, values: {} };
    renderFx();
    $("btn-eval").disabled = false;

    // values of references
    return Excel.run(function (ctx) {
      var refNodes = F.collectRefs(parsed.ast);
      return resolveRefs(ctx, refNodes, a.sheet, a.address).then(function (entries) {
        return loadInfo(ctx, entries).then(function () {
          if (gen !== state.gen || !state.fx) return;
          var byText = {};
          entries.forEach(function (e) { byText[e.label.replace(/\$/g, "").toUpperCase()] = e; });
          walkFx(root, function (x) {
            if (x.node.type === "ref" || x.node.type === "name" || x.node.type === "struct") {
              x.entry = byText[x.node.text.replace(/\$/g, "").toUpperCase()];
            }
          });
          state.fx.sheetScoped = {};
          entries.forEach(function (e) { if (e.kind === "name" && e.scope === "sheet") state.fx.sheetScoped[e.node.name.toUpperCase()] = true; });
          renderFx();
        });
      });
    }).catch(showError);
  }

  function walkFx(x, fn) { fn(x); x.kids.forEach(function (k) { walkFx(k, fn); }); }

  function highlight(text) {
    try {
      var toks = F.tokenize("=" + text), out = "", pos = 1, s = "=" + text;
      toks.forEach(function (t) {
        out += esc(s.slice(pos, t.start));
        var cls = t.type === "ref" ? "t-ref" : t.type === "name" ? "t-name" : t.type === "struct" ? "t-struct" : t.type === "num" ? "t-num" : t.type === "str" ? "t-str" : t.type === "func" ? "t-func" : "";
        out += cls ? '<span class="' + cls + '">' + esc(t.text) + "</span>" : esc(t.text);
        pos = t.end;
      });
      return out + esc(s.slice(pos));
    } catch (e) { return esc(text); }
  }

  function fxValue(x) {
    var v = state.fx.values[x.id];
    if (v !== undefined) return v;
    var t = x.node.type;
    if (x.entry) return x.entry.text || (x.entry.cells > 1 ? x.entry.cells + " cells" : "");
    if (t === "num" || t === "bool" || t === "err") return x.node.text;
    if (t === "str") return x.node.text.slice(1, -1).replace(/""/g, '"');
    return "";
  }

  function renderFx() {
    var out = [];
    (function rec(x) { out.push(x); if (x.expanded) x.kids.forEach(rec); })(state.fx.root);
    state.fxFlat = out;
    var parsed = state.fx.parsed;
    var html = out.map(function (x, i) {
      var caret = x.kids.length ? (x.expanded ? "▾" : "▸") : "";
      var text = F.nodeText(parsed, x.node);
      var shown = text.length > 90 ? text.slice(0, 87) + "…" : text;
      var lbl = highlight(shown);
      var note = x.entry && x.entry.note ? '<span class="badge warn" title="' + esc(x.entry.note) + '">!</span>' : "";
      return '<div class="row' + (i === state.fxSel ? " sel" : "") + '" data-i="' + i + '" style="padding-left:' + (x.depth * 14 + 4) + 'px">' +
        '<span class="caret" data-caret="1">' + caret + '</span>' + (x.prefix ? '<span class="op mono">' + esc(x.prefix) + "</span>" : "") + '<span class="lbl mono">' + lbl + "</span>" + note +
        '<span class="rv">' + esc(fxValue(x)) + "</span></div>";
    }).join("");
    var box = $("fx-tree");
    box.innerHTML = html;
    box.querySelectorAll(".row").forEach(function (row) {
      row.addEventListener("click", function (ev) {
        var i = +row.getAttribute("data-i"), x = state.fxFlat[i];
        if (ev.target.getAttribute("data-caret") && x.kids.length) { x.expanded = !x.expanded; renderFx(); return; }
        fxSelect(i, true);
      });
    });
    renderFxDetail();
  }

  function fxSelect(i, nav) {
    if (i < 0 || i >= state.fxFlat.length) return;
    state.fxSel = i;
    $("fx-tree").querySelectorAll(".row").forEach(function (r) { r.classList.toggle("sel", +r.getAttribute("data-i") === i); });
    var sel = $("fx-tree").querySelector(".row.sel"); if (sel) sel.scrollIntoView({ block: "nearest" });
    var x = state.fxFlat[i];
    renderHeader(x.node.start, x.node.end);
    renderFxDetail();
    if (nav && x.entry && x.entry.sheet) navigate(x.entry.sheet, x.entry.address);
  }

  function renderFxDetail() {
    var x = state.fxFlat[state.fxSel], d = $("detail");
    if (!x) { d.innerHTML = '<div class="d-f">Click a part of the formula to see it highlighted above. References jump to their cells.</div>'; return; }
    var v = fxValue(x);
    var where = x.entry && x.entry.sheet ? sheetRef(x.entry.sheet) + "!" + x.entry.address : "";
    d.innerHTML = '<div class="d-addr mono">' + highlight(F.nodeText(state.fx.parsed, x.node)) + "</div>" +
      (v !== "" ? '<div class="mono">= ' + esc(v) + "</div>" : "") +
      (where ? '<div class="d-f mono">→ ' + esc(where) + "</div>" : "") +
      (x.entry && x.entry.formula ? '<div class="d-f mono">' + esc(x.entry.formula) + "</div>" : "");
  }

  function onFxKey(ev) {
    var k = ev.key, x = state.fxFlat[state.fxSel];
    if (k === "ArrowDown") fxSelect(Math.min(state.fxSel + 1, state.fxFlat.length - 1), true);
    else if (k === "ArrowUp") fxSelect(Math.max(state.fxSel - 1, 0), true);
    else if (k === "ArrowRight" && x && x.kids.length && !x.expanded) { x.expanded = true; renderFx(); }
    else if (k === "ArrowLeft" && x && x.expanded) { x.expanded = false; renderFx(); }
    else if (k === "Enter" && x && x.entry && x.entry.sheet && x.entry.cells === 1) { navigate(x.entry.sheet, x.entry.address); setAnchor(x.entry.sheet, x.entry.address, true); }
    else if (k === "Escape") goOrigin();
    else if (k === "Backspace") goBack();
    else return;
    ev.preventDefault();
  }

  /** Evaluate every non-leaf sub-expression in a temporary very-hidden sheet. */
  function evaluateFx() {
    var fx = state.fx, a = state.anchor;
    if (!fx) return;
    var targets = [];
    walkFx(fx.root, function (x) { if (!LEAF[x.node.type]) targets.push(x); });
    targets = targets.slice(0, 300);
    if (!targets.length) { setStatus("Nothing to evaluate."); return; }
    var row = F.parseArea(a.address).r1 - 1;
    var btn = $("btn-eval"); btn.disabled = true; setStatus("Evaluating…");
    var formulas = targets.map(function (x) {
      var q = F.qualifiedText(fx.parsed, x.node, a.sheet, fx.sheetScoped);
      return '=LET(fxtv,(' + q + '),IF(IFERROR(ROWS(fxtv)*COLUMNS(fxtv),1)>1,"{array "&ROWS(fxtv)&"×"&COLUMNS(fxtv)&"}",fxtv))';
    });
    return Excel.run(function (ctx) {
      var ws = ctx.workbook.worksheets.add("FxTrace_tmp_" + Date.now().toString(36));
      ws.visibility = Excel.SheetVisibility.veryHidden;
      var rng = ws.getRangeByIndexes(row, 0, 1, formulas.length);
      rng.formulas = [formulas];
      rng.load("values");
      return ctx.sync().then(function () {
        targets.forEach(function (x, i) { fx.values[x.id] = fmtVal(rng.values[0][i]); });
      }).finally(function () {
        ws.delete();
        return ctx.sync();
      });
    }).then(function () {
      if (state.fx === fx) { renderFx(); setStatus("Evaluated " + targets.length + " parts. Values of LET variables can't be shown on their own."); }
    }).catch(showError).then(function () { btn.disabled = false; });
  }

  // ---------------------------------------------------------------- keyboard shortcuts in the grid
  // Defined in shortcuts.json; Mac: Cmd+Option+Arrows, Windows: Ctrl+Alt+Arrows.
  var nav = { stack: [], queue: Promise.resolve() };

  function registerShortcuts() {
    if (!Office.actions || !Office.actions.associate) return;
    var map = {
      FxPrecedents: function () { return navDir("prec"); },
      FxDependents: function () { return navDir("dep"); },
      FxNext: function () { return navCycle(1); },
      FxPrevious: function () { return navCycle(-1); },
      FxBack: function () { return navBack(); },
      FxShowPane: function () { return Office.addin.showAsTaskpane(); },
    };
    Object.keys(map).forEach(function (id) {
      Office.actions.associate(id, function () {
        // run one shortcut at a time, in the order pressed
        nav.queue = nav.queue.then(map[id]).catch(showError);
        return nav.queue;
      });
    });
    try {
      if (Office.context.requirements.isSetSupported("SharedRuntime", "1.1")) {
        // load FxTrace (and its shortcuts) automatically whenever a workbook opens
        Office.addin.setStartupBehavior(Office.StartupBehavior.load);
      }
    } catch (ignore) { /* not available */ }
  }

  function readSelection() {
    return Excel.run(function (ctx) {
      var cell = ctx.workbook.getActiveCell(); cell.load("address");
      var sel = ctx.workbook.getSelectedRange(); sel.load("address");
      return ctx.sync().then(function () {
        var c = F.splitAddress(cell.address), r = F.splitAddress(sel.address);
        return { sheet: c.sheet, address: c.address, selKey: norm(r.sheet, r.address) };
      });
    });
  }

  function navDir(dir) {
    return readSelection().then(function (cur) {
      var top = nav.stack[nav.stack.length - 1];
      var item = top && top.list[top.idx];
      var onItem = item && norm(item.sheet, item.address) === cur.selKey;
      if (!onItem) nav.stack = [];
      var src = onItem ? { sheet: item.sheet, address: item.address, cells: item.cells || 1 } : { sheet: cur.sheet, address: cur.address, cells: 1 };
      var prevMode = state.mode;
      state.mode = dir; // expandRange / loadInfo use the mode
      var p = src.cells > 1 ? expandRange(src).then(function (kids) { return { entries: kids }; })
        : dir === "prec" ? getPrecedents(src.sheet, src.address) : getDependents(src.sheet, src.address);
      return p.then(function (res) {
        var list = res.entries.filter(function (e) { return e.sheet && e.address; });
        if (!list.length) {
          state.mode = prevMode;
          setStatus((dir === "prec" ? "No precedents" : "No dependents") + " for " + labelFor(src.sheet, src.address, null) + (res.empty ? " – " + res.empty : ""));
          return;
        }
        var level = { sheet: src.sheet, address: src.address, list: list, idx: 0, dir: dir };
        nav.stack.push(level);
        return navigate(list[0].sheet, list[0].address).then(function () { mirror(level, res.note); });
      });
    });
  }

  function navCycle(delta) {
    var top = nav.stack[nav.stack.length - 1];
    if (!top) { setStatus("Press Cmd+Option+← (precedents) or Cmd+Option+→ (dependents) first."); return Promise.resolve(); }
    top.idx = (top.idx + delta + top.list.length) % top.list.length;
    var e = top.list[top.idx];
    return navigate(e.sheet, e.address).then(function () { mirror(top); });
  }

  function navBack() {
    var top = nav.stack.pop();
    if (!top) return goOrigin() || Promise.resolve();
    return navigate(top.sheet, top.address).then(function () {
      var prev = nav.stack[nav.stack.length - 1];
      if (prev) mirror(prev);
      else { state.lastNav = null; return setAnchor(top.sheet, top.address, false); }
    });
  }

  /** Show the current shortcut level in the panel (if it is open). */
  function mirror(level, note) {
    state.gen++; // cancel any refresh in flight
    state.mode = level.dir;
    document.querySelectorAll(".tab").forEach(function (b) { b.classList.toggle("active", b.getAttribute("data-mode") === level.dir); });
    $("tree").hidden = false; $("fx").hidden = true;
    var same = state.anchor && norm(state.anchor.sheet, state.anchor.address) === norm(level.sheet, level.address);
    if (!same) {
      state.anchor = { sheet: level.sheet, address: level.address.replace(/\$/g, "") };
      Excel.run(function (ctx) {
        var r = ctx.workbook.worksheets.getItem(level.sheet).getRange(level.address);
        r.load("formulas,text");
        return ctx.sync().then(function () {
          var f = r.formulas[0][0];
          state.anchor.formula = typeof f === "string" && f[0] === "=" ? f : "";
          state.anchor.text = r.text[0][0];
          renderHeader();
        });
      }).catch(function () { /* header only */ });
    }
    level.list.forEach(function (e) { e.depth = 0; e.parent = null; e.expanded = false; });
    state.roots = level.list;
    state.selIdx = level.idx;
    renderTree();
    var depth = nav.stack.indexOf(level) + 1;
    setStatus((level.dir === "prec" ? "Precedent " : "Dependent ") + (level.idx + 1) + " of " + level.list.length +
      (depth > 1 ? " · level " + depth : "") + (note ? " · " + note : ""));
  }

  function showError(e) {
    console.error(e);
    setStatus("Error: " + (e && e.message ? e.message : e), true);
  }
})();
