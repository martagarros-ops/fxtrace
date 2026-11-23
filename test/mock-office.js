/* Minimal mock of Office.js / Excel JS API for local UI testing */
(function () {
  var F = null; // formula helpers loaded later
  var WB = window.__WB;
  var SUPPORT = window.__SUPPORT || { "1.15": true };
  var selHandlers = [];
  window.__log = [];

  function colToNum(c) { var n = 0; for (var k = 0; k < c.length; k++) n = n * 26 + (c.charCodeAt(k) - 64); return n; }
  function numToCol(n) { var s = ""; while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); } return s; }
  function area(a) {
    a = a.replace(/\$/g, "").toUpperCase();
    var p = a.split(":"), m1 = /^([A-Z]+)(\d+)$/.exec(p[0]), m2 = /^([A-Z]+)(\d+)$/.exec(p[1] || p[0]);
    return { r1: +m1[2], c1: colToNum(m1[1]), r2: +m2[2], c2: colToNum(m2[1]) };
  }
  function q(s) { return /^[A-Za-z0-9_]+$/.test(s) ? s : "'" + s.replace(/'/g, "''") + "'"; }

  function NullObj() { return { isNullObject: true, load: function () { return this; } }; }

  function Range(sheet, addr) {
    var ar = area(addr), self = this;
    this.sheet = sheet; this.ar = ar;
    var a1 = numToCol(ar.c1) + ar.r1, a2 = numToCol(ar.c2) + ar.r2;
    this.address = q(sheet) + "!" + (a1 === a2 ? a1 : a1 + ":" + a2);
    this.worksheet = new Worksheet(sheet);
    this.isNullObject = false;
    function grid(fn) {
      var out = [];
      for (var r = ar.r1; r <= ar.r2; r++) { var row = []; for (var c = ar.c1; c <= ar.c2; c++) row.push(fn(WB.sheets[sheet][numToCol(c) + r])); out.push(row); }
      return out;
    }
    Object.defineProperty(this, "formulas", {
      get: function () { return grid(function (x) { return x ? (x.f || x.v) : ""; }); },
      set: function (v) { window.__log.push(["write", sheet, v]); },
    });
    Object.defineProperty(this, "values", { get: function () {
      if (/^FxTrace_tmp/.test(sheet)) return [window.__log[window.__log.length - 1][2][0].map(function (_, i) { return i * 1.5; })];
      return grid(function (x) { return x ? x.v : ""; }); } });
    Object.defineProperty(this, "text", { get: function () { return grid(function (x) { return x ? (x.t != null ? x.t : String(x.v)) : ""; }); } });
  }
  Range.prototype.load = function () { return this; };
  Range.prototype.select = function () {
    WB.active = { sheet: this.sheet, address: this.address.split("!")[1] };
    window.__log.push(["select", this.address]);
    setTimeout(function () { selHandlers.forEach(function (h) { h(); }); }, 10);
  };
  Range.prototype.getTables = function () {
    var self = this;
    return { getFirstOrNullObject: function () {
      var t = (WB.tables || []).filter(function (t) { var a = area(t.range); return t.sheet === self.sheet && self.ar.r1 >= a.r1 && self.ar.r1 <= a.r2 && self.ar.c1 >= a.c1 && self.ar.c1 <= a.c2; })[0];
      return t ? new Table(t) : NullObj();
    } };
  };
  Range.prototype.getDirectDependents = function () {
    var self = this, res = { addresses: [], load: function () { return this; } };
    // naive: any formula containing this cell address text
    var key = self.address.split("!")[1];
    Object.keys(WB.sheets).forEach(function (s) {
      Object.keys(WB.sheets[s]).forEach(function (a) {
        var f = WB.sheets[s][a].f || "";
        var hit = (s === self.sheet && new RegExp("(^|[^A-Z!])\\$?" + key.replace(/(\D+)(\d+)/, "$1\\$?$2") + "(?!\\d)").test(f)) ||
          f.indexOf(q(self.sheet) + "!" + key) >= 0 || f.indexOf(q(self.sheet) + "!$" + key.replace(/(\D+)(\d+)/, "$1$$$2")) >= 0;
        (WB.names || []).forEach(function (n) { if (n.formula.replace(/\$/g, "") === "=" + q(self.sheet) + "!" + key && f.indexOf(n.name) >= 0) hit = true; });
        if (hit) res.addresses.push(q(s) + "!" + a);
      });
    });
    res.__empty = !res.addresses.length;
    window.__pendingDep = res;
    return res;
  };

  function Table(t) {
    this.name = t.name; this.showHeaders = t.showHeaders !== false; this.showTotals = !!t.showTotals; this.isNullObject = false;
    this.columns = { items: t.columns.map(function (c) { return { name: c }; }), load: function () { return this; } };
    this._t = t;
  }
  Table.prototype.load = function () { return this; };
  Table.prototype.getRange = function () { return new Range(this._t.sheet, this._t.range); };

  function NamedItem(n) { this.name = n.name; this.formula = n.formula; this.type = n.type || "Range"; this.value = n.value || n.formula; this.isNullObject = false; }
  NamedItem.prototype.load = function () { return this; };

  function Names(scope) {
    var list = (WB.names || []).filter(function (n) { return (n.scope || null) === scope; });
    this.items = list.map(function (n) { return new NamedItem(n); });
    this.getItemOrNullObject = function (nm) {
      var n = list.filter(function (x) { return x.name.toUpperCase() === nm.toUpperCase(); })[0];
      return n ? new NamedItem(n) : NullObj();
    };
  }
  Names.prototype.load = function () { return this; };

  function Worksheet(name) { this.name = name; this.names = new Names(name); }
  Worksheet.prototype.load = function () { return this; };
  Worksheet.prototype.getRange = function (a) { if (!WB.sheets[this.name]) throw new Error("ItemNotFound sheet " + this.name); return new Range(this.name, a); };
  Worksheet.prototype.getRangeByIndexes = function (r, c, rows, cols) { return new Range(this.name, numToCol(c + 1) + (r + 1) + ":" + numToCol(c + cols) + (r + rows)); };
  Worksheet.prototype.activate = function () { window.__log.push(["activate", this.name]); };
  Worksheet.prototype.delete = function () { delete WB.sheets[this.name]; window.__log.push(["delete", this.name]); };
  Worksheet.prototype.getUsedRangeOrNullObject = function () {
    var keys = Object.keys(WB.sheets[this.name]);
    if (!keys.length) return NullObj();
    var r1 = 1e9, r2 = 0, c1 = 1e9, c2 = 0;
    keys.forEach(function (k) { var a = area(k); r1 = Math.min(r1, a.r1); r2 = Math.max(r2, a.r1); c1 = Math.min(c1, a.c1); c2 = Math.max(c2, a.c1); });
    return new Range(this.name, numToCol(c1) + r1 + ":" + numToCol(c2) + r2);
  };

  function Ctx() {
    var wb = {
      getActiveCell: function () { return new Range(WB.active.sheet, WB.active.address.split(":")[0]); },
      getSelectedRange: function () { return new Range(WB.active.sheet, WB.active.address); },
      worksheets: {
        get items() { return Object.keys(WB.sheets).map(function (s) { return new Worksheet(s); }); },
        load: function () { return this; },
        getItem: function (n) { if (!WB.sheets[n]) throw new Error("ItemNotFound: " + n); return new Worksheet(n); },
        add: function (n) { WB.sheets[n] = {}; window.__log.push(["add", n]); return new Worksheet(n); },
        onChanged: { add: function () {} },
      },
      names: new Names(null),
      tables: { getItemOrNullObject: function (n) { var t = (WB.tables || []).filter(function (t) { return t.name.toUpperCase() === n.toUpperCase(); })[0]; return t ? new Table(t) : NullObj(); } },
    };
    this.workbook = wb;
  }
  Ctx.prototype.sync = function () {
    var p = window.__pendingDep; window.__pendingDep = null;
    if (p && p.__empty) return Promise.reject({ code: "ItemNotFound", message: "No cells were found." });
    return new Promise(function (res) { setTimeout(res, 5); });
  };

  window.Excel = { run: function (fn) { return Promise.resolve().then(function () { return fn(new Ctx()); }); }, SheetVisibility: { veryHidden: "VeryHidden" } };
  window.Office = {
    HostType: { Excel: "Excel" },
    EventType: { DocumentSelectionChanged: "sel" },
    context: {
      requirements: { isSetSupported: function (s, v) { return s === "SharedRuntime" || !!SUPPORT[v]; } },
      document: { addHandlerAsync: function (t, h) { selHandlers.push(h); } },
    },
    actions: { associate: function (id, fn) { (window.__actions = window.__actions || {})[id] = fn; } },
    addin: { setStartupBehavior: function () { window.__startup = true; }, showAsTaskpane: function () { return Promise.resolve(); } },
    StartupBehavior: { load: "Load" },
    onReady: function (cb) { setTimeout(function () { cb({ host: "Excel" }); }, 0); },
  };
})();
