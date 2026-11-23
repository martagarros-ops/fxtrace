const { chromium } = require("/opt/npm-tools/node_modules/playwright");
const fs = require("fs"), path = require("path");
const SRC = path.join(__dirname, "../docs");
const MOCK = fs.readFileSync(path.join(__dirname, "mock-office.js"), "utf8");

const WB = {
  active: { sheet: "P&L", address: "D10" },
  sheets: {
    Assumptions: { B3: { v: 0.05, t: "5.0%" }, B4: { v: 0.3, t: "30.0%" } },
    "P&L": {
      C5: { v: 1000, t: "1,000" },
      D5: { f: "=C5*(1+Growth)", v: 1050, t: "1,050" },
      D6: { f: "=D5*Assumptions!$B$4", v: 315, t: "315" },
      D8: { f: "=SUMPRODUCT(Sales[Units],Sales[Price])", v: 460, t: "460" },
      D10: { f: "=IF(D5>0,D8/D5+SUM(D5:D6)-TaxRate,NA())", v: 1365.21, t: "1,365.2" },
    },
    Data: {
      B2: { v: "Region" }, C2: { v: "Units" }, D2: { v: "Price" },
      B3: { v: "North" }, C3: { v: 10 }, D3: { v: 20 },
      B4: { v: "South" }, C4: { v: 13 }, D4: { v: 20 },
      B5: { v: "East" }, C5: { v: 4 }, D5: { v: 0 },
    },
  },
  names: [
    { name: "Growth", formula: "=Assumptions!$B$3" },
    { name: "TaxRate", formula: "=0.21", type: "Double", value: 0.21 },
  ],
  tables: [{ name: "Sales", sheet: "Data", range: "B2:D5", columns: ["Region", "Units", "Price"] }],
};

(async () => {
  const browser = await chromium.launch({ executablePath: "/usr/bin/google-chrome" });
  for (const support of [{ "1.15": true }, {}]) {
    const tag = support["1.15"] ? "api" : "scan";
    const page = await browser.newPage({ viewport: { width: 340, height: 640 } });
    const errors = [];
    page.on("pageerror", e => errors.push(e.message));
    page.on("console", m => { if (m.type() === "error") errors.push("console: " + m.text()); });
    await page.addInitScript(([wb, sup]) => { window.__WB = wb; window.__SUPPORT = sup; }, [JSON.parse(JSON.stringify(WB)), support]);
    await page.route("https://appsforoffice.microsoft.com/**", r => r.fulfill({ contentType: "application/javascript", body: MOCK }));
    await page.goto("file://" + SRC + "/taskpane.html");
    await page.waitForTimeout(400);
    const rows = async sel => page.$$eval(sel + " .row", rs => rs.map(r => r.innerText.replace(/\s+/g, " ").trim()));
    console.log(`\n=== [${tag}] Precedents of P&L!D10`);
    console.log(await rows("#tree"));
    console.log("status:", await page.textContent("#status"));
    await page.screenshot({ path: `/home/claude/fxtrace/test/shot-${tag}-prec.png` });

    // expand D5 (first row) via keyboard
    await page.click("#tree .row >> nth=0");
    await page.keyboard.press("ArrowRight");
    await page.waitForTimeout(300);
    await page.keyboard.press("ArrowDown");
    await page.waitForTimeout(300);
    console.log("after expand:", await rows("#tree"));
    console.log("log:", JSON.stringify(await page.evaluate(() => window.__log.slice(-4))));
    console.log("anchor still:", await page.textContent("#btn-origin"));
    await page.screenshot({ path: `/home/claude/fxtrace/test/shot-${tag}-expand.png` });

    // structured refs: precedents of D8, expand first
    await page.evaluate(() => { window.__WB.active = { sheet: "P&L", address: "D8" }; });
    await page.click("#btn-trace");
    await page.waitForTimeout(400);
    await page.click("#tree .row >> nth=0");
    await page.keyboard.press("ArrowRight");
    await page.waitForTimeout(400);
    console.log("D8 precedents:", await rows("#tree"));
    console.log("detail:", (await page.textContent("#detail")).trim());
    // dependents of C5 : re-anchor via Enter on row "D5"? set active cell & trace
    await page.evaluate(() => { window.__WB.active = { sheet: "P&L", address: "D5" }; });
    await page.click("#btn-trace");
    await page.waitForTimeout(300);
    await page.click('.tab[data-mode="dep"]');
    await page.waitForTimeout(500);
    console.log(`=== [${tag}] Dependents of P&L!D5`);
    console.log(await rows("#tree"));
    console.log("status:", await page.textContent("#status"));
    await page.screenshot({ path: `/home/claude/fxtrace/test/shot-${tag}-dep.png` });

    // dependents of Assumptions!B3 (via name)
    await page.evaluate(() => { window.__WB.active = { sheet: "Assumptions", address: "B3" }; });
    await page.click("#btn-trace");
    await page.waitForTimeout(500);
    console.log("deps of Assumptions!B3:", await rows("#tree"));

    // back button
    await page.click("#btn-back");
    await page.waitForTimeout(400);
    console.log("after back anchor:", await page.textContent("#btn-origin"));

    // Formula explorer on D10
    await page.evaluate(() => { window.__WB.active = { sheet: "P&L", address: "D10" }; });
    await page.click("#btn-trace");
    await page.waitForTimeout(300);
    await page.click('.tab[data-mode="fx"]');
    await page.waitForTimeout(400);
    console.log(`=== [${tag}] Formula explorer`);
    console.log(await rows("#fx-tree"));
    await page.click("#btn-eval");
    await page.waitForTimeout(400);
    console.log("after eval:", (await rows("#fx-tree")).slice(0, 4));
    console.log("eval log:", JSON.stringify(await page.evaluate(() => window.__log.filter(l => l[0] !== "select" && l[0] !== "activate"))).slice(0, 600));
    await page.click("#fx-tree .row >> nth=3");
    await page.waitForTimeout(200);
    console.log("header mark:", await page.innerHTML("#anchor-formula"));
    await page.screenshot({ path: `/home/claude/fxtrace/test/shot-${tag}-fx.png` });
    console.log("ERRORS:", errors);
    await page.close();
  }
  await browser.close();
})();
