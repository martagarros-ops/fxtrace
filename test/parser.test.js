const F = require("../docs/formula.js");
const cases = [
  "=SUM(A1:B10)*2",
  "=IF(Assumptions!$C$5>0,'P&L 2026'!D12/Revenue,NA())",
  "=-2^2",
  "=XLOOKUP(A2,Table1[Key],Table1[[#Data],[Value]],\"n/a\")",
  "=[@Revenue]*(1+Growth)",
  "=SUM(Sheet1:Sheet3!B2)",
  "=A:A+3:3",
  "=SUMPRODUCT((B2:B9=\"x\")*C2:C9)",
  "=IF(A1,,1)",
  "=LOG10(100)+_xlfn.LET(x,A1,x*2)",
  "=SORT(A2#)",
  "=OFFSET(A1,1,0)+INDIRECT(\"B\"&2)",
  "={1,2;3,4}",
  "=[Budget.xlsx]Sheet1!$A$1+'[Budget 2.xlsx]My Sheet'!B2",
  "=SUM((A1,B1,C1))",
  "=A1%*10",
  "=Sheet2!#REF!+1",
  "=TRUE+FALSE()",
  "=@A1:A10",
  "=INDEX(A1:C3,2,2):C3",
];
let fail = 0;
for (const f of cases) {
  try {
    const p = F.parse(f);
    const refs = F.collectRefs(p.ast).map(r => r.type + ":" + (r.sheet ? r.sheet + "!" : "") + (r.address || r.name || r.text) + (r.book ? "@" + r.book : ""));
    console.log("OK ", f, "\n    root=" + p.ast.type + (p.ast.op || p.ast.name || ""), "refs=", refs.join(" | "), "dyn=", F.dynamicFunctions(p.ast).join(","));
  } catch (e) { fail++; console.log("ERR", f, e.message); }
}
// qualified text
const p = F.parse("=IF(Assumptions!$C$5>0,D12/Revenue,A1:B2)");
console.log(F.qualifiedText(p, p.ast.children[1], "P&L"));
console.log(F.qualifiedText(p, p.ast, "It's"));
console.log(JSON.stringify(F.parseArea("$B$2:A10")), JSON.stringify(F.parseArea("C:C")), JSON.stringify(F.parseArea("4:5")));
console.log(JSON.stringify(F.splitAddress("'My ''Sheet'''!A1:B2")));
process.exit(fail ? 1 : 0);
