module.exports = {
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
