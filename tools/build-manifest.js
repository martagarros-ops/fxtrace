// Usage: node tools/build-manifest.js https://<user>.github.io/fxtrace
const fs = require("fs"), path = require("path");
const base = (process.argv[2] || "").replace(/\/+$/, "");
if (!/^https:\/\//.test(base)) { console.error("Give the https base URL, e.g. https://myuser.github.io/fxtrace"); process.exit(1); }
const tpl = fs.readFileSync(path.join(__dirname, "../manifest.template.xml"), "utf8");
fs.writeFileSync(path.join(__dirname, "../manifest.xml"), tpl.replace(/\{\{BASE\}\}/g, base));
console.log("manifest.xml written for " + base);
