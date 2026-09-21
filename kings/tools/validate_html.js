const fs = require("fs");

const html = fs.readFileSync("outputs/kings_reign_timeline.html", "utf8");
console.log(JSON.stringify({
  hasTitle: html.includes("열왕기"),
  hasData: html.includes("const DATA"),
  bytes: Buffer.byteLength(html, "utf8"),
}));

const scriptMatch = html.match(/<script>([\s\S]*)<\/script>/);
if (!scriptMatch) {
  throw new Error("No inline script found.");
}

new Function(scriptMatch[1]);
console.log("script ok");
