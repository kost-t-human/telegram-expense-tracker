// Checks that the markup and the script still agree: every inline handler in
// index.html resolves to something app.js defines, and every element app.js
// reaches for exists in the markup.
// Run: node test/wiring.test.js
const fs = require('fs');

const root = __dirname + '/..';
const html = fs.readFileSync(root + '/index.html', 'utf8');
const js = fs.readFileSync(root + '/app.js', 'utf8');

const matchAll = (src, re) => [...src.matchAll(re)].map(m => m[1]);

const JS_KEYWORDS = new Set(['if', 'for', 'while', 'switch', 'catch', 'return', 'typeof']);

const defined = new Set(matchAll(js, /^(?:async )?function (\w+)/gm)
  .concat(matchAll(js, /^const (\w+)\s*=/gm)));

// Every call inside an inline handler, not just the first one: attributes like
// onkeydown="if(...) addCategory()" hide the real handler behind a keyword.
const handlers = new Set(
  [...html.matchAll(/\bon\w+="([^"]*)"/g)]
    .flatMap(m => matchAll(m[1], /\b(\w+)\s*\(/g))
    .filter(name => !JS_KEYWORDS.has(name)));

// Handlers app.js builds into markup strings are wired the same way.
for (const name of matchAll(js, /onclick="(\w+)\(/g)) handlers.add(name);

const htmlIds = new Set(matchAll(html, /\bid="([\w-]+)"/g));
const usedIds = new Set(matchAll(js, /\$\('([\w-]+)'\)/g));

let failed = 0;
const report = (label, items) => {
  if (!items.length) return;
  console.error(`FAIL: ${label}: ${items.join(', ')}`);
  failed++;
};

report('inline handlers with no definition in app.js', [...handlers].filter(h => !defined.has(h)));
report('elements app.js looks up but index.html does not define', [...usedIds].filter(id => !htmlIds.has(id)));

if (failed) process.exit(1);
console.log(`wiring: ${handlers.size} handlers, ${usedIds.size} element ids OK`);
