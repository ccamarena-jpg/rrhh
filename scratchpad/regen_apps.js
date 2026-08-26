// Regenera la constante embebida APPS_SCRIPT en index.html desde gas/Codigo.gs
const fs = require('fs');
const htmlPath = process.argv[2];
const gasPath  = process.argv[3];

let html = fs.readFileSync(htmlPath, 'utf8');
let gas  = fs.readFileSync(gasPath, 'utf8').replace(/\r\n/g, '\n').replace(/\n+$/, '');

const esc = s => s.replace(/\\/g, '\\\\').replace(/`/g, '\\`').replace(/\$\{/g, '\\${');

const startStr = 'const APPS_SCRIPT = `';
const si = html.indexOf(startStr);
if (si < 0) { console.error('No se encontró el marcador de inicio'); process.exit(1); }
const contentStart = si + startStr.length;

// Buscar el backtick de cierre no escapado
let ei = -1;
for (let i = contentStart; i < html.length; i++) {
  if (html[i] === '`') {
    // contar backslashes previos
    let bs = 0, j = i - 1;
    while (j >= 0 && html[j] === '\\') { bs++; j--; }
    if (bs % 2 === 0) { ei = i; break; }
  }
}
if (ei < 0) { console.error('No se encontró el backtick de cierre'); process.exit(1); }

const before = html.slice(0, si);
const after  = html.slice(ei + 1);
const newHtml = before + startStr + esc(gas) + '\n`' + after;

fs.writeFileSync(htmlPath, newHtml, 'utf8');
console.log('APPS_SCRIPT regenerado OK. Bytes GAS:', gas.length, '· offset inicio:', si, '· offset cierre viejo:', ei);
