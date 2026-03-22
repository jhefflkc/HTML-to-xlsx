const assert = require('assert');
const fs = require('fs');
const path = require('path');
const app = require('../app.js');

const htmlPath = path.join(__dirname, '..', 'Matricula UNI | Horarios.html');
const html = fs.readFileSync(htmlPath, 'utf8');
const rows = app.extractRowsFromHtml(html);

assert.strictEqual(rows.length, 487, 'Debe extraer las 487 filas del HTML de ejemplo.');
assert.deepStrictEqual(rows[0], {
  Cod: 'AA215',
  Secc: 'E',
  Curso: 'GEOLOGIA',
  Tipo: 'TEORIA',
  Horario: 'MA 10-12',
  Aula: 'D2-351',
  Docente: 'ROJAS LEON CARLOS ALBERTO'
});

const workbook = app.buildWorkbook([{ name: 'Matricula UNI | Horarios.html', rows: rows.length }], rows);
assert.ok(workbook.length > 0, 'El XLSX generado no debe estar vacío.');

const signature = Buffer.from(workbook.slice(0, 4)).toString('hex');
assert.strictEqual(signature, '504b0304', 'El archivo debe empezar con la firma ZIP correcta.');

console.log('Smoke test OK:', rows.length, 'filas y XLSX válido.');
