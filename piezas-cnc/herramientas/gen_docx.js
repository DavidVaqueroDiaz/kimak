const fs = require('fs');
const d = require('docx');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, PageBreak, AlignmentType,
        Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageOrientation } = d;

const data = JSON.parse(fs.readFileSync('materiales.json', 'utf8'));
const AZUL = '1F3864', GRIS = '595959', FONDO = 'DCE6F1';
const nf = n => n.toLocaleString('es-ES');

const hijos = [];

// ---------- Portada ----------
hijos.push(new Paragraph({ spacing: { before: 2400, after: 120 }, alignment: AlignmentType.CENTER,
  children: [new TextRun({ text: 'PIEZAS POR MATERIAL', bold: true, size: 72, color: AZUL })] }));
hijos.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 600 },
  children: [new TextRun({ text: 'Servidor M2  ·  solo archivos .ard  ·  una copia por pieza', size: 26, color: GRIS })] }));
[[nf(data.total) + ' piezas distintas', true],
 [nf(data.clasificadas) + ' con material identificado en el nombre', false],
 [nf(data.sin_material) + ' con el material sin indicar (solo el codigo)', false],
 [nf(data.sin_clasificar) + ' con un texto que no se reconoce como material', false],
 [(data.grupos.length - 2) + ' materiales distintos', false]].forEach(([t, b]) =>
  hijos.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 },
    children: [new TextRun({ text: t, size: 24, bold: b })] })));
hijos.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 600 },
  children: [new TextRun({ text: 'Cada material empieza en una pagina nueva.', size: 22, italics: true, color: GRIS })] }));

// ---------- Indice ----------
hijos.push(new Paragraph({ children: [new PageBreak()] }));
hijos.push(new Paragraph({ spacing: { after: 240 },
  children: [new TextRun({ text: 'INDICE DE MATERIALES', bold: true, size: 40, color: AZUL })] }));

const ANCHOS = [6200, 1800, 1600];
const celda = (txt, opts = {}) => new TableCell({
  width: { size: opts.w, type: WidthType.DXA },
  shading: opts.cab ? { type: ShadingType.CLEAR, fill: FONDO } : undefined,
  margins: { top: 40, bottom: 40, left: 80, right: 80 },
  children: [new Paragraph({ alignment: opts.der ? AlignmentType.RIGHT : AlignmentType.LEFT,
    children: [new TextRun({ text: txt, bold: !!opts.cab, size: 20 })] })] });

const filas = [new TableRow({ tableHeader: true, children: [
  celda('MATERIAL', { w: ANCHOS[0], cab: true }), celda('ESPESOR', { w: ANCHOS[1], cab: true }),
  celda('PIEZAS', { w: ANCHOS[2], cab: true, der: true })] })];
data.grupos.forEach(g => filas.push(new TableRow({ children: [
  celda(g.familia === 'ZZ' ? g.titulo.split('  (')[0] : g.familia, { w: ANCHOS[0] }),
  celda(g.familia === 'ZZ' ? '' : g.espesor, { w: ANCHOS[1] }),
  celda(nf(g.piezas.length), { w: ANCHOS[2], der: true })] })));
hijos.push(new Table({ columnWidths: ANCHOS, rows: filas,
  borders: { top: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' }, bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
             left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
             insideHorizontal: { style: BorderStyle.SINGLE, size: 2, color: 'D9D9D9' }, insideVertical: { style: BorderStyle.NONE } } }));

// ---------- Un bloque por material ----------
data.grupos.forEach(g => {
  hijos.push(new Paragraph({ children: [new PageBreak()] }));
  hijos.push(new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 0, after: 60 },
    shading: { type: ShadingType.CLEAR, fill: FONDO },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: AZUL, space: 4 } },
    children: [new TextRun({ text: g.titulo, bold: true, size: 40, color: AZUL })] }));
  hijos.push(new Paragraph({ spacing: { after: 200 },
    children: [new TextRun({ text: nf(g.piezas.length) + (g.piezas.length === 1 ? ' pieza' : ' piezas'), size: 22, bold: true, color: GRIS })] }));
  g.piezas.forEach(p => hijos.push(new Paragraph({ spacing: { before: 0, after: 0, line: 230 },
    children: [new TextRun({ text: p, size: 18, font: 'Consolas' })] })));
});

const doc = new Document({
  creator: 'KIMAK', title: 'Piezas por material',
  styles: { default: { document: { run: { font: 'Calibri', size: 22 } } } },
  sections: [{
    properties: { page: { margin: { top: 720, right: 720, bottom: 720, left: 900 } } },
    footers: { default: new d.Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      children: [new TextRun({ children: ['Pag. ', d.PageNumber.CURRENT, ' de ', d.PageNumber.TOTAL_PAGES], size: 18, color: GRIS })] })] }) },
    children: hijos }] });

Packer.toBuffer(doc).then(b => { fs.writeFileSync('piezas_por_material.docx', b);
  console.log('OK  ' + (b.length / 1048576).toFixed(2) + ' MB  ·  ' + hijos.length + ' parrafos'); });
