// Synthetic files for local preview tests. No user data is read.
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { deflateSync } from 'node:zlib';
import {
  PDFDocument, StandardFonts, rgb, PDFString, PDFHexString,
  beginText, endText, setFontAndSize, setTextMatrix, showText,
} from 'pdf-lib';
import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun, PageBreak, WidthType } from 'docx';
import * as XLSX from 'xlsx';
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(root, 'tests', 'fixtures');
await mkdir(output, { recursive: true });

// A tiny deterministic PNG chart, embedded in Word, PDF, and PowerPoint.
function pngChunk(type, data) {
  const content = Buffer.concat([Buffer.from(type), data]);
  let crc = 0xffffffff;
  for (const byte of content) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  const size = Buffer.alloc(4);
  size.writeUInt32BE(data.length);
  const checksum = Buffer.alloc(4);
  checksum.writeUInt32BE((crc ^ 0xffffffff) >>> 0);
  return Buffer.concat([size, content, checksum]);
}
const imageWidth = 240;
const imageHeight = 120;
const pixels = Buffer.alloc((imageWidth * 3 + 1) * imageHeight, 255);
for (let y = 0; y < imageHeight; y++) {
  pixels[y * (imageWidth * 3 + 1)] = 0;
  for (let x = 0; x < imageWidth; x++) {
    const bar = (x >= 30 && x < 70 && y >= 70) || (x >= 100 && x < 140 && y >= 40) || (x >= 170 && x < 210 && y >= 20);
    if (bar && y < 105) pixels.fill(55, y * (imageWidth * 3 + 1) + 1 + x * 3, y * (imageWidth * 3 + 1) + 4 + x * 3);
  }
}
const imageHeader = Buffer.alloc(13);
imageHeader.writeUInt32BE(imageWidth, 0);
imageHeader.writeUInt32BE(imageHeight, 4);
imageHeader[8] = 8;
imageHeader[9] = 2;
const image = Buffer.concat([
  Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
  pngChunk('IHDR', imageHeader), pngChunk('IDAT', deflateSync(pixels)), pngChunk('IEND', Buffer.alloc(0)),
]);

// PDF: exactly two pages; English text, Chinese text, numbers, and an image.
const pdf = await PDFDocument.create();
pdf.setTitle('Office Viewer synthetic PDF');
const latin = await pdf.embedFont(StandardFonts.Helvetica);
const cjkDescriptor = pdf.context.register(pdf.context.obj({
  Type: 'FontDescriptor', FontName: 'STSong-Light', Flags: 4,
  FontBBox: [0, -200, 1000, 1000], ItalicAngle: 0, Ascent: 1000, Descent: -200, CapHeight: 800, StemV: 80,
}));
const cjkDescendant = pdf.context.register(pdf.context.obj({
  Type: 'Font', Subtype: 'CIDFontType0', BaseFont: 'STSong-Light',
  FontDescriptor: cjkDescriptor,
  CIDSystemInfo: { Registry: PDFString.of('Adobe'), Ordering: PDFString.of('GB1'), Supplement: 4 },
}));
const cjkFont = pdf.context.register(pdf.context.obj({
  Type: 'Font', Subtype: 'Type0', BaseFont: 'STSong-Light', Encoding: 'UniGB-UCS2-H', DescendantFonts: [cjkDescendant],
}));
const pdfImage = await pdf.embedPng(image);
for (let number = 1; number <= 2; number++) {
  const page = pdf.addPage([595.28, 841.89]);
  page.drawText(`Office Viewer PDF - Page ${number}`, { x: 48, y: 775, size: 22, font: latin });
  const cjkKey = page.node.newFontDictionary('Chinese', cjkFont);
  const chineseText = number === 1 ? '本地文档预览测试' : '第二页中文内容';
  const encodedChinese = [...chineseText].map(character => character.charCodeAt(0).toString(16).padStart(4, '0')).join('');
  page.pushOperators(beginText(), setFontAndSize(cjkKey, 18), setTextMatrix(1, 0, 0, 1, 48, 727), showText(PDFHexString.of(encodedChinese)), endText());
  page.drawText(number === 1 ? 'Amount: -1234.50   Rate: 12.5%   Date: 2026-10-17' : 'Second-page text is selectable and searchable.', { x: 48, y: 685, size: 13, font: latin });
  page.drawImage(pdfImage, { x: 48, y: 480, width: 360, height: 180 });
  page.drawLine({ start: { x: 48, y: 64 }, end: { x: 547, y: 64 }, thickness: 0.5, color: rgb(0.7, 0.7, 0.7) });
  page.drawText(`Synthetic fixture | ${number} / 2`, { x: 48, y: 44, size: 10, font: latin });
}
const pdfBytes = await pdf.save();
await writeFile(path.join(output, '示例文档.pdf'), pdfBytes);
assert.equal((await PDFDocument.load(pdfBytes)).getPageCount(), 2);

// Word: explicit page break; shared styles have no paragraph pagination flags.
const wordStyles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="楷体"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="160"/></w:pPr></w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style>
  <w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>
</w:styles>`;
const word = new Document({
  title: 'Office Viewer synthetic Word document', creator: 'Office Viewer tests', externalStyles: wordStyles,
  sections: [{ children: [
    new Paragraph({ text: '本地文档预览测试', style: 'Title' }),
    new Paragraph('This synthetic document checks Chinese text, tables, numbers, images, and page breaks.'),
    new Paragraph({ children: [new TextRun('正文包含中文，'), new TextRun({ text: '粗体', bold: true }), new TextRun('和 '), new TextRun({ text: 'italic text', italics: true }), new TextRun('。')] }),
    new Table({ width: { size: 100, type: WidthType.PERCENTAGE }, rows: [
      new TableRow({ tableHeader: true, children: ['项目', '数值', '单位'].map(text => new TableCell({ children: [new Paragraph(text)] })) }),
      new TableRow({ children: ['金额', '-1,234.50', '元'].map(text => new TableCell({ children: [new Paragraph(text)] })) }),
      new TableRow({ children: ['比例', '12.5%', '百分比'].map(text => new TableCell({ children: [new Paragraph(text)] })) }),
    ] }),
    new Paragraph({ children: [new PageBreak()] }),
    new Paragraph({ text: '第二页图片测试', style: 'Title' }),
    new Paragraph('A simple three-bar image appears below.'),
    new Paragraph({ children: [new ImageRun({ type: 'png', data: image, transformation: { width: 360, height: 180 } })] }),
    new Paragraph('End of the synthetic Word document.'),
  ] }],
});
const wordBytes = await Packer.toBuffer(word);
await writeFile(path.join(output, '示例报告.docx'), wordBytes);
const wordZip = await JSZip.loadAsync(wordBytes);
for (const entry of Object.values(wordZip.files).filter(entry => /^word\/.*\.xml$/.test(entry.name))) {
  const xml = await entry.async('string');
  // Check actual enabled properties, while permitting explicit false values.
  for (const property of xml.matchAll(/<w:(keepNext|keepLines|pageBreakBefore|suppressLineNumbers)\b([^>]*)\/?\s*>/g)) {
    assert.match(property[2], /w:val="(?:0|false|off)"/, `${entry.name} enables ${property[1]}`);
  }
}

// Excel: two sheets, cached formula, numeric date, merged title, signed values.
const workbook = XLSX.utils.book_new();
const results = XLSX.utils.aoa_to_sheet([
  ['结果预览测试', null, null, null],
  ['项目', '金额', '比例', '日期'],
  ['收入', 1200.5, 0.125, 46312],
  ['支出', -34.75, -0.025, 46313],
  ['合计', null, null, null],
]);
results.B3.z = results.B4.z = '#,##0.00;[Red]-#,##0.00';
results.C3.z = results.C4.z = '0.0%';
results.D3.z = results.D4.z = 'yyyy-mm-dd';
results.B5 = { t: 'n', f: 'SUM(B3:B4)', v: 1165.75, z: '#,##0.00' };
results['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
results['!cols'] = [{ wch: 24 }, { wch: 18 }, { wch: 16 }, { wch: 20 }];
results['!rows'] = [{ hpt: 36 }, { hpt: 36 }, { hpt: 30 }, { hpt: 30 }, { hpt: 30 }];
XLSX.utils.book_append_sheet(workbook, results, '结果表');
const notes = XLSX.utils.aoa_to_sheet([
  ['检查项', '预期结果'], ['公式', 'SUM(B3:B4) = 1165.75'], ['负数', -34.75], ['中文', '中文内容完整显示'], ['第二工作表', 'Sheet switching works'],
]);
notes['!cols'] = [{ wch: 24 }, { wch: 44 }];
notes['!rows'] = Array.from({ length: 5 }, () => ({ hpt: 30 }));
XLSX.utils.book_append_sheet(workbook, notes, '说明');
const workbookBytes = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
await writeFile(path.join(output, '示例数据.xlsx'), workbookBytes);
const checkedWorkbook = XLSX.read(workbookBytes, { type: 'buffer', cellDates: true });
assert.deepEqual(checkedWorkbook.SheetNames, ['结果表', '说明']);
assert.equal(checkedWorkbook.Sheets['结果表'].B5.v, 1165.75);
assert.equal(checkedWorkbook.Sheets['结果表'].B5.f, 'SUM(B3:B4)');
assert.equal(checkedWorkbook.Sheets['结果表'].D3.v.toISOString().slice(0, 10), '2026-10-17');

// PowerPoint: two static slides with editable text, table, and embedded image.
const presentation = new PptxGenJS();
presentation.layout = 'LAYOUT_WIDE';
presentation.author = 'Office Viewer tests';
presentation.subject = 'Synthetic preview validation';
presentation.title = '本地演示文稿预览测试';
presentation.lang = 'zh-CN';
const firstSlide = presentation.addSlide();
firstSlide.addText('本地演示文稿预览测试', { x: 0.7, y: 0.6, w: 12, h: 0.7, fontFace: 'Microsoft YaHei', fontSize: 28, bold: true, color: '222222' });
firstSlide.addText('Office Viewer · Slide 1 of 2', { x: 0.7, y: 1.5, w: 12, h: 0.5, fontSize: 18, color: '444444' });
firstSlide.addTable([['项目', '数值'], ['金额', '-1,234.50'], ['比例', '12.5%']], { x: 0.7, y: 2.4, w: 5.4, h: 2.2, fontFace: 'Microsoft YaHei', fontSize: 18, color: '222222', border: { type: 'solid', pt: 1, color: '999999' }, fill: 'FFFFFF', margin: 0.15 });
firstSlide.addImage({ data: `image/png;base64,${image.toString('base64')}`, x: 7, y: 2.4, w: 5.2, h: 2.6 });
const secondSlide = presentation.addSlide();
secondSlide.addText('第二页中文内容', { x: 0.7, y: 0.6, w: 12, h: 0.7, fontFace: 'Microsoft YaHei', fontSize: 28, bold: true, color: '222222' });
secondSlide.addText('Slide switching works. The original file stays unchanged.', { x: 0.7, y: 1.7, w: 11.8, h: 0.8, fontSize: 22, color: '222222' });
secondSlide.addText('日期 2026-10-17', { x: 0.7, y: 3, w: 10, h: 0.6, fontFace: 'Microsoft YaHei', fontSize: 22, color: '222222' });
await presentation.writeFile({ fileName: path.join(output, '示例演示.pptx') });
const slideZip = await JSZip.loadAsync(await readFile(path.join(output, '示例演示.pptx')));
assert.equal(Object.keys(slideZip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length, 2);

await mkdir(path.join(output, '子目录'), { recursive: true });
await writeFile(path.join(output, '子目录', '摘要.pdf'), pdfBytes);
console.log('Created synthetic PDF (2 pages), DOCX (explicit page break), XLSX (2 sheets), PPTX (2 slides), and a nested PDF.');
