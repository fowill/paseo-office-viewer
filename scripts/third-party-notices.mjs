import { copyFile, mkdir, readFile, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const output = path.join(root, 'licenses');

// Direct browser imports, followed by dependencies shipped inside JSZip's browser bundle.
const packages = [
  ['pdfjs-dist', ['LICENSE', 'cmaps/LICENSE', 'standard_fonts/LICENSE_FOXIT', 'standard_fonts/LICENSE_LIBERATION', 'wasm/LICENSE_JBIG2', 'wasm/LICENSE_OPENJPEG', 'wasm/LICENSE_PDFJS_JBIG2', 'wasm/LICENSE_PDFJS_OPENJPEG', 'wasm/LICENSE_PDFJS_QCMS', 'wasm/LICENSE_QCMS']],
  ['docx-preview', ['LICENSE']],
  ['pptxviewjs', ['LICENSE']],
  ['xlsx', ['LICENSE']],
  ['chart.js', ['LICENSE.md']],
  ['@kurkle/color', ['LICENSE.md']],
  ['jszip', ['LICENSE.markdown']],
  ['pako', ['LICENSE']],
  ['lie', ['license.md']],
  ['immediate', ['LICENSE.txt']],
  ['readable-stream', ['LICENSE']],
  ['core-util-is', ['LICENSE']],
  ['inherits', ['LICENSE']],
  ['isarray', ['README.md']],
  ['process-nextick-args', ['license.md']],
  ['safe-buffer', ['LICENSE']],
  ['setimmediate', ['LICENSE.txt']],
  ['string_decoder', ['LICENSE']],
  ['util-deprecate', ['LICENSE']],
  ['esbuild', ['LICENSE.md']],
];

const browserRows = [];
const dependencyRows = [];
const texts = ['Paseo Office Viewer third-party licenses\n\nThese notices apply to third-party components. They do not license the original project source.\nJSZip is used under its MIT license option.\n'];
let copiedFiles = 0;

for (const [name, files] of packages) {
  const source = path.join(root, 'node_modules', name);
  const metadata = JSON.parse(await readFile(path.join(source, 'package.json'), 'utf8'));
  const directory = name.replace('@', '').replaceAll('/', '-');
  const links = [];
  for (const file of files) {
    const relativeOutput = `licenses/${directory}/${file}`;
    const destination = path.join(root, relativeOutput);
    await mkdir(path.dirname(destination), { recursive: true });
    await copyFile(path.join(source, file), destination);
    const text = await readFile(destination, 'utf8');
    if (!text.trim()) throw new Error(`Empty third-party notice: ${name}/${file}`);
    texts.push(`\n\n===== ${name}: ${file} =====\n\n${text}`);
    links.push(`[${file}](${relativeOutput})`);
    copiedFiles += 1;
  }
  const license = name === 'jszip' ? 'MIT (selected from upstream MIT OR GPL)' : metadata.license;
  const row = `| ${name} | ${metadata.version} | ${license} | ${links.join(', ')} |`;
  if (browserRows.length < 7) browserRows.push(row);
  else dependencyRows.push(row);
}

// Pako's root LICENSE contains MIT only; the zlib source has an additional notice.
const pakoSource = await readFile(path.join(root, 'node_modules/pako/lib/zlib/inflate.js'), 'utf8');
const zlibNotice = pakoSource.match(/^\/\/ \(C\)[\s\S]*?^\/\/ 3\.[^\r\n]*/m)?.[0];
if (!zlibNotice) throw new Error('Pako zlib notice was not found in the installed package.');
await writeFile(path.join(output, 'pako/ZLIB-NOTICE.txt'), zlibNotice + '\n');
texts.push(`\n\n===== pako: zlib source notice =====\n\n${zlibNotice}\n`);

// SheetJS carries its copyright in the source header, separately from Apache-2.0.
const sheetjsSource = await readFile(path.join(root, 'node_modules/xlsx/xlsx.mjs'), 'utf8');
const sheetjsCopyright = sheetjsSource.split(/\r?\n/, 1)[0];
if (!sheetjsCopyright.includes('SheetJS')) throw new Error('SheetJS copyright header was not found.');
await writeFile(path.join(output, 'xlsx/COPYRIGHT.txt'), sheetjsCopyright + '\n');
texts.push(`\n\n===== xlsx: source copyright =====\n\n${sheetjsCopyright}\n`);

await writeFile(path.join(output, 'THIRD_PARTY_LICENSES.txt'), texts.join(''));
const notice = `# Third Party Notices

Paseo Office Viewer uses the components below. Their original copyright notices and license texts are copied from the installed packages into [licenses](licenses/). The complete text is also available in [THIRD_PARTY_LICENSES.txt](licenses/THIRD_PARTY_LICENSES.txt). Keep these notices with redistributed viewer builds.

These licenses apply to third-party components. They do not grant a license to this project's original source code. No license has been selected for the original source code.

## Browser viewer components

The versions below are the actual installed versions used by the browser build. JavaScript is bundled and minified by esbuild; retained legal comments remain at the end of the generated bundle. PDF.js worker code, fonts, CMaps, and WebAssembly files are copied without modification.

| Component | Installed version | Package license | Original notices |
| --- | --- | --- | --- |
${browserRows.join('\n')}

## Embedded dependencies and build tool

JSZip distributes a prebuilt browser bundle that includes compression, Promise, stream, and scheduling helpers. Their notices are included below. Versions in this table identify the installed dependency packages used to obtain the notices; they do not claim to identify every vendored copy's version inside upstream bundles. The last row, esbuild, is the build tool rather than browser runtime code.

| Component | Installed version | Package license | Original notice |
| --- | --- | --- | --- |
${dependencyRows.join('\n')}

JSZip is used under the MIT option of its dual license. Its original license file is retained in full, including the alternative GPL text.

Pako includes zlib-derived code. Its additional copyright and restrictions are retained in [ZLIB-NOTICE.txt](licenses/pako/ZLIB-NOTICE.txt). The SheetJS source copyright is retained in [COPYRIGHT.txt](licenses/xlsx/COPYRIGHT.txt). The isarray package keeps its MIT license in its README, so that file is included in full.

## PDF.js supporting assets

The PDF.js package license does not replace the individual notices for its supporting assets. The copies above include Adobe CMap terms; Foxit font terms; the SIL Open Font License for Liberation fonts; and PDFium JBIG2, OpenJPEG, qcms, and Mozilla wrapper notices. All six upstream WebAssembly notice files are retained.

## Rebuild the notices

After installing the locked dependencies, run \`node scripts/third-party-notices.mjs\`. The script copies the listed license files and regenerates this index and the complete text. Review the package list whenever browser imports, dependency versions, or copied PDF.js assets change.
`;
await writeFile(path.join(root, 'THIRD_PARTY_NOTICES.md'), notice);
console.log(`Collected ${copiedFiles + 2} notice files from ${packages.length} installed packages.`);
