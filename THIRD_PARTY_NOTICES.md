# Third Party Notices

Paseo Office Viewer uses the components below. Their original copyright notices and license texts are copied from the installed packages into [licenses](licenses/). The complete text is also available in [THIRD_PARTY_LICENSES.txt](licenses/THIRD_PARTY_LICENSES.txt). Keep these notices with redistributed viewer builds.

These licenses apply to third-party components. They do not grant a license to this project's original source code. No license has been selected for the original source code.

## Browser viewer components

The versions below are the actual installed versions used by the browser build. JavaScript is bundled and minified by esbuild; retained legal comments remain at the end of the generated bundle. PDF.js worker code, fonts, CMaps, and WebAssembly files are copied without modification.

| Component | Installed version | Package license | Original notices |
| --- | --- | --- | --- |
| pdfjs-dist | 5.4.624 | Apache-2.0 | [LICENSE](licenses/pdfjs-dist/LICENSE), [cmaps/LICENSE](licenses/pdfjs-dist/cmaps/LICENSE), [standard_fonts/LICENSE_FOXIT](licenses/pdfjs-dist/standard_fonts/LICENSE_FOXIT), [standard_fonts/LICENSE_LIBERATION](licenses/pdfjs-dist/standard_fonts/LICENSE_LIBERATION), [wasm/LICENSE_JBIG2](licenses/pdfjs-dist/wasm/LICENSE_JBIG2), [wasm/LICENSE_OPENJPEG](licenses/pdfjs-dist/wasm/LICENSE_OPENJPEG), [wasm/LICENSE_PDFJS_JBIG2](licenses/pdfjs-dist/wasm/LICENSE_PDFJS_JBIG2), [wasm/LICENSE_PDFJS_OPENJPEG](licenses/pdfjs-dist/wasm/LICENSE_PDFJS_OPENJPEG), [wasm/LICENSE_PDFJS_QCMS](licenses/pdfjs-dist/wasm/LICENSE_PDFJS_QCMS), [wasm/LICENSE_QCMS](licenses/pdfjs-dist/wasm/LICENSE_QCMS) |
| docx-preview | 0.4.1 | Apache-2.0 | [LICENSE](licenses/docx-preview/LICENSE) |
| pptxviewjs | 1.1.9 | MIT | [LICENSE](licenses/pptxviewjs/LICENSE) |
| xlsx | 0.20.3 | Apache-2.0 | [LICENSE](licenses/xlsx/LICENSE) |
| chart.js | 4.5.1 | MIT | [LICENSE.md](licenses/chart.js/LICENSE.md) |
| @kurkle/color | 0.3.4 | MIT | [LICENSE.md](licenses/kurkle-color/LICENSE.md) |
| jszip | 3.10.1 | MIT (selected from upstream MIT OR GPL) | [LICENSE.markdown](licenses/jszip/LICENSE.markdown) |

## Embedded dependencies and build tool

JSZip distributes a prebuilt browser bundle that includes compression, Promise, stream, and scheduling helpers. Their notices are included below. Versions in this table identify the installed dependency packages used to obtain the notices; they do not claim to identify every vendored copy's version inside upstream bundles. The last row, esbuild, is the build tool rather than browser runtime code.

| Component | Installed version | Package license | Original notice |
| --- | --- | --- | --- |
| pako | 1.0.11 | (MIT AND Zlib) | [LICENSE](licenses/pako/LICENSE) |
| lie | 3.3.0 | MIT | [license.md](licenses/lie/license.md) |
| immediate | 3.0.6 | MIT | [LICENSE.txt](licenses/immediate/LICENSE.txt) |
| readable-stream | 2.3.8 | MIT | [LICENSE](licenses/readable-stream/LICENSE) |
| core-util-is | 1.0.3 | MIT | [LICENSE](licenses/core-util-is/LICENSE) |
| inherits | 2.0.4 | ISC | [LICENSE](licenses/inherits/LICENSE) |
| isarray | 1.0.0 | MIT | [README.md](licenses/isarray/README.md) |
| process-nextick-args | 2.0.1 | MIT | [license.md](licenses/process-nextick-args/license.md) |
| safe-buffer | 5.1.2 | MIT | [LICENSE](licenses/safe-buffer/LICENSE) |
| setimmediate | 1.0.5 | MIT | [LICENSE.txt](licenses/setimmediate/LICENSE.txt) |
| string_decoder | 1.1.1 | MIT | [LICENSE](licenses/string_decoder/LICENSE) |
| util-deprecate | 1.0.2 | MIT | [LICENSE](licenses/util-deprecate/LICENSE) |
| esbuild | 0.28.2 | MIT | [LICENSE.md](licenses/esbuild/LICENSE.md) |

JSZip is used under the MIT option of its dual license. Its original license file is retained in full, including the alternative GPL text.

Pako includes zlib-derived code. Its additional copyright and restrictions are retained in [ZLIB-NOTICE.txt](licenses/pako/ZLIB-NOTICE.txt). The SheetJS source copyright is retained in [COPYRIGHT.txt](licenses/xlsx/COPYRIGHT.txt). The isarray package keeps its MIT license in its README, so that file is included in full.

## PDF.js supporting assets

The PDF.js package license does not replace the individual notices for its supporting assets. The copies above include Adobe CMap terms; Foxit font terms; the SIL Open Font License for Liberation fonts; and PDFium JBIG2, OpenJPEG, qcms, and Mozilla wrapper notices. All six upstream WebAssembly notice files are retained.

## Rebuild the notices

After installing the locked dependencies, run `node scripts/third-party-notices.mjs`. The script copies the listed license files and regenerates this index and the complete text. Review the package list whenever browser imports, dependency versions, or copied PDF.js assets change.
