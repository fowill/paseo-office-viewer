import { parseArgs } from 'node:util';
import path from 'node:path';
import { createPreviewServer } from './http';

const { values } = parseArgs({ options: { root: { type: 'string', default: process.cwd() }, help: { type: 'boolean', short: 'h' } } });
if (values.help) {
  console.log('Usage: paseo-office-viewer --root <project-directory>\nLocal, read-only PDF / DOCX / XLSX / PPTX previews. Stop with Ctrl+C.');
} else {
  const preview = await createPreviewServer();
  try {
    const url = await preview.openWorkspace(path.resolve(values.root!));
    console.log(`Paseo Office Viewer 1.1.0\nWorkspace: ${path.resolve(values.root!)}\nOpen: ${url}\nStop: Ctrl+C`);
    for (const signal of ['SIGINT', 'SIGTERM'] as const) {
      process.once(signal, async () => { await preview.close(); process.exit(0); });
    }
  } catch (error) {
    await preview.close();
    console.error((error as Error).message);
    process.exitCode = 1;
  }
}
