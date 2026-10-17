// Paseo calls this from the installed checkout before loading the plugin.
import { access } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
let installed = true;
try {
  for (const name of ['esbuild', 'pdfjs-dist', 'docx-preview', 'pptxviewjs', 'xlsx']) {
    await access(path.join(root, 'node_modules', name, 'package.json'));
  }
} catch { installed = false; }
if (!installed) {
  // Only fixed arguments reach cmd.exe; no file names or user input are interpolated.
  const command = process.platform === 'win32' ? 'cmd.exe' : 'npm';
  const args = process.platform === 'win32' ? ['/d', '/s', '/c', 'npm ci --omit=dev --ignore-scripts'] : ['ci', '--omit=dev', '--ignore-scripts'];
  const result = spawnSync(command, args, { cwd: root, stdio: 'inherit', windowsHide: true });
  if (result.error) throw result.error;
  if (result.status !== 0) process.exit(result.status ?? 1);
}
await import('./build.mjs');
