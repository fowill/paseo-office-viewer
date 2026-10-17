// Validate the installed plugin through Paseo's daemon RPC and a fresh headless browser.
// Reads only this project's fixtures. Never prints daemon credentials or preview tokens.
import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { homedir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { randomUUID } from 'node:crypto';
import { createPaseoApi } from '@getpaseo/client';
import { DaemonClient } from '@getpaseo/client/internal/daemon-client';
import { chromium, expect } from '@playwright/test';

const project = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const paseoHome = process.env.PASEO_HOME || path.join(homedir(), '.paseo');
const daemonUrl = process.env.PASEO_SMOKE_URL || 'ws://127.0.0.1:6767/ws';
assert.equal(new URL(daemonUrl).hostname, '127.0.0.1', 'This smoke test is restricted to the local daemon.');
const daemon = new DaemonClient({
  url: daemonUrl,
  clientId: `office-viewer-smoke-${randomUUID()}`,
  clientType: 'cli',
  localCredential: async () => (await readFile(path.join(paseoHome, 'local-credential'), 'utf8')).trim(),
  reconnect: { enabled: false },
  connectTimeoutMs: 10_000,
});
let browser;
let previewUrl = '';
let token = '';
const report = { testedAt: new Date().toISOString(), workspace: project, status: 'failed', checks: [] };

try {
  await daemon.connect();
  const paseo = createPaseoApi(daemon);
  // open() reuses this directory's existing workspace; it creates no agent.
  const workspace = await paseo.workspaces.open(project);
  assert.equal(path.resolve(workspace.directory), project);
  const plugin = (await daemon.listPlugins()).find(entry => entry.id === 'office-viewer');
  assert.ok(plugin, 'office-viewer is not installed in the selected local daemon.');
  assert.equal(plugin.status, 'running');
  report.checks.push('Installed plugin is running in the live local Paseo daemon.');

  const result = await daemon.invokePluginRpc('office-viewer', 'preview.open', { workspaceId: workspace.id });
  assert.equal(typeof result.url, 'string');
  previewUrl = result.url;
  const preview = new URL(previewUrl);
  assert.equal(preview.hostname, '127.0.0.1');
  token = new URLSearchParams(preview.hash.slice(1)).get('token');
  assert.ok(token);
  const licenses = await fetch(`${preview.origin}/third-party-licenses.txt`);
  assert.equal(licenses.status, 200);
  assert.match(licenses.headers.get('content-type'), /text\/plain/);
  assert.ok((await licenses.text()).length > 100);
  report.checks.push('Bundled third-party licenses are served by the installed plugin.');
  const response = await fetch(`${preview.origin}/api/files?path=tests%2Ffixtures`, { headers: { Authorization: `Bearer ${token}` } });
  assert.equal(response.status, 200);
  const listing = await response.json();
  assert.equal(listing.path, 'tests/fixtures');
  const filenames = ['示例文档.pdf', '示例报告.docx', '示例数据.xlsx', '示例演示.pptx'];
  assert.deepEqual(listing.entries.filter(entry => !entry.directory).map(entry => entry.name).sort(), [...filenames].sort());
  for (const filename of filenames) {
    const endpoint = new URL('/api/file', preview.origin);
    endpoint.searchParams.set('path', `tests/fixtures/${filename}`);
    const document = await fetch(endpoint, { headers: { Authorization: `Bearer ${token}` } });
    assert.equal(document.status, 200, filename);
    assert.deepEqual(Buffer.from(await document.arrayBuffer()), await readFile(path.join(project, 'tests', 'fixtures', filename)), filename);
  }
  report.checks.push('preview.open RPC resolves the project workspace and serves all four fixture files unchanged.');

  browser = await chromium.launch({ channel: process.platform === 'win32' ? 'msedge' : undefined, headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  page.setDefaultTimeout(20_000);
  const errors = [];
  const externalRequests = [];
  page.on('pageerror', error => errors.push(error.message));
  page.on('request', request => {
    if (!request.url().startsWith(preview.origin + '/') && !request.url().startsWith('data:') && !request.url().startsWith('blob:')) externalRequests.push(new URL(request.url()).origin);
  });
  await page.goto(previewUrl);
  await page.getByRole('button', { name: '打开文件夹 tests', exact: true }).click();
  await page.getByRole('button', { name: '打开文件夹 fixtures', exact: true }).click();
  for (const filename of filenames) {
    await page.getByRole('button', { name: `预览 ${filename}`, exact: true }).click();
    await expect(page.locator('#file-name')).toHaveText(filename);
    await expect(page.locator('#preview-status')).toContainText('已加载');
    await expect(page.locator('#preview-error')).toBeHidden();
    if (filename.endsWith('.pdf')) {
      await expect(page.locator('.textLayer')).toContainText('本地文档预览测试');
      await page.locator('#next-page').click();
      await expect(page.locator('.textLayer')).toContainText('第二页中文内容');
    } else if (filename.endsWith('.docx')) {
      await expect(page.locator('section.docx')).toHaveCount(2);
      await expect(page.locator('#preview table')).toContainText('-1,234.50');
      await expect(page.locator('#preview img')).toHaveCount(1);
      await expect.poll(() => page.locator('#preview img').evaluate(image => image.complete && image.naturalWidth > 0)).toBe(true);
    } else if (filename.endsWith('.xlsx')) {
      await expect(page.locator('td[data-address="B5"]')).toHaveText('1,165.75');
      await expect(page.locator('td[data-address="D3"]')).toHaveText('2026-10-17');
      await page.getByRole('tab', { name: '说明', exact: true }).click();
      await expect(page.locator('td[data-address="B5"]')).toHaveText('Sheet switching works');
    } else {
      await expect(page.locator('#preview canvas')).toHaveCount(1);
      const first = await page.locator('#preview canvas').evaluate(canvas => canvas.toDataURL());
      await page.locator('#next-page').click();
      await expect(page.locator('#preview-status')).toContainText('第 2 / 2 页');
      const second = await page.locator('#preview canvas').evaluate(canvas => canvas.toDataURL());
      assert.notEqual(first, second);
    }
    report.checks.push(`Headless rendering passed through the installed plugin: ${filename}`);
  }
  await page.getByRole('button', { name: '预览 示例报告.docx', exact: true }).click();
  await expect(page.locator('#preview-status')).toContainText('已加载');
  await expect.poll(() => page.locator('#preview img').evaluate(image => image.complete && image.naturalWidth > 0)).toBe(true);
  report.checks.push('Word images load after switching from PowerPoint back to Word.');
  assert.deepEqual(errors, []);
  assert.deepEqual(externalRequests, []);
  report.checks.push('No browser page errors and no external document/network requests.');
  report.status = 'passed';
  console.log('Live Paseo plugin smoke test passed: RPC, workspace restriction, unchanged bytes, and all four renderers.');
} catch (error) {
  let message = error instanceof Error ? error.message : String(error);
  if (previewUrl) message = message.replaceAll(previewUrl, '[preview URL redacted]');
  if (token) message = message.replaceAll(token, '[preview token redacted]');
  report.error = message;
  console.error(message);
  process.exitCode = 1;
} finally {
  await browser?.close();
  await daemon.close();
  await mkdir(path.join(project, 'audit'), { recursive: true });
  await writeFile(path.join(project, 'audit', 'paseo-smoke.json'), JSON.stringify(report, null, 2) + '\n');
}
