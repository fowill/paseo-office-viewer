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
const theme = { surface0: '#1b1d21', surface1: '#202327', surface2: '#2a2e34', border: '#30353c', foreground: '#e6e8eb', foregroundMuted: '#9aa3ad', accent: '#7fb3a4', accentForeground: '#1b1d21', statusSuccess: '#70b480', statusWarning: '#e5b567', statusDanger: '#e57878' };
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

  const result = await daemon.invokePluginRpc('office-viewer', 'preview.open', { workspaceId: workspace.id, theme });
  assert.equal(typeof result.url, 'string');
  previewUrl = result.url;
  const preview = new URL(previewUrl);
  assert.equal(preview.hostname, '127.0.0.1');
  const fragment = new URLSearchParams(preview.hash.slice(1));
  token = fragment.get('token');
  assert.ok(token);
  assert.equal(fragment.get('embed'), '1');
  const licenses = await fetch(`${preview.origin}/third-party-licenses.txt`);
  assert.equal(licenses.status, 200);
  assert.match(licenses.headers.get('content-type'), /text\/plain/);
  assert.ok((await licenses.text()).length > 100);
  report.checks.push('Bundled third-party licenses are served by the installed plugin.');
  const listing = await daemon.invokePluginRpc('office-viewer', 'preview.list', { workspaceId: workspace.id, path: 'tests/fixtures' });
  const filenames = ['示例文档.pdf', '示例报告.docx', '示例数据.xlsx', '示例演示.pptx'];
  assert.deepEqual(listing.entries.filter(entry => !entry.directory).map(entry => entry.name).sort(), [...filenames].sort());
  assert.ok(listing.entries.every(entry => entry.path.startsWith('tests/fixtures/')));
  for (const filename of filenames) {
    const endpoint = new URL('/api/file', preview.origin);
    endpoint.searchParams.set('path', `tests/fixtures/${filename}`);
    const document = await fetch(endpoint, { headers: { Authorization: `Bearer ${token}` } });
    assert.equal(document.status, 200, filename);
    assert.deepEqual(Buffer.from(await document.arrayBuffer()), await readFile(path.join(project, 'tests', 'fixtures', filename)), filename);
  }
  report.checks.push('preview.open and preview.list RPCs resolve the project workspace and serve all four fixture files unchanged.');

  const firstSelection = await daemon.invokePluginRpc('office-viewer', 'preview.select', { workspaceId: workspace.id, path: `tests/fixtures/${filenames[0]}` });
  assert.equal(firstSelection.viewerConnected, false);
  assert.equal(firstSelection.url, previewUrl);
  report.checks.push('The first preview.select reports no live viewer and returns the embedded viewer URL.');

  browser = await chromium.launch({ channel: process.platform === 'win32' ? 'msedge' : undefined, headless: true });
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  page.setDefaultTimeout(20_000);
  const errors = [];
  const externalRequests = [];
  page.on('pageerror', error => errors.push(error.message));
  page.on('request', request => {
    if (!request.url().startsWith(preview.origin + '/') && !request.url().startsWith('data:') && !request.url().startsWith('blob:')) externalRequests.push(new URL(request.url()).origin);
  });
  await page.goto(firstSelection.url);
  await expect(page.locator('body')).toHaveClass(/embed/);
  await expect(page.locator('.sidebar')).toBeHidden();
  await expect.poll(() => page.evaluate(() => document.documentElement.style.getPropertyValue('--bg'))).toBe(theme.surface0);
  await expect.poll(() => page.evaluate(() => document.documentElement.style.colorScheme)).toBe('dark');
  for (const filename of filenames) {
    if (filename !== filenames[0]) {
      const selected = await daemon.invokePluginRpc('office-viewer', 'preview.select', { workspaceId: workspace.id, path: `tests/fixtures/${filename}` });
      assert.equal(selected.viewerConnected, true);
      assert.equal(selected.url, previewUrl);
    }
    await expect(page.locator('#file-name')).toHaveText(filename);
    await expect(page.locator('#preview-status')).toContainText('已加载');
    await expect(page.locator('#preview-error')).toBeHidden();
    if (filename.endsWith('.pdf')) {
      await expect(page.locator('.textLayer')).toContainText('本地文档预览测试');
      await expect(page.locator('.page-stage')).toHaveCSS('background-color', 'rgb(255, 255, 255)');
      const updatedTheme = { ...theme, surface0: '#22252a', accent: '#80b8c8' };
      await daemon.invokePluginRpc('office-viewer', 'preview.open', { workspaceId: workspace.id, theme: updatedTheme });
      await expect.poll(() => page.evaluate(() => document.documentElement.style.getPropertyValue('--bg'))).toBe(updatedTheme.surface0);
      await expect.poll(() => page.evaluate(() => document.documentElement.style.getPropertyValue('--accent'))).toBe(updatedTheme.accent);
      await expect(page.locator('.page-stage')).toHaveCSS('background-color', 'rgb(255, 255, 255)');
      report.checks.push('The embedded page receives initial and updated theme colors through real RPC and SSE while document paper stays white.');
      await page.locator('#next-page').click();
      await expect(page.locator('.textLayer')).toContainText('第二页中文内容');
    } else if (filename.endsWith('.docx')) {
      await expect(page.locator('section.docx')).toHaveCount(2);
      await expect(page.locator('#preview table')).toContainText('-1,234.50');
      await expect(page.locator('#preview img')).toHaveCount(1);
      await expect.poll(() => page.locator('#preview img').evaluate(image => image.complete && image.naturalWidth > 0)).toBe(true);
    } else if (filename.endsWith('.xlsx')) {
      await expect(page.locator('td[data-address="B4"]')).toHaveText('-34.75');
      await expect(page.locator('td[data-address="C3"]')).toHaveText('12.5%');
      await expect(page.locator('td[data-address="B5"]')).toHaveText('1,165.75');
      await expect(page.locator('td[data-address="D3"]')).toHaveText('2026-10-17');
      await page.locator('td[data-address="B5"]').click();
      await expect(page.locator('#cell-value')).toContainText('公式：=SUM(B3:B4)');
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
  const wordAgain = await daemon.invokePluginRpc('office-viewer', 'preview.select', { workspaceId: workspace.id, path: 'tests/fixtures/示例报告.docx' });
  assert.equal(wordAgain.viewerConnected, true);
  await expect(page.locator('#file-name')).toHaveText('示例报告.docx');
  await expect(page.locator('#preview-status')).toContainText('已加载');
  await expect.poll(() => page.locator('#preview img').evaluate(image => image.complete && image.naturalWidth > 0)).toBe(true);
  report.checks.push('Word images load after switching from PowerPoint back to Word.');
  assert.deepEqual(errors, []);
  assert.deepEqual(externalRequests, []);
  report.checks.push('No browser page errors and no external document/network requests.');
  await page.close();
  await expect.poll(async () => {
    const selected = await daemon.invokePluginRpc('office-viewer', 'preview.select', { workspaceId: workspace.id, path: `tests/fixtures/${filenames[0]}` });
    return selected.viewerConnected;
  }).toBe(false);
  report.checks.push('Closing the preview page disconnects SSE; the next selection reports viewerConnected=false.');
  report.status = 'passed';
  console.log('Live Paseo plugin smoke test passed: open/list/select RPC, SSE switching, theme sync, unchanged bytes, and all four renderers.');
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
