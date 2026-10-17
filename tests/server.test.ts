import { after, before, test } from 'node:test';
import assert from 'node:assert/strict';
import { copyFile, mkdir, mkdtemp, open, readFile, rm, stat, symlink, utimes, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { request } from 'node:http';
import { createPreviewServer, MAX_FILE_BYTES } from '../server/http';

const project = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const cache = path.join(project, '.cache');
const fixtureNames = ['示例文档.pdf', '示例报告.docx', '示例数据.xlsx', '示例演示.pptx'];
let sandbox: string;
let workspace: string;
let outside: string;
let origin: string;
let token: string;
let preview: Awaited<ReturnType<typeof createPreviewServer>>;

function api(endpoint: string, filename = '', options: RequestInit = {}) {
  const url = new URL(`/api/${endpoint}`, origin);
  url.searchParams.set('path', filename);
  return fetch(url, { ...options, headers: { Authorization: `Bearer ${token}`, ...options.headers } });
}

before(async () => {
  await mkdir(cache, { recursive: true });
  sandbox = await mkdtemp(path.join(cache, 'server-tests-'));
  workspace = path.join(sandbox, 'workspace');
  outside = path.join(sandbox, 'outside');
  await mkdir(workspace);
  await mkdir(outside);
  await Promise.all(fixtureNames.map(name => copyFile(path.join(project, 'tests', 'fixtures', name), path.join(workspace, name))));
  await mkdir(path.join(workspace, '子目录'));
  await copyFile(path.join(workspace, fixtureNames[0]), path.join(workspace, '子目录', '中文 文件.pdf'));
  await writeFile(path.join(workspace, 'unsupported.txt'), 'Synthetic unsupported file');
  await writeFile(path.join(workspace, '.hidden.pdf'), 'Hidden synthetic file');
  await mkdir(path.join(workspace, 'node_modules'));
  await writeFile(path.join(outside, 'outside.pdf'), 'Synthetic external file');
  preview = await createPreviewServer();
  const link = new URL(await preview.openWorkspace(workspace));
  origin = link.origin;
  token = new URLSearchParams(link.hash.slice(1)).get('token')!;
});

after(async () => {
  if (preview) await preview.close();
  // Deletion is restricted to the one temporary directory created by this test.
  if (sandbox) {
    assert.equal(path.dirname(path.resolve(sandbox)), cache);
    assert.match(path.basename(sandbox), /^server-tests-/);
    await rm(sandbox, { recursive: true, force: true });
  }
});

test('viewer HTML and script are available with security headers', async () => {
  const response = await fetch(origin);
  assert.equal(response.status, 200);
  assert.match(response.headers.get('content-type')!, /text\/html/);
  assert.equal(response.headers.get('cache-control'), 'no-store');
  assert.equal(response.headers.get('referrer-policy'), 'no-referrer');
  assert.equal(response.headers.get('x-content-type-options'), 'nosniff');
  assert.match(response.headers.get('content-security-policy')!, /frame-ancestors 'none'/);
  assert.match(await response.text(), /app\.js/);
  assert.equal((await fetch(`${origin}/app.js`)).status, 200);
});

test('API requires a valid session token', async () => {
  assert.equal((await fetch(`${origin}/api/files`)).status, 401);
  assert.equal((await api('files', '', { headers: { Authorization: 'Bearer invalid' } })).status, 401);
  assert.equal((await api('files')).status, 200);
});

test('foreign Origin and Host headers are rejected', async () => {
  assert.equal((await api('files', '', { headers: { Origin: 'https://example.invalid' } })).status, 403);
  assert.equal((await api('files', '', { headers: { Origin: origin } })).status, 200);
  const status = await new Promise<number | undefined>((resolve, reject) => {
    const pending = request(`${origin}/api/files`, { headers: { Host: 'example.invalid', Authorization: `Bearer ${token}` } }, response => {
      response.resume();
      resolve(response.statusCode);
    });
    pending.on('error', reject);
    pending.end();
  });
  assert.equal(status, 403);
});

test('only GET requests are accepted', async () => {
  const response = await api('files', '', { method: 'POST' });
  assert.equal(response.status, 405);
  assert.equal(response.headers.get('allow'), 'GET');
});

test('listing contains supported formats and directories, preserves Unicode paths', async () => {
  const response = await api('files');
  const data = await response.json();
  assert.equal(response.status, 200);
  assert.equal(data.workspaceName, 'workspace');
  assert.equal(data.path, '');
  assert.equal(data.entries[0].name, '子目录');
  assert.equal(data.entries[0].directory, true);
  assert.deepEqual(data.entries.filter((entry: { directory: boolean }) => !entry.directory).map((entry: { name: string }) => entry.name).sort(), [...fixtureNames].sort());
  const nested = await (await api('files', '子目录')).json();
  assert.equal(nested.entries[0].path, '子目录/中文 文件.pdf');
  assert.equal(nested.path, '子目录');
});

test('every format returns the complete original bytes and expected MIME type', async () => {
  const expectedTypes = ['application/pdf', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'];
  for (let index = 0; index < fixtureNames.length; index++) {
    const name = fixtureNames[index];
    const expected = await readFile(path.join(workspace, name));
    const response = await api('file', name);
    assert.equal(response.status, 200, name);
    assert.equal(response.headers.get('content-type'), expectedTypes[index]);
    assert.equal(Number(response.headers.get('content-length')), expected.length);
    assert.deepEqual(Buffer.from(await response.arrayBuffer()), expected, name);
  }
  assert.equal((await api('file', '子目录/中文 文件.pdf')).status, 200);
});

test('stat detects file size and timestamp changes', async () => {
  const target = path.join(workspace, 'refresh.pdf');
  await writeFile(target, 'first');
  const initial = await (await api('stat', 'refresh.pdf')).json();
  await writeFile(target, 'updated synthetic contents');
  const future = new Date(Date.now() + 5000);
  await utimes(target, future, future);
  const updated = await (await api('stat', 'refresh.pdf')).json();
  assert.equal(initial.size, 5);
  assert.equal(updated.size, 26);
  assert.ok(updated.modified > initial.modified);
  assert.equal(updated.size, (await stat(target)).size);
});

test('parent traversal, absolute paths, and junction escapes cannot read outside the workspace', async t => {
  assert.equal((await api('file', '../outside/outside.pdf')).status, 403);
  assert.equal((await api('file', '..\\outside\\outside.pdf')).status, process.platform === 'win32' ? 403 : 404);
  assert.equal((await api('file', path.join(outside, 'outside.pdf'))).status, 403);
  try {
    await symlink(outside, path.join(workspace, 'external'), process.platform === 'win32' ? 'junction' : 'dir');
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === 'EPERM') {
      t.diagnostic('Junction/symlink creation is unavailable; lexical escape assertions still ran.');
      return;
    }
    throw error;
  }
  assert.equal((await api('file', 'external/outside.pdf')).status, 403);
  const listing = await (await api('files')).json();
  assert.ok(!listing.entries.some((entry: { name: string }) => entry.name === 'external'));
});

test('missing files, unsupported formats, directories, and unknown routes return clear errors', async () => {
  assert.equal((await api('file', 'missing.pdf')).status, 404);
  assert.equal((await api('file', 'unsupported.txt')).status, 415);
  assert.equal((await api('file', '子目录')).status, 415);
  assert.equal((await api('files', fixtureNames[0])).status, 400);
  assert.equal((await api('unknown')).status, 404);
  assert.equal((await fetch(`${origin}/missing.js`)).status, 404);
});

test('files larger than 100 MiB are rejected without reading them into memory', async () => {
  const oversized = await open(path.join(workspace, 'oversized.pdf'), 'w');
  try { await oversized.truncate(MAX_FILE_BYTES + 1); } finally { await oversized.close(); }
  assert.equal((await api('file', 'oversized.pdf')).status, 413);
  assert.equal((await api('stat', 'oversized.pdf')).status, 413);
});

test('workspace sessions are isolated and the same root reuses its session', async () => {
  const same = new URL(await preview.openWorkspace(workspace));
  assert.equal(new URLSearchParams(same.hash.slice(1)).get('token'), token);
  const other = new URL(await preview.openWorkspace(outside));
  const otherToken = new URLSearchParams(other.hash.slice(1)).get('token');
  assert.notEqual(otherToken, token);
  const response = await fetch(`${origin}/api/files`, { headers: { Authorization: `Bearer ${otherToken}` } });
  const listing = await response.json();
  assert.deepEqual(listing.entries.map((entry: { name: string }) => entry.name), ['outside.pdf']);
  await assert.rejects(preview.openWorkspace(path.join(workspace, fixtureNames[0])), /工作区目录/);
});
