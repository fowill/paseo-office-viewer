import { createServer, type ServerResponse } from 'node:http';
import { randomBytes } from 'node:crypto';
import { createReadStream } from 'node:fs';
import { realpath, readdir, stat } from 'node:fs/promises';
import path from 'node:path';
import { assets } from './assets.generated';
import type { PreviewTheme } from '../shared/preview';

const formats: Record<string, string> = {
  '.pdf': 'application/pdf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
};
export const MAX_FILE_BYTES = 100 * 1024 * 1024;

interface PreviewSession {
  root: string;
  theme?: PreviewTheme;
  selected?: { path: string; seq: number };
  viewers: Set<ServerResponse>;
}

// Resolve both paths before checking containment, including Windows junctions.
async function workspacePath(root: string, relative: string) {
  const filename = await realpath(path.resolve(root, relative));
  const inside = path.relative(root, filename);
  if (inside === '..' || inside.startsWith(`..${path.sep}`) || path.isAbsolute(inside)) {
    throw Object.assign(new Error('文件不在当前工作区内。'), { status: 403 });
  }
  return filename;
}

async function listEntries(root: string, relative: string) {
  const filename = await workspacePath(root, relative);
  if (!(await stat(filename)).isDirectory()) {
    throw Object.assign(new Error('请选择文件夹。'), { status: 400 });
  }
  const directoryEntries = await readdir(filename, { withFileTypes: true });
  const visibleEntries = directoryEntries.filter(entry =>
    !entry.name.startsWith('.') && !['node_modules', '$RECYCLE.BIN', 'System Volume Information'].includes(entry.name) &&
    (entry.isDirectory() || Boolean(formats[path.extname(entry.name).toLowerCase()])));
  const entries = [];
  for (const entry of visibleEntries) {
    try {
      const entryPath = path.relative(root, path.join(filename, entry.name));
      const resolved = await workspacePath(root, entryPath);
      const item = await stat(resolved);
      entries.push({ name: entry.name, path: entryPath.split(path.sep).join('/'), directory: item.isDirectory(), size: item.size, modified: item.mtimeMs });
    } catch { /* A removed file or an external junction is not a selectable document. */ }
  }
  entries.sort((a, b) => Number(b.directory) - Number(a.directory) || a.name.localeCompare(b.name, 'zh-CN', { numeric: true }));
  return { path: path.relative(root, filename).split(path.sep).join('/'), workspaceName: path.basename(root), entries };
}

async function previewFile(root: string, relative: string) {
  const filename = await workspacePath(root, relative);
  const info = await stat(filename);
  const mime = formats[path.extname(filename).toLowerCase()];
  if (!info.isFile() || !mime) {
    throw Object.assign(new Error('支持 PDF、DOCX、XLSX 和 PPTX。'), { status: 415 });
  }
  if (info.size > MAX_FILE_BYTES) {
    throw Object.assign(new Error('文件超过 100 MiB 限制。'), { status: 413 });
  }
  return { filename, info, mime };
}

function broadcast(session: PreviewSession, event: string, data: unknown) {
  const message = `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`;
  for (const viewer of session.viewers) {
    if (viewer.destroyed || viewer.writableEnded) session.viewers.delete(viewer);
    else viewer.write(message);
  }
}

export async function createPreviewServer() {
  const sessions = new Map<string, PreviewSession>();
  let origin = '';

  async function sessionFor(directory: string) {
    const root = await realpath(directory);
    if (!(await stat(root)).isDirectory()) throw new Error('工作区目录不存在。');
    const existing = [...sessions].find(([, value]) => value.root === root);
    if (existing) return { token: existing[0], session: existing[1] };
    const token = randomBytes(32).toString('base64url');
    const session: PreviewSession = { root, viewers: new Set() };
    sessions.set(token, session);
    return { token, session };
  }

  const server = createServer(async (request, response) => {
    response.setHeader('X-Content-Type-Options', 'nosniff');
    response.setHeader('Referrer-Policy', 'no-referrer');
    response.setHeader('Cache-Control', 'no-store');
    response.setHeader('Content-Security-Policy', "default-src 'self'; script-src 'self' 'wasm-unsafe-eval'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self' data: blob:; worker-src 'self' blob:; connect-src 'self' blob:; object-src 'none'; frame-src 'none'; base-uri 'none'; frame-ancestors 'none'");
    const json = (status: number, value: unknown) => {
      response.writeHead(status, { 'Content-Type': 'application/json; charset=utf-8' });
      response.end(JSON.stringify(value));
    };
    try {
      if (`http://${request.headers.host}` !== origin ||
          (request.headers.origin && request.headers.origin !== origin)) {
        json(403, { error: '仅允许本机预览页面访问。' });
        return;
      }
      if (request.method !== 'GET') {
        response.setHeader('Allow', 'GET');
        json(405, { error: '预览服务只支持读取。' });
        return;
      }
      const url = new URL(request.url ?? '/', origin);
      if (!url.pathname.startsWith('/api/')) {
        const asset = assets[url.pathname === '/' ? '/index.html' : url.pathname];
        if (!asset) { json(404, { error: '页面不存在。' }); return; }
        response.writeHead(200, { 'Content-Type': asset.type });
        response.end(Buffer.from(asset.base64, 'base64'));
        return;
      }
      // EventSource cannot attach an Authorization header. Only this route accepts a query token.
      const token = request.headers.authorization?.replace(/^Bearer /, '') ??
        (url.pathname === '/api/events' ? url.searchParams.get('token') : undefined);
      const session = token ? sessions.get(token) : undefined;
      if (!session) { json(401, { error: '预览连接已失效。请从 Paseo 重新打开。' }); return; }
      if (url.pathname === '/api/session') {
        json(200, { workspaceName: path.basename(session.root), selected: session.selected, theme: session.theme });
        return;
      }
      if (url.pathname === '/api/events') {
        response.writeHead(200, { 'Content-Type': 'text/event-stream; charset=utf-8', Connection: 'keep-alive' });
        session.viewers.add(response);
        response.write(': connected\n\n');
        const heartbeat = setInterval(() => response.write(': heartbeat\n\n'), 25_000);
        heartbeat.unref();
        response.on('close', () => {
          clearInterval(heartbeat);
          session.viewers.delete(response);
        });
        return;
      }
      if (!['/api/files', '/api/file', '/api/stat'].includes(url.pathname)) {
        json(404, { error: '接口不存在。' }); return;
      }
      const relative = url.searchParams.get('path') ?? '';
      if (url.pathname === '/api/files') {
        json(200, await listEntries(session.root, relative));
        return;
      }
      const { filename, info, mime } = await previewFile(session.root, relative);
      if (url.pathname === '/api/stat') {
        json(200, { size: info.size, modified: info.mtimeMs });
        return;
      }
      response.writeHead(200, { 'Content-Type': mime, 'Content-Length': info.size });
      const stream = createReadStream(filename);
      stream.on('error', () => response.destroy());
      response.on('close', () => stream.destroy());
      stream.pipe(response);
    } catch (error) {
      const problem = error as Error & { code?: string; status?: number };
      const status = problem.status ?? (problem.code === 'ENOENT' ? 404 : problem.code === 'EACCES' ? 403 : 400);
      json(status, { error: problem.status ? problem.message : status === 404 ? '文件已移动或删除。请刷新列表。' : status === 403 ? '没有权限读取此文件。' : '无法读取文件或文件夹。' });
    }
  });
  await new Promise<void>((resolve, reject) => {
    server.once('error', reject);
    server.listen(0, '127.0.0.1', () => { server.off('error', reject); resolve(); });
  });
  const address = server.address();
  if (!address || typeof address === 'string') throw new Error('无法启动本机预览服务。');
  origin = `http://127.0.0.1:${address.port}`;
  return {
    async openWorkspace(directory: string, options: { embed?: boolean; theme?: PreviewTheme } = {}) {
      const { token, session } = await sessionFor(directory);
      if (options.theme) {
        session.theme = options.theme;
        broadcast(session, 'theme', options.theme);
      }
      return `${origin}/#token=${token}${options.embed ? '&embed=1' : ''}`;
    },
    async list(directory: string, relative: string) {
      const { session } = await sessionFor(directory);
      const { entries } = await listEntries(session.root, relative);
      return { entries };
    },
    async select(directory: string, relative: string) {
      const { token, session } = await sessionFor(directory);
      const { filename } = await previewFile(session.root, relative);
      session.selected = { path: path.relative(session.root, filename).split(path.sep).join('/'), seq: (session.selected?.seq ?? 0) + 1 };
      broadcast(session, 'select', session.selected);
      return { url: `${origin}/#token=${token}&embed=1`, viewerConnected: session.viewers.size > 0 };
    },
    async close() {
      for (const session of sessions.values()) {
        for (const viewer of session.viewers) viewer.end();
      }
      sessions.clear();
      server.closeAllConnections();
      await new Promise<void>((resolve, reject) => server.close(error => error ? reject(error) : resolve()));
    },
  };
}
