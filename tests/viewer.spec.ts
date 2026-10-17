import { test, expect, type Page } from '@playwright/test';
import assert from 'node:assert/strict';
import { copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as XLSX from 'xlsx';
import { createPreviewServer } from '../server/http';

const project = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const fixtures = path.join(project, 'tests', 'fixtures');
const cache = path.join(project, '.cache');
const fixtureNames = ['示例文档.pdf', '示例报告.docx', '示例数据.xlsx', '示例演示.pptx'];
let workspace: string;
let url: string;
let server: Awaited<ReturnType<typeof createPreviewServer>>;
let pageErrors: string[];
let externalRequests: string[];

async function openFile(page: Page, name: string) {
  await page.getByRole('button', { name: `预览 ${name}`, exact: true }).click();
  await expect(page.locator('#preview-status')).toContainText('已加载');
  await expect(page.locator('#preview-error')).toBeHidden();
  await expect(page.locator('#file-name')).toHaveText(name);
}

async function expectPopulatedCanvas(page: Page) {
  await expect(page.locator('#preview canvas')).toHaveCount(1);
  const pixels = await page.locator('#preview canvas').evaluate((node: HTMLCanvasElement) => {
    const data = node.getContext('2d')!.getImageData(0, 0, node.width, node.height).data;
    let dark = 0;
    for (let i = 0; i < data.length; i += 4) {
      if (data[i + 3] && data[i] < 200 && data[i + 1] < 200 && data[i + 2] < 200) dark++;
    }
    return { width: node.width, height: node.height, dark };
  });
  expect(pixels.width).toBeGreaterThan(500);
  expect(pixels.height).toBeGreaterThan(300);
  expect(pixels.dark).toBeGreaterThan(500);
}

test.beforeAll(async () => {
  await mkdir(cache, { recursive: true });
  workspace = await mkdtemp(path.join(cache, 'viewer-tests-'));
  await Promise.all(fixtureNames.map(name => copyFile(path.join(fixtures, name), path.join(workspace, name))));
  await mkdir(path.join(workspace, '子目录'));
  await copyFile(path.join(fixtures, '示例文档.pdf'), path.join(workspace, '子目录', '摘要.pdf'));
  await Promise.all(['pdf', 'docx', 'xlsx', 'pptx'].map(extension => writeFile(path.join(workspace, `损坏.${extension}`), 'This is not an Office or PDF document.')));

  // A large sheet exercises both row and column windows without expanding the DOM.
  const large = XLSX.utils.book_new();
  const data = Array.from({ length: 240 }, (_, row) => Array.from({ length: 40 }, (_, column) => `R${row + 1}C${column + 1}`));
  const sheet = XLSX.utils.aoa_to_sheet(data);
  sheet['!merges'] = [{ s: { r: 98, c: 27 }, e: { r: 101, c: 32 } }];
  XLSX.utils.book_append_sheet(large, sheet, '分页数据');
  XLSX.utils.book_append_sheet(large, XLSX.utils.aoa_to_sheet([]), '空白工作表');
  await writeFile(path.join(workspace, '大表与空表.xlsx'), XLSX.write(large, { type: 'buffer', bookType: 'xlsx' }));

  server = await createPreviewServer();
  url = await server.openWorkspace(workspace);
});

test.afterAll(async () => {
  if (server) await server.close();
  if (workspace) {
    assert.equal(path.dirname(path.resolve(workspace)), cache);
    assert.match(path.basename(workspace), /^viewer-tests-/);
    await rm(workspace, { recursive: true, force: true });
  }
});

test.beforeEach(async ({ page }) => {
  pageErrors = [];
  externalRequests = [];
  page.on('pageerror', error => pageErrors.push(error.message));
  page.on('request', request => {
    const requestUrl = request.url();
    if (!requestUrl.startsWith(new URL(url).origin + '/') && !requestUrl.startsWith('data:') && !requestUrl.startsWith('blob:')) externalRequests.push(requestUrl);
  });
  await page.goto(url);
  await expect(page.getByRole('button', { name: '预览 示例文档.pdf', exact: true })).toBeVisible();
});

test.afterEach(async () => {
  expect(pageErrors, 'No uncaught browser errors').toEqual([]);
  expect(externalRequests, 'Previews must not contact external services').toEqual([]);
});

test('PDF renders Chinese, selectable text, image, page controls, and zoom', async ({ page }) => {
  await openFile(page, '示例文档.pdf');
  await expect(page.locator('.textLayer')).toContainText('本地文档预览测试');
  await expect(page.locator('.textLayer')).toContainText('Amount: -1234.50');
  await expectPopulatedCanvas(page);
  await page.locator('.textLayer').evaluate(node => {
    const range = document.createRange();
    range.selectNodeContents(node);
    const selection = window.getSelection()!;
    selection.removeAllRanges();
    selection.addRange(range);
  });
  expect(await page.evaluate(() => window.getSelection()!.toString())).toContain('本地文档预览测试');
  await page.locator('#next-page').click();
  await expect(page.locator('#preview-status')).toContainText('第 2 / 2 页');
  await expect(page.locator('.textLayer')).toContainText('第二页中文内容');
  await expect(page.locator('#next-page')).toBeDisabled();
  await page.locator('#zoom-in').click();
  await expect(page.locator('#zoom-value')).toHaveText('110%');
  await expect(page.locator('#zoom-in')).toBeEnabled();
  await page.locator('#page-number').fill('1');
  await page.locator('#page-number').press('Enter');
  await expect(page.locator('#preview-status')).toContainText('第 1 / 2 页');
});

test('Word renders both pages, table, image, and inline formatting', async ({ page }) => {
  await openFile(page, '示例报告.docx');
  await expect(page.locator('section.docx')).toHaveCount(2);
  await expect(page.locator('section.docx').first()).toContainText('本地文档预览测试');
  await expect(page.locator('#preview table')).toContainText('-1,234.50');
  await expect(page.locator('section.docx').nth(1)).toContainText('第二页图片测试');
  await expect(page.locator('#preview img')).toHaveCount(1);
  await expect.poll(() => page.locator('#preview img').evaluate((image: HTMLImageElement) => image.complete && image.naturalWidth > 0)).toBe(true);
  await expect(page.getByText('粗体', { exact: true })).toHaveCSS('font-weight', '700');
  await page.locator('#zoom-out').click();
  await expect(page.locator('#zoom-value')).toHaveText('90%');
  await expect(page.locator('.docx-host')).toHaveCSS('zoom', '0.9');
});

test('Excel retains displayed values, dates, merges, cached formulas, and sheet switching', async ({ page }) => {
  await openFile(page, '示例数据.xlsx');
  await expect(page.locator('td[data-address="A1"]')).toHaveAttribute('colspan', '4');
  await expect(page.locator('td[data-address="B3"]')).toHaveText('1,200.50');
  await expect(page.locator('td[data-address="B4"]')).toHaveText('-34.75');
  await expect(page.locator('td[data-address="C3"]')).toHaveText('12.5%');
  await expect(page.locator('td[data-address="D3"]')).toHaveText('2026-10-17');
  await page.locator('td[data-address="B5"]').click();
  await expect(page.locator('#cell-value')).toContainText('公式：=SUM(B3:B4)');
  await expect(page.locator('#cell-value')).toContainText('已保存的值：1,165.75');
  await page.getByRole('tab', { name: '说明', exact: true }).click();
  await expect(page.locator('td[data-address="B5"]')).toHaveText('Sheet switching works');
  await page.locator('#zoom-in').click();
  await expect(page.locator('.sheet-grid')).toHaveCSS('zoom', '1.1');
});

test('large Excel uses bounded row and column windows, clipped merges, cell navigation, and blank sheets', async ({ page }) => {
  await openFile(page, '大表与空表.xlsx');
  expect(await page.locator('.sheet-grid td').count()).toBeLessThanOrEqual(3000);
  await expect(page.locator('#row-range')).toHaveText('1–100 / 240 行');
  await expect(page.locator('#column-range')).toHaveText('A–AD / AN 列');
  await expect(page.locator('td[data-address="AB99"]')).toHaveAttribute('rowspan', '2');
  await expect(page.locator('td[data-address="AB99"]')).toHaveAttribute('colspan', '3');
  await page.locator('#next-rows').click();
  await expect(page.locator('#row-range')).toHaveText('101–200 / 240 行');
  await expect(page.locator('td[data-address="AB99"]')).toHaveText('R99C28');
  await page.locator('#next-columns').click();
  await expect(page.locator('#column-range')).toHaveText('AE–AN / AN 列');
  await expect(page.locator('td[data-address="AB99"]')).toHaveAttribute('colspan', '3');
  await page.locator('#cell-address').fill('AN240');
  await page.locator('#go-to-cell').click();
  await expect(page.locator('#selected-cell')).toHaveText('AN240');
  await expect(page.locator('#cell-value')).toHaveText('R240C40');
  await expect(page.locator('#row-range')).toHaveText('201–240 / 240 行');
  await page.getByRole('tab', { name: '空白工作表', exact: true }).click();
  await expect(page.locator('#preview')).toHaveText('此工作表为空。');
  await expect(page.locator('#next-rows')).toBeDisabled();
  await expect(page.locator('#next-columns')).toBeDisabled();
});

test('PowerPoint renders nonblank slides, navigation, zoom, and reload', async ({ page }) => {
  await openFile(page, '示例演示.pptx');
  await expectPopulatedCanvas(page);
  const firstPixels = await page.locator('#preview canvas').evaluate((canvas: HTMLCanvasElement) => canvas.toDataURL());
  await page.locator('#next-page').click();
  await expect(page.locator('#preview-status')).toContainText('第 2 / 2 页');
  await expectPopulatedCanvas(page);
  const secondPixels = await page.locator('#preview canvas').evaluate((canvas: HTMLCanvasElement) => canvas.toDataURL());
  expect(secondPixels).not.toBe(firstPixels);
  await page.locator('#zoom-in').click();
  await expect(page.locator('#zoom-in')).toBeEnabled();
  await expect(page.locator('#preview canvas')).toHaveAttribute('width', '1056');
  await page.locator('#reload-file').click();
  await expect(page.locator('#preview-status')).toContainText('第 1 / 2 页');
  await expect(page.locator('#zoom-value')).toHaveText('100%');
  await expectPopulatedCanvas(page);
});

test('corrupt files report errors and a valid file can be opened afterwards', async ({ page }) => {
  for (const extension of ['pdf', 'docx', 'xlsx', 'pptx']) {
    await page.getByRole('button', { name: `预览 损坏.${extension}`, exact: true }).click();
    await expect(page.locator('#preview-error')).toBeVisible();
    await expect(page.locator('#preview-error')).toContainText('无法预览');
    await expect(page.locator('#preview-status')).toBeEmpty();
    await expect(page.locator('#toolbar')).toBeHidden();
  }
  await openFile(page, '示例文档.pdf');
});

test('rapid file changes keep the final selected document', async ({ page }) => {
  for (const name of ['示例演示.pptx', '示例报告.docx', '示例演示.pptx', '示例文档.pdf', '示例数据.xlsx', '示例文档.pdf']) {
    await page.getByRole('button', { name: `预览 ${name}`, exact: true }).dispatchEvent('click');
  }
  await expect(page.locator('#preview-status')).toContainText('第 1 / 2 页');
  await expect(page.locator('#file-name')).toHaveText('示例文档.pdf');
  await expect(page.locator('.textLayer')).toContainText('本地文档预览测试');
  await expect(page.locator('#preview-error')).toBeHidden();
  await page.locator('#next-page').click();
  await expect(page.locator('#preview-status')).toContainText('第 2 / 2 页');
});

test('session survives page reload and file filtering, folders, and refresh work', async ({ page }) => {
  expect(new URL(page.url()).hash).toBe('');
  await page.reload();
  await page.locator('#file-filter').fill('示例数据');
  await expect(page.locator('.file-row')).toHaveCount(1);
  await page.locator('#file-filter').fill('');
  await page.getByRole('button', { name: '打开文件夹 子目录', exact: true }).click();
  await openFile(page, '摘要.pdf');
  await page.getByRole('button', { name: '项目根目录', exact: true }).click();
  await copyFile(path.join(fixtures, '示例文档.pdf'), path.join(workspace, '新增文件.pdf'));
  await page.locator('#refresh-files').click();
  await openFile(page, '新增文件.pdf');
});

test('reloading a selected file reads its updated contents', async ({ page }) => {
  await openFile(page, '示例数据.xlsx');
  const filename = path.join(workspace, '示例数据.xlsx');
  const original = await readFile(filename);
  const workbook = XLSX.read(original, { type: 'buffer' });
  workbook.Sheets['结果表'].B3 = { t: 'n', v: 9876.5, z: '#,##0.00' };
  await writeFile(filename, XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }));
  try {
    await page.locator('#reload-file').click();
    await expect(page.locator('td[data-address="B3"]')).toHaveText('9,876.50');
  } finally { await writeFile(filename, original); }
});

test('mobile layout keeps controls reachable and documents scroll within the preview', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await openFile(page, '示例演示.pptx');
  await expect(page.locator('#reload-file')).toBeVisible();
  await expect(page.locator('#next-page')).toBeVisible();
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBe(390);
  expect(await page.locator('#preview').evaluate(node => node.scrollWidth > node.clientWidth)).toBe(true);
  await page.locator('#next-page').click();
  await expect(page.locator('#preview-status')).toContainText('第 2 / 2 页');
});

test('missing authentication gives an actionable message', async ({ page }) => {
  await page.evaluate(() => sessionStorage.clear());
  await page.goto(new URL(url).origin);
  await expect(page.locator('#preview-error')).toContainText('缺少本地预览会话');
  await expect(page.locator('#list-status')).toContainText('请从 Paseo 插件打开此页面');
});
