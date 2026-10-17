# Paseo Office Viewer

在 Paseo 桌面版中只读预览 PDF、Word、Excel 和 PowerPoint。点击工作区头部的 Office Viewer 按钮，打开使用 Paseo 原生组件的紧凑文件树，再点击文件预览。预览页使用 Paseo 工作区浏览器，并跟随宿主主题。文档不会发送到第三方预览服务。

V1.1 将文件导航移到插件面板。嵌入预览页不再显示第二套侧栏、品牌区和启动卡片。文件树按需展开目录；已打开的预览页随文件选择更新。插件不拦截聊天消息或 Markdown 中的文件链接。

## V1.1 支持范围

| 格式 | 支持的操作 | 边界 |
| --- | --- | --- |
| `.pdf` | 翻页、缩放、选择文字 | 扫描件没有 OCR |
| `.docx` | 文字、表格、图片、缩放 | 复杂分页、浮动对象、公式可能与 Word 不同 |
| `.xlsx` | 工作表切换、合并单元格、数值格式、单元格定位、公式查看 | 使用已保存的公式结果，不重算；不显示图表、图片、条件格式；不保证完整字体样式 |
| `.pptx` | 静态幻灯片、翻页、缩放 | 动画、复杂母版和字体不保证与 PowerPoint 一致 |

只预览，不编辑原文件。单个文件上限为 100 MiB。旧版 `.doc`、`.xls`、`.ppt` 和密码保护文件不在支持范围内。

## 安装

需要 Node.js 22.14 或更新版本、npm，以及 Paseo 桌面版。支持的 Paseo 版本范围为 `>=0.11.4 <0.13.0`；验证范围见 [测试记录](docs/TESTING.md)。文档预览要求客户端与 Paseo 主机在同一台电脑上运行。手机端只提供文件树和桌面预览提示。

在 Paseo 设置中启用 Plugins，然后安装：

```powershell
paseo plugin add git:fowill/paseo-office-viewer
```

首次安装会按 `package-lock.json` 从 npm 和 SheetJS 官方源安装预览库，并构建本地资源。使用时不需要下载预览库，也不要求安装 Microsoft Office。

也可以从源码安装：

```powershell
git clone https://github.com/fowill/paseo-office-viewer.git
cd paseo-office-viewer
npm ci
npm run build
npm run typecheck
paseo plugin install "$((Get-Location).Path)"
```

## 在 Paseo 中使用

1. 打开一个本机工作区。
2. 点击工作区头部的 Office Viewer 图标，打开文件树面板。
3. 展开文件夹，点击 PDF、DOCX、XLSX 或 PPTX 文件。
4. 点击另一个文件，已连接的预览页会更新内容。
5. 文件修改后，点击预览页的“刷新”重读内容。点击文件树头部的刷新按钮，更新目录列表。

插件请求在 Paseo 的 `explorer` 位置打开面板，并同时注册 `workspace` 位置作为入口。实际位置和面板切换方式由 Paseo 决定。`Ctrl+K` 中的 `Office Viewer · 文档预览` 命令仍可使用，但日常操作不需要它。

PDF、Word 和 PPT 默认适应可用宽度，也可手动缩放。Excel 提供底部工作表标签、分段浏览和单元格定位；点击单元格可查看值和公式。缩放 Excel 时保留已有表格。工具栏跟随 Paseo 主题，文档纸面和 Excel 网格保持白色。Excel 表格以外的空区域使用宿主背景色。

### Paseo API 的边界

- 预览仍在 Paseo 内置浏览器中。切到聊天等其他视图后，再点文件会在后台更新内容；插件无法主动聚焦已有的浏览器标签页。需要时点击文件树头部的“打开预览”。宿主可能复用或新开标签页。
- 插件通过 SSE（服务器推送连接）判断预览页是否仍在连接。首次打开后有 10 秒保护期，避免连续点击反复打开页面；收到已连接状态后立即解除保护。若在这 10 秒内立即关闭预览，再点文件可能不会自动重开，可点击文件树头部的“打开预览”。保护期结束后，无连接时会自动重新打开。这是连接状态推断，不是 Paseo 提供的标签页关闭事件。
- 手机端和非 Electron 客户端没有 `openBrowser` 接口。预览服务只监听本机 `127.0.0.1`。这些客户端可以查看文件树，但本版不提供文档预览。

Paseo 原生窗口中的按钮位置、面板布局和点击流程仍需人工验收。桌面操作工具未取得窗口访问授权，不能用自动测试结果代替这项检查。

## 独立运行与 Windows PATH

在源码目录运行：

```powershell
npm start -- --root "D:\YourProject"
```

在浏览器打开终端显示的 `Open:` 地址。独立模式保留紧凑的左侧文件列表和面包屑导航，并跟随系统明暗主题。使用 `Ctrl+C` 停止服务。

若希望在任意目录使用命令，运行一次：

```powershell
.\scripts\setup-path.ps1
```

此脚本仅向当前用户的 PATH 添加项目 `bin` 和已有的 Paseo `resources\bin`。不覆盖原 PATH。重开终端后可运行：

```powershell
paseo --version
paseo-office-viewer --root "D:\YourProject"
```

自定义安装路径可使用 `-PaseoBin "C:\YourPath\Paseo\resources\bin"`。独立命令依赖源码目录的开发依赖，因此须先运行 `npm ci`。

## 开发与测试

```powershell
npm ci
npm run check
```

Windows 测试使用已安装的 Microsoft Edge；其他平台需先运行 `npx playwright install chromium`。测试文件由脚本生成，全为合成数据，不读取个人文档。测试产物不提交到 Git。

修改预览代码后，执行：

```powershell
npm run build
npm run typecheck
paseo plugin reload office-viewer
```

重新从插件入口打开预览页。插件重载后，旧页面的临时连接会失效。

## 文件与网络访问

预览服务只监听 `127.0.0.1`，使用随机端口和临时访问令牌。文件读取受当前工作区目录限制；目录连接和符号链接不能越界读取。文件列表按需读取单层目录，不进行全盘扫描。页面禁止加载外部文档资源。

Paseo 插件本身是可信任的本机代码，具有主机用户权限。公开仓库不包含账号令牌、个人文件或测试缓存。

## 项目结构

- `client/`、`shared/`、`index.*`：Paseo 面板与 RPC。
- `server/http.ts`：本机文件访问与预览服务。
- `web/`：四种格式的显示界面。
- `scripts/build.mjs`：将网页与依赖资源打包到服务端模块。
- `scripts/create-fixtures.mjs`、`tests/`：可重复的合成文件和测试。

GitHub 初始化步骤见 [GITHUB_SETUP.md](docs/GITHUB_SETUP.md)。第三方组件及许可证见 [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md)。

V1.1 的交互决定、实现边界和延期内容见 [DESIGN_V1.1.md](docs/DESIGN_V1.1.md)。
