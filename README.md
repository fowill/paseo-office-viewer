# Paseo Office Viewer

在 Paseo 桌面版中只读预览 PDF、Word、Excel 和 PowerPoint。本插件在当前电脑启动预览服务，并在 Paseo 工作区浏览器中打开文件。文档不会发送到第三方预览服务。

![Excel 预览示例，使用合成测试数据](docs/screenshots/excel.png)

## V1.0 支持范围

| 格式 | 支持的操作 | 边界 |
| --- | --- | --- |
| `.pdf` | 翻页、缩放、选择文字 | 扫描件没有 OCR |
| `.docx` | 文字、表格、图片、缩放 | 复杂分页、浮动对象、公式可能与 Word 不同 |
| `.xlsx` | 工作表切换、合并单元格、数值格式、单元格定位、公式查看 | 使用已保存的公式结果，不重算；不显示图表、图片、条件格式；不保证完整字体样式 |
| `.pptx` | 静态幻灯片、翻页、缩放 | 动画、复杂母版和字体不保证与 PowerPoint 一致 |

只预览，不编辑原文件。单个文件上限为 100 MiB。旧版 `.doc`、`.xls`、`.ppt` 和密码保护文件不在 V1.0 支持范围内。

## 安装

需要 Node.js 22.14 或更新版本、npm，以及 Paseo 桌面版。支持的 Paseo 版本范围为 `>=0.11.4 <0.13.0`；实际桌面验证环境见 [测试记录](docs/TESTING.md)。当前只支持客户端与 Paseo 主机在同一台电脑上运行。手机和远程主机预览暂未接入。

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
2. 按 `Ctrl+K`，选择 `Office Viewer · 文档预览`。
3. 在插件面板中点击“打开文档预览”。
4. 在预览页左侧选择文件或进入子文件夹。
5. 文件修改后，点击“重新读取”读取新内容。刷新列表可以发现新增文件。

PDF、PPT 提供翻页和缩放。Excel 提供工作表标签、分段浏览和单元格定位；点击单元格可查看值和公式。V1.0 不接管 Paseo 原生文件树的双击行为。

## 独立运行与 Windows PATH

在源码目录运行：

```powershell
npm start -- --root "D:\YourProject"
```

在浏览器打开终端显示的 `Open:` 地址。使用 `Ctrl+C` 停止服务。

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
