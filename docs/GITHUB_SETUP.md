# GitHub 发布设置

目标仓库名称是 `paseo-office-viewer`，可见性为公开。仓库保存插件源码和测试代码。依赖、构建产物、本地工具、测试输出和实际办公文件不进入 Git。

## 登录 GitHub

在项目根目录打开 PowerShell。此电脑已配置 Git 的提交者姓名和邮箱。提交者身份不等于 GitHub 登录。

本机准备的 GitHub CLI 位于 `.tools\gh\bin\gh.exe`。该目录已被 Git 忽略。其他电脑可以从 [GitHub CLI 官网](https://cli.github.com/)安装 `gh`，并将下面命令中的路径替换为 `gh`。

```powershell
& .\.tools\gh\bin\gh.exe auth login --hostname github.com --git-protocol https --web
& .\.tools\gh\bin\gh.exe auth status
& .\.tools\gh\bin\gh.exe api user --jq .login
```

第一条命令显示一次性验证码和 GitHub 登录地址。本人在浏览器完成登录和授权即可。不要把密码或访问令牌写进项目文件。

## 首次发布

先完成本地构建和测试。首次创建本地仓库时执行：

```powershell
git init -b main
git status --short
git add .
git diff --cached --stat
git commit -m "Release Office Viewer v1.0.0"
```

确认暂存区只包含计划公开的代码和文档。随后创建公开仓库并推送当前提交：

```powershell
& .\.tools\gh\bin\gh.exe repo create paseo-office-viewer --public --source . --remote origin --push
& .\.tools\gh\bin\gh.exe repo view --json nameWithOwner,url,visibility
git status --short
```

如果仓库已经创建或已经设置 `origin`，不要重复执行 `repo create`。用 `git remote -v` 检查现有地址，再推送。

后续更新：

```powershell
git add <需要发布的文件>
git commit -m "说明本次修改"
git push
```

公开仓库不自动授予开源许可。添加许可证前，应确认许可证选择和第三方依赖要求。

## 设置本机 PATH

PATH 是终端查找命令的目录列表。本项目的脚本只把项目 `bin` 和 Paseo 的 `resources\bin` 加入当前用户 PATH，不修改系统 PATH。重复执行不会重复添加目录。

```powershell
# 先显示计划添加的目录，不修改设置。
& .\scripts\setup-path.ps1 -WhatIf

# 写入当前用户 PATH。
& .\scripts\setup-path.ps1
```

Paseo 默认目录为 `$env:LOCALAPPDATA\Programs\Paseo\resources\bin`。自定义安装位置可以用 `-PaseoBin '实际目录'` 指定。项目 `bin` 目录和 Paseo 目录均须存在。

设置后重新打开终端。已运行的终端应用可能需要重启，才能继承更新后的 PATH。移动项目目录后，需要更新 PATH 中的旧目录。

官方说明：[GitHub CLI 登录](https://cli.github.com/manual/gh_auth_login)、[从本地项目创建仓库](https://cli.github.com/manual/gh_repo_create)。
