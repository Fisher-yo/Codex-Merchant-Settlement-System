# 开发日志

这里记录我整理这个项目的过程。它不是只给别人看的说明书，也是给未来的自己看的学习轨迹。

## 2026-06-01

### 今天做了什么

- 优化 Windows 版周结算操作台界面，补充桌面快捷方式图标支持。
- 精简日常使用流程，只保留新建目录、导入总对账单、生成结算结果和打开结果目录等高频入口。
- 新增 macOS 版入口 `启动周结算操作台.command`。
- 新增 macOS 终端菜单脚本 `run-settlement-console-macos.ps1`，复用现有拆分和汇总逻辑。
- 新增 `MACOS使用说明.md`，记录 MacBook 上第一次安装 PowerShell、授权启动文件和日常操作步骤。
- 调整拆分脚本中的 xlsx 临时目录路径拼接，让生成 Excel 文件的逻辑同时适配 Windows 和 macOS。
- 更新 `README.md`，同时说明 Windows 和 macOS 两个版本的启动方式。

### 学到的点

- Windows Forms 依赖 Windows 桌面环境，macOS 上不能直接运行同一套 GUI。
- PowerShell Core 可以跨平台运行脚本，适合先做一个稳定的 macOS 终端菜单版。
- 跨平台脚本里不要把 `\` 写死进路径，优先使用 `Join-Path`。
- Mac 终端里拖入文件时，路径里的空格可能会以 `\ ` 形式出现，需要在脚本里做一次处理。

### 验证记录

- 对 macOS 新增启动脚本做了 PowerShell 语法检查。
- 对原有拆分脚本和新建目录脚本做了 PowerShell 语法检查。
- 确认 GitHub 远端为 `https://github.com/Fisher-yo/Codex-Merchant-Settlement-System.git`。

### 下一步计划

- 在真正的 macOS 设备上双击 `启动周结算操作台.command` 做一次端到端试运行。
- 用一份脱敏总对账单验证 Windows 和 macOS 生成结果是否一致。
- 如果 macOS 终端菜单使用顺手，再考虑是否需要做原生图形界面。

## 2026-05-31

### 今天做了什么

- 把本地目录初始化成 Git 仓库。
- 创建 `.gitignore`，排除不适合进入版本管理的文件。
- 创建 `.gitattributes`，让 Git 更清楚地区分文本文件和二进制文件。
- 暂存项目文件，准备进行第一次提交。
- 新增 `README.md`、`DEVLOG.md` 和 `CHANGELOG.md`，开始用 Markdown 记录项目。

### 学到的 Git 概念

- Git 仓库：一个被 Git 管理的项目目录。
- 暂存区：准备放进下一次提交的文件列表。
- 提交：把当前项目状态保存成一个版本记录。
- `.gitignore`：告诉 Git 哪些文件不需要追踪。
- `.gitattributes`：告诉 Git 某些文件应该按文本还是二进制处理。

### 遇到的问题

在普通用户目录 `C:\Users\fuyu0` 执行 Git 命令时，出现了：

```text
fatal: not in a git directory
```

原因是当时没有进入项目目录。Git 命令需要在仓库目录里执行，或者使用全局配置。

正确做法：

```cmd
D:
cd \codex\fuyu-codex\codex-list
```

然后再执行：

```cmd
git status
git commit -m "Initial project import"
```
