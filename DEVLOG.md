# 开发日志

这里记录我作为 Git 新手整理这个项目的过程。它不是只给别人看的说明书，也是给未来的自己看的学习轨迹。

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

### 下一步计划

- 配置 Git 用户名和邮箱。
- 完成第一次提交。
- 在 GitHub 创建仓库。
- 优先考虑私有仓库，避免业务资料被公开。
- 学习如何把本地仓库推送到 GitHub。

## GitHub 同步准备

第一次同步前，我需要确认几件事：

- 是否已经配置 Git 用户名和邮箱。
- 是否已经完成本地第一次提交。
- GitHub 仓库是否应该设为私有。
- 是否有合同、证件、发票、结算截图等敏感文件需要脱敏或移除。

建议命令记录：

```cmd
git config --global user.name "fuyu0"
git config --global user.email "你的邮箱"
```

完成提交：

```cmd
git commit -m "Initial project import"
```

连接 GitHub 远程仓库：

```cmd
git remote add origin https://github.com/你的用户名/你的仓库名.git
git push -u origin main
```

