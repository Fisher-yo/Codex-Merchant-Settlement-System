# macOS 使用说明

这个版本给 MacBook 使用，入口是：

```text
启动周结算操作台.command
```

## 第一次使用

1. 安装 PowerShell for macOS。

   如果你已经装了 Homebrew，可以在终端运行：

   ```bash
   brew install --cask powershell
   ```

2. 让启动文件可双击打开。

   在本项目根目录打开终端，运行：

   ```bash
   chmod +x 启动周结算操作台.command
   ```

3. 双击 `启动周结算操作台.command`。

## 日常流程

1. 选择 `1`，确认年份和对账日期。
2. 选择 `2`，准备本周结算目录。
3. 选择 `3`，把后台总对账单 `.xlsx` 文件拖进终端窗口后回车。
4. 选择 `4`，生成结算结果。
5. 用 `5` 到 `8` 打开对应目录查看结果。

## 说明

- Windows 版入口仍然保留，macOS 版不会影响原来的 `.exe`、`.cmd` 和 Windows 图形操作台。
- macOS 版是终端菜单，因为原 Windows 图形界面依赖 Windows Forms，不能直接在 macOS 运行。
- 生成逻辑复用原来的结算脚本，输出目录和表格规则保持一致。
