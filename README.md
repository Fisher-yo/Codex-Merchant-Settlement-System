# 周结算操作台

这是一个用于商家周结算的本地工具，保留日常使用必需能力，并提供 Windows 和 macOS 两套入口。

## 保留功能

- 新建本周结算目录
- 导入后台总对账单
- 按供应商拆分商家对账单
- 生成财务汇总表和商家确认台账
- 合并手工补充订单
- 输出异常待处理文件

## Windows 使用

双击根目录中的：

```text
启动周结算操作台.cmd
```

需要桌面入口时，双击：

```text
创建桌面快捷方式.cmd
```

## macOS 使用

第一次使用请先安装 PowerShell for macOS：

```bash
brew install --cask powershell
chmod +x 启动周结算操作台.command
```

之后双击根目录中的：

```text
启动周结算操作台.command
```

更多说明见：

```text
MACOS使用说明.md
```

## 保留文件

- `周结算操作台.exe`：Windows 无控制台窗口启动入口
- `启动周结算操作台.cmd`：Windows 备用启动入口
- `启动周结算操作台.command`：macOS 终端菜单入口
- `run-settlement-console.ps1`：Windows 启动脚本
- `run-settlement-console-macos.ps1`：macOS 启动脚本
- `workspace/04_商家周结算`：操作台、拆分脚本和必要模板
