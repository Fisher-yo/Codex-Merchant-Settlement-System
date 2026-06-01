#!/bin/zsh
set -e

SCRIPT_DIR="${0:A:h}"
cd "$SCRIPT_DIR"

if ! command -v pwsh >/dev/null 2>&1; then
  osascript -e 'display dialog "请先安装 PowerShell for macOS，然后再打开周结算操作台。\n\n推荐安装命令：brew install --cask powershell" buttons {"好"} default button "好"' >/dev/null
  exit 1
fi

pwsh -NoProfile -ExecutionPolicy Bypass -File "$SCRIPT_DIR/run-settlement-console-macos.ps1"

echo ""
echo "周结算操作台已关闭。按回车退出。"
read
