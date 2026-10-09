#!/usr/bin/env bash
# install-mac-integration.sh
# 把"AI 改文档"安装成 macOS 原生体验：
#   1. OfficeCopilot.app 复制到 /Applications 并注册 office-copilot:// URL Scheme
#   2. AI改文档.app 复制到 /Applications
#   3. 打印在 Finder/WPS 右键和全局快捷键里启用的步骤
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
APP_SUPPORT="$HOME/Library/Services"

mkdir -p /Applications
mkdir -p "$APP_SUPPORT"

echo "==> 安装 OfficeCopilot.app 到 /Applications"
rm -rf "/Applications/OfficeCopilot.app"
cp -R "$SCRIPT_DIR/OfficeCopilot.app" /Applications/
codesign --force --deep --sign - /Applications/OfficeCopilot.app >/dev/null 2>&1 || true
/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f /Applications/OfficeCopilot.app
echo "    ✓ office-copilot:// 已注册"

echo "==> 安装 AI改文档.app 到 /Applications"
rm -rf "/Applications/AI改文档.app"
if [[ -d "$SCRIPT_DIR/AI改文档.app" ]]; then
  cp -R "$SCRIPT_DIR/AI改文档.app" /Applications/
  codesign --force --deep --sign - "/Applications/AI改文档.app" >/dev/null 2>&1 || true
  echo "    ✓ 双击即可使用"
else
  echo "    ! AI改文档.app 未找到，跳过"
fi

echo ""
echo "================================================================"
echo "✅ 基础安装完成。剩下两步需要你在图形界面里点一下（macOS 不允许脚本代劳）"
echo ""
echo "【步骤 A】Finder/WPS 右键菜单 → 快速操作"
echo "  1. 打开 /Applications/Automator.app"
echo "  2. 菜单 File → Open（⌘O），选择本目录下的 'AI改文档.workflow'"
echo "     路径: $SCRIPT_DIR/AI改文档.workflow"
echo "  3. 右上角点 'Play' 一次让系统信任，然后 ⌘S 保存（名字保持 'AI改文档'）"
echo "  4. 系统设置 → 键盘 → 键盘快捷键 → 服务 → 常规，找到 'AI改文档' 勾上"
echo "     之后任何 .docx 在 Finder/WPS 右键 → 快速操作 → AI改文档"
echo ""
echo "【步骤 B】给它绑个全局快捷键（可选但推荐）"
echo "  同一个快捷键设置面板，双击 'AI改文档' 右边的 'none'，按 ⌃⌥⌘W"
echo "  在 WPS 写文档时按下快捷键就会弹窗问指令。"
echo ""
echo "【步骤 C】文档内链接触发（可选）"
echo "  任何 .docx 里插入一个超链接，地址写："
echo "    office-copilot://ai?file=/绝对路径/xxx.docx&instruction=审核这份合同"
echo "  在 WPS/Word 里点这个链接就会启动助手。"
echo "================================================================"
