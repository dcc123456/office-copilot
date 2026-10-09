#!/usr/bin/env osascript
# Finder Quick Action: "AI 改文档（Office Copilot）"
#
# 工作流：
#   1. 从 Finder 右键拿到选中的 .docx
#   2. 弹窗让用户输入自然语言指令
#   3. 调用 oc ai -y 直接写入（原地修改，CLI 会自动生成 .bak）
#   4. 如果 WPS 正在前台显示该文件，让它重新加载
#   5. 否则用 WPS 打开
#
# 在 Automator 里：新建 "Quick Action"，工作流选择 "receives files or folders
# in Finder"，拖一个 "Run AppleScript" 动作，粘贴本文件的全部内容，保存为
# "AI 改文档"。之后任何 .docx 右键 → 快速操作 → AI 改文档。

on run {input, parameters}
    if (count of input) is 0 then return
    set docxFile to item 1 of input
    set docxPath to POSIX path of docxFile

    set projectRoot to "/Users/dcc/deepseek-harness-workspece/office-copilot"
    set ocBin to projectRoot & "/apps/cli/dist/index.js"
    set envFile to projectRoot & "/.env"

    tell application "System Events"
        set ext to name extension of (disk item docxPath)
    end tell
    if ext is not "docx" then
        display dialog "只支持 .docx 文件。" buttons {"好"} default button "好" with icon stop
        return
    end if

    set instruction to text returned of ¬
        (display dialog "告诉助手怎么改这份文档：" & linefeed & linefeed & "📄 " & (name of (info for docxFile)) default answer "" buttons {"取消", "开始"} default button "开始")

    if instruction is "" then return

    set shellCmd to ¬
        "set -a; [ -f '" & envFile & "' ] && . '" & envFile & "'; set +a; " & ¬
        "cd '" & projectRoot & "' && " & ¬
        "node '" & ocBin & "' ai " & my shq(docxPath) & " " & my shq(instruction) & " -y 2>&1; " & ¬
        "echo \"\n[oc-exit:$?]\""

    set output to do shell script shellCmd
    set exitCode to last word of (last paragraph of output)

    if exitCode is "0" then
        display notification "已写入：" & (name of (info for docxFile)) with title "Office 智能助手"
        my openOrReloadInWps(docxPath)
    else
        display dialog "修改失败：" & linefeed & linefeed & output buttons {"复制错误", "好"} default button "好" with icon stop
    end if
end run

on openOrReloadInWps(docxPath)
    tell application "System Events"
        set wpsRunning to (exists process "wpsoffice")
    end tell

    if not wpsRunning then
        do shell script "open -a wpsoffice " & my shq(docxPath)
        return
    end if

    -- WPS Mac 没有公开的 AppleScript 字典来"重新加载"。我们先激活它，
    -- 如果它显示的是同一个文件，WPS 会在窗口回到前台时检测到外部改动并提示重载；
    -- 如果不是同一个文件，就用 open 打开新文件。
    tell application "wpsoffice" to activate
    delay 0.4
    do shell script "open -a wpsoffice " & my shq(docxPath)
end openOrReloadInWps

on shq(s)
    set s2 to ""
    repeat with c in s
        set ch to c as string
        if ch is "'" then
            set s2 to s2 & "'\\''"
        else
            set s2 to s2 & ch
        end if
    end repeat
    return "'" & s2 & "'"
end shq
