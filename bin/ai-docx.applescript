#!/usr/bin/env osascript
# Double-clickable AI helper: asks for a .docx file and a natural-language
# instruction, runs `oc ai`, and opens the result in WPS Office.
#
# Install:
#   1. Save this file as an Application via Script Editor (or just double-click
#      this .applescript and choose "Run" once to allow osascript).
#   2. Or create a Finder Quick Action that runs this script (see README).

use AppleScript version "2.4"
use scripting additions

set projectRoot to "/Users/dcc/deepseek-harness-workspece/office-copilot"
set ocBin to projectRoot & "/apps/cli/dist/index.js"
set envFile to projectRoot & "/.env"

set docxFile to choose file with prompt "选择要修改的 .docx 文件：" of type {"org.openxmlformats.wordprocessingml.document", "com.microsoft.word.doc", "public.data"}
set docxPath to POSIX path of docxFile

set instruction to text returned of ¬
    (display dialog "告诉助手怎么改这份文档：" default answer "" buttons {"取消", "开始"} default button "开始")

if instruction is "" then
    display notification "没有输入指令，已取消。" with title "Office 智能助手"
    return
end if

-- Load .env so LLM_API_KEY is present; oc ai -y 会自动确认。
set shellCmd to "set -a; [ -f '" & envFile & "' ] && . '" & envFile & "'; set +a; " & ¬
    "cd '" & projectRoot & "' && node '" & ocBin & "' ai '" & ¬
    my shq(docxPath) & "' " & my shq(instruction) & " -y 2>&1; echo \"\n[exit:$?]\""

set output to do shell script shellCmd

display dialog "修改完成，正在用 WPS 打开。" buttons {"好"} default button "好" with title "Office 智能助手"

do shell script "open -a wpsoffice " & my shq(docxPath)

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
