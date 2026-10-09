#!/usr/bin/env osascript
# Global hot-key helper: prompts for an instruction, then runs `oc ai` on the
# document currently open in WPS (by reading WPS's window title to resolve
# the file path). Falls back to a file picker.

on run
    set projectRoot to "/Users/dcc/deepseek-harness-workspece/office-copilot"
    set ocBin to projectRoot & "/apps/cli/dist/index.js"
    set envFile to projectRoot & "/.env"

    set docxPath to missing value

    tell application "System Events"
        if exists process "wpsoffice" then
            tell process "wpsoffice"
                set frontmost to true
                try
                    set winTitle to value of attribute "AXTitle" of window 1
                on error
                    set winTitle to ""
                end try
            end tell
        end if
    end tell

    -- WPS 窗口标题一般是 "文件名.docx - WPS Office"，但 macOS 不提供可靠路径。
    -- 我们让用户确认 / 选择文件，以避免歧义。
    set dlgText to "当前将修改 WPS 最前方窗口的文档。" & linefeed & "如果不是该文件，点"选取文件"选择。"
    if winTitle is not "" then
        set dlgText to dlgText & linefeed & linefeed & "窗口标题：" & winTitle
    end if

    set chooseResult to display dialog dlgText buttons {"选取文件", "取消", "就是它"} default button "就是它" with title "Office 智能助手"
    if button returned of chooseResult is "取消" then return
    if button returned of chooseResult is "选取文件" then
        set picked to choose file with prompt "选择要修改的 .docx：" of type {"org.openxmlformats.wordprocessingml.document", "com.microsoft.word.doc", "public.data"}
        set docxPath to POSIX path of picked
    end if

    if docxPath is missing value then
        -- Try to resolve the window title to a file under common locations.
        set fname to ""
        if winTitle contains " - " then
            set AppleScript's text item delimiters to " - "
            set parts to text items of winTitle
            set AppleScript's text item delimiters to ""
            set fname to first item of parts
        else
            set fname to winTitle
        end if
        if fname does not end with ".docx" then set fname to fname & ".docx"

        -- Spotlight lookup
        try
            set found to do shell script "mdfind -name " & my shq(fname) & " | head -1"
            if found is not "" then set docxPath to found
        end try
    end if

    if docxPath is missing value then
        display dialog "无法识别当前 WPS 文档，请用"选取文件"再试。" buttons {"好"} default button "好"
        return
    end if

    set instruction to text returned of ¬
        (display dialog "告诉助手怎么改：" default answer "" buttons {"取消", "开始"} default button "开始")
    if instruction is "" then return

    set shellCmd to ¬
        "set -a; [ -f '" & envFile & "' ] && . '" & envFile & "'; set +a; " & ¬
        "cd '" & projectRoot & "' && node '" & ocBin & "' ai " & my shq(docxPath) & " " & my shq(instruction) & " -y 2>&1; echo \"\n[oc-exit:$?]\""
    set output to do shell script shellCmd

    display dialog "已写入：" & docxPath buttons {"好"} default button "好" giving up after 5
    tell application "wpsoffice" to activate
end run

on shq(s)
    set s2 to ""
    repeat with c in s
        set ch to c as string
        if ch is "'" then set s2 to s2 & "'\\''" else set s2 to s2 & ch
    end repeat
    return "'" & s2 & "'"
end shq
