-- URL dispatcher for OfficeCopilot.app.
--
-- Supports:
--   office-copilot://ai?file=<urlencoded path>&instruction=<urlencoded text>
--   office-copilot://ai?instruction=<text>    (uses WPS frontmost doc)
--   office-copilot://pick                     (file picker)
--
-- WPS does not register custom URL schemes itself, but hyperlinks inside a
-- Word/WPS document DO open through macOS Launch Services, so you can keep a
-- "AI 改文档" link on the first page of your template and click it.

on open location this_URL
    set projectRoot to "/Users/dcc/deepseek-harness-workspece/office-copilot"
    set ocBin to projectRoot & "/apps/cli/dist/index.js"
    set envFile to projectRoot & "/.env"

    set AppleScript's text item delimiters to "office-copilot://"
    set rest to text item 2 of this_URL
    set AppleScript's text item delimiters to "?"
    set action to text item 1 of rest
    set qs to text item 2 of rest
    set AppleScript's text item delimiters to ""

    set params to my parseQuery(qs)
    set file_ to my mapGet(params, "file")
    set instruction to my mapGet(params, "instruction")

    if file_ is not "" then
        set docxPath to my urldecode(file_)
    else
        try
            set picked to choose file with prompt "选择要修改的 .docx：" of type {"org.openxmlformats.wordprocessingml.document", "com.microsoft.word.doc", "public.data"}
            set docxPath to POSIX path of picked
        on error
            return
        end try
    end if

    if instruction is "" then
        set q to display dialog "告诉助手怎么改：" default answer "" buttons {"取消", "开始"} default button "开始"
        set instruction to text returned of q
    end if

    set shellCmd to ¬
        "set -a; [ -f '" & envFile & "' ] && . '" & envFile & "'; set +a; " & ¬
        "cd '" & projectRoot & "' && " & ¬
        "node '" & ocBin & "' ai " & my shq(my urldecode(docxPath)) & " " & my shq(my urldecode(instruction)) & " -y 2>&1; echo \"\n[oc-exit:$?]\""
    set output to do shell script shellCmd

    display dialog output buttons {"复制", "好"} default button "好"
    do shell script "open -a wpsoffice " & my shq(my urldecode(docxPath))
end open location

on run
    display dialog "OfficeCopilot 正在运行。\n\n在文档里使用 office-copilot:// 链接触发。" buttons {"好"} default button "好"
end run

on parseQuery(qs)
    set map to ""
    set AppleScript's text item delimiters to "&"
    set parts to text items of qs
    set AppleScript's text item delimiters to ""
    set out to {}
    repeat with p in parts
        set AppleScript's text item delimiters to "="
        set kv to text items of p
        set AppleScript's text item delimiters to ""
        if (count of kv) ≥ 2 then
            set end of out to {my urldecode(item 1 of kv), my urldecode(item 2 of kv)}
        end if
    end repeat
    return out
end parseQuery

on mapGet(m, k)
    repeat with pair in m
        if item 1 of pair is k then return item 2 of pair
    end repeat
    return ""
end mapGet

on urldecode(s)
    set s to my replaceText(s, "+", " ")
    set s to my replaceText(s, "%20", " ")
    set s to my replaceText(s, "%2F", "/")
    set s to my replaceText(s, "%3A", ":")
    set s to my replaceText(s, "%2C", ",")
    set s to my replaceText(s, "%E2%80%9C", "“")
    set s to my replaceText(s, "%E2%80%9D", "”")
    return s
end urldecode

on replaceText(s, f, r)
    set AppleScript's text item delimiters to f
    set parts to text items of s
    set AppleScript's text item delimiters to r
    set s to parts as text
    set AppleScript's text item delimiters to ""
    return s
end replaceText

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
