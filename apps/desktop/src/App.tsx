import { useEditor } from "@tiptap/react";
import { EDITOR_EXTENSIONS } from "./editor/extensions/editorExtensions.js";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { FileDown } from "lucide-react";
import { EditorContent } from "@tiptap/react";
import { Toolbar, type ToolbarAction } from "./components/Toolbar.js";
import { Outline } from "./components/Outline.js";
import { PageSetupDialog, geometryFor } from "./components/PageSetupDialog.js";
import { TemplatePickerDialog } from "./components/TemplatePickerDialog.js";
import { TocDialog } from "./components/TocDialog.js";
import { CaptionDialog } from "./components/CaptionDialog.js";
import {
  insertToc,
  watchTocPageNumbers,
  type TocOptions,
  type TocReject,
} from "./editor/toc.js";
import { insertCaption, type CaptionOptions, type CaptionReject } from "./editor/captions.js";
import type {
  FootnoteBodies,
  LinkStyle,
  PaneStyle,
  ResolvedStyle,
  SectionGeometry,
  StyleEdit,
} from "@office-copilot/docx-io";
import { docStyleSheet } from "./editor/docStyleSheet.js";
import { StylePane } from "./components/StylePane.js";
import { FootnotePane } from "./components/FootnotePane.js";
import { AIPanel } from "./ai/AIPanel.js";
import { StatusBar } from "./components/StatusBar.js";
import { FindBar } from "./components/FindBar.js";
import { TableBubbleMenu } from "./components/TableBubbleMenu.js";
import { TableContextMenu } from "./components/TableContextMenu.js";
import { EditorContextMenu } from "./components/EditorContextMenu.js";
import { AISuggestionBubble } from "./ai/AISuggestionBubble.js";
import { PageChrome } from "./components/PageChrome.js";
import { SettingsModal } from "./components/SettingsModal.js";
import { HistoryPanel } from "./components/HistoryPanel.js";
import { Welcome } from "./components/Welcome.js";
import { EdgeResizer } from "./components/EdgeResizer.js";
import { DialogHost, confirmDialog } from "./components/ui/Dialog.js";
import { ToastHost, toast } from "./components/ui/Toast.js";
import { usePanelResize } from "./hooks/usePanelResize.js";
import { CommandPalette } from "./components/CommandPalette.js";
import type { CommandContext } from "./commands/registry.js";
import { useUiStore } from "./uiStore.js";
import {
  documentTitle,
  listTemplateRuns,
  revertTemplateRuns,
  runTemplate,
  runTypography,
  type DocLevelState,
} from "./editor/formatDocument.js";
import { buildFooterDoc, buildHeaderDoc } from "./editor/templateChrome.js";
import { mmToPx, type DocTemplate } from "./editor/templates.js";
import { buildPrintHtml } from "./editor/printHtml.js";
import { renderHeaderFooterHtml } from "./editor/renderHeaderFooter.js";

// The exact shape of a ProseMirror document accepted by editor.commands.setContent.
// We intentionally type loosely; TipTap accepts JSON content directly.
type PMDoc = Record<string, unknown>;

/** A4 at 96dpi — what the editor renders until 页面设置 supplies a section. */
const FALLBACK_PAGE = { pageWidth: 794, pageHeight: 1123, header: 48, footer: 48 };

const STARTER_CONTENT: PMDoc = {
  type: "doc",
  content: [
    {
      type: "heading",
      attrs: { level: 1 },
      content: [{ type: "text", text: "欢迎使用 Office Copilot" }],
    },
    {
      type: "paragraph",
      content: [
        {
          type: "text",
          text: "这是一个独立的桌面文档编辑器，",
        },
        {
          type: "text",
          marks: [{ type: "bold" }],
          text: "保存出的 .docx 可以直接在 Word 和 WPS 中打开。",
        },
      ],
    },
    {
      type: "paragraph",
      content: [
        { type: "text", text: "试试顶部工具栏：加粗、标题、列表、表格；或者选中一段文字点右侧 AI 面板进行扩写/润色。" },
      ],
    },
  ],
};

export function App() {
  const [filePath, setFilePath] = useState<string | null>(null);
  const [dirty, setDirty] = useState(false);
  const [started, setStarted] = useState(false);
  const [showFind, setShowFind] = useState(false);
  const [outlineCollapsed, setOutlineCollapsed] = useState(false);
  const [stylePaneOpen, setStylePaneOpen] = useState(false);
  // 脚注窗格：正文与各部件条目的对应关系跟着文档走，换文档就整份换掉。
  const [footnotePaneOpen, setFootnotePaneOpen] = useState(false);
  const [footnoteBodies, setFootnoteBodies] = useState<FootnoteBodies>({});
  // 样式窗格的数据随文档走：打开 / 新建时从主进程带回来的 styles.xml 清单换一次。
  const [docStyles, setDocStyles] = useState<PaneStyle[]>([]);
  // 同一份 styles.xml 的另一面：每条样式合并出来的可见格式，拼成文档作用域的样式表。
  const [docStyleFormats, setDocStyleFormats] = useState<ResolvedStyle[]>([]);
  const docStyleCss = useMemo(
    () => docStyleSheet(docStyleFormats),
    [docStyleFormats],
  );
  const [pageSetupOpen, setPageSetupOpen] = useState(false);
  const [tocOpen, setTocOpen] = useState(false);
  const [captionOpen, setCaptionOpen] = useState(false);
  const [templateOpen, setTemplateOpen] = useState(false);
  const [aiCollapsed, setAiCollapsed] = useState(false);
  const [showAI, setShowAI] = useState(true);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [editorEl, setEditorEl] = useState<HTMLElement | null>(null);
  const [headerDoc, setHeaderDoc] = useState<Record<string, unknown> | null>(null);
  const [footerDoc, setFooterDoc] = useState<Record<string, unknown> | null>(null);
  const [section, setSection] = useState<SectionGeometry | null>(null);
  // The 美化排版 template's hyperlink colours, in the same document-level slot
  // as 页面设置: renderer through these two variables, saved into run rPr.
  const [linkStyle, setLinkStyle] = useState<LinkStyle | null>(null);
  // 「还原排版」按钮的可用态：本文档是否还留有未撤销的美化排版记录。
  const [canRevert, setCanRevert] = useState(false);

  /** 重读修改历史，刷新「还原排版」可用态。docKey 为文件路径或未命名文档名。 */
  const refreshCanRevert = useCallback(async (docKey: string) => {
    const plan = await listTemplateRuns(docKey);
    setCanRevert(plan.runs.length > 0);
  }, []);

  // 打开/另存/新建都会换 docKey，历史记录跟着文档走，所以按 docKey 重查一次。
  useEffect(() => {
    void refreshCanRevert(filePath ?? "未命名文档");
  }, [filePath, refreshCanRevert]);

  useEffect(() => {
    const open = () => setSettingsOpen(true);
    const openHistory = () => setHistoryOpen(true);
    window.addEventListener("open-settings", open);
    window.addEventListener("open-history", openHistory);
    return () => {
      window.removeEventListener("open-settings", open);
      window.removeEventListener("open-history", openHistory);
    };
  }, []);

  const editor = useEditor({
    extensions: EDITOR_EXTENSIONS,
    content: STARTER_CONTENT,
    editorProps: { attributes: { spellcheck: "false" } },
  });

  useEffect(() => {
    if (!editor) return;
    setEditorEl(editor.view.dom);
    const onUpdate = () => {
      setDirty(true);
      setStarted(true);
    };
    editor.on("update", onUpdate);
    return () => {
      editor.off("update", onUpdate);
    };
  }, [editor]);

  // Keep a ref so menu handlers always see the current filePath.
  const filePathRef = useRef<string | null>(null);  useEffect(() => {
    filePathRef.current = filePath;
  }, [filePath]);
  const dirtyRef = useRef(false);
  useEffect(() => {
    dirtyRef.current = dirty;
  }, [dirty]);

  // 样式窗格攒下的「新建 / 更新样式」编辑，保存时一次性交给主进程落进 styles.xml。
  // 用 ref 不用 state：没有 UI 读它，而菜单、快捷键、工具栏拿到的 handler 必须始终
  // 看到最新一份（handleSave 的依赖里塞 state 会让每个菜单项跟着重建）。
  const styleEditsRef = useRef<StyleEdit[]>([]);

  // Reusable file actions (used by both menu and toolbar).
  const handleNew = useCallback(() => {
    if (!editor) return;
    editor.commands.setContent({
      type: "doc",
      content: [
        { type: "heading", attrs: { level: 1 }, content: [{ type: "text", text: "无标题" }] },
        { type: "paragraph" },
      ],
    });
    setHeaderDoc(null);
    setFooterDoc(null);
    setSection(null);
    setLinkStyle(null);
    setDocStyles([]);
    setDocStyleFormats([]);
    setFootnoteBodies({});
    // 换文档 = 攒下的样式编辑作废：它们指向的是上一份文档的 styles.xml。
    styleEditsRef.current = [];
    setFilePath(null);
    setDirty(false);
    setStarted(true);
  }, [editor]);

  const handleOpen = useCallback(async () => {
    if (!editor) return;
    const r = await window.officeCopilot.openFile();
    if (!r.ok) {
      if (!r.canceled && r.message) toast(r.message, "error");
      return;
    }
    editor.commands.setContent(r.doc as never);
    setHeaderDoc((r.header as never) ?? null);
    setFooterDoc((r.footer as never) ?? null);
    setSection((r.section as never) ?? null);
    setLinkStyle(null);
    setDocStyles((r.styles as PaneStyle[] | undefined) ?? []);
    setDocStyleFormats((r.styleFormats as ResolvedStyle[] | undefined) ?? []);
    setFootnoteBodies((r.footnotes as FootnoteBodies | undefined) ?? {});
    styleEditsRef.current = [];
    setFilePath(r.path ?? null);
    setDirty(false);
    setStarted(true);
    if (r.path) window.officeCopilot.addRecent(r.path);
    if (r.warnings && r.warnings.length) toast(r.warnings.join("\n"));
  }, [editor]);

  const openPath = useCallback(
    async (p: string, guardDirty = false) => {
      if (!editor) return;
      if (guardDirty && dirtyRef.current) {
        const ok = await confirmDialog({
          title: "打开新文档？",
          message: "当前文档有未保存的修改，打开后这些修改将丢失。",
          okLabel: "放弃并打开",
          danger: true,
        });
        if (!ok) return;
      }
      const r = await window.officeCopilot.readDocx(p);
      if (!r.ok) {
        if (r.message) toast(r.message, "error");
        return;
      }
      editor.commands.setContent(r.doc as never);
      setHeaderDoc((r.header as never) ?? null);
      setFooterDoc((r.footer as never) ?? null);
      setSection((r.section as never) ?? null);
      setLinkStyle(null);
      setDocStyles((r.styles as PaneStyle[] | undefined) ?? []);
      setDocStyleFormats((r.styleFormats as ResolvedStyle[] | undefined) ?? []);
      setFootnoteBodies((r.footnotes as FootnoteBodies | undefined) ?? {});
      styleEditsRef.current = [];
      setFilePath(r.path ?? p);
      setDirty(false);
      setStarted(true);
      window.officeCopilot.addRecent(p);
      if (r.warnings && r.warnings.length) toast(r.warnings.join("\n"));
    },
    [editor],
  );

  // Drag a .docx from Finder/Explorer anywhere onto the window to open it.
  const dragDepthRef = useRef(0);
  const [dropActive, setDropActive] = useState(false);
  const hasFiles = (e: React.DragEvent) =>
    Array.from(e.dataTransfer.types).includes("Files");
  const onDragEnter = useCallback((e: React.DragEvent) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepthRef.current += 1;
    setDropActive(true);
  }, []);
  const onDragOver = useCallback((e: React.DragEvent) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  }, []);
  const onDragLeave = useCallback(() => {
    dragDepthRef.current = Math.max(0, dragDepthRef.current - 1);
    if (dragDepthRef.current === 0) setDropActive(false);
  }, []);
  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      dragDepthRef.current = 0;
      setDropActive(false);
      const file = Array.from(e.dataTransfer.files).find((f) =>
        /\.(docx|json)$/i.test(f.name),
      );
      if (!file) {
        toast("只能打开 .docx 文档。", "error");
        return;
      }
      const p = window.officeCopilot.pathForFile(file);
      if (p) void openPath(p, true);
    },
    [openPath],
  );

  /**
   * 收下一条样式窗格交出来的编辑：同一条样式只留最新一份，然后向主进程要一次预览投影。
   *
   * 预览走的是保存时同一个 `applyStyleEdits`（只是不落盘），所以窗格列表与注入编辑器
   * 的样式表就是保存之后 Word 里的那份定义，渲染层不必自己再实现一遍样式合并。
   *
   * @param edit 窗格组装好的编辑，`styleId` 已定稿（新建的 id 也由 `nextStyleId` 分好了）。
   */
  const submitStyleEdit = useCallback(async (edit: StyleEdit) => {
    const merged = [
      ...styleEditsRef.current.filter((previous) => previous.styleId !== edit.styleId),
      edit,
    ];
    styleEditsRef.current = merged;
    const r = await window.officeCopilot.previewStyles({
      path: filePathRef.current,
      styles: merged,
    });
    if (!r.ok) {
      toast(r.message ?? "样式预览失败。", "error");
      return;
    }
    setDocStyles((r.styles as PaneStyle[] | undefined) ?? []);
    setDocStyleFormats((r.styleFormats as ResolvedStyle[] | undefined) ?? []);
  }, []);

  const handleSave = useCallback(async () => {
    if (!editor) return;
    const json = editor.getJSON() as PMDoc;
    const r = await window.officeCopilot.saveFile({
      path: filePathRef.current,
      doc: json,
      header: headerDoc as never,
      footer: footerDoc as never,
      section: section as never,
      link: linkStyle as never,
      styles: styleEditsRef.current,
      footnotes: footnoteBodies,
    });
    if (r.ok && r.path) {
      setFilePath(r.path);
      setDirty(false);
      window.officeCopilot.addRecent(r.path);
    } else if (!r.canceled && r.message) {
      toast(r.message, "error");
    }
  }, [editor, headerDoc, footerDoc, section, linkStyle, footnoteBodies]);

  const handleSaveAs = useCallback(async () => {
    if (!editor) return;
    const json = editor.getJSON() as PMDoc;
    const defaultName = filePathRef.current
      ? filePathRef.current.split(/[\\/]/).pop()
      : "document.docx";
    const r = await window.officeCopilot.saveFileAs({
      doc: json,
      defaultName,
      header: headerDoc as never,
      footer: footerDoc as never,
      section: section as never,
      link: linkStyle as never,
      styles: styleEditsRef.current,
      footnotes: footnoteBodies,
    });
    if (r.ok && r.path) {
      setFilePath(r.path);
      setDirty(false);
      window.officeCopilot.addRecent(r.path);
    } else if (!r.canceled && r.message) {
      toast(r.message, "error");
    }
  }, [editor, headerDoc, footerDoc, section, linkStyle, footnoteBodies]);

  const handlePrint = useCallback(async () => {
    if (!editor) return;
    const g = section ?? {
      pageWidth: 794, pageHeight: 1123,
      marginTop: 96, marginRight: 96, marginBottom: 96, marginLeft: 96,
      header: 48, footer: 48, contentWidth: 602, contentHeight: 931,
    };
    const html = buildPrintHtml({
      bodyHtml: editor.getHTML(),
      headerHtml: renderHeaderFooterHtml(headerDoc as never, 1, 1),
      footerHtml: renderHeaderFooterHtml(footerDoc as never, 1, 1),
      pageWidthPx: g.pageWidth,
      pageHeightPx: g.pageHeight,
      marginTopPx: g.marginTop,
      marginRightPx: g.marginRight,
      marginBottomPx: g.marginBottom,
      marginLeftPx: g.marginLeft,
      headerPx: g.header,
      footerPx: g.footer,
      link: linkStyle,
    });
    const defaultName = filePath
      ? filePath.split(/[\\/]/).pop()?.replace(/\.docx$/i, "") + ".pdf"
      : "document.pdf";
    const r = await window.officeCopilot.printToPdf(html, defaultName);
    if (r.ok) toast(`已导出：${r.path}`, "success");
    else if (!r.canceled && r.message) toast(r.message, "error");
  }, [editor, section, headerDoc, footerDoc, filePath, linkStyle]);

  const lastTemplateId = useUiStore((s) => s.lastTemplateId);
  const setLastTemplate = useUiStore((s) => s.setLastTemplate);

  /** 一键还原：撤销本文档全部未撤销的美化排版，并把页面级设置写回套模板前的留底。 */
  const revertBeautify = useCallback(() => {
    if (!editor) return;
    void (async () => {
      const docKey = filePath ?? "未命名文档";
      const plan = await listTemplateRuns(docKey);
      if (!plan.runs.length) {
        setCanRevert(false);
        toast("没有可还原的美化排版记录。");
        return;
      }
      const ok = await confirmDialog({
        title: "还原排版",
        message: `将撤销 ${plan.runs.length} 次美化排版对正文的改动${
          plan.docLevel ? "，并恢复套模板前的页边距、页眉页脚与超链接配色" : ""
        }。`,
        okLabel: "还原",
        cancelLabel: "取消",
      });
      if (!ok) return;
      const result = await revertTemplateRuns(editor, docKey);
      const restored = result.docLevel;
      if (restored) {
        setSection(restored.page as never);
        setHeaderDoc(restored.header as never);
        setFooterDoc(restored.footer as never);
        setLinkStyle(restored.link);
      }
      setLastTemplate(undefined);
      setDirty(true);
      const parts: string[] = [`已撤销 ${result.reverted} 次美化排版`];
      parts.push(restored ? "页面设置已复原" : "页边距/页眉页脚未留底，保持现状");
      if (result.failed.length)
        parts.push(`${result.failed.length} 处保留现状（${result.failed[0]}）`);
      toast(parts.join("；") + "。", result.reverted ? "success" : "error");
      void refreshCanRevert(docKey);
    })();
  }, [editor, filePath, setLastTemplate, refreshCanRevert]);

  /** 上一次目录页码观察器的停止函数：再插一次目录前先掐掉旧的循环。 */
  const tocWatch = useRef<(() => void) | null>(null);
  useEffect(() => () => tocWatch.current?.(), []);

  /**
   * 「插入目录」对话框点确定：落 op、写修改历史，然后盯分页把页码刷对。
   *
   * @param opts 对话框选项；`contentWidthPx` 是 CSS px，磅的换算只在 docx-io 边界发生。
   * @returns 成功返回 null（对话框据此关闭）；失败返回引擎的拒绝码，中文文案在对话框里拼。
   */
  const handleInsertToc = useCallback(
    async (opts: TocOptions): Promise<TocReject | null> => {
      if (!editor) return "ops-failed";
      const result = await insertToc(editor, filePath ?? "未命名文档", opts);
      if (!result.ok) return result.reason ?? "ops-failed";
      setDirty(true);
      toast(`已插入目录（${result.entries} 条）。`, "success");
      // 目录自己占了版面，标题会被往后推：分页收敛之前读到的页码都是旧的。
      tocWatch.current?.();
      tocWatch.current = watchTocPageNumbers(editor);
      return null;
    },
    [editor, filePath],
  );

  /**
   * 「插入题注」对话框点确定：落 op、写修改历史。编号不用这里操心——
   * `CaptionNumbering` 插件会在事务落地后顺手把后面同标签的题注一起重排。
   *
   * @param opts 对话框选项。
   * @returns 成功返回 null（对话框据此关闭）；失败返回引擎的拒绝码。
   */
  const handleInsertCaption = useCallback(
    async (opts: CaptionOptions): Promise<CaptionReject | null> => {
      if (!editor) return "ops-failed";
      const result = await insertCaption(editor, filePath ?? "未命名文档", opts);
      if (!result.ok) return result.reason ?? "ops-failed";
      setDirty(true);
      toast(`已插入题注（${opts.label.trim()} ${result.number}）。`, "success");
      return null;
    },
    [editor, filePath],
  );

  const handleToolbarAction = useCallback(
    (action: ToolbarAction) => {
      switch (action) {
        case "new":
          handleNew();
          break;
        case "open":
          handleOpen();
          break;
        case "save":
          handleSave();
          break;
        case "save-as":
          handleSaveAs();
          break;
        case "print":
          handlePrint();
          break;
        case "find-replace":
          setShowFind(true);
          break;
        case "page-setup":
          setPageSetupOpen(true);
          break;
        case "insert-toc":
          setTocOpen(true);
          break;
        case "insert-caption":
          setCaptionOpen(true);
          break;
        case "styles":
          setStylePaneOpen((v) => !v);
          break;
        case "footnotes":
          setFootnotePaneOpen((v) => !v);
          break;
        case "beautify":
          setTemplateOpen(true);
          break;
        case "revert-beautify":
          revertBeautify();
          break;
      }
    },
    [handleNew, handleOpen, handleSave, handleSaveAs, handlePrint, revertBeautify],
  );

  const runFormatDocument = useCallback(() => {
    if (!editor) return;
    void runTypography(editor, filePath ?? "未命名文档").then((r) => {
      if (!r.ok) {
        toast(r.message ?? "排版失败。", r.message?.includes("无需") ? undefined : "error");
        return;
      }
      setDirty(true);
      const skipped = r.skipped.length ? `；${r.skipped.length} 段保留原样（${r.skipped[0]}）` : "";
      toast(`已规范化 ${r.touched} 个段落，可在修改历史里查看 diff 或单独撤销${skipped}`, "success");
    });
  }, [editor, filePath]);

  const applyTemplate = useCallback(
    (t: DocTemplate) => {
      if (!editor) return;
      setTemplateOpen(false);
      void (async () => {
        const docKey = filePath ?? "未命名文档";
        // 页面级设置随后会被模板覆盖，先把当前值留底进这条历史记录。
        const docLevel: DocLevelState = {
          page: section,
          header: headerDoc,
          footer: footerDoc,
          link: linkStyle,
        };
        const r = await runTemplate(t, editor, docKey, docLevel);
        const alreadyTyped = !r.ok && Boolean(r.message?.includes("无需"));
        if (!r.ok && !alreadyTyped) {
          toast(r.message ?? "美化排版失败。", "error");
          return;
        }

        // 页边距与页眉页脚和「页面设置」一样属于文档级设置，跟随 App state，
        // 不进修改历史；正文样式那部分才是逐条可撤销的。
        if (t.page) {
          const base = section ?? FALLBACK_PAGE;
          setSection(
            geometryFor(
              base.pageWidth,
              base.pageHeight,
              {
                top: mmToPx(t.page.top),
                right: mmToPx(t.page.right),
                bottom: mmToPx(t.page.bottom),
                left: mmToPx(t.page.left),
              },
              base.header,
              base.footer,
            ) as never,
          );
          setDirty(true);
        }

        const nextFooter = buildFooterDoc(t);
        const nextHeader = buildHeaderDoc(t, documentTitle(editor));
        let keptChrome = false;
        if ((nextFooter || nextHeader) && (footerDoc || headerDoc)) {
          const replace = await confirmDialog({
            title: `美化排版 · ${t.label}`,
            message: "该模板将替换现有的页眉/页脚内容（正文样式不受影响）。",
            okLabel: "替换",
            cancelLabel: "保留原有",
          });
          keptChrome = !replace;
        }
        if (!keptChrome) {
          if (nextFooter) setFooterDoc(nextFooter as never);
          if (nextHeader) setHeaderDoc(nextHeader as never);
          if (nextFooter || nextHeader) setDirty(true);
        }

        // 超链接的配色属于文档而不是段落：它不改正文，所以不进修改历史。
        const nextLink = t.link ?? null;
        const linkChanged =
          (nextLink?.color ?? null) !== (linkStyle?.color ?? null) ||
          (nextLink?.underline ?? true) !== (linkStyle?.underline ?? true);
        if (linkChanged) {
          setLinkStyle(nextLink);
          setDirty(true);
        }
        setLastTemplate(t.id);

        const parts: string[] = [
          r.ok
            ? `已按「${t.label}」调整 ${r.touched} 个段落`
            : `正文已符合「${t.label}」`,
        ];
        if (t.page) parts.push("页边距已套用");
        if (nextFooter || nextHeader)
          parts.push(keptChrome ? "页眉页脚保留原样" : "页眉页脚已套用");
        if (linkChanged) parts.push("超链接样式已套用");
        if (r.ok && r.skipped.length)
          parts.push(`${r.skipped.length} 段保留原样（${r.skipped[0]}）`);
        toast(
          `${parts.join("；")}。正文改动可在修改历史里逐条撤销。`,
          "success",
        );
        void refreshCanRevert(docKey);
      })();
    },
    [
      editor,
      filePath,
      section,
      headerDoc,
      footerDoc,
      linkStyle,
      setLastTemplate,
      refreshCanRevert,
    ],
  );

  useEffect(() => {
    if (!window.officeCopilot || !editor) return;
    const subs = [
      window.officeCopilot.onMenu("menu:new", handleNew),
      window.officeCopilot.onMenu("menu:open", handleOpen),
      window.officeCopilot.onMenu("menu:save", handleSave),
      window.officeCopilot.onMenu("menu:save-as", handleSaveAs),
      window.officeCopilot.onMenu("menu:print", handlePrint),
      window.officeCopilot.onMenu("menu:find", () => setShowFind((v) => !v)),
      window.officeCopilot.onMenu("ai:toggle-panel", () => {
        if (!showAI) setShowAI(true);
        else setAiCollapsed((v) => !v);
      }),
      window.officeCopilot.onMenu("outline:toggle", () =>
        setOutlineCollapsed((v) => !v),
      ),
      window.officeCopilot.onMenu("ai:format-document", runFormatDocument),
      window.officeCopilot.onMenu("ai:beautify-document", () =>
        setTemplateOpen(true),
      ),
      window.officeCopilot.onMenu("ai:expand", () =>
        window.dispatchEvent(new CustomEvent("ai:run", { detail: "扩写" })),
      ),
      window.officeCopilot.onMenu("ai:shorten", () =>
        window.dispatchEvent(new CustomEvent("ai:run", { detail: "缩写" })),
      ),
      window.officeCopilot.onMenu("ai:polish", () =>
        window.dispatchEvent(new CustomEvent("ai:run", { detail: "润色" })),
      ),
    ];
    return () => {
      subs.forEach((u) => u());
    };
  }, [
    editor,
    handleNew,
    handleOpen,
    handleSave,
    handleSaveAs,
    handlePrint,
    runFormatDocument,
    showAI,
  ]);

  // The AI panel chip and the command palette are renderer-side, so they
  // dispatch a DOM event rather than going through the menu IPC channel.
  useEffect(() => {
    window.addEventListener("ai:format-document", runFormatDocument);
    return () =>
      window.removeEventListener("ai:format-document", runFormatDocument);
  }, [runFormatDocument]);

  // Open-recent from the application menu.
  useEffect(() => {
    if (!window.officeCopilot?.onOpenRecent || !editor) return;
    const off = window.officeCopilot.onOpenRecent((p) => openPath(p));
    return off;
  }, [editor, openPath]);

  const insertImage = useCallback(async () => {    const r = await window.officeCopilot.pickImage();
    if (r.ok)
      editor
        ?.chain()
        .focus()
        .insertContent({
          type: "image",
          attrs: { src: r.dataUrl, layout: "inline" },
        })
        .run();
  }, [editor]);

  const insertTable = useCallback(() => {
    editor
      ?.chain()
      .focus()
      .insertTable({ rows: 3, cols: 3, withHeaderRow: true })
      .run();
  }, [editor]);

  const wordCount = editor?.storage.characterCount.words() ?? 0;
  const charCount = editor?.storage.characterCount.characters() ?? 0;
  const filename = useMemo(() => {
    if (!filePath) return "未命名文档";
    return filePath.split(/[\\/]/).pop() ?? filePath;
  }, [filePath]);

  const outlineWidth = useUiStore((s) => s.outlineWidth);
  const aiWidth = useUiStore((s) => s.aiWidth);
  const setOutlineWidth = useUiStore((s) => s.setOutlineWidth);
  const setAiWidth = useUiStore((s) => s.setAiWidth);
  const outlineResize = usePanelResize({
    side: "left",
    width: outlineWidth,
    setWidth: setOutlineWidth,
    min: 180,
    max: 480,
    resetTo: 220,
  });
  const aiResize = usePanelResize({
    side: "right",
    width: aiWidth,
    setWidth: setAiWidth,
    min: 280,
    max: 640,
    resetTo: 340,
  });

  const [paletteOpen, setPaletteOpen] = useState(false);
  const setTheme = useUiStore((s) => s.setTheme);
  const commandCtx = useMemo<CommandContext>(
    () => ({
      handleNew,
      handleOpen,
      handleSave: () => void handleSave(),
      handleSaveAs: () => void handleSaveAs(),
      handlePrint: () => void handlePrint(),
      openPageSetup: () => setPageSetupOpen(true),
      openToc: () => setTocOpen(true),
      showFindBar: () => setShowFind(true),
      toggleOutline: () => setOutlineCollapsed((v) => !v),
      toggleAI: () => {
        setShowAI(true);
        setAiCollapsed((v) => !v);
      },
      openHistory: () => setHistoryOpen(true),
      openSettings: () => setSettingsOpen(true),
      runAI: (action) => {
        setShowAI(true);
        window.dispatchEvent(new CustomEvent("ai:run", { detail: action }));
      },
      formatDocument: runFormatDocument,
      openTemplatePicker: () => setTemplateOpen(true),
      setTheme,
    }),
    [
      handleNew,
      handleOpen,
      handleSave,
      handleSaveAs,
      handlePrint,
      runFormatDocument,
      setTheme,
    ],
  );

  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === "k") {
        e.preventDefault();
        setPaletteOpen((v) => !v);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  if (!editor) return <div className="app">启动中…</div>;

  return (
    <div
      className="app"
      onDragEnter={onDragEnter}
      onDragOver={onDragOver}
      onDragLeave={onDragLeave}
      onDrop={onDrop}
    >
      <div className="titlebar">
        <div className="title">
          {dirty ? <span className="dirty-dot" aria-hidden>•</span> : null}
          {filename} — Office Copilot
        </div>
      </div>
      <div
        className="main"
        style={
          {
            "--outline-w": `${outlineWidth}px`,
            "--ai-w": `${aiWidth}px`,
            "--oc-link-color": linkStyle?.color ?? "#0563c1",
            "--oc-link-decoration":
              linkStyle?.underline === false ? "none" : "underline",
          } as React.CSSProperties
        }
      >
        <Outline
          editor={editor}
          collapsed={outlineCollapsed}
          onToggle={() => setOutlineCollapsed((v) => !v)}
        />
        {!outlineCollapsed && (
          <div
            className="resizer"
            title="拖动调整宽度，双击复位"
            onPointerDown={outlineResize.onPointerDown}
            onDoubleClick={outlineResize.onDoubleClick}
          />
        )}
        {stylePaneOpen && (
          <StylePane
            editor={editor}
            styles={docStyles}
            onEdit={(edit) => void submitStyleEdit(edit)}
            onClose={() => setStylePaneOpen(false)}
          />
        )}
        {footnotePaneOpen && editor && (
          <FootnotePane
            editor={editor}
            bodies={footnoteBodies}
            onBodies={(next) => {
              setFootnoteBodies(next);
              setDirty(true);
            }}
            onClose={() => setFootnotePaneOpen(false)}
          />
        )}
        <div className="editor-area">
          {/* 文档作用域的样式表：这份 .docx 的 styles.xml 合并结果，只命中编辑器里的段落。 */}
          {docStyleCss ? <style>{docStyleCss}</style> : null}
          {started ? (
            <Toolbar
              editor={editor}
              onInsertImage={insertImage}
              onInsertTable={insertTable}
              onAction={handleToolbarAction}
              canRevert={canRevert}
              stylesOpen={stylePaneOpen}
              footnotesOpen={footnotePaneOpen}
              dirty={dirty}
            />
          ) : null}
          {showFind && <FindBar editor={editor} onClose={() => setShowFind(false)} />}
          {started ? (
            editor && (
              <PageChrome
                editor={editor}
                headerDoc={headerDoc}
                footerDoc={footerDoc}
                section={section}
                onHeaderChange={setHeaderDoc}
                onFooterChange={setFooterDoc}
              >
                <TableBubbleMenu editor={editor} />
                <AISuggestionBubble
                  editor={editor}
                  docKey={filePath ?? "未命名文档"}
                  onApplied={() => setDirty(true)}
                />
              </PageChrome>
            )
          ) : (
            <Welcome
              onNew={handleNew}
              onOpen={handleOpen}
              onOpenRecent={(p) => openPath(p, true)}
            />
          )}
          <StatusBar wordCount={wordCount} charCount={charCount} editorEl={editorEl} />
        </div>
        {showAI && !aiCollapsed && (
          <div
            className="resizer"
            title="拖动调整宽度，双击复位"
            onPointerDown={aiResize.onPointerDown}
            onDoubleClick={aiResize.onDoubleClick}
          />
        )}
        {showAI && (
          <AIPanel
            editor={editor}
            docKey={filePath ?? "未命名文档"}
            collapsed={aiCollapsed}
            onToggle={() => setAiCollapsed((v) => !v)}
            onOpenHistory={() => setHistoryOpen(true)}
            onChanged={() => setDirty(true)}
          />
        )}
      </div>
      <SettingsModal
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
      />
      <PageSetupDialog
        open={pageSetupOpen}
        section={section as SectionGeometry | null}
        onClose={() => setPageSetupOpen(false)}
        onApply={(s) => {
          setSection(s as never);
          setDirty(true);
        }}
      />
      <TemplatePickerDialog
        open={templateOpen}
        lastTemplateId={lastTemplateId}
        onClose={() => setTemplateOpen(false)}
        onPick={applyTemplate}
      />
      <TocDialog
        open={tocOpen}
        editor={editor}
        contentWidthPx={section?.contentWidth ?? 602}
        onClose={() => setTocOpen(false)}
        onInsert={handleInsertToc}
      />
      <CaptionDialog
        open={captionOpen}
        editor={editor}
        onClose={() => setCaptionOpen(false)}
        onInsert={handleInsertCaption}
      />
      <HistoryPanel
        open={historyOpen}
        onClose={() => setHistoryOpen(false)}
        editor={editor}
        docKey={filePath ?? "未命名文档"}
        onChanged={() => {
          setDirty(true);
          // 历史面板里逐条撤销后，可还原的美化排版可能已经清零。
          void refreshCanRevert(filePath ?? "未命名文档");
        }}
      />
      {dropActive && (
        <div className="drop-overlay">
          <div className="drop-overlay-card">
            <FileDown size={28} />
            松开以打开 .docx
          </div>
        </div>
      )}
      <CommandPalette
        open={paletteOpen}
        onClose={() => setPaletteOpen(false)}
        ctx={commandCtx}
      />
      {/* 右键菜单挂在页面之外：.page-stage 上有 zoom 的 transform，
          放进舞台里 position: fixed 会以缩放后的容器为基准，菜单就跑到鼠标外。 */}
      <EditorContextMenu editor={editor} />
      <TableContextMenu editor={editor} />
      <DialogHost />
      <ToastHost />
      <EdgeResizer />
    </div>
  );
}
