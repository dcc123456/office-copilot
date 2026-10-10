# Word 功能补齐规格（Word parity spec）

本文件是**实现契约**，不是愿望清单。每节都要能直接翻译成代码与验收条目；任何一节落地时，
对应的实现提交必须引用本节标题。范围来自 `README.md` 第五节的差距表 1–7 名（用户确认全做）。

新增 `docs/` 目录的理由：这份规格要按功能分节、含 OOXML 落点与验收判据，塞进 `README.md`
会把用户手册撑坏，而 `AGENTS.md` 是约束清单不该混入设计。

## 零、全局约束（每一节都受它管）

1. **单位换算只在 `packages/docx-io` 进出口**（AGENTS §2.1）。编辑器里一律 pt / 倍数 / CSS px；
   OOXML 里是 twips（1pt = 20 twips）。因此本规格里所有"写进 `w:*` 属性"的换算都发生在
   `ooxml-pm.ts`，`apps/desktop` 里禁止出现第二份换算。
2. **属性必须显式建模，`rawPpr` 只做兜底**。改动前的问题是"读侧只解析三四个属性，写侧
   `cleanRaw` 把整个 `<w:ind>` 抹掉再按 attrs 重建"，于是**没建模的属性在保存时静默消失**
   ——右缩进就是这么丢的（实测见 §1.1）。新属性走的路线统一是：读侧解析进 attrs → 渲染层用
   `data-*` 呈现 → 写侧从 attrs 生成。
   **已落地**：`cleanRaw` 与整元素重建都删掉了，现在 `mergedElementXml` 做属性级合并（只删
   真正被改值的那对同义属性），`paragraphPprXml` 按 CT_PPr 顺序重建 `<w:pPr>`；暂时无法建模
   的复杂结构（如整段 `w:pBdr`）继续靠 rawPpr 原样透传。
3. **一个功能一个提交**（用户要求 + AGENTS §1.4/1.6）。每个提交的验收按 §7 那张表逐条打分。
4. **新增节点/标记必须注册进 `EDITOR_EXTENSIONS`**（`apps/desktop/src/editor/extensions/editorExtensions.ts`），
   并同步补 `docBlocks.live.test.ts` 真实 schema 用例 + `docx-io` 往返用例（AGENTS §5.3）。
5. **Word 兼容性判据**：每个功能都要用真实 `.docx` 往返验证（生成 → 读回 → 比对 attrs）。
   本机没有 Word/WPS，**"Word 里打开是否长这样"一律标 BLOCKED**，禁止写成已验证。

### 0.1 落点约定：一份 `w:ind` 生成器

现在有两份 pt→twips 的 `w:ind` 写法且行为不一致：`synthesizedFormatXml`（`:1027-1040`，
负 `indentFirst` 会转成 `w:hanging`）与 `pPrXml`（`:1255-1260`，负值原样写进 `w:firstLine`）。
§1 的第一件事就是把它们合并成**一个**函数 `paragraphIndAttrs(attrs): string[]`，两份调用它，
并按 AGENTS §3.2 补一条"两处输出必须一致"的用例。

---

## 一、排名 1：段落对话框全字段

Word 的「段落」对话框给的是：常规（对齐方式 + 大纲级别 + 缩进：左侧/右侧/特殊格式[无/首行/悬挂]）、
缩进正文与对齐方式见 Microsoft 官方
[调整缩进和间距](https://support.microsoft.com/zh-cn/office/%E8%B0%83%E6%95%B4%E7%BC%A9%E8%BF%9B%E5%92%8C%E9%97%B4%E8%B7%9D-36239d98-14c5-411e-a880-1ddf25d65cd6)；
间距（段前/段后/行距）；以及"如果只在 Web 视图查看"之下的换行分组：孤行控制、与下段同页、
段中不分页、段前分页。OOXML 侧对应 `w:pPr` 的子元素，`w:ind` 的属性集是
`w:left`/`w:right`/`w:firstLine`/`w:hanging`（+ 字符单位版 `w:*Chars`，+ 2010 起的 `w:start`/`w:end`
同义写法），单位"twentieths of a point"，见
[Indentation Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.indentation?view=openxml-3.0.1)。

### 1.1 先修保真缺陷（缺陷修复，单独一个提交）

实测（一次性往返探针，`w:ind w:left="300" w:right="200"` 读入再写出）：

- 读侧 `ooxml-pm.ts:437-455` 只取 `w:left`/`w:firstLine`/`w:hanging`，attrs 得到
  `{ indent: 15, rawPpr: '<w:ind w:left="300" w:right="200"/>' }`——**右缩进没有属性承载**。
- 写侧 `cleanRaw`（`:980-990`）把 rawPpr 里的 `<w:ind>` 整个删掉，`synthesizedFormatXml` 重建出
  `<w:ind w:left="300"/>` → **`w:right` 永久丢失**。
- 悬挂缩进往返是正确的（`w:hanging="420"` ↔ `indentFirst: -21` ↔ `w:hanging="420"`），
  但段落面板的数值框写死 `Math.max(0, Number(v))`（`components/ParagraphPopover.tsx` 的 `Num`），
  用户在面板里碰一下首行缩进就把悬挂清零。
- `horizontalRule` 写出 `<w:pStyle w:val="HorizontalRule"/>`（`:933`），本仓不产出该样式定义 →
  推断为缺陷，**Word 里实测 BLOCKED**；正确落点是段落底部边框 `w:pBdr/w:bottom`。

**修复判据**：上面四条各补一条往返用例（右缩进保留、悬挂保留且面板不篡改、`w:start`/`w:end` 与
`w:left`/`w:right` 等价读入、`horizontalRule` 导出为 `w:pBdr`）。

### 1.2 编辑器属性与 UI

新增全局属性（`extensions/ParagraphFormatting.ts`，types: paragraph/heading）：

| 属性 | 单位 | 渲染 | OOXML |
| --- | --- | --- | --- |
| `indentRight` | pt | `margin-right` | `w:ind/@w:right` |
| `indentFirst` 语义扩展 | pt；**负值=悬挂**（保持现有约定，不再新增一个 boolean） | `text-indent` 可为负 | 负→`@w:hanging`，正→`@w:firstLine` |
| `outlineLevel` | 0–9 或 null（Word 的"级别 1–10"= 0–9） | 无视觉效果，仅 `data-outline-level` + 大纲侧栏读取 | `w:outlineLvl/@w:val` |
| `keepWithNext` | boolean | 无 | `w:keepNext` |
| `keepLines` | boolean | 无 | `w:keepLines`（段中不分页） |
| `widowControl` | boolean \| null（null=继承样式，Word 默认开） | 无 | `w:widowControl`（CT_OnOff：关写 `w:val="0"`，缺省即 true） |
| `snapToGrid` | boolean \| null | 无 | `w:snapToGrid` |

UI：`components/ParagraphPopover.tsx` 的 `ParagraphFields` 改成 Word 的两段式——
"常规"（对齐方式 4 段钮 + 大纲级别下拉 + 特殊格式下拉[无/首行缩进/悬挂缩进] + 左侧/右侧缩进 +
特殊格式缩进值）、"间距"（现有段前/段后/行距不动）、"分页与换行"（5 个开关：孤行控制、
与下段同页、段中不分页、段前分页、对齐到网格）。数值框**必须允许负值**（悬挂走的是"特殊格式"下拉 +
正数宽度，内部再取负），并删掉 `Math.max(0, …)`。段落属性入口同时挂进右键菜单已有的「段落(P)…」，
不新建第二个面板。

> 实测差异：开关用的是既有的 `.pp-seg-btn` 而不是 checkbox——`styles.css` 里 `input[type=checkbox]`
> 没有 `width:auto`，新增一行复选框需要先改那份公共样式表。

**不做**：`w:leftChars`/`w:firstLineChars` 等字符单位版属性——本仓一律按 pt 建模，
读到就保留：`mergedElementXml` 做的是属性级合并，`*Chars`/`*Lines` 同义属性只在本轮**确实改了
那个值**时才删除，否则原样透传（实测见 `docx.test.ts` 的"两次保存字节一致"用例）。

### 1.3 提交切分（排名 1 已全部落地）

1. `fix(docx-io): 段落右缩进与 w:start/w:end 在往返中丢失` —— 含 §0.1 的 `w:ind` 生成器合并。（`35ee7a6`）
2. `fix(desktop): 段落面板把悬挂缩进清零` —— 特殊格式下拉 + 磅值取绝对值。（`da82107`）
3. `feat(docx-io): 段落属性补大纲级别/同页/不分页/孤行控制/对齐网格` —— 读写 + 往返用例。（`857f184`）
4. `feat(desktop): 段落对话框补齐 Word 字段` —— 全局属性 + `ParagraphFields` 改版 + live 用例。（`ed57c87`）
5. `fix(docx-io): 分隔线导出为段落底部边框` —— `horizontalRule` 不再引用本仓没写进
   styles.xml 的 `HorizontalRule` 样式，改成 `<w:pBdr><w:bottom w:val="single" w:sz="6"
   w:space="1" w:color="auto"/>`（`w:sz` 单位 1/8 磅，ECMA-376 §17.3.4 的例子里 sz=18 即
   2.25 磅）。导入侧不新增识别：带 `w:pBdr` 的段落继续按 rawPpr 原样透传，两次保存字节一致。
   （`688e2e6`）

> 遗留（不在排名 1 范围内，留给后续排名）：`codeBlockXml` 仍引用未定义的 `Code` 样式；
> 标题/引用块整体丢弃 rawPpr；`w:tab` 制表符不会重新写出。

---

## 二、排名 2：具名样式库 / 样式窗格

Word 的样式是**文档结构的载体**：目录、交叉引用、多级编号、比较文档都靠样式名。
官方描述见 [使用 Word 撰写优秀论文](https://support.microsoft.com/zh-hk/topic/%E4%BD%BF%E7%94%A8-microsoft-word-%E6%92%B0%E5%AF%AB%E5%84%AA%E7%A7%80%E7%9A%84%E8%AB%96%E6%96%87-52cc22e6-7a7a-4fd9-834d-fb6521451af1)
与 [格式化 Word 文档](https://support.microsoft.com/zh-tw/word/format-your-word-document)。

OOXML：`word/styles.xml` 里 `<w:style w:type="paragraph|character" w:styleId=".."><w:name w:val=".."/>`
+ 自己的 `w:pPr`/`w:rPr`；段落通过 `<w:pPr><w:pStyle w:val="styleId"/></w:pPr>` 引用，
run 通过 `<w:rStyle>`。**样式链**（`w:basedOn`）与 `w:next`、`w/link`、`w:default` 都要考虑。

目标形态：段落/标题节点新增可选 `styleId` 属性；`docx-io` 读侧从 `StyleMap` 解析样式清单
（名称、类型、basedOn、自身的 pPr/rPr），写侧**把样式定义合并回 `word/styles.xml`** 而不再是
只写 `w:pStyle` 引用。样式窗格 = 列出内建 + 自定义样式、应用、"新建样式"（从当前选区抓取格式）、
"修改样式"、"清除格式"。

依赖：`StyleMap` 已有的能力边界（见 §附录 A 盘点）。

提交切分：① 读侧样式清单结构化 + 往返；② 编辑器 `styleId` 属性 + 应用/清除命令；
③ 样式窗格 UI；④ 写侧 styles.xml 合并生成。

## 三、排名 3：目录 / 题注 / 交叉引用 / 脚注尾注

域代码语义以官方为准：
[字段代码：TOC（目录）字段](https://support.microsoft.com/zh-cn/office/%E5%9F%9F%E4%BB%A3%E7%A0%81%EF%BC%9Atoc%EF%BC%88%E7%9B%AE%E5%BD%95%EF%BC%89%E5%9F%9F-1f538bc4-60e6-4854-9f64-67754d78d05c)
（`\o "1-3"` 收大纲级别、`\h` 超链接、`\z` 隐藏 Web 视图下的页码、`\u` 收用了"大纲级别"段落格式但未用标题样式的段落）；
文档部件结构见 [WordprocessingML 文档的结构](https://learn.microsoft.com/zh-cn/office/open-xml/word/structure-of-a-wordprocessingml-document)。

要点：
- 目录 = `w:sdt` 或 `w:p` 内 `w:fldSimple w:instr=" TOC \o &quot;1-3&quot; \h \z \u "`，**缓存结果**
  存在域内容里；Word 打开后"更新域"才重算。我们要能读进来（保留原样，不许丢）、能插入、
  能在本地按当前标题/大纲级别重新生成缓存结果（这也是 §2 排名 2 与排名 1 `outlineLevel` 的下游）。
- 题注 = `SEQ` 域 + 标签文本；书签 = `w:bookmarkStart/End`；交叉引用 = `REF`/`PAGEREF` 域。
- 脚注/尾注 = 独立的 `word/footnotes.xml`、`word/endnotes.xml` 部件 + 正文里的
  `w:footnoteReference`/`w:endnoteReference`。**这是本规格里唯一需要新增 zip 部件的一项**，
  要一并补 `[Content_Types].xml` 与 `document.xml.rels`。

提交切分：① 书签（`w:bookmarkStart/End` 往返 + 编辑器锚点）；② 域的读侧保留 + `Field.ts` 扩展
支持 `instr` 文本与缓存结果；③ 目录生成与插入；④ 脚注/尾注部件；⑤ 题注 SEQ + 交叉引用。

## 四、排名 4：批注 / 修订跟踪

- 修订 = `w:ins`/`w:del`（run 级）、`w:pPrChange`/`w:rPrChange`（格式改动）、段落删除用
  `w:del` 包住段落标记；`w:delText` 替代 `w:t`。接受 = 落地并删标记，拒绝 = 反向。
- 批注 = `word/comments.xml` 部件 + 正文 `w:commentRangeStart`/`w:commentRangeEnd`/
  `w:commentReference`。
- 与现有机制的关系：`userData/modlog/<文档名>.json` + 内容寻址撤销（AGENTS §2.5）是**本地改动历史**，
  不能当修订用——它不进文件。本节的修订标记必须是文档内容的一部分，且要与 modlog 互不干扰
  （modlog 继续记录 op，撤销一个带修订标记的 op 时逆操作要把标记一起撤销）。
- 渲染层：`w:delText` 要有删除线呈现；`w:ins` 下划线呈现；接受/拒绝入口进右键菜单。
- 需新增 mark（`trackedChange`）与节点（`commentAnchor`?）→ 一律走 `EDITOR_EXTENSIONS` + live 用例。

提交切分：① 修订读写（run 级 ins/del/delText）；② 编辑器呈现 + 接受/拒绝命令；③ 批注部件读写；
④ 批注侧栏 UI 与右键菜单入口。

## 五、排名 5：制表位 + 前导符

官方：[在制表位之间显示点或前导符](https://support.microsoft.com/zh-cn/office/c80b62d0-2244-4078-95bd-8c82f91440ba)。
OOXML：`w:pPr/w:tabs/w:tab w:val="left|center|right|decimal|bar|list|clear" w:leader="none|dot|hyphen|dash|line|heavy|middleDot" w:pos="twips"`。
`[TabStop Class](https://learn.microsoft.com/zh-cn/dotnet/api/documentformat.openxml.wordprocessing.tabstop?view=openxml-2.15.0)` 给出属性集。

要点：段落属性 `tabStops: {positionPt, alignment, leader}[]`（位置用 pt，导出 ×20）；
编辑器里 `Tab` 键已经能产生 `w:tab` 字符，缺的是**制表位定义**与"Tab 字符按制表位对齐"的
渲染（Tiptap 里用 `white-space: pre-wrap` + 每段生成 `background-image`/`text-indent` 代价太大 →
决定：制表位只做数据层与导出，编辑器内 Tab 仍按浏览器默认宽度呈现，并在规格里如实写明这个差距）。
UI 挂在「段落」面板下方（对齐 Word 的"制表位…"按钮），支持"清除/全部清除"。

## 六、排名 6：上标 / 下标

OOXML：`w:rPr/w:vertAlign w:val="superscript|subscript"`。
实现：新增 Tiptap mark（不装 `@tiptap/extension-superscript`，因为它渲染成 `<sup>` 且不带
`data-*`，与 `TextStyle` 的属性白名单模式不一致；改为把 `vertAlign` 作为 `TextStyle` mark 的
一个属性，与 `color`/`fontSize` 同构）→ 这样字体面板、右键「清除所有格式」、docx 往返都自动复用现有链路。
入口：工具栏 + 右键「字体(F)…」面板。**注意 AGENTS §2.3**：部分文字带上下标时禁止整段拍平。

## 七、排名 7：分栏 / 水印 / 页面边框 / 首字下沉 / 行号

- 分节：`w:p/w:pPr/w:sectPr`（分节符）与 body 末尾的 `w:sectPr`。现在 `pmToOoxml`（`:234-243`）
  只在 body 末尾写**一个**固定 `w:sectPr`，读侧遇到 `w:sectPr` 直接 `continue` 丢弃
  （`:215`）——所以多节文档保存后会被拍成单节。这是本节第一个要修的保真问题。
- 分栏：`w:cols w:num="2" w:space="twips"` + 每栏 `w:col`；编辑器需要 `section` 级属性承载，
  决定用「页面设置」面板（已有页面尺寸/边距的地方）。
- 页面边框：`w:pgBorders`；水印：`w:headerReference` 里的 `v:shape` 文字（Word 用 WordArt VML，
  本仓可以退化成 `w:pict` 文本框，需在规格里明确"Word 打开效果待验"）。
- 首字下沉：`w:pPr/w:framePr w:dropRows="3" w:line="…" w:hAnchor="text" w:vAnchor="text"`。
- 行号：`w:lnNumType w:countBy="1" w:restart="newPage"`（节属性）。

提交切分：① 多节 sectPr 读写保真；② 分栏；③ 页面边框 + 行号；④ 首字下沉；⑤ 水印。

## 八、总验收表（每个功能提交都要填满）

| 条目 | 判据 |
| --- | --- |
| 往返用例 | `packages/docx-io` 新增用例：给定 OOXML → attrs → OOXML，属性不丢不变 |
| live 用例 | `docBlocks.live.test.ts` 用真实 `EDITOR_EXTENSIONS` schema 覆盖新属性/mark |
| 局部格式 | 若支持整段应用，必须拒绝"一段内部分文字有不同格式"的拍平（AGENTS §2.3） |
| 撤销 | 该功能产生的每个 op 有逆操作，`modlog` 记一条 |
| 类型 + 测试 | `pnpm build`、`pnpm test` 全绿，且**逐包核对通过数**（AGENTS §8） |
| UI | `pnpm dev:desktop` 实跑：选中→改→撤销→导出 `.docx`→读回比对（AGENTS §7.2） |
| Word 实测 | 本机无 Word/WPS → 标 BLOCKED；不得写"已验证" |
| 提交 | 中文 `<type>(<scope>): 做了什么+为什么`，一提交一件事，推送后 `git ls-remote origin main` 复核 |

### 8.1 排名 1 验收记录（`35ee7a6` `da82107` `857f184` `ed57c87` `688e2e6`）

- 往返用例：`docx.test.ts` 的"keeps every indentation spelling through a real save"、
  "keeps the pagination switches and the paragraph mark run props"、
  "keeps an imported paragraph border through two real saves"、
  "saves the editor's horizontal rule as a paragraph bottom border"（两次保存字节一致）。
- live 用例：`docBlocks.live.test.ts` 的"段落对话框补齐的字段（真实 schema）"3 条，含
  "新段落六个新属性都是 null，套版不会凭空写出 `w:keepNext`"。
- 局部格式：整段字符样式仍走 `applyStyleToNode` 的部分覆盖拒绝；新增字段都是段落级，无拍平面。
- 撤销：`indentRight` 与 5 个开关进 `STYLE_KEYS`，`setStyle` op 的逆操作由既有快照机制给出；
  开关与 `outlineLevel` **刻意不上报给 LLM**（`agent.ts` 提示词只列缩进/间距），避免模型乱设分页。
- 类型 + 测试：`pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包
  docx-io 103 / desktop 119 / word-client 8 / word-tools 6 / server 2；`apps/cli`、
  `packages/shared` 无测试文件（根命令 `passWithNoTests`，故必须逐包核对）。
- UI：`[!] BLOCKED`——本仓没有 Electron 的 CDP harness，渲染层 dev 也无 mock bridge，
  新面板未实跑走一遍；仅有真实 schema 的 live 测试作为替代证据。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。

## 附录 A：现状盘点（`path:line` 为证）

- 段落格式属性与命令：`apps/desktop/src/editor/extensions/ParagraphFormatting.ts:34-133`
  （`setParagraphFormat` 逐块 `setNodeMarkup`）。
- 段落面板：`apps/desktop/src/components/ParagraphPopover.tsx`（`ParagraphFields` 被工具栏与
  右键「段落(P)…」共用）；字体面板 `components/FontFields.tsx`；
  菜单模型 `apps/desktop/src/editor/contextMenu.ts`。
- OOXML 读侧段落属性：`packages/docx-io/src/ooxml-pm.ts:307-383`（pPr 采集 + rawPpr 透传）、
  `:385-`（`readParagraphFormat`：`w:ind` 含 left/right/start/end/firstLine/hanging、`w:spacing`
  行距、`w:outlineLvl` 与 4 个 CT_OnOff 开关）。
- OOXML 写侧：`paragraphXml` `:1001-`、`PPR_CHILD_ORDER` `:1029`（CT_PPr 子元素顺序）、
  `modeledPprParts` `:1102`、`paragraphPprXml` `:1137`（导入与新建段落共用）、
  `mergedElementXml` `:1216`（属性级合并）、`horizontalRuleXml` `:1422`、
  `pmToOoxml` 尾部单 sectPr。
- 样式：`packages/docx-io/src/styles.ts`（`StyleMap`）、`docx.ts:119`（`word/styles.xml`）。
- 域：`apps/desktop/src/editor/extensions/Field.ts`；编号：`packages/docx-io/src/numbering.ts`。
- 模板里的首行缩进按 pt 计算：`apps/desktop/src/editor/templates.ts:26`（`2 字符 = 2 × 字号磅值`）。
