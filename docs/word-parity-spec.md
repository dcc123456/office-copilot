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

提交切分（实际落地顺序与提交号）：

| # | 内容 | 提交 |
| --- | --- | --- |
| ① | 读侧样式清单结构化（`StyleMap` 解析 `w:name`/`w:type`/`w:basedOn`/`w:default` + 往返） | `21cb2e9` `aeed0be` |
| ② | 写侧把 `attrs.styleId` 写成 `<w:pStyle>`（string=套用、null=删引用、缺失=沿用源文档写法） | `58220ff` |
| ③ | 编辑器 `styleId` 全局属性（默认值只能是 `undefined`，否则导入文档的引用会被整批删掉） | `91131d0` |
| ③.5 | 写侧按样式表复核引用：改了级别就换成这一级声明的那条，标题转正文就删引用 | `aa0d848` |
| ④ | 样式窗格：引擎半（`paneStyles()` + 唯一的对齐规则 `fitStyleId`）`7cb35b5`；UI 半（清单经 IPC 进渲染层、`applyNamedStyle` 命令、窗格）`481c4ec` | `7cb35b5` `481c4ec` |
| ⑤ | 样式链解析成可见格式（文档作用域生成的样式表，换算仍只在 `docx-io`）：引擎半 `26e5783`，渲染半 `3a9a79b`；前置整理 `5171dcb`（`@w:val` 读取收进 `xml.ts`，第四处副本） | `26e5783` `3a9a79b` |
| ⑥ | 新建样式 / 修改样式（从选区抓格式）：抓取半 `b6dd2be`，窗格入口与保存装配 `84d4e80` | `b6dd2be` `84d4e80` |
| ⑦ | 写侧 `styles.xml` 合并生成（splice-and-keep，保住 `docDefaults`/`latentStyles`）+ 真实保存往返：引擎半 `8970b77`，装配半 `84d4e80` | `8970b77` `84d4e80` |

窗格这一版的取舍（都是按 Word 的行为定的，不是省事）：

- **只写引用，不解析样式链**。Word 里套样式不会把样式定义盖到段落上，直接格式永远赢；
  解析成段落属性等于毁掉用户的红字、加粗和手工缩进。
- **清除样式 = 不写引用**（Word 里等同该文档的默认段落样式），且标题节点退回 `paragraph`——
  留着 heading 却没有标题样式引用，重新导入时这段会被读成正文，级别白丢。
- **新建文档的窗格是空的**。`fileHandlers.ts` 的空白模板根本不写 `styles.xml`，硬塞一份 Word
  预置样式清单会写出文档没声明的 `styleId`，Word 照正文排版——正是早期 `HorizontalRule` 犯过的错。
- 窗格高亮与写侧共用 `fitStyleId`，因此不会出现"窗格显示标题 1、Word 排成标题 2"。

⑤（样式链解析 + 文档作用域样式表）这一版的取舍，一手依据是 ISO/IEC 29500-1 §17.7.2.5：
`w:docDefaults` 的属性"applied first in the style hierarchy; therefore they are superseded by
any further conflicting formatting"，`w:pPrDefault`/`w:rPrDefault` 各带一份 pPr/rPr。

- **合并顺序 = docDefaults → `w:basedOn` 链（根在前）→ 这条样式自身**；直接格式不进这里，
  它由段落的 attrs 出行内 `style`，CSS 的行内优先级正好停在 Word 的"直接格式最高"上。
  否决方案"导入时把解析结果铺进段落 attrs"：那会把样式定义盖成用户从没做过的直接格式，
  导出后再分不清哪部分属于用户（AGENTS §2.3）。
- **引擎产出结构化数据、不产出 CSS，且用的是编辑器那套词汇**（pt / 倍数 / `pt:N` / `atLeast:N` /
  CSS px / #rrggbb）。候选方案"引擎直接生成样式表文本"被否：行距与 pt→px 的映射在渲染层已经各有一份
  （`ParagraphFormatting.ts`），引擎再生成一份就是第二套换算（§2.1）。
- **`w:val="0"` 与"没提过"必须分开**：`readRunProperties` 的开关是三态，孙子样式用 `<w:b w:val="0"/>`
  取消父样式的加粗要能取消；`w:pStyle` 引用了本文档没声明的 id 时退回默认段落样式而不是"无格式"
  （真实简历里就有 `w:pStyle w:val="Heading1"` 而 styles.xml 声明的是 `1`）。
- **主题字体不猜**：Word 写的 docDefaults 是 `w:asciiTheme`/`w:eastAsiaTheme`，没有 `theme1.xml`
  就求不出字体名，读不出就不产出（WPS 写的是字面 `w:ascii`/`w:eastAsia`，能读）。
- **样式定义里 `w:pPr/w:rPr` 那份"段落标记"格式不计入**：那是 Word 给段落标记本身的，不影响正文。
- `w:hidden` 的样式**照旧解析**——那只管窗格列不列它；链上一个属性都没写过的样式不产出规则。
- 特异度：`.ProseMirror [data-style-id="X"]`（0,2,0）> 基线 `.ProseMirror p`/`h1`（0,1,1）；
  兜底规则用 `:not(:is(...))` 排除已声明的 id，`:is()` 取参数里最高的特异度所以与条数无关，
  同为 0,2,0 时靠"兜底在前、具名在后"的源码顺序让具名样式获胜，同时未声明的引用也拿到了默认格式，
  没有基于 Normal 的根样式也不会被兜底漏灌。
- 样式表文本来自用户文件，按系统边界校验：`styleId` 只允许 `[A-Za-z0-9_-]{1,64}` 才进选择器，
  字体名/颜色/长度按形状白名单，读不懂的按"这份文档没写过"丢弃（含 `<`、`}` 的值能把 `<style>`
  里的规则截断）。
- 已知差距（不修）：引用未声明 id 的**标题节点**仍带基线 `h1` 的粗体与外边距（Word 会整体按默认样式排）；
  `buildPrintHtml` 是另一份 CSS，PDF/打印暂不含文档样式。

`481c4ec` 顺带修掉的一个真实缺陷：`setParagraphFormat` 原先逐块各发一个基于旧 `doc` 的事务，
选中多段时第二段起 ProseMirror 抛 `Applying a mismatched transaction`，只有第一段改得上
（实测：`docBlocks.live.test.ts` 跨段用例在旧实现下必失败）。现在两条段落命令共用
`mapFormatBlocks`，一次操作攒成一个事务。

⑥⑦（新建 / 更新样式，并把样式定义真的写回文档）这一版的取舍：

- **预览与保存共用同一个 `DocxFile.applyStyleEdits`，"窗格预览 == 保存结果"不是巧合而是同一段
  代码跑两遍**。为此在主进程加了 `styles:preview`：拿当前文档（没有文件就用空白模板）在内存里
  合并攒下的编辑，返回 `paneStyles()` 与样式链投影，**不落盘**。否决方案"渲染层自己算样式链"：
  那是第二套样式解析，与 §5 引擎层优先直接冲突。
- **样式表先于 `replaceProseMirror` 落地**。正文里的 `<w:pStyle w:val="X"/>` 引用的是 styles.xml
  声明过的 id；先写引用再写声明，中间失败就留下"引用了未声明样式"的文档，Word 会静默按正文排。
  `docx-io` 的用例"saves both the declaration and the body reference of a created style"钉的就是
  这一条顺序（同一次保存里 `document.xml` 有引用、`styles.xml` 有声明）。
- **新建 / 更新只写引用，不改块类型**（`applyStyleReference` 用 `setParagraphFormat({ styleId })`）。
  不能用 `applyNamedStyle`：那条命令按 Word 的套用规则把标题段转成正文段，用户只是想给这段起个名。
- **同一条样式只留最新一份编辑**（`styleEditsRef` 按 `styleId` 去重后追加），且用 ref 不用 state：
  菜单、快捷键、工具栏拿到的保存 handler 必须始终看到最新一份，把 state 塞进依赖会让每个菜单项跟着重建。
  换文档（新建 / 打开）时清空——那些编辑指向的是上一份文档的 `styles.xml`。
- **合并是 splice-and-keep**：同 id 的旧 `<w:style>` 整条换成新的，`docDefaults`、`latentStyles`
  与本文档其它样式原样保留；部件缺失时同时补 `[Content_Types].xml` 的 Override 与
  `word/_rels/document.xml.rels` 的 Relationship（Word 是经关系找到 styles.xml 的，只登记内容类型不够）。
- **局部格式不拍平**（AGENTS §2.3）：`captureStyleFromSelection` 对"各段属性不一致"和"某项字符格式
  只覆盖部分文字"两种情况返回失败码，窗格按码拼中文提示（§5.7 引擎不给文案）。
  `[!]` 这条拒绝分支在真实 app 里没能手工触发（探针按 `title === "加粗"` 找按钮，实际 DOM 的 title 不同），
  由 `styleCapture.test.ts` 的 7 条用例覆盖。

排名 2 的遗留（不在本排名的验收范围内）：编辑器里多块一起改标题级别时，Tiptap 的
`setNode`/`setHeading` 只在 `$anchor.sameParent($head)` 时才复制原块属性，跨块会丢 `styleId`
（一手依据：`@tiptap/core/dist/index.js:3307-3331`）；`fitStyleId` 改为"没建模过的标题先问样式表"
之后，真实文档里这个丢失不再改变导出结果，故记为遗留不单独修。

另一条已知差距（`loadForEdit` 引入前就存在，不是这版带来的）：**另存为到一个还不存在的路径时，
`saveDocx` 从空白模板重建文档**，源文件的 `styles.xml`、`numbering.xml` 等部件不会跟过去，正文里
`Heading1`/`Heading2` 这类引用因此在副本里成了未声明样式。这版保证的是"用户自己新建的样式，声明与
引用一定同时到位"；要把整份保真复制到新路径，得改成从**当前内存里的那份 `DocxFile`** 写盘，属于保存
链路的重构，留给后续排名处理。

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

提交切分（实际落地顺序与提交号）：

| # | 内容 | 提交 |
| --- | --- | --- |
| ① | 书签：`w:bookmarkStart/End` 往返（段内直子元素，落单的一半在序列化前丢掉）+ 编辑器锚点节点 | `ba17b00` |
| ② | 域读侧：同段 begin/end 配对成一条 `field` 节点，保留 `instr` 与**字面**缓存结果（`\t`⇄`<w:tab/>`、`\n`⇄`<w:br/>`），域内套域按 fieldDepth 折进外层结果；`Field.ts` 属性补 parse/render | `16a0e89` |
| ③.1 | 跨段域的边界模型：段落内 begin 而本段无 end 时吐 `fieldBegin`/`fieldEnd` 一对节点（`fldCharTypes`/`closesInParagraph` 预扫判定），begin 后域深度归零让条目里的 `PAGEREF` 继续结构化；写侧 `fieldBoundaryXml` 还原 begin+instrText+separate 三件套；编辑器两个行内 atom | `bc06382` |
| ③.2 | 目录生成与插入：按当前标题/`outlineLevel` 造 `TOC1…` 条目段（书签 + 制表位 + `PAGEREF`），页码取自本地分页。前提是排名 5 ① 的制表位数据层，已落地 `8c37fb4` | `2dadd0a` |
| ③.3 | 目录的入口与刷新：插入目录弹窗（级别数/是否显示页码/前导符/插入位置取自选区），落地后等分页收敛再刷一次 `PAGEREF` 缓存页码 | `e9f2806` |
| ④.1 | 脚注/尾注部件保真（`footnotes.xml`/`endnotes.xml` + Content_Types + rels）+ 正文引用点的行内 atom | `013d8ca` |
| ④.2 | 脚注正文的编辑与编号：在编辑器里改脚注内容、插入/删除脚注时重排 `w:id` 与 Word 的显示编号 | `f005b3c` |
| ⑤.1 | 题注 `SEQ`：`SEQ 图 \* ARABIC` 域 + `_Ref` 书签包裹缓存编号，标签各自计数、编号随正文派生重排；对话框选标签/位置/说明文字 | `51750c3` |
| ⑤.2 | 交叉引用 `REF`/`PAGEREF`：指向 `_Ref…`/`_Toc…` 书签的引用弹窗，新 op `insertInline`/`removeInline`；前置的读侧保真 `2e71a6f` | `2e71a6f` `4284046` |

③.2 的取舍（都是刻意的，不要当成待修项）：

- **不写 `w:pStyle`**：文档里没有 `TOC1`/`TOC2` 样式定义，写过去就是一个悬空引用，Word 会回落到默认样式并
  判定文档可疑。条目段的缩进用 `w:ind`、制表位用 `w:tabs`，都走排名 5 ① 已经打通的 attrs。
- **不包 `w:hyperlink`**：`\h` 是"更新域时由 Word 生成超链接"，我们此刻手写的链接只会让每一条都变蓝带下划线。
- **前导符在编辑器里还看不见**：制表位只呈现为 `data-tab-stops`，没有对应的 CSS，所以目录行在编辑器里是
  "文字 + 一串空格似的东西"，写进文件才是点线。排名 5 ② 落地的是**编辑入口**（段落浮层里的「制表位…」），
  呈现这条差距原样保留，见五节"验收与差距"。
- **`deleteBlock` 的逆操作会降级 json 段**：撤销一个插入目录的 op 时，`newFromBlock()` 按纯文本重建，
  `\t` 仍能序列化成 `<w:tab/>`，但 `PAGEREF` 域与 `tabStops` attrs 会丢。整块撤销的语义仍然是"文字回到原样"，
  结构回到原样要等 json 段自己的逆操作。

③.3 的验收与差距：

- 真实 app 实测（构建产物 + CDP 驱动，`--user-data-dir` 走独立目录，不碰用户自己在跑的实例）：工具栏「插入目录」
  开弹窗 → 默认"光标所在段之后"、收录 1 个标题 → 确定后 toast「已插入目录（1 条）。」，条目段落
  `data-tab-stops` 为 `{positionPt: 451.5, alignment: "right", leader: "dot"}`，文字 `无标题⇥1`，脏标记亮起。
- **多页刷新路径未走通**：实测样例只有一页，`watchTocPageNumbers` 的"分页收敛后页码变了"这条分支没被真正触发过；
  它的正确性目前只有 `tocPageNumberRefresh` 的单测背书。（这两个函数在 ⑤.2 改名
  `watchFieldPageNumbers`/`fieldPageNumberRefresh`，因为目录条目与交叉引用共用同一套页码刷新。）
- **"用 Word 打开一遍"仍 `[!] BLOCKED`**：本机没有 Word / WPS / LibreOffice，`.docx` 只做到读回一致。
- 前导符在编辑器里依旧不可见（见上面 ③.2 的取舍）；排名 5 ② 只补了编辑入口，这条没有因此关掉。

④.1 的做法（脚注/尾注部件）：这一项修的是一个**既有缺陷**——正文里的 `w:footnoteReference` 运行块
在写侧根本没有对应分支，保存时整段被丢，用户的脚注引用凭空消失。取**透明部件**策略：正文只留一个引用
原子，条目本身按 id 存在 `FootnoteStore` 侧表里，没改过的条目整段 `rawXml` 原样带走（连 `w:footnote`
自己身上那些没建模的属性一起），只有 `dirty` 条目才用 PM⇄OOXML 片段转换重建；`separator` 与
`continuationSeparator`（负数与 0 号 id）是排版家具，**永不因为正文没引用它们而被删**；部件里一旦出现
非数字 id，整个部件标记为 `unmodeled` 且一次都不重写，宁可放弃功能也不静默丢内容。登记复用
`ensurePartReference`（Content_Types Override + rels，幂等），表头页脚则迁到共享的片段转换入口，
不再各自留一份实现（§3.2）。编辑器侧 `FootnoteReference`/`EndnoteReference` 两个行内 atom 注册进
`EDITOR_EXTENSIONS`；芯片当时只显示"注/尾"，因为 `renderHTML` 拿不到节点在文档里的位置——编号与正文的编辑
都由 ④.2 补上。

④.1 的验收与差距：

- 真实 app 实测（构建产物 + CDP 驱动，独立 `--user-data-dir`，不碰用户在跑的实例）：含脚注与尾注各一条的
  样例打开后，芯片是蓝色上标（`verticalAlign: "super"`，文字"注"/"尾"），无白屏；点保存后读回部件，
  8 项结构判定全过（引用点带上标 `rPr`、正文条目与 `separator` 家具齐备、Content_Types 有 Override、
  rels 未重复）。
- **"用 Word 打开一遍"仍 `[!] BLOCKED`**：本机没有 Word / WPS / LibreOffice，只做到读回一致。
- 差距：脚注正文在编辑器里**只读**（侧表不进 doc，改了也不会写回），Word 的显示编号没有回写；
  这两条由 ④.2 关掉。`w:customMarkFollows` 标过的引用刻意不读；脚注正文里的图片不解析关系，整段原样透传。

④.2 的做法（正文编辑与编号）：**编号是派生显示值，绝不进文件**——Word 按引用点在正文里的出现顺序算号
（ECMA-376 §17.11，`w:id` 只是内部键），所以 `FootnoteNumbering` 插件在 `docChanged` 时遍历一遍正文，
把号写回芯片的 `number` attrs，并用 `setMeta("addToHistory", false)` 让"重排号数"不掺进用户的撤销栈。
同一次遍历顺手记下 `kind:id → 第一次出现的 pos`，窗格的排序、跳正文、删引用都读这一份索引，不再各扫
一遍树（§3.2）。正文的编辑走 `FootnoteBodies` 侧表，随页眉页脚**同一份 payload** 进出主进程，落盘仍然只有
`DocxFile.setFootnoteBodies` → `applyFootnoteBodies` 一个写手；文字没变的条目由 `sameFragment` 判定后继续
按原始 XML 字节透传，`separator` 版式件不接受编辑，`unmodeled` 的部件整体无视窗格送来的正文。插入取号用
`nextFootnoteId`（正常条目里最大的数字 id 加一，版式件不参与，空表从 1 起）。**删引用只删芯片、不删缓存正文**：
用户接着撤销时字还在，真正没人引用的条目由保存时的 `syncFootnoteIds` 清掉，窗格里那条排在最后、号显示 "?"
且不可点。脚注正文的嵌套编辑器复用 `NESTED_EXTENSIONS`——页眉页脚是它的第一处使用，这里是第二处（§3.2）。

④.2 的验收与差距：

- 真实 app 实测（构建产物 + CDP 驱动，独立 `--user-data-dir`，不碰用户在跑的实例），样例是脚注两条 + 尾注一条：
  打开后芯片即 `1 / 2 / 1`（脚注按正文顺序、尾注自己数），窗格列出 3 行带预览；选中第一条后嵌套编辑器给出
  部件里读出的正文；改成新句子后列表预览与落盘同步；「插入脚注」得到号 2 并把原第二条顶成 3——**派生号跟着
  正文重排**；删除第一条的引用后号再次重排（原 2 变 1），那条正文变成孤立行显示 "?"；点工具栏「撤销」后芯片
  回来且正文缓存还在。
- 保存后读回 `word/footnotes.xml` 的 8 项判定全过：条目 id 齐 `1/2/3`，两条 `separator` 家具保留，
  三个 `w:footnoteRef` 记号都在，没改的第二条逐字节透传（连 `xml:space="preserve"` 一起），新条目是带记号的
  空正文，`endnotes.xml` 一字未动，Content_Types 只登记一次。
- 差距：撤销实测走的是工具栏「撤销」按钮（与 `editor.can().undo()` 同一条路径），CDP 注入的 Cmd+Z 没有送达
  PM 的 keymap，**键盘路径只有 PM history 背书，不是实测**。同理一次撤销会把 500ms 内的多处改动并成一个历史组
  （PM history 的 `newGroupDelay`），探针必须拉开间隔才看得清单步语义——真实手速下用户看到的也是并组后的结果。
- 差距：`w:numberFormat`（`i, ii, ①` 之类）与 `w:customMarkFollows` 仍不读不写，原样透传；脚注正文里的图片仍不
  解析关系；窗格没有"脚注转尾注"与编号格式入口。
- **"用 Word 打开一遍"仍 `[!] BLOCKED`**：本机没有 Word / WPS / LibreOffice，只做到读回一致。

②的已知差距：一个 `field` 节点只表示**同一段内**的 begin/end 配对，Word 生成的多段目录从第二段起读成纯文本
——已由 ③.1 解决。缓存结果里的嵌套域指令（如目录条目的 `PAGEREF`）②只保留文本，③.1 起改由边界节点之间的
正文承载，重新生成目录（③.2）因此必须先有这一对边界。
未闭合的域（有 begin 无 end）③.1 选择**原样透传**：不替用户编一个 end，Word 重算时会自己补齐。

⑤.1 的做法（题注 `SEQ`）：Word 的题注编号是**域的缓存结果**，不是正文文字（ECMA-376 §17.9.4.15、
微软「字段代码：SEQ 字段」），所以 `editor/captions.ts` 造的段落是
`text("图 ") → bookmarkStart(_RefN) → field(SEQ) → bookmarkEnd → text(" 说明")` 五件套：编号写进
`field.display`，指令 `SEQ 图 \* ARABIC` 写进 `instr`（标签含空格才加引号），序列化仍走 ②/③.1 打通的
`fldChar`/`instrText` 路径，`docx-io` 一行没改。取号按"目标边界之前有几条同标签的域"来数
（`sequenceCountBefore`），图/表/公式各自计数，上题注与下题注因此能拿到不同起始号。编号是**派生显示值**：
`CaptionNumbering` 插件在 `docChanged` 后按标签重扫一遍、把号写回节点 attrs，并用
`setMeta("addToHistory", false)` 让重排不掺进撤销栈——模式与 ④.2 的脚注编号相同（第二处同一逻辑，§3.2 只是
照搬结构，没有为此新建基础设施）。书签 id 复用目录那份 `nextBookmarkId`（全文档最大数字 id 加一），
`_Toc…` 与 `_Ref…` 不会撞号。落文档走既有 op：`planCaptionInsertion` 先按大纲模拟，`runOps` 有跳过的段就
整批放弃并返回 `ops-failed`；块模型寻址用的文字由 `jsonContentText(json)` 现算——那份实现原本私有在
`toc.ts`，这是第二处使用，已上收到 `docBlocks.ts`（§3.2）。

⑤.1 的取舍（都是刻意的，不要当成待修项）：

- **不写 `w:pStyle Caption`**：与 ③.2 同理，文档里没有这个样式定义，写过去就是悬空引用；题注段的 attrs 一律留空。
- **块模型的地址文字里是双空格**（`图  系统架构`）：域是原子节点，不贡献 `textContent`，而 `buildJsonBlock`
  拿现算的期望文字核对节点，所以标签后那个空格必须留在 text 节点里，否则撤销时按内容寻址找不到回来那条路。
- **不手工包 `w:hyperlink`**：交给 Word 更新域时生成，同 ③.2。
- **`_Ref` 书签只保证存在，不做跳转**：引用入口是 ⑤.2。

⑤.1 的验收与差距：

- 真实 app 实测（构建产物 + CDP 驱动，独立 `--user-data-dir`，不碰用户在跑的实例），样例是两段正文
  （`正文一 / 正文二`）：工具栏「插入题注」开弹窗，预览即 `将插入「图 1 架构图」`；确定后芯片 `["1"]`、
  书签 `["_Ref1"]`、新段 `图 1 架构图`。**第二次插在前面那条题注之上**，预览给 `图 1 流程图`，落地后芯片
  `["1","2"]`、书签 `["_Ref2","_Ref1"]`、段落顺序 `图 1 流程图 / 图 2 架构图`——**编号跟着正文重排**，
  新号占 1、旧那条自动变 2；点工具栏「撤销」回到 `["1"]` 与 `图 1 架构图`。
- 保存后读回 `word/document.xml`：`SEQ` 指令 2 条、缓存编号 `["1","2"]`，`w:bookmarkStart` 是段直子元素
  （`w:id="2" w:name="_Ref2"` / `w:id="1" w:name="_Ref1"`），且指令文字没有出现在任何 `w:t` 里
  （这是要求，不是缺陷）。`docx-io` 读回题注 2 段，首段节点序
  `text:图 → bookmarkStart:_Ref2 → field:1 → bookmarkEnd → text: 流程图`，写回没有被挡下的段。
- **保存产物与读入的 `document.xml` 逐字节不一致（2095 → 2149），根因不在题注**：另跑了一个"打开后不编辑、
  直接保存"的隔离探针，结果 `<w:sectPr>` **整块消失**（打开时有，原样保存后没有）——页面设置被应用的保存
  payload 丢了，而 `docx-io` 的直接读写路径会保留它。这是 §2.7 的既有缺陷，属排名 7 ①（多节 sectPr 读写保真），
  **没有夹进本次题注提交**。
- 差距：题注段没有 `Caption` 样式，在 Word 里不会自动带出题注的字号与间距；编号只在编辑器里派生，
  落盘是域的缓存值，Word 按 F9 重算时以自己为准。
- **"用 Word 打开一遍"仍 `[!] BLOCKED`**：本机没有 Word / WPS / LibreOffice，只做到读回一致。

⑤.2 的做法（交叉引用 `REF`/`PAGEREF`）：一手依据是微软「字段代码：Ref 域」（形状是
`{ REF 书签名 [开关] }`，该文列出的开关只有 `\d \f \h \n \p \r \t \w`；不带开关取"书签所代表的那段文字"，
`\p` 给 above/below，`\h` 是"更新域时做成超链接"），页码那一档走 `PAGEREF`（③.2 的目录条目已经在写它）。
可引用目标 = 文档里**成对出现**的书签（`collectRefTargets` 按 `w:id` 配对，落单的半边不列，与 Word 的
对话框一致）；"书签文字"取书签实际裹住的那段行内内容，而域是原子节点、它对外贡献的是自己的 `display`，
所以 `_Ref…` 裹住 `SEQ` 域时书签文字就是编号——这与 Word 的列法一致。落文档新加一对 op：
`insertInline`/`removeInline`（`ai/docEdit.ts` 模拟 + `editor/docBlocks.ts` 落地）。为什么不复用 `setText`：
撤一条引用会连带整段重排，把用户同段里的图片、域与字符样式一起拍平（§2.3）；`insertInline` 只放一个行内
原子域、不贡献 `textContent`，块的地址文字因此不变，逆操作 `removeInline` 按**指令文字**认域（内容寻址、
不含坐标），因此撤引用不牵连别处后来的手工编辑（§2.5）。插入点用块模型的文字偏移
（`inlinePosForTextOffset` 只在文字上计数，域与书签不占位）。页码档取自本地分页 `pageNumberOf`，与 ③.3
同一个来源：`bookmarkPages` 去掉 `_Toc` 前缀过滤（`PAGEREF _Ref…` 也要查得到页码），
`tocPageNumberRefresh`/`watchTocPageNumbers` 因此改名 `fieldPageNumberRefresh`/`watchFieldPageNumbers`。
编号档：`captionNumberPatches` 在同一次遍历里记下"书签名 → 那条 `SEQ` 的号数"，走完文档再回头刷
`REF _Ref…` 的缓存结果，所以"引用在前、题注在后"（Word 的"以下"）也能跟着重排。
读侧另有一处**既有缺陷**被本次实测抓出来并单独提交（`2e71a6f`）：`w:hyperlink` 里只按 run 抽文字，
Word 自己写的 `{ PAGEREF _Ref2 \h }` 指令被丢掉、存回去就是死文字；改成链接内部走同一套域状态机。
同一提交还修了 `w:anchor` 越存越长的 `#`（读侧折成 `#名字`，写侧没剥掉）。

⑤.2 的取舍（都是刻意的，不要当成待修项）：

- **`\h` 只写进指令，不手写 `w:hyperlink`**：同 ③.2。链接只落在文字上（`addMark` 的既有约定），
  所以整条 `w:hyperlink` 里只有域、没有文字时，存回去不再有链接——指令保住了，超链接让 Word 更新域时补。
- **"以上/以下"的中文措辞是推断**：位置判定（域在书签之前算 below、之后算 above）按官方文档实现且实测写出，
  但字面语言由更新方的界面语言决定，本机没有 Word，观察不到它重算后的措辞。
- **不做跳转**：点芯片不导航到书签；`\h` 生成的链接要在 Word 里按 F9 之后才可点。
- **引用块模型的地址文字仍带题注那个双空格**（`图  系统架构`），与 ⑤.1 同因；`setText` 会把它归一成一个空格，
  单测里那条断言按归一化后的文字写。

⑤.2 的验收与差距：

- 真实 app 实测（构建产物 + CDP 驱动，独立 `--user-data-dir`）：样例是"带 `_Toc1` 书签的标题 +
  一条正文 + Word 写的题注（`_Ref10` 裹 `SEQ`，缓存 1）+ Word 写的交叉引用（`{ PAGEREF _Ref10 \h }` 整条被
  `w:hyperlink` 包着）+ 一条正文"。打开后芯片 `["1","2"]`——**Word 那条交叉引用读成了域**
  （`2e71a6f` 之前只有 `["1"]`，链接里的指令被折成死文字）。
- 弹窗列出 `["目录条目：第一章 绪论","题注：图  系统架构"]`；选"题注 + 页码"提示
  `域指令：{ PAGEREF _Ref10 \h }`，确定后域落在光标处（`正文一` 变成 `[域]正文一`）；切"书签文字"提示变
  `{ REF _Ref10 \h }，当前内容为「1」`；切"以上或下方 + 纯文字"并把目标换成目录条目时提示 `{ REF _Toc1 \p }`，
  缓存结果写 `以上`（光标在标题之后，实测）。
- **点对话框里的分段按钮一度会把对话框关掉**：`CrossRefDialog` 漏了既有约定——`.modal` 上要挂
  `onClick={(e) => e.stopPropagation()}`，否则点击冒泡到 `modal-backdrop` 的关闭处理（实测第二次打开后
  「确定」按钮已经不在 DOM 里）。补上后连点三档对话框仍在。**这是本次探针才暴露出来的，⑤.1 没踩到**：
  那次的探针用 JS 直接写 input 值，没触发过段按钮的真实点击。
- 保存后读回 `word/document.xml`：`w:instrText` 五条
  `[" PAGEREF _Ref10 \\h "," REF _Toc1 \\p "," REF _Ref10 \\h "," SEQ 图 \\* ARABIC "," PAGEREF _Ref10 \\h "]`，
  指令没有落进任何 `w:t`，`_Toc1`/`_Ref10` 两个 `w:bookmarkStart` 配对还在；Word 那条的缓存页码从 2 被本地
  分页刷成 1（`fieldPageNumberRefresh` 的既定行为）。`<w:sectPr>` 依旧在保存 payload 里消失——与 ⑤.1 同一条
  既有缺陷（排名 7 ①），再次复现，未夹进本次提交。
- 撤销：一次「撤销」把三条引用一起退掉（芯片 5→2，`正文一` 回到没有域的样子，题注与 Word 那条不受牵连）——
  三次插入落在 ProseMirror 历史的同一个分组里（`newGroupDelay` 默认 500ms，脚本往返之间没跨过窗口），
  不是撤销粒度出错；逐条粒度由 `crossrefs.test.ts`、`docEdit.test.ts` 与 `docBlocks.live.test.ts` 的真实
  schema 用例背书。**再退一步会越过"打开文件"这一步、整篇换回欢迎页**：`App.tsx` 用
  `editor.commands.setContent` 装载文档，这一步进了 ProseMirror 的撤销栈。属既有缺陷（任何一次插入都踩得到，
  与交叉引用无关），未夹进本次提交，登记为排名 3 的收尾项：装载文档应发一条
  `setMeta("addToHistory", false)` 的整篇替换事务。
- **"用 Word 打开一遍"仍 `[!] BLOCKED`**：本机没有 Word / WPS / LibreOffice，只做到读回一致。

## 四、排名 4：批注 / 修订跟踪

一手来源（每条都实际打开过，`learn.microsoft.com` 按元素名猜 URL 会 404，别再试）：

- ECMA-376 Part 4 的镜像页给出内容模型：`commentRangeStart`/`commentRangeEnd` 属于
  `EG_RangeMarkupElements` ⊂ `EG_RunLevelElts` ⊂ `EG_PContent`，**与 `w:r` 平级、排在 `w:pPr` 之后**，
  而且 body 层也允许，所以一个批注区间可以跨段
  （[commentRangeStart](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_commentRangeStart_topic_ID0EFJMV.html)）；
  `commentReference` 属于 `EG_RunInnerContent`，**必须待在 `w:r` 里**
  （[commentReference](https://c-rex.net/samples/ooxml/e1/Part4/OOXML_P4_DOCX_commentReference_topic_ID0E4PNV.html)）。
  规范原文的样例就是我们要写回去的形状：
  ```xml
  <w:p>
    <w:r><w:t xml:space="preserve">Some </w:t></w:r>
    <w:commentRangeStart w:id="0" />
    <w:r><w:t>text.</w:t></w:r>
    <w:commentRangeEnd w:id="0" />
    <w:r><w:commentReference w:id="0" /></w:r>
  </w:p>
  ```
- `w:comment` 是 `CT_Comment` extends `CT_TrackChange`：`w:id` **必填**、`w:author` **必填**（≤255 字）、
  `w:date` 可选、`w:initials` 可选（≤9 字），内容是 `EG_BlockLevelElts`——**批注体可以放整段甚至表格**
  （[Comment Class](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.comment?view=openxml-3.0.1)）。
- 部件注册：content type `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml`，
  关系类型 `…/officeDocument/2006/relationships/comments`，根元素 `w:comments`
  （[Open-XML-SDK 的部件清单](https://raw.githubusercontent.com/dotnet/Open-XML-SDK/main/data/parts/WordprocessingCommentsPart.json)）。
- 修订：`w:ins`/`w:del` 是 `CT_RunTrackChange`，包的是**运行块而不是裸文字**；删除的文字换成 `w:delText`
  （同一份 `CT_Text`，`xml:space` 是约定不是强制）
  （[DeletedRun](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.deletedrun?view=openxml-3.0.1)）。
  段落标记自己的增删写成 `w:pPr/w:rPr/w:ins|w:del`，`CT_ParaRPr` 里它们排在 `EG_RprBase` **之前**；
  接受一次段落标记删除 = 把本段并进**后一段**（本段 `pPr` 留下）
  （[wml_xsd 内容模型](https://schemas.liquid-technologies.com/officeopenxml/2006/wml_xsd.html)、
  [Accept all revisions](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-accept-all-revisions-in-a-word-processing-document)）。
- `w:ins`/`w:del` 属于 `EG_RunLevelElts`，而 `EG_RunLevelElts` ⊂ `EG_PContent` ⊂ `CT_Hyperlink`，
  所以 `<w:hyperlink><w:ins><w:r>` 这个方向有内容模型依据；**反方向（`w:hyperlink` 出现在 `w:ins` 里）
  查不到一手来源**（ECMA-376 在线镜像对本机返回 403），因此写回一律取有依据的一侧。
- 注解 id 是**整个文档共享的一个编号空间**（批注、书签、`customXmlInsRangeStart` 都用它），规范对流只有一条
  硬约束："同一个 id 出现多条批注时 Word 只加载一条，其余忽略"。`w:date` 缺失合法；重复 id 的容错行为**无文档可查**。
- `w15:people` / `w15:commentEx`（已解决状态）/ `w16cid:commentId` 都是**微软扩展不是 ECMA**。Word 自己会生成，
  但"第三方生产者省略它们会怎样"没有任何官方说明——**未证实**，所以我们的取舍是：读到的原样带走，不主动生成。

判据与取舍（这些是刻意的，不要当成待修项）：

- **批注在编辑器里建模成三个 inline 节点，不是 `comment` mark**：`commentRangeStart {id}` /
  `commentRangeEnd {id}` / `commentReference {id}`，与既有 `bookmarkStart`/`bookmarkEnd`、
  `footnoteReference` 同形（上面内容模型那条里它们本来就是 `w:p` 的直接子元素、与 `w:r` 平级）。
  改用 mark 是错的三条理由：mark 会被复制/剪切带走，一份批注被粘到别处就成了同 id 的第二个区间，
  而 id 是文档级唯一命名空间（见上面"注解 id 是整个文档共享的一个编号空间"那条）；按 mark 集合干活的既有功能（格式刷、样式抓取）
  会被一个"不是字符样式"的 mark 污染；PM 的 mark 跨块边界本就别扭，批注区间恰恰常常跨段。
  代价是区间归属要在读侧自己配对（start/end 各落在自己那段），写侧照原位置各写一个空元素——
  这与书签已经做过的事完全一样，不算新增复杂度。**空区间**（start 紧跟 end）也能自然表示，
  不需要退化处理。
- **`w:commentReference` 落 PM，形状照 `footnoteReference`**：它同样是"引用点"而不是文字，
  读侧消费掉外面那个 `w:r`（那份 `w:rPr` 是 Word 的呈现样式，不进模型），写侧固定输出
  `<w:r><w:commentReference w:id="N"/></w:r>`。批注编号的呈现由编辑器自己算（脚注那边有
  `FootnoteNumbering` 插件的先例），不跟 Word 的 `AnnotationReference` 样式较劲；② 只把芯片与
  区间端点画出来，编号与跨段高亮要跟侧栏一起在 ⑤ 做——那一步才需要"点条目跳到锚点"的位置索引。
- **批注体走脚注那套透明部件**（`footnotes.ts` 的形状：按 id 存 `{ rawXml, dirty }`，没改过的整段原样带走，
  改过的才用 `pmToOoxml` 片段重建），复用 `DocxFile.ensurePartReference` 注册 content type 与关系
  （§3.3/§3.4：引擎层加一个 `comments.ts`，不新建包、不改保存主链路）。
- **modlog 不是修订**（AGENTS §2.5 的本地历史不进文件）。两者互不干扰：`w:ins`/`w:del` 是文档内容，
  op 的逆操作要把标记一起撤销；"接受修订"是一次普通改动块，照样可撤销、照样进修改历史。
- 修订的**读侧保真先于写侧**：③ 之前 `w:ins`/`w:del`/`w:delText` 在正文遍历里根本没有分支，导入一份带修订的
  `.docx` 会把删除的文字整块丢掉或把插入的文字当普通文字——**这是既有缺陷**，比"能不能开修订"更严重，
  所以 ③ 排在呈现与命令之前。
- 段落标记的 `w:ins`/`w:del` 在 `w:pPr/w:rPr` 里，**没有运行可走**，只遍历 `w:r` 的读侧一定会漏
  （见上面"段落标记自己的增删"那条）。③ 的处理是**继续靠 `rawPpr` 原样透传**，不做成可见的段落分隔符，
  也不给编辑器加"本段的段落标记是插入/删除的"这一层——但透传只活在 docx-io 层，**过了编辑器就没了**：
  真实 app 探针里 `w:pPr/w:rPr/w:ins w:id="10"` 打开→保存一次后消失（`ooxml-pm.test.ts` 的同名往返用例
  仍然绿，说明丢在编辑器边界而不是转换层）。文字内容不受影响，丢的是"这段的段落标记也是插进来的"，
  后果局限在 Word 里**拒绝**那次修订时留下一个空段而不是并掉两段。
- `rawPpr` 到编辑器就断是**一片独立的缺口**，不是修订特有的：Tiptap 的段落 schema 没有这个属性，
  PM 建节点时不认识的就直接不保留，所以任何"靠透传保真"的段落内容都过不了"打开→保存"这一趟。
  补法是把 `rawPpr` 注册成段落的非渲染全局属性，但它会同时改变排名 1 / 5 / 7 各条的透传结局，
  按 §1.4 单独成片，不夹进修订的提交里。
- **修订建模成标记而不是节点**：一次修订只覆盖段落里的一段文字，文字必须留在原位；做成 `ins`/`del`
  节点会把段落劈开，还得重新定义跨块语义。标记形状 `{kind, id, author, date}` 进 `PMMark` 联合（§5.3），
  四者都 `rendered: false`——归属信息不是 DOM 契约，靠 `editor.getJSON()` 落盘。
  `parseHTML` 返回空数组：粘贴进来的内容不该带着别人的修订归属。
- **一次修订 = 一个 `w:ins`/`w:del` 元素**：Word 常把一次修订拆成几个 run（不同 `w:rPr`），写回时按
  attrs 全等的相邻分组并回同一个元素、内部留多个 `w:r`。不并的话同一份文档里会出现两个同 `w:id` 的
  元素，而重复 id 下 Word 的"接受此修订"会不会只清掉一半**无文档可查**（见上面"注解 id 共享编号空间"那条），
  这里按"不制造重复 id"保守处理。
- **嵌套取外层**：`del` 里套 `ins` 只留外层（外层是后发生的操作），内层归属丢弃、层级不建模；
  两档同时挂在一个字上，OOXML 的这一层无处表达。
- **`w:author`/`w:date`/`w:id` 读进来原样带走，一个都不编**：文档没写作者就只显示"插入"两个字，
  不补一个名字；`w:id` 保持字符串，不在这里做数字换算（§2.1）。
- **接受 / 拒绝按整个文档的修订区间生效**：一次修订可能被 Word 按段落拆开，只处理选区那一段的话，
  接受完还剩半条。选区没碰到任何修订时命令返回 `false`，菜单据此置灰。

提交切分（实现顺序，逐条落地后回填提交号）：

| # | 内容 | 提交 |
| --- | --- | --- |
| ① | docx-io：`word/comments.xml` 透明部件（按 id 存、没动过原样带走）+ 正文三个锚点元素读写往返 | `5b32ea8` |
| ② | desktop：三个批注节点注册进 `EDITOR_EXTENSIONS`（区间端点画〔〕、引用记号画"批"字芯片），live 用例 | `52b66da` |
| ③ | docx-io：`w:ins`/`w:del`/`w:delText` 读成一个 `revision` 标记、写回按修订区间合并成单元素；段落标记的增删仍靠 `rawPpr` | `5197192` |
| ④a | desktop：标记注册进 `EDITOR_EXTENSIONS`（插入下划线 / 删除删除线 + 作者与时间）+ 接受 / 拒绝命令 | `bee11e4` |
| ④b | desktop：接受 / 拒绝的右键菜单入口 | `ced2499` |
| ⑤ | 批注侧栏 UI：列出、新增（选区右键）、删除、跳转到锚点；写侧的 `w:date`/`w:author` 取设置里的作者名 | 待做 |
| ⑥ | 修订开关：开了就把本次编辑写成 `w:ins`/`w:del` | 待做 |

③④a 的实测（同一套 CDP 探针，独立 `--user-data-dir`）：一份含一次插入（内部两个 run，一个带 `w:b`）
与两处删除的 `.docx` 让 app 打开后，编辑器里 4 个 `span[data-revision]` 全在位——
插入是绿色下划线、删除是红色删除线，`title` 分别写着"插入：张三（2026-01-02T03:04:05Z）"和
"删除：李四（2026-01-03T00:00:00Z）"；点保存后 `document.xml` 里 `<w:ins w:id="1" …>` 与
`<w:del w:id="2" …>` 各回来一条、被删的字仍是 `w:delText`、插入那条内部两个 `w:r`（`<w:b/>` 那位没被
拆散），`w:id="1"` 只出现一次（没有重复 id）。第二段那条 `<w:del w:id="2">` 与第一段同 id 是** fixture
故意造的"一处修订被 Word 按段落拆开"**，同 id 才让"接受一次全清"成立，不是合并漏了。
**唯一丢的是段落标记那条 `w:pPr/w:rPr/w:ins w:id="10"`**，原因与影响见上面"段落标记"那条判据。
用 Word / WPS 打开验证：本机没装，**BLOCKED**。

③④a 之后仍然没建模的：移动类修订（`w:moveFrom`/`w:moveTo`/`w:moveFromRangeStart`）读侧没有分支，
移动过的文字按普通文字收下、写回不再带移动标记——真要补是给同一个标记加一个 `kind` 值，不新开机制。

④b 的实测（同一套 CDP 探针，右击用 DOM 上的 `contextmenu` 事件合成）：右击落在插入的文字上，菜单里
"接受修订 / 拒绝修订"两条都可用，点接受后**两个 `ins` 片段一起清掉**（同一次修订的两个 run），删除线
那两条原样留着；右击落在没有修订的正文上，两条都在但 `disabled`；右击落在删除的文字上点拒绝，
被删的字**留下**且标记清空。存盘后 `document.xml` 里 `<w:ins>`/`<w:del>` 一条都不剩、正文一字未少，
带 `w:b` 的那个 run 仍带着自己的 `w:rPr`。用 Word / WPS 打开验证：本机没装，**BLOCKED**。

①② 的实测（本机 macOS，桌面 app 独立 `--user-data-dir` + CDP）：打一份含两条批注（其中一条区间跨段、
一条正文里有 `w:br`）的 `.docx` 让 app 打开，编辑器里 6 个锚点全在位（`〔文字〕批` 跨两段各一组），
点保存后 `word/comments.xml` **逐字节等于原部件**，正文里 6 个锚点元素重写回来，`[Content_Types].xml`
补上 `comments+xml` 登记。用 Word / WPS 打开验证：本机没装，**BLOCKED**。


## 五、排名 5：制表位 + 前导符

官方：[在制表位之间显示点或前导符](https://support.microsoft.com/zh-cn/office/c80b62d0-2244-4078-95bd-8c82f91440ba)。
OOXML：`w:pPr/w:tabs/w:tab w:val="left|center|right|decimal|bar|list|clear" w:leader="none|dot|hyphen|dash|line|heavy|middleDot" w:pos="twips"`。
`[TabStop Class](https://learn.microsoft.com/zh-cn/dotnet/api/documentformat.openxml.wordprocessing.tabstop?view=openxml-2.15.0)` 给出属性集。

要点：段落属性 `tabStops: {positionPt, alignment, leader}[]`（位置用 pt，导出 ×20）；
编辑器里 `Tab` 键已经能产生 `w:tab` 字符，缺的是**制表位定义**与"Tab 字符按制表位对齐"的
渲染（Tiptap 里用 `white-space: pre-wrap` + 每段生成 `background-image`/`text-indent` 代价太大 →
决定：制表位只做数据层与导出，编辑器内 Tab 仍按浏览器默认宽度呈现，并在规格里如实写明这个差距）。
UI 挂在「段落」面板下方（对齐 Word 的"制表位…"按钮），支持"清除/全部清除"。

提交切分（实际落地顺序与提交号）：

| # | 内容 | 提交 |
| --- | --- | --- |
| ① | 数据层：`TabStop` 类型 + `readTabStops`/`tabsElementXml` 往返（读不懂整组退回透传，空数组=全部清除）+ 编辑器段落属性 `tabStops` | `8c37fb4` |
| ② | docx-io：`cmToPt`/`ptToCm` 厘米口径（换算仍只在进出口） | `e13c970` |
| ② | desktop：段落浮层里的「制表位…」编辑区（列出/新增/修改/清除/全部清除）+ 三条命令 + 与目录对话框共用的前导符与对齐词表 | `3dfe85f` |
| ③ | 「更新样式以匹配所选内容」把段落上直改的制表位写进样式定义 | `f3c4324` |

①排在排名 3 的目录生成（③.2）之前落地：Word 的目录条目靠**右对齐 + 点前导符**制表位把页码
顶到右边距，没有这一层，我们生成的目录在 Word 里就是一行挤在一起的"标题页码"。

②的做法（与最初设想的差别）：
- Word 把它做成**独立对话框**，这里折进「段落」浮层底部展开：另开一层窗口会抢走编辑器选区，
  而制表位恰恰要求用户看得见自己在给哪一段定位。
- 位置按**厘米**录入（Word 的口径），存进 attrs 仍是磅；换算用 docx-io 的 `cmToPt`/`ptToCm`，
  与既有的 `cssPxToPt` 同一处（§2.1）。厘米取两位小数就是 Word 的显示精度，
  415.3pt ↔ 14.65cm 双向不漂。
- 「设置」是**逐段并入**（`withTabStop`：同位置整条替换，其余留着），不是拿面板里那张清单整份覆盖
  ——后者会静默删掉别段落在别处的制表位，属 §2.3 的拍平。「清除」删那一条，遇到没建模过的段落仍写
  `undefined`（不凭空删用户没写过的 `w:tabs`）；只有「全部清除」落成空清单，导出据此删掉段落自己的
  `w:tabs`。三条命令都走 `mapFormatBlocks`，光标没拉选区也作用于当前段（live 用例覆盖）。
- 前导符选「无」时**不写** `w:leader`：Word 自己也是省略的，写 `"none"` 会让两次保存的字节不一致。
- 对齐方式显示全七档（`left/center/right/decimal/bar/list/clear`），候选只给五档：`list` 是编号生成的
  列表制表位、`clear` 是"取消的制表位"，都只在导入里出现，界面显示得出来却不给造（选中项 `disabled`）。
- 位置认位用 0.1 磅取整比较（`sameTabStopPosition`）：读写两侧都是 0.1 磅精度，不需要容差。

③的做法：`styleCapture` 的建模键补一个 `tabStops`，样式定义那侧**零新代码**——它和段落共用
`modeledPprParts`，所以抓来的 `w:tabs` 顶掉样式原有的那条、选区是「全部清除」时空清单把它删掉，
两条语义与段落完全一致。代价是段落属性的比较要从按引用比改成按内容比：制表位是对象数组，
PM 里每段各持一份，按引用比会把两段一模一样的清单判成冲突而拒绝更新。

验收与差距：
- 实跑（真实 app + CDP 真实鼠标事件，独立 `--user-data-dir`）：导入的 `w:pos="8306"` 显示成
  "14.65 厘米 右对齐 点线 …"；载入清单后三个控件回填 `14.65/right/dot`；「清除」→ `[]`、⌘Z 复原；
  「签名行」段（原本没建模过）设 2 厘米 → `[{56.7, right, dot}]`，清单同步；列表位显示"0.74 厘米
  列表制表位 （无）"且候选项禁用；保存出的 XML 是 `<w:tab w:val="right" w:leader="dot" w:pos="1134"/>`
  （2cm = 1134 twips），全清那一段连 `pPr` 都不留。产物再用 docx-io 读回：`positionPt 56.7` 原值回来，
  警告为空。
- `[!] BLOCKED`：本机没有 Word / WPS，"用 Word 打开一遍"仍未实测；前导符与制表位对齐的**视觉**效果只到
  "文档里写对了"这一层，编辑器内 Tab 仍按浏览器默认宽度呈现（① 就定下的差距，没有因为这次加 UI 而变）。
- ③ 实跑（真实 app + 带 `styles.xml` 的样例）：光标停在有直改 `w:tabs` 的那一段 → 样式窗格点「签名行」
  → 「更新样式以匹配所选内容」→ toast「已按所选内容更新样式「签名行」，保存后写进文档。」；
  保存出的 `styles.xml` 里那条样式的 pPr 变成 `<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="8306"/></w:tabs>`
  （原本的 `left`/`720` 被顶掉，`w:basedOn` 与段落自己的直改制表位都留着）。
- 抓取口径的一处已知粗糙：段落把制表位留在没建模过的透传里时（`rawPpr` 带着 `w:tabs`），
  这一项抓不出来——与其它没建模的段落属性同一口径，不是这条切片专门要修的东西。
- 已知保真边界：`[]`（全部清除）导出后段落没有 `pPr`，再读回来是 `undefined`——两者对 Word 等价
  （样式给的制表位照常生效），但 UI 上第二次打开这一段的清单显示为"（这一段没有制表位）"。


## 六、排名 6：上标 / 下标

OOXML：`w:rPr/w:vertAlign w:val="baseline|superscript|subscript"`，ECMA-376 Part 4 §17.3.2.27 的
`ST_VerticalAlignRun`（本机经 c-rex.net 的 ECMA-376 镜像页取回原文；`learn.microsoft.com` 上按元素名猜的那条
URL 是 404，别再试）。两档抬起/沉下的语义含"有则缩小"，`baseline` 是缺省档位**但会被显式写出来**用来取消样式
继承的上标——所以读侧是分三种情况的：写过 baseline / 没写过 / 写了另两档。

做法（与本节最初设想的差别，说清免得下次又按旧文案实现）：
- 原计划"把 `vertAlign` 当 `TextStyle` 的一个属性"，落地时改成**独立的 `vertAlign` mark 带一个 `val` 枚举**：
  OOXML 那边就一个元素一个值，两个开关 mark 能同时挂成"又上标又下标"；`TextStyle` 的属性都是能并存的维度
  （字体、字号、颜色），塞进去反而要额外写互斥。也不装 `@tiptap/extension-superscript`/`subscript`（§6.4 的问答：
  现有依赖能渲染，但它是"mark 名即档位"的两个 mark，且不带 `data-*`，覆盖不了这个语义）。互斥由 `tr.addMark`
  同类型顶掉白送。
- 档位 `val` 是 `rendered: false` 的属性：DOM 里只有标准的 `<sup>`/`<sub>`，复制/打印出去不带非标准属性，
  落盘走 `editor.getJSON()`，attrs 照样带得出。
- 读侧只把 `superscript`/`subscript` 落成标记，`baseline` 等于"没有这个标记"；`raised` 之类不认识的值整个丢弃
  （这份 rPr 还会进样式定义，塞一个 Word 不认的值比重丢更糟）。
- 写入仍只有 `runPropertiesParts` 一个点，`RPR_ORDER` 把 `vertAlign` 排在 `shd` 之后（CT_RPr 的元素顺序），于是
  直接格式、样式定义、页眉页脚预览三条链路共用同一份实现（§3.2/§3.3）；页眉页脚预览里原来那两条重复的
  if 链顺手收成一张 `inlineTagOf` 表（净代码下降）。

提交切分（实际落地顺序与提交号）：

| # | 内容 | 提交 |
| --- | --- | --- |
| ① | docx-io：`VerticalAlign` 类型 + `w:vertAlign` 读写往返（三档、未知档位丢弃、两档互不顶掉、元素顺序） | `0cd9254` |
| ② | desktop：`vertAlign` mark + 工具栏「上标/下标」+ 注册进 `EDITOR_EXTENSIONS`，格式刷、样式抓取、页眉页脚预览都认它 | `1bf35a2` |

取舍：
- **不写 `w:sz`**。缩小是排版行为（实测：正文 14px 的段落里角标算得 11.6667px），文档里发明一个字号等于改用户
  的字号，也违反 §2.1 的换算单点。代价是 Word 里缩多少由它的排字规则决定，不由我们写死。
- `baseline` 不进编辑器，也不产出 `w:val="baseline"`：编辑器里没有"取消继承"这个概念。样式窗格那条链路将来要
  显式取消时再补。
- AI 提案与模板套版的字符样式词表（`CHAR_STYLE_KEYS`）**不加**角标：套版对它是"既不投票也不写"，所以整段套版
  不会拍平用户手打的 H₂O（live 用例覆盖）。这是有意的最小范围（§3.6），不是遗漏。
- **样式定义里的 `vertAlign` 预览不出来**：CSS 的 `vertical-align` 不继承，文档作用域样式表只能预览可继承的
  字体/字号/颜色。这条差距记录在案，不在本节修。
- 没做键盘加速器（Word 是 `Ctrl+=` 上标、`Ctrl+Shift+=` 下标），本节验收面是工具栏入口与文档保真。

验收与差距：
- 往返用例：`packages/docx-io/src/ooxml-pm.test.ts` 新增 3 条（三档读法 + 未知档位丢弃；往返字节一致且两档互不
  顶掉；`vertAlign` 排在 rPr 末尾）。
- live 用例：`apps/desktop/src/editor/docBlocks.live.test.ts` 新增 2 条（真实 schema 收下两个档位、档位不被
  mark 缺省值抹平；整段套版只补字体颜色，下标原样留着）。
- 标记单测：`apps/desktop/src/editor/extensions/VertAlign.test.ts` 7 条（缺省档位、`<sup>`/`<sub>` 各认一档、
  渲染选标签且不挂 `val`、换档后只剩一条、同档再点是取消）。
- 样式抓取：`styleCapture.test.ts` 新增 2 条（下标带档位进样式定义；上标与下标混排报 `mixed-run-format`，
  detail 是 `vertAlign`，即 §2.3 的不拍平）。
- 门禁（2026-10-11 实测）：`pnpm build` 退出 0，`pnpm test` 退出 0，逐包核对通过数——docx-io 214、desktop 231、
  server 2、word-client 8、word-tools 6；`packages/shared`、`apps/cli`、`apps/task-pane` 仍无测试文件（`passWithNoTests`）。
- 真实 app 实跑（独立 `--user-data-dir` + CDP 探针，2026-10-11 实测，探针产物只在 `tmp/` 且已删）：打开含
  `subscript` 与 `superscript` 两段样例 → 屏幕上分别是 `sub(2)`/`sup(2)`，算得角标 11.6667px、正文 14px、
  `vertical-align: sub`；选中下标点「上标」→ 按钮类名带 `active` 且变成 `sup(2)`，再点 → 类名回到 `icon`（取消），
  第三点 → 又 `sup(2)`；另一段点「下标」→ `sub(2)`；存盘后 `word/document.xml` 里是
  `<w:vertAlign w:val="superscript"/>` 与 `<w:vertAlign w:val="subscript"/>`，全文没有 `w:sz`，DOM 里也没有残留的 `val=`。
- `[!]` 用 Word / WPS 实际打开一遍：本机没装，**BLOCKED**。上面是"读回字节一致 + 浏览器真实渲染"，不是 Word 实测。


## 七、排名 7：分栏 / 水印 / 页面边框 / 首字下沉 / 行号

- 分节：`w:p/w:pPr/w:sectPr`（分节符）与 body 末尾的 `w:sectPr`。`pmToOoxml`（`:291-301`）在 body
  末尾写一个默认 `w:sectPr`，读侧遇到 body 级 `w:sectPr` 直接 `continue` 丢弃（`:264`）——
  节属性走的是 `section` 载荷而不是 PM 文档，所以单节文档靠 `replaceProseMirror` 把原节点换回来。
  段内分节符（`w:pPr/w:sectPr`）目前只靠 `rawPpr` 透传，经编辑器会丢（见下面"①的实测"）。
- 分栏（**②已落地**）：`w:cols w:num="2" w:space="twips"` + 每栏 `w:col`，承载在 `SectionGeometry.columns`
  上（`{count, gap, equalWidth, widths}`，gap/widths 是 CSS px @96DPI），入口就是「页面设置」面板。
- 页面边框（**③已落地**）：`w:pgBorders`（`CT_PageBorders`，父元素 `sectPr`，§17.6.10）承载在
  `SectionGeometry.pageBorders` 上：`{offsetFrom, top, left, bottom, right}`，每条边
  `{style, sizePt, spacePt, color}`。线宽 `w:sz` 是**八分之一 pt**、间距 `w:space` 是 **pt**（§17.3.4），
  两者就按 pt 存在模型里，不再换算成 cm。水印：`w:headerReference` 里的 `v:shape` 文字（Word 用
  WordArt VML，本仓可以退化成 `w:pict` 文本框，需在规格里明确"Word 打开效果待验"）。
- 首字下沉：`w:pPr/w:framePr w:dropRows="3" w:line="…" w:hAnchor="text" w:vAnchor="text"`。
- 行号（**③已落地**）：`w:lnNumType w:countBy w:distance w:start`（`CT_LineNumber`，父元素 `sectPr`，
  在 §17.6.1 的子元素序列里排在 `pgBorders` 之后、`cols` 之前），
  承载在 `SectionGeometry.lineNumbers` = `{countBy, distance(CSS px), start}`；`w:restart` 本片**不建模**，
  靠"值等价就整段不重建"原样留着。

提交切分：① 节属性读写保真（`bba064e`）；② 分栏（`b93c1ef` `a806d69`）；
③ 页面边框 + 行号（`587a9c9` `96afc34`）；④ 首字下沉；⑤ 水印。

> ①的判据与实测（`bba064e`）：⑤.1 期间探针看到的"**打开后直接保存，`<w:sectPr>` 整块消失**"
> 根因不在读侧丢弃，而在 `buildSectPr`：它用 `{ ...existing }` 复制解析出来的节点，而
> `parseDocument` 的节点只在**原型**上带 `type`，展开拿不到，`serialize()` 的 `isElement`
> 判不出来就返回空串。桌面 `saveDocx` 每次保存都会把打开时读到的页面设置
> `setSectionGeometry()` 带回去（`apps/desktop/electron/fileHandlers.ts:285`），于是这一步
> 每次都走到那条坏分支——整块节属性被一个"无名元素"替换掉。这也解释了为什么
> `docx-io` 只调 `replaceProseMirror` 的直接读写路径看着是好的：那条路径不重建节属性。
> 顺带修的第二处：`w:pgMar` 原本缺失时**追加在末尾**，会排到 `w:cols` / `w:docGrid` 之后，
> 违反 `CT_SectPr` 的顺序（ECMA-376 §17.6.1 把 `pgMar` 定在 `pgSz` 之后、`pgBorders`/`lnNumType`/
> `cols`/`docGrid` 之前），现在插在 `pgSz` 后面。
> 实测两份：`docx.test.ts` 里复刻 app 保存路径（open → toProseMirror → setSectionGeometry →
> replaceProseMirror → toBytes）的两条用例，断言未建模的 `w:cols` / `w:pgBorders` /
> `w:lnNumType` / `w:docGrid` 逐字留下、顺序不动、二次保存字节一致；真实桌面 app（独立
> `--user-data-dir` + CDP，`tmp/sect-app-probe.mjs`）打开一份带全字段的 `.docx` 直接点保存，
> 读回来的 `document.xml` 里
> `<w:sectPr><w:pgSz w:w="11906" …/><w:pgMar …/><w:pgBorders …>…</w:pgBorders><w:lnNumType …/>`
> `<w:cols w:space="425" w:num="2"/><w:docGrid …/></w:sectPr>` **与打开时一致**，正文两段一字不缺。
> 修之前的同一探针：`document.xml` 里 `<w:sectPr` 命中 0 次、`w:pgSz` 也 0 次（段内分节符那条
> 因为不走 `buildSectPr` 才活着）。
> 仍然没解决的：段内分节符（分节符那一段的 `w:pPr/w:sectPr`）靠 `rawPpr` 透传，而
> `rawPpr` 在 Tiptap 的段落 schema 里没有注册，所以**经编辑器**的多节文档仍会被拍成单节——
> 这条与排名 4 记下的 `rawPpr` 缺口是同一个，属独立一片，不夹进本节。
> 用 Word / WPS 实际打开验证：本机未安装，**BLOCKED**。

> ②的判据与实测（`b93c1ef` `a806d69`）：分栏建模成 `SectionGeometry.columns`
> （`{count, gap, equalWidth, widths}`，`gap`/`widths` 是 CSS px @96DPI，进出口才换算 twips），
> 沿节属性走 `section` 载荷，不进 PM 文档。读侧只在真的存在 `w:cols` 时才赋值，**不拿
> `count: 1` 冒充单栏**——否则一份没有 `w:cols` 的文档保存后会被凭空加一节属性。写侧
> `w:num` / `w:space` 必写，`w:equalWidth="0"` 与每栏 `w:col w:w` 只在真的是不等宽时才写。
> 节属性插入位置从 ① 的"pgMar 紧跟 pgSz"特例统一成 `placeInSection` + `SECT_PR_ORDER`
> （ECMA-376 §17.6.1 的 `CT_SectPr` 子元素序列），③ 要加 `w:pgBorders` / `w:lnNumType` 时不必再改这段。
> `w:cols` 上还有个没建模的属性 `w:sep`（栏间分隔线），处理规则是**值等价就整段不重建**
> （`sameColumns`，0.5px 容差吃掉 px↔twips 的往返误差），只有用户真的改了分栏才重写，
> 与 Word 自己的行为一致（AGENTS §2.7：不擅自改动用户已有的版式意图）。
> 桌面入口就在「页面设置」面板：栏数 1–13、栏间距按 cm 显示，1 栏时间距控件禁用；
> `geometryFor()` 多接一个 `columns` 参数，顺手删掉 `PageSetupDialog` 与 `PageChrome` 里各自
> 复制的 `SectionGeometryLike`（同一段结构第三份复制，AGENTS §3.2）和 `App.tsx` 里两处
> `as never`。套模板/美化时把现有 `columns` 原样带过去——模板只改页边距。
> 编辑视图**只画栏边界导引线**（虚线，`pointer-events: none`），不做文字重排：分页是
> `PagedLayout` 逐个量顶层块高度算出来的，给可编辑层加 CSS 多栏会让测量全部失效。
> 真正的分栏落在打印/PDF 链路（`printHtml` 的 `column-count` + `column-fill: auto` + 固定栏高），
> 那才是交付物。
> 实测（`tmp/cols-app-probe.mjs`，独立 `--user-data-dir` + CDP 起真实 app）：打开单栏文档时面板
> 初始值 `栏数=1 / 间距=1.27cm / 间距禁用`；改成 3 栏、间距 0.8cm 后提示"每栏 ≈ 4.35 cm"，
> 画布上两条导引线 `left 314.711px / 509.222px`（= marginLeft 120 + 194.7 / 389.2），保存读回
> `<w:cols w:space="453" w:num="3"/>`（0.8cm→30.2px→453twips 对得上）。
> 顺带发现（**没有改，交给你决定**）：桌面原地保存不产生 `.bak`，只有 CLI 会备份
> （`apps/cli/src/fsutil.ts:38`），和 AGENTS §8 安全项 4"原地写回默认先出 `.bak`"不一致。
> 另一个观测：页面设置重应用一次后 `w:pgMar` 的 1800 变成 1803 twips，是 cm↔px 往返的
> 半 twips 误差，属预期，不影响 Word 打开。
> 用 Word / WPS 实际打开验证：本机未安装，**BLOCKED**。

> ③的判据与实测（`587a9c9` `96afc34`）：先补外部取证（Microsoft Learn 的 OpenXML XML-Schema 参考，
> 你已明示许可）确认三件事才动手——`CT_PageBorders` 的父元素与 `w:offsetFrom` 取值、
> `CT_Border` 里 `w:sz` 是**八分之一 pt**而 `w:space` 是**pt**（不是 twips）、`ST_Border` 的字面量清单
> （`single/thick/double/dotted/dashed/dotDash/dotDotDash/triple/thinThickSmallGap/thickThinSmallGap/wave…`），
> 以及 `CT_LineNumber` 只有 `countBy/distance/restart/start` 四个属性、ISO 示例
> `<w:lnNumType w:countBy="1"/>` 说明 `restart`/`start` 都有缺省。
> 模型：`SectionGeometry.pageBorders`（四条边各 `{style, sizePt, spacePt, color}`，`w:val` 为
> `none`/`nil` 的边读成 `null`）与 `lineNumbers`（`distance` 是 CSS px，`w:distance` 的 twips 换算
> 只在 docx-io 进出口发生，§2.1）。三态语义沿用 ②：**`undefined` = 这次保存不涉及，元素一动不动；
> `null` = 用户关掉，删掉这条元素；有值 = 值等价就不重建**。这条规则这次被抽成通用的一处
> `placeOrClear(arr, name, wanted, parse, same, build)`——它是 ② 那条 cols 分支的第三次出现（§3.2），
> 顺带删掉 `buildSectPr` 里"从头造一节"的重复分支，`existing` 有无只影响初始数组。
> 没建模的属性（`w:pgBorders` 上的 `w:display`/`w:zOrder`、每条边的 `w:shadow`/`w:frame`/主题色、
> `w:lnNumType` 的 `w:restart`）靠这条规则原样穿过保存；一旦用户真改了边框就整条重建，`w:display`
> 会随之消失——和 Word 改边框时重写该元素的行为一致，与 ② 的 `w:sep` 同一取舍。
> 颜色复用 `ooxml-pm` 的 `colorVal`（改成导出，不再写第二份 OOXML hex→CSS 转换，§3.1）；
> `samePageBorders` 比色时忽略大小写，因为 `<input type=color>` 交回 `#ff0000` 而文件里常写 `FF0000`，
> 这不算用户改动。容差：线宽 1/16 pt、间距 0.5pt、行号距离 0.5px，都是单位往返的误差量级。
> 桌面侧把线型→CSS、pt→px、偏移基准（距页面 / 距正文）三份换算收进
> `apps/desktop/src/editor/pageBorders.ts` 一处，对话框、画布与 `printHtml` 共用（§3.3），
> `geometryFor` 的第 6 个参数由 `columns` 改成 `extras` 对象（`{columns, pageBorders, lineNumbers}`），
> 套版路径把三项一起带走——模板只改页边距（§2.7）。行号**不在画布上逐行编号**：分页是
> `PagedLayout` 按块高度算的，逐行编号要把每个文本行映射到屏幕行，成本远超本片收益；
> 数据仍完整写进 `.docx`，由 Word/WPS 渲染。
> 实测（`tmp/border-app-probe.mjs`，独立 `--user-data-dir` + CDP 起真实 app）：一份既无
> `w:pgBorders` 也无 `w:lnNumType` 的文档，打开时画布 0 条边框、对话框四条边都是"无"、
> 行号未勾选且"每几行一个号"禁用；给上边选"虚线"（线宽/间距留缺省 0.5pt/1pt）+ 勾行号 +
> 每 2 行 + 距正文 0.5cm → 确定后画布出现 1 条 `.oc-page-border`
> （`top 1.33333px`、`width 793.733px`、计算样式 `dashed 1px rgb(0,0,0)`），保存读回
> `<w:pgBorders w:offsetFrom="page"><w:top w:val="dashed" w:sz="4" w:space="1" w:color="auto"/></w:pgBorders>`
> `<w:lnNumType w:countBy="2" w:distance="284"/>`（½pt→`sz=4`、0.5cm→18.9px→284twips 都对得上），
> 正文两段一字不缺；**再点一次保存整份文件字节一致**且两条元素各只有一条，说明读回的值被判为
> "没改"、没有重复重建。探针里同样观测到 ② 已记录的副作用：点确定会写 `<w:cols w:num="1"/>`、
> `w:pgMar` 的 1800 变 1803。
> 用 Word / WPS 实际打开验证：本机未安装，**BLOCKED**（线宽/间距的单位与 `ST_Border` 字面量
> 取自微软 XML-Schema 参考页，不是本机实测）。

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

### 8.2 排名 2（窗格部分）验收记录（`21cb2e9` `aeed0be` `58220ff` `91131d0` `aa0d848` `7cb35b5` `481c4ec`）

- 往返用例：`docx.test.ts` 的"hands the style pane the paragraph styles this document declares"
  （有 `styles.xml` → 过滤 character/hidden/semiHidden、按 `uiPriority` 排、带上 `headingLevel` 与
  `isDefault`；没有 → `[]`）；`ooxml-pm.test.ts` 的"编辑器里新建的标题也先问样式表，别写本文档没声明的
  `Heading2`"；`styles.test.ts` 的 `paneStyles()` 4 条 + `fitStyleId` 换级/转正文/无样式表三种结局。
- live 用例：`docBlocks.live.test.ts` 的"具名样式引用 styleId（真实 schema）"6 条——三态往返、
  只补引用不牵连缩进/分页开关/红字、套标题样式（正文→heading + level，红字与加粗原样）、
  清除样式（标题退回正文、`level` 不残留）、跨段一次改完、列表项不被改坏结构。
- 局部格式：`applyNamedStyle` 只改块级属性与节点类型，run 上的 mark 一个都不碰（用例逐条断言）。
- 撤销：窗格是编辑器命令，撤销走 Tiptap 的事务历史（⌘Z），不产生 `setStyle` op。`styleId`
  **刻意没有**进 `STYLE_KEYS`（`ai/docEdit.ts`）：样式引用是给结构用的（目录、交叉引用认它），
  不该让模型按"这段落看起来像标题"去猜着改。
- 类型 + 测试：`pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包
  docx-io 132 / desktop 127 / word-client 8 / word-tools 6 / server 2；`apps/cli`、`packages/shared`
  仍无测试文件。
- UI：`[!] BLOCKED`——本仓没有 Electron 的 CDP harness，样式窗格未实跑走一遍。当时样式表样式**不会**
  在编辑器里渲染成最终外观（样式链解析是 ⑤），这是刻意的：解析会把样式定义盖成直接格式，
  与 Word 的优先级相反。⑤ 之后编辑器里能看到样式效果，见 §8.3。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS；写侧结论来自 OOXML 与往返用例，不是打开验证。
- 提交：`481c4ec` 已推送，`git ls-remote origin main` 复核一致。

### 8.3 排名 2 ⑤（样式链解析 + 文档作用域样式表）验收记录（`5171dcb` `26e5783` `3a9a79b`）

- 真实文件取证（两份本机 .docx）：`丁长长-高级前端开发2026.docx` 的 docDefaults 带
  `w:rPrDefault/w:rPr/w:rFonts`（字面 `w:ascii`="Times New Roman"、`w:eastAsia`="宋体"）与
  **空的 `<w:pPrDefault/>`**；Normal 是 `styleId="1" @w:default="1"`（`widowControl=0`、`jc=both`、`sz=21`），
  标题 1 是 `styleId="2"` `basedOn=1`（`spacing before=340 after=330 line=576 lineRule=auto`、
  `outlineLvl=0`、`b`、`sz=44`）。`self-introduction.docx` 只有主题字体（`w:asciiTheme`），
  解析结果里就没有字体名——不猜。
- 引擎用例：`packages/docx-io/src/style-format.test.ts` 9 条（默认条目在最前、标题 1 的链继承、
  `<w:b w:val="0"/>`→false 与缺失→undefined、只有主题字体时不产出、无 styles.xml→`[]`、
  字符样式不参与、`w:hidden` 仍解析、`w:basedOn` 成环在重复处截断、段落标记的 pPr/rPr 不计入）；
  `docx.test.ts` 新增接线用例（`styleFormats` 三条：`null`/`"1"`/`"2"`，默认样式出现两次是刻意的，
  真实文档会写 `<w:pStyle w:val="Normal"/>`；空白文档 `[]`）。
- 渲染用例：`apps/desktop/src/editor/docStyleSheet.test.ts` 9 条（空输入不注入、兜底选择器、
  pt/倍数→CSS、`font-weight: normal` 取消基线粗体、下划线+删除线合一条、不安全 styleId 整条跳过、
  形状不认识的值逐条丢弃其余生效、未声明 id 由兜底接管且具名规则排在其后）。
- Chromium 实测（**不是** Electron 窗口）：用真实 .docx 生成的 CSS + 应用原样的 `styles.css`
  组成探针页，读 `getComputedStyle`：样式 2 的 `<h1>` 得 29.3px / 22.7px / 22px / 行距 70.32px
  （= 2.4×29.3）/ 加粗，压过基线的 26px；`data-style-id="12"` 的正文首行缩进 28px（= 21pt）；
  无引用、空串引用、未声明 `Heading1` 引用三种段落都落到默认样式（14px、两端对齐、
  Times New Roman+宋体）；行内 `font-size: 20px` 的正文仍盖住样式的 14px（§2.3 的直接格式优先级）。
- app 内实跑：`[!] BLOCKED`——没有 Electron 的 CDP harness，且 `Welcome` 的最近文件走的是
  用户自己的 `userData`，不能拿第二个实例去动它。上面那条浏览器探针是本轮能看到的最直接证据。
- 不变量自查：单位换算仍只在 `docx-io`（渲染层复用既有 `ptToPx`/`lineHeightCss`，未新增映射）；
  颜色仍 #rrggbb（形状不符直接丢）；不拍平局部格式（不写 attrs，行内格式原样保留）；
  不改文档内容，纯显示层；写侧一行未动，导出结果与本排名 ④ 时一致。
- 复用：`@w:val` 读取收进 `xml.ts`（第四处副本，`5171dcb`）；字符属性表抽成
  `readRunProperties` 供正文与样式定义共用；`runStyleDeclarations` 从 `TextStyle` 抽出，
  标记渲染与样式表同一份。遗留：`components/TemplatePickerDialog.tsx:12` 还有一份私有 `ptToPx`
  副本（第三处），留给单独的重构提交。
- 类型 + 测试：`pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包
  docx-io 142 / desktop 136 / word-client 8 / word-tools 6 / server 2；`apps/cli`、`packages/shared`、
  `apps/task-pane` 仍无测试文件。
- 提交：`26e5783` `3a9a79b` 已推送，`git ls-remote origin main` = `3a9a79b`。

### 8.4 排名 4 ③④a④b（修订跟踪读写保真 + 编辑器标记、接受/拒绝与菜单入口）验收记录（`5197192` `bee11e4` `ced2499`）

- 往返用例：`ooxml-pm.test.ts` 的"修订跟踪（`w:ins` / `w:del` / `w:delText`）"8 条——读成标记且
  id/author/date 逐字带走、写回逐字节相同且二次写出一致、一次修订并成一个元素内部两个 `w:r`、
  `del` 套 `ins` 只留外层且不产出 `<w:ins`、超链接在外修订在内、段落标记的 `w:pPr/w:rPr/w:ins` 透传、
  文档没作者时不编名字、被删的制表符留在同一条删除里；再加 `DocxFile` 的
  "keeps tracked insertions and deletions across open → save"。
- live 用例：`docBlocks.live.test.ts` 的"修订跟踪（真实 schema）"6 条——docx-io 的产出进真实
  `EDITOR_EXTENSIONS` 不被丢、写回仍是 `w:ins`/`w:delText`、部分改色不拍平同段的加粗与红字、
  接受/拒绝四种结局各一条、一处修订被按段落拆开时接受一次全清、选区没碰到修订时返回 `false`，
  再加"菜单落点事实与命令同一个判据"一条（`selectionTouchesRevision` 与 `acceptRevision`
  不能一个说有一个说无）。菜单数据层另有 `contextMenu.test.ts` 的一条置灰用例，
  呈现另有 `extensions/Revision.test.ts` 3 条（`data-revision` 与 `title`、缺作者时不编名、
  `parseHTML` 返回空）。
- 局部格式：标记合并只按 `type` 与 attrs 分组，不重建 `w:rPr`，同段其它字符格式原样留在自己的 run 里。
- 撤销：接受 / 拒绝是**一个事务**（所有区间按位置倒序攒在同一个 `tr` 上），所以 ⌘Z 一次回到改动前；
  本节没有新增 op 类型，modlog 与 AI 改动块的逆操作机制不受影响。接到右键菜单后仍是同一条命令，
  用户在菜单里点的是一次普通编辑。
- 类型 + 测试：`pnpm build` EXIT=0（含 `tsc --noEmit`，8 个包）；`pnpm test` EXIT=0，逐包
  docx-io 249 / desktop 264（24 个文件）/ word-client 8 / word-tools 6 / server 2；`apps/cli`、
  `packages/shared`、`apps/task-pane` 仍无测试文件。
  **踩到的坑要记下来**：`apps/desktop` 解析 `@office-copilot/docx-io` 走的是 `dist/`，改了 docx-io
  源码不重新构建的话，桌面测试看到的还是旧转换器（症状是插入的文字被并回普通 run、`delText` 直接消失）。
- UI：真实 app 探针两轮（见上面"③④a 的实测"与"④b 的实测"）——4 个 `span[data-revision]` 的配色与
  `title` 正确，保存后 `w:ins`/`w:del`/`w:delText` 回来且无重复 id；右击落在修订上两条命令可用、
  落在正文上置灰，点接受/拒绝后标记按整条修订清空而文字一字不缺。
  段落标记那条修订经编辑器丢失，已记为缺口。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。
- 提交：`5197192` `bee11e4` `ced2499` 已推送，`git ls-remote origin main` = `ced2499`。

### 8.5 排名 7 ①（节属性在桌面保存链路的保真）验收记录（`bba064e`）

- 往返用例：`docx.test.ts` 新增 2 条——"keeps every child of sectPr when the app re-applies its page
  setup"（复刻 `saveDocx` 的顺序，断言 `w:cols`/`w:pgBorders`/`w:lnNumType`/`w:docGrid` 逐字留下、
  `pgSz`→`pgMar`→`pgBorders`→`lnNumType`→`cols` 的相对顺序不动、二次保存字节一致）、
  "inserts w:pgMar right after w:pgSz when the source section has no margins"。
  公共夹具 `saveTwice()` 扩了 `{ sectPr, applyPageSetup }` 两个入参而不是另起一个平行 helper
  （AGENTS §3.2：同一段"打开→保存→重开"的逻辑已经有第二处在长）。
- live 用例：本片**不涉及**编辑器节点或标记——节属性走 `section` 载荷（`HeaderFooterPayload`），
  不进 PM 文档，所以 §5.3 的注册要求不适用；相应的 `docBlocks.live.test.ts` 没有新条目。
- 局部格式：不涉及（不碰 `w:rPr`）。
- 撤销：不涉及——修的是序列化环节，没新增 op；保存仍走 `.tmp` → `rename`（§2.4）。
- 类型 + 测试：`pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包 docx-io **251**（+2）/
  desktop 264（24 个文件）/ word-client 8 / word-tools 6 / server 2；`apps/cli`、
  `packages/shared`、`apps/task-pane` 仍无测试文件。
- UI：真实 app 探针（`tmp/sect-app-probe.mjs`，独立 `--user-data-dir` + CDP）打开带全字段节属性的
  `.docx`、不编辑直接点保存，读回 `document.xml`：`w:pgSz`/`w:pgMar`/`w:pgBorders`/`w:lnNumType`/
  `w:cols`/`w:docGrid` 六项全在且与打开时一致，正文两段一字不缺。修前同一探针是
  `<w:sectPr` 0 次、`w:pgSz` 0 次。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。顺序合规性按 ECMA-376 §17.6.1 的
  `CT_SectPr` 序列推断，不是实测。
- 遗留（写清楚，不算本片做完）：段内分节符仍靠 `rawPpr` 透传，经编辑器会丢，多节文档在桌面上
  仍被拍成单节；与排名 4 记下的 `rawPpr` 注册缺口同属一片独立改动。
- 提交：`bba064e` 已推送，`git ls-remote origin main` = `bba064e`。

### 8.6 排名 7 ②（分栏）验收记录（`b93c1ef` `a806d69`）

- 往返用例：`docx.test.ts` 新增 7 条——读侧 3 条（`w:cols` 解析成 CSS px 的 count/gap/widths、
  源节没有 `w:cols` 时 `columns` 保持 `undefined`、不等宽分栏读出每栏 `w:col w:w`）、写侧/保真 4 条
  （未改动的 `w:cols` 连没建模的 `w:sep` 一起逐字留下且二次保存字节一致、改了栏数才重写并生成
  `w:space`/`w:num`、不等宽写出 `w:equalWidth="0"` + 每栏 `w:col`、单栏文档不凭空造 `w:cols`）。
  公共夹具仍是 ① 扩出来的 `saveTwice({ sectPr, applyPageSetup, editSection })`（AGENTS §3.2）。
- live 用例：本片**不涉及**——分栏走 `section` 载荷不进 PM 文档，§5.3 的注册要求不适用，
  `docBlocks.live.test.ts` 没有新条目。
- 局部格式：不涉及（不碰 `w:rPr`）。
- 撤销：不涉及——没有新增 op，页面设置是文档级；保存仍走 `.tmp` → `rename`（§2.4）。
- 类型 + 测试：`pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包 docx-io **258**（+7）/
  desktop **269**（26 个文件，+5：`PageSetupDialog.test.ts` 3 条 + `printHtml.test.ts` 2 条）/
  word-client 8 / word-tools 6 / server 2；`apps/cli`、`packages/shared`、`apps/task-pane` 仍无测试文件。
- UI：真实 app 探针（`tmp/cols-app-probe.mjs`，独立 `--user-data-dir` + CDP）实跑读回的数字——面板
  初始 `栏数=1 / 间距=1.27cm / 间距控件禁用`，改 3 栏 + 0.8cm 后"每栏 ≈ 4.35 cm"，导引线
  `left 314.711px / 509.222px`，保存后的 `document.xml` 里 `<w:cols w:space="453" w:num="3"/>`。
  编辑视图故意不做文字重排（会让 `PagedLayout` 的高度测量失效），真实分栏在打印/PDF 链路。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。属性形态与顺序按 ECMA-376 §17.6.1 推断，
  加上观测到的现存文档属性顺序，不是实测。
- 遗留（写清楚，不算本片做完）：① 面板不提供逐栏宽度编辑，不等宽文档只在**栏数不变**时保住
  `widths`，改栏数会拉平成等宽；② 编辑视图只有导引线，没有栏内回流；③ 段内 `w:pPr/w:cols`
  仍受 `rawPpr` 注册缺口影响（与 §8.5 同一条）。
- 提交：`b93c1ef` + `a806d69` 已推送，`git ls-remote origin main` = `a806d69`。

### 8.7 排名 7 ③（页面边框 + 行号）验收记录（`587a9c9` `96afc34`）

- 往返用例：`docx.test.ts` 新增 7 条——读侧 3 条（`reads w:pgBorders into the section geometry in
  points`：八分之一 pt 的 `w:sz` 与 pt 的 `w:space` 落成 `sizePt`/`spacePt`，不做第二次换算；
  `leaves page borders and line numbers unset when the section has neither`：源节没有这两项时
  `pageBorders`/`lineNumbers` 必须是 `undefined` 而不是 `null`，否则保存会凭空删掉用户的东西；
  `reads w:lnNumType as CSS px and ignores the unmodelled w:restart`：`w:distance` twips→px，
  `w:restart` 不进模型）、写侧/保真 4 条（`keeps an untouched w:pgBorders verbatim including
  w:display and w:shadow`：值等价就整段不重建，没建模的属性逐字留下且二次保存字节一致；
  `rewrites w:pgBorders in top-left-bottom-right order when the user changes it`；
  `writes w:lnNumType between pgMar and cols and drops it when line numbers are off`：位置按
  `SECT_PR_ORDER`（`docx.ts:1296-1310`）；`removes w:pgBorders when every side is switched off`：
  四条边全关 = `null` = 删元素）。公共夹具仍是 `saveTwice({ sectPr, applyPageSetup, editSection })`。
- live 用例：不涉及——页面边框与行号都走 `section` 载荷不进 PM 文档，§5.3 的注册要求不适用。
- 局部格式：不涉及（不碰 `w:rPr`；`w:pgBorders` 是节属性）。
- 撤销：不涉及——没有新增 op，页面设置是文档级；保存仍走 `.tmp` → `rename`（§2.4）。
- 类型 + 测试：`pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包 docx-io **265**（6 个文件，+7）/
  desktop **282**（27 个文件，+13：`pageBorders.test.ts` 新增 6 条 + `PageSetupDialog.test.ts` +5 +
  `printHtml.test.ts` +2）/ word-client 8 / word-tools 6 / server 2；`apps/cli`、`packages/shared`、
  `apps/task-pane` 仍无测试文件。
- UI：真实 app 探针（`tmp/border-app-probe.mjs`，独立 `--user-data-dir` + CDP）实跑读回的数字——
  画布初始 `.oc-page-border` 0 条，对话框初始四条边都是"无"、行号未启用且"每几行一个号"禁用；
  把上边改成虚线 0.5pt 间距 1pt + 启用行号每 2 行后，画布出现 1 条
  `left 0px / top 1.33333px / width 793.733px`、计算样式 `dashed 1px rgb(0, 0, 0)`；保存后的
  `document.xml` 里
  `<w:pgBorders w:offsetFrom="page"><w:top w:val="dashed" w:sz="4" w:space="1" w:color="auto"/></w:pgBorders>`
  与 `<w:lnNumType w:countBy="2" w:distance="284"/>`，两条各 1 次、正文两段一字不缺、二次保存字节一致。
  编辑视图**故意不做逐行行号**：`PagedLayout` 按顶层块量高度，给每行插行号会让分页测量失效；
  数据照样往返，真实行号由 Word 与打印/PDF 链路渲染。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。`w:pgBorders`/`w:lnNumType` 的形态、单位
  （`w:sz` 八分之一 pt、`w:space` pt、`w:distance` twips）与子元素顺序按 ECMA-376 §17.6.10 / §17.3.4
  推断，加上观测到的现存文档属性顺序，不是实测。
- 遗留（写清楚，不算本片做完）：① 编辑视图无逐行行号（只有数据与打印链路有）；② 改任一条边会
  整段重建 `w:pgBorders`，没建模的 `w:display`/`w:shadow`/`w:zOrder` 在那一次会掉（与 ② 的 `w:sep`
  同一个取舍）；③ 对话框的颜色只能选 hex，没法把已改过的颜色重置回 `auto`；④ `w:restart` 不建模，
  只在整段未动时保住；⑤ 段内分节符仍受 `rawPpr` 注册缺口影响（与 §8.5、§8.6 同一条）；
  ⑥ 桌面原地保存没有产出 `.bak`（只有 `apps/cli/src/fsutil.ts:38` 有），与 §8 安全红线第 4 条对不上，
  等你决定是补桌面侧还是改文档。
- 提交：`587a9c9` + `96afc34` 已推送，`git ls-remote origin main` = `96afc34`。

## 附录 A：现状盘点（`path:line` 为证）

- 段落格式属性与命令：`apps/desktop/src/editor/extensions/ParagraphFormatting.ts`——段落属性表在
  `addGlobalAttributes`，两条命令都经 `mapFormatBlocks`（`function mapFormatBlocks:56`）攒成一个事务
  落地；`applyNamedStyle` 在 `addCommands` 里。样式窗格：`components/StylePane.tsx`
  （清单来自 `fileHandlers.ts` 的 `OpenResult.styles`）。
- 段落面板：`apps/desktop/src/components/ParagraphPopover.tsx`（`ParagraphFields` 被工具栏与
  右键「段落(P)…」共用）；字体面板 `components/FontFields.tsx`；
  菜单模型 `apps/desktop/src/editor/contextMenu.ts`。
- OOXML 读侧段落属性：`packages/docx-io/src/ooxml-pm.ts:303-`（`paragraphFromXml`：pPr 采集 + rawPpr 透传）、
  `:399`（`readParagraphFormat`：`w:ind` 含 left/right/start/end/firstLine/hanging、`w:spacing`
  行距、`w:outlineLvl` 与 4 个 CT_OnOff 开关；样式定义里的 pPr 走同一个函数）、
  `:511`（`jcToTextAlign`）、`:714`（`readRunProperties`：正文 run 与样式定义共用的字符属性表）。
- OOXML 写侧：`paragraphXml` `:1052-`、`PPR_CHILD_ORDER` `:1085`（CT_PPr 子元素顺序）、
  `modeledPprParts` `:1158`、`paragraphPprXml` `:1196`（导入与新建段落共用）、
  `mergedElementXml` `:1304`（属性级合并）、`horizontalRuleXml` `:1512`、
  `pmToOoxml` `:246` 尾部单 sectPr。
- 样式：`packages/docx-io/src/styles.ts`（`StyleMap`，含 `docDefaultsPPrInner:229`、
  `defaultParagraphStyleId:236`）、`style-format.ts:143`（`resolvedStyles`：按 §17.7.2.5 合并样式链）、
  `docx.ts:132`（`word/styles.xml`）。
- 编辑器侧的文档样式表：`apps/desktop/src/editor/docStyleSheet.ts:169`（`docStyleSheet`）、
  `:31`（`fallbackSelector`）；共用的映射在 `extensions/TextStyle.ts:21`（`runStyleDeclarations`）、
  `extensions/ParagraphFormatting.ts:276`（`ptToPx`）与 `:321`（`lineHeightCss`）；
  注入点 `apps/desktop/src/App.tsx`（`editor-area` 里的 `<style>`，数据来自 `fileHandlers.ts` 的
  `OpenResult.styleFormats`）。
- 域：`apps/desktop/src/editor/extensions/Field.ts`；编号：`packages/docx-io/src/numbering.ts`。
- 模板里的首行缩进按 pt 计算：`apps/desktop/src/editor/templates.ts:26`（`2 字符 = 2 × 字号磅值`）。
