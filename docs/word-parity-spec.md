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
  **（这片已经做掉：④ 的 `d30939a` 把 `rawPpr` 注册成隐式段落属性，见 §7 的 ④ 判据与 §8.8。
  段落标记自己的 `w:ins`/`w:del` 现在也跟着 rawPpr 活过编辑器，但那条场景没有单测覆盖：
  ⑤ 是侧栏、碰不到段落标记，⑥ 的开关明确不记结构步骤（见下面 ⑥ 的判据），所以这条要等真给
  段落标记建模修订时一起补。）**
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

⑤ 的判据（侧栏一片，实现时定下来的）：

- **`w:author` 取设置里的作者名，留空兜底成 `Office Copilot`**：规范里 `w:author` 是**必填**属性
  （见上面 `CT_Comment` 那条），不写就是 Word 判定文档损坏的红线，所以"文档没写作者就不编"那条
  只对**读侧透传**成立；我们**新造**的条目必须有值。署名每次新建时现读 `getSettings()`，
  改完名字不用重启、不用重开文档，下一条批注就跟着变。
- **`w:date` 按规范到秒、带 `Z`，不带小数秒**：`ST_Datetime` 允许两种写法，Word 自己写的是
  `2026-01-02T10:30:00Z`，取观测到的这一种；`commentStamp()` 用 `toISOString().slice(0,19)+"Z"`
  而不是整串（小数秒会让 Word 侧栏的时间戳显示成"未知"，至少没有官方说明支持它）。
- **id 只在批注内部取最大值 +1**（`nextCommentId`）：注解 id 虽然是文档级共享编号空间（见上面那条），
  但"批注与书签的编号是否必须互不重复"**无官方说明**，所以按 Word 实际写出的文件一样各部件各算各的。
  这是**推断**，不是实测。
- **右键菜单里"新建批注"按选区置灰，不做成"光标也能批"**：Word 的"新建批注"没选区时会锚到当前字，
  我们的锚点建模是 `start`/`end` 两个区间端点，空区间的语义在编辑器里没有可见呈现（画出来是
  〔〕挨在一起），与其给一个看不见的锚点，不如把命令按住并要求选区，判据与菜单里其它"不可用看得见"
  的条目一致。
- **删掉聚焦行以后接着聚焦相邻那条**：侧栏的"删除批注"只出现在聚焦行的编辑器下方，
  删完不留下一条就等于把这条命令一起删了（第一次实测就是这么坏的：删掉最后一条以后
  `AFTER_DELETE_ALL={rows:1,chips:3}`，用户再也点不到删除）。Word 删批注后同样把焦点落到邻近条。
- **最后一条批注删光时连 `word/comments.xml` 一起摘掉**：只"不写部件"会把旧字节留包里，
  Word 把已删的批注重新显示出来；空的 `<w:comments>` 又会让 Word 以为文档有批注区。
  所以按 Word 的做法三处一起清（部件文件、`[Content_Types].xml` 的 Override、文档关系），
  包里本来没这个部件时直接返回，避免无批注文档被反复改写 rels 字节。
- **侧栏只做装配，判据在引擎层**（§3.3）：`src/editor/comments.ts` 出 `insertCommentOnSelection` /
  `deleteCommentById` / `commentStamp`，拒绝只回错误码（`no-selection`），中文提示在渲染层组装（§5.7）；
  `CommentPane.tsx` 不含任何 id 分配或写回判定。新建条目要的空正文与第一条预览文字都复用现成实现
  （`emptyPmDoc` 从 docx-io 包口出去、`bodyPreview` 加在 `docBlocks.ts` 里给脚注与批注两个窗格共用），
  不再各写一份（§3.1/§3.2）。

⑥ 的判据（修订开关，实现时定下来的）：

- **拦截点选 `appendTransaction` 而不是输入通道**：Word 的"修订"对所有编辑通道生效——打字、退格、删除、
  粘贴、拖拽、格式化命令。在 `handleTextInput` 上拦要每个通道写一遍，而且**拿不到被删的文字**（输入处理器
  只看到新字）。`appendTransaction` 看到的是已经算好的 steps，删掉的内容从 `tr.docs[i]`（该步骤**之前**那份
  文档，prosemirror-transform 的 `Transform` 逐步存）里取，一份实现覆盖全部通道。
- **必须跳过历史事务**：撤销/重做本身不是"一次新修订"，否则按一次 ⌘Z 就多出一条 `w:ins`。判据取
  prosemirror-history@1.5.0 的 `isHistoryTransaction`——它是**单参数**版本，实质就是 `tr.getMeta(historyKey) != null`，
  这里直接读 meta，不依赖那个导出（社区实现里踩过两参签名的版本，`apply(tr, old, new)` 那套是旧版）。
- **只认 `ReplaceStep`，格式与结构改动明确不记**：加/摘标记（改粗体、改字号）在 Word 里是 `w:rPrChange`、
  回车与变列表改的是段落结构（段落标记的 `w:ins`/`w:del` 在 `w:pPr/w:rPr` 里），这两类本仓都没有建模，
  所以整批不记。**宁可少记，不许记错**——记成 `w:ins`/`w:del` 的语义是"这段文字是新来的/被删的"，
  把格式改动写成文字修订会让用户在 Word 里拒绝一次修订时丢掉正文。
- **坐标要按步骤映射，jobs 按位置降序落地**：每一步的区间经 `new Mapping(maps.slice(index + 1))` 才落到
  最终文档上（后面的步骤会把坐标推走）；同位置先做"还原被删文字"再做"加插入标记"，
  落出来才是 Word 的 `<w:del>旧</w:del><w:ins>新</w:ins>` 顺序。
- **开关命令只 `setMeta`，绝不调 `dispatch`**（实测 `@tiptap/core/dist/index.js` 的 `buildProps` 与
  `CommandManager.commands`）：命令拿到的 `dispatch` 是 `() => undefined`，只用来告诉命令"这次会落地"，
  真正派发的是 CommandManager 在命令返回 true 以后的 `view.dispatch(tr)`。自己再调一次就是把同一份事务
  派两遍。**这条也解释了手搓测试 harness 的写法**：断言前必须自己 `dispatch(tr)`。
- **`w:author` 留空是可以的，不复用批注那个兜底名**：`CT_TrackChange` 的 `w:author`/`w:date` 是**可选**属性
  （`CT_Comment` 才把 `w:author` 收成必填，见上面 `CT_Comment` 那条），所以设置里没填署名就写空，
  不像批注那样兜底成 `Office Copilot`——给一条修订编个不存在的作者，比留空更糟。
- **连打五个字是一条 `w:ins` 而不是五条**：ProseMirror 会把标记相同的相邻文字节点合并，同一批 `revision`
  attrs 落进同一个 run，写回就是一个 `<w:ins w:id="1">`，与 Word 的行为一致。
- **一次撤销覆盖"用户编辑 + 我们补的标记"**：`appendTransaction` 挂出来的事务带 `appendedTransaction` meta，
  prosemirror-history 据此把 `newGroup` 判为 false，两步落进同一个撤销组；撤销事务自己带 `historyKey`，
  又被上面第一条挡掉，所以 ⌘Z 之后既不多出一条修订、也不会留下半条。
- **光标只在"还原被删文字"造成位移时重摆**：`TextSelection.near` 只在有 restore job 且其位置不晚于光标时才发，
  纯打字的场合不动光标——输入法组合期间挪光标会打断候选窗。
- **记不了的三类就地放弃**：原子节点（图片、引用点）上没有地方放标记，删图片记不出 `w:del`；选区跨段的删除
  整批不记（`w:del` 是 run 级的，跨块要拆成段落标记修订，见上面"不记结构改动"）；别的插件在同一个事务里
  append 出来的 `ReplaceStep` 会被当成用户编辑记下来——目前 `EDITOR_EXTENSIONS` 里没有这种插件，
  真出现时按 `tr.getMeta` 白名单挡，不再加一层通用判定。
- **编号复用 `nextAnnotationId`（docx-io）**：修订与批注/脚注同一套"命名空间内最大值 +1"，`w:id` 从 1 起，
  不新开计数器（§3.2；四处并五处的整理在 `d281eb3` 单独成片）。标记形状直接复用 ③ 的 `revision`
  （`kind`/`id`/`author`/`date`），所以 §5.3 不需要新的 docx-io 往返用例——写侧形状与 ③ 完全同一条路径。
- **换文档前必须关开关**（`App.tsx` 的 `stopTracking()`）：开关是插件状态、不随文档走，开着时
  `setContent` 会把整份新文档的每个字都记成"本次插入"。新建 / 打开 / 拖入 / 最近文件四条路径都调它。
- **`Fragment` 不可迭代**：prosemirror-model 的 `Fragment` 没有 `Symbol.iterator`，遍历 slice 内容只能
  `childCount` + `child(i)`（第一版直接 `[...slice.content]` 报 `not iterable`）。

提交切分（实现顺序，逐条落地后回填提交号）：

| # | 内容 | 提交 |
| --- | --- | --- |
| ① | docx-io：`word/comments.xml` 透明部件（按 id 存、没动过原样带走）+ 正文三个锚点元素读写往返 | `5b32ea8` |
| ② | desktop：三个批注节点注册进 `EDITOR_EXTENSIONS`（区间端点画〔〕、引用记号画"批"字芯片），live 用例 | `52b66da` |
| ③ | docx-io：`w:ins`/`w:del`/`w:delText` 读成一个 `revision` 标记、写回按修订区间合并成单元素；段落标记的增删仍靠 `rawPpr` | `5197192` |
| ④a | desktop：标记注册进 `EDITOR_EXTENSIONS`（插入下划线 / 删除删除线 + 作者与时间）+ 接受 / 拒绝命令 | `bee11e4` |
| ④b | desktop：接受 / 拒绝的右键菜单入口 | `ced2499` |
| ⑤ | 批注侧栏 UI：列出、新增（选区右键）、删除、跳转到锚点；写侧的 `w:date`/`w:author` 取设置里的作者名 | `3a0a771` `9dafc61` `1588561` `aa589be` |
| ⑥ | 修订开关：开了就把本次编辑写成 `w:ins`/`w:del`；前置整理 `d281eb3`（四处 max+1 编号器收进 `nextAnnotationId`） | `d281eb3` `27169f2` |

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


⑥ 的实测（本机 macOS，桌面 app 独立 `--user-data-dir` + CDP，`tmp/track-app-probe.mjs`）：打开一份
**没有任何修订**的两段 `.docx`，设置里署名写"修订者甲"。点工具栏「修订」后按钮 `active` 亮起；在第一段句末
用真实输入通道打五个字，编辑器里出现 1 个 `span[data-revision="ins"]`，`title` 写着
"插入：修订者甲（2026-10-10T23:01:15Z）"；拖选第二段里的"几个字"按 Backspace，出现第二个芯片
`kind="del"`，**被删的字仍在树里**（`TEXT_IN_EDITOR` 两段文字一字不少）。存盘后 `document.xml` 里
`<w:ins w:id="1" w:author="修订者甲" w:date="…"><w:r><w:t>补了五个字</w:t></w:r></w:ins>` 与
`<w:del w:id="2" …><w:r><w:delText>几个字</w:delText></w:r></w:del>` 各一条——连打五个字只有**一个** id，
删除走的是 `w:delText`。⌘Z 一次：del 芯片消失、文字回来、ins 那条不受牵连（撤销没被记成新修订）。
再点按钮关掉（`active` 归 false）打六个字，芯片数量不变、二次存盘 `w:id` 只有 `1` 那一条，
新打的字是**裸 `<w:t>`**（`SECOND_PARA_PLAIN` 为真）。右击落在插入的字上，菜单里"接受修订/拒绝修订"都可用，
点接受后芯片清空、文字留着，三次存盘后 `<w:ins>`/`<w:del>` 一条不剩、正文四句完整。
截图 `tmp/track-changes.png`：插入是绿色下划线、删除是红色删除线。用 Word / WPS 打开验证：本机没装，**BLOCKED**。

⑥ 之后仍然没建模的（都是上面判据里点明的取舍，不是漏项）：格式修订（`w:rPrChange`/`w:pPrChange`）、
段落标记修订（回车/合并段落）、图片等原子节点的删除、跨段删除、移动类修订（`w:moveFrom`/`w:moveTo`）。

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
  段内分节符（`w:pPr/w:sectPr`）靠 `rawPpr` 透传；`rawPpr` 到 ④ 才注册进段落 schema，此前经编辑器会丢（见下面"①的实测"）。
- 分栏（**②已落地**）：`w:cols w:num="2" w:space="twips"` + 每栏 `w:col`，承载在 `SectionGeometry.columns`
  上（`{count, gap, equalWidth, widths}`，gap/widths 是 CSS px @96DPI），入口就是「页面设置」面板。
- 页面边框（**③已落地**）：`w:pgBorders`（`CT_PageBorders`，父元素 `sectPr`，§17.6.10）承载在
  `SectionGeometry.pageBorders` 上：`{offsetFrom, top, left, bottom, right}`，每条边
  `{style, sizePt, spacePt, color}`。线宽 `w:sz` 是**八分之一 pt**、间距 `w:space` 是 **pt**（§17.3.4），
  两者就按 pt 存在模型里，不再换算成 cm。
- 水印（**⑤已落地**）：**不是节属性**，规格里原先"`w:headerReference` 里的 `v:shape`"这句也不准确——
  Word 把它写成**默认页眉部件里的一段艺术字**：`w:hdr/w:p/w:r/w:pict/v:shape type="#_x0000_t136"`，
  文字在 `v:textpath@string`，颜色在 `v:shape@fillcolor`，半透明在 `v:fill@opacity`（`ST_Percentage`），
  对角是 `v:shape@rotation="315"`（度），"自动"字号是 `fitshape="t"` + `font-size:1pt`。
  承载在 `SectionGeometry.watermark` = `{text, font, sizePt, color, semiTransparent, layout}`
  （`sizePt: null` 即"自动"，`layout` 是 `diagonal|horizontal`），三态语义同 ②③④。
  退化成 `w:pict` 文本框那条路**没走**：文本框形状 Word 认，但"水印"在 Word 里就是 t136 艺术字，
  换成文本框会让用户在自己的文档里认不出这条对象，也拿不到 `v:fill` 的透明度。
- 首字下沉（**④已落地**）：`w:pPr/w:framePr` 的 `w:dropCap="none|drop|margin"` + `w:lines`
  （ECMA-376 §17.3.1.11 `CT_FramePr`；`w:lines` 缺省 **1**，规格里原先写的 `w:dropRows` 这个属性
  **不存在**）。承载在段落属性 `dropCap` 上：`{position, lines?}`；三态沿用 ②③——
  `undefined` = 这一段没建模过、原文一动不动，`null` = 用户选"无"，有值 = 值等价就不重建。
- 行号（**③已落地**）：`w:lnNumType w:countBy w:distance w:start`（`CT_LineNumber`，父元素 `sectPr`，
  在 §17.6.1 的子元素序列里排在 `pgBorders` 之后、`cols` 之前），
  承载在 `SectionGeometry.lineNumbers` = `{countBy, distance(CSS px), start}`；`w:restart` 本片**不建模**，
  靠"值等价就整段不重建"原样留着。

提交切分：① 节属性读写保真（`bba064e`）；② 分栏（`b93c1ef` `a806d69`）；
③ 页面边框 + 行号（`587a9c9` `96afc34`）；④ 首字下沉（`2cfe911` `d30939a`）；
⑤ 水印（`2044594` 数据层 + `1a44229` 桌面入口）；
⑥ 你实测"水印与分栏不可用"后的可用性修复（`d079287` 画布+PDF 画水印、`177f21b` 导出不再截成 1 页，
判据与实测见 §8.13）；⑦ 横向宽高约定按 Word 的写法改正 + 其余档位实跑（`57bcccd`，见 §8.14）；
⑧ 行号画进编辑画布（每页从起始号重新数、按行归页）+ 分栏与行号的生效范围说明（`12216f0` `6856b9a`，见 §8.15）。

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
> **（④ 的 `d30939a` 把 `rawPpr` 注册成隐式段落属性，这条已经解决，见 ④ 的判据与 §8.8。）**
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

> ④的判据与实测（`2cfe911` `d30939a`）：外部取证先纠正了本规格写错的一处属性名——`CT_FramePr`
> （§17.3.1.11）里**没有** `w:dropRows`，首字下沉是 `w:dropCap`（`none`/`drop`/`margin`）+
> `w:lines`，且 `w:lines` 的 Schema 缺省是 **1 行**。缺省是 1 就决定了写侧的取舍：用户在面板上
> 做过任何一次选择都要显式写行数，否则"3 行"导出会变成 1 行；反过来没动过的段落绝不凭空补
> `w:lines`，仍是"值等价就整段不重建"。另一处不能想当然：`w:framePr` 只要存在就表示"这一段是
> 文本框"，所以选"无"不能退化成留下 `<w:framePr/>`——纯下沉框架整条删掉，带边距/锚点的真文本框
> 只把 `w:dropCap` 写成 `none` 并去掉 `w:lines`。
> 顺带补上 ①②③ 一路挂在遗留里的 `rawPpr` 缺口：docx-io 早就把 `w:pPr` 里没建模的那半截塞进
> `rawPpr`，但 Tiptap 的段落 schema 没注册这个属性，于是**任何经编辑器的保存都会把它洗掉**——
> 真文本框的 `w:hSpace`/锚点、段内 `w:pBdr`、`w:pPr/w:sectPr` 全都掉。④ 把它注册成隐式属性
> （`parseHTML` 恒 `undefined`、`renderHTML` 什么都不写：HTML 里没有原文可复原，也绝不把 OOXML
> 泄进 DOM 让粘贴带出垃圾），②的遗留③ 与 ③的遗留⑤ 一起解决。
> 浏览器侧（Electron 32 / Chromium 128 实测）：`initial-letter` 认 `3` 也认 `3 1`，**负值与
> `-webkit-initial-letter` 都不支持**，所以"突出到页边距外"只能映射成两值语法 `N 1`（跨 N 行、
> 只下沉 1 行）；该属性不继承也不认行内 style，所以整份代码只有一条
> `.oc-drop-cap::first-letter { initial-letter: var(--oc-drop-cap, 3) }`，段落用 CSS 变量把行数传进来，
> `printHtml` 按 `?raw` 引用同一个文件——画布与纸上是同一条声明（§3.2）。行数走变量而不是复制规则，
> 是因为 `initial-letter` 无法写进行内 style。
> 已知差距（实测数字）：Chromium 只在正文行数够时才把下沉字占的行位撑出来——240px 宽、
> `14px/1.5` 的夹具里 3 行正文配 3 行下沉仍是 63px、配 4 行才 84px，而**单行正文配 3 行下沉只有
> 21px**，下沉字于是压到下一段（Word 会按 `w:lines` 留空）。`1lh` 单位要 Chrome 133 才有，
> 本机拿不到；为了让画布撑高而再造一套行盒计算既会跟 `PagedLayout` 的高度测量打架，也不是交付物
> 关心的事，所以画布保持近似，差距同时记在样式文件与 §8.8。

> ⑤的判据与实测（`2044594` 数据层 + `1a44229` 桌面入口）：动手前先把 Word 自己写出来的那段
> XML 当一手来源（`tmp/wm-fixture.mjs` 留了一份），确认三件与本仓判断直接相关的事——
> ① 水印**不在 `sectPr` 里**，它是默认页眉部件 `word/header1.xml` 中一段 `w:p/w:r/w:pict/v:shape`，
> 所以本规格原来那句"`w:headerReference` 里的 `v:shape`"是错的（`w:headerReference` 只是节属性指向
> 页眉的那条引用）；② 艺术字形状类型 `#_x0000_t136` 的 `path` 与 14 条 `v:f` 公式是 WordArt 的
> **固定样板**，逐字照抄即可，只有 `rotation`（度，对角 315）、`fillcolor`、`v:fill@opacity`
> （`ST_Percentage`，`50%` 或定点数 `50000f`）、`v:textpath@string` 与 `style` 里的
> `font-family`/`font-size` 是变量，"自动"字号就是 `fitshape="t"` + `font-size:1pt`；
> ③ VML 命名空间 `urn:schemas-microsoft-com:vml`（外加 `o`、`w10`）必须声明在 `w:hdr` 根上，
> 缺了 Word 直接不认这段形状。据此否掉了"退化成 `w:pict` 文本框"的原计划：文本框 Word 打得开，
> 但用户在 Word 的对象树里看到的就不是"水印"，也拿不到 `v:fill` 的透明度（§6.1 要写否决理由）。
> 模型：`SectionGeometry.watermark` = `{text, font, sizePt, color, semiTransparent, layout}`，
> `sizePt: null` 表示"自动"；三态沿用 ②③④（`undefined` 一动不动 / `null` 删掉 / 有值且不等价才重建）。
> 判"是不是水印"只认 `type="#_x0000_t136"` **且**带 `v:textpath@string` 的形状，且带
> `w:txbxContent` 的段落一律不碰——用户页眉里自己放的装饰性艺术字不许被"清理"掉（§2.7）。
> 因为水印落在页眉部件，保存链路必须能**只建页眉不建页脚**：原来的
> `ensureDefaultHeaderFooter()` 一次建两条，改成按 kind 取用的 `ensureHeaderFooterPart(bundle, kind)`，
> 顺带删掉"编辑页眉就凭空写一个空页脚部件"这条副作用（§3.5）。
> 触碰同一段代码时修掉两个真会损坏文档的缺陷：① 页眉与页脚共用同一个 `rIdHf1`，两条一起建时
> 后者覆盖前者的关系；② `word/_rels/document.xml.rels` 在没有任何关系时序列化成自闭合
> `<Relationships/>`，而旧代码用 `rels.replace("</Relationships>", …)` 追加——匹配不上就**静默丢掉**
> 页眉关系，Word 会报"文件损坏"。改成解析后 `addRelationship` 再序列化，并从
> `ensurePartReference` 里抽出这一份复用（同一段逻辑第二次出现，§3.2）。
> 另一处不能想当然：`v:textpath@style` 里是 `font-family:'黑体'`，用通用的 `escapeAttr` 会把单引号
> 写成 `&apos;`，Word 读不出字体名，所以属性值转义分成"style 用只转 `&<>` 与双引号"和
> "`@string` 用 `escapeAttr`"两条（§2.2 的取舍：读不懂不如不写）。
> 桌面入口就是「页面设置」的水印区：启用 / 文字 / 字体 / 字号（pt，0 为自动）/ 颜色 / 半透明 /
> 版式（对角·横向，复用 `.ps-seg`，未启用时整排禁用），`watermarkFor()` 与
> `pageBordersFor()`/`lineNumbersFor()` 同一条三态规则；`geometryFor()` 的 `extras` 多接一项，
> 套模板路径把文档原有水印原样带走（§2.7）。
> 实测（`tmp/wm-app-probe.mjs`，独立 `--user-data-dir` + CDP 起真实 app，一份**没有页眉部件**、
> `document.xml.rels` 是空表根的 `.docx`）：打开时对话框"启用水印"未勾、文字/半透明/版式全部禁用；
> 勾上→填"内部资料"、黑体、字号 0、颜色 `#dddddd`、对角→确定→保存，读回
> `word/header1.xml` 里 `v:shape` 与 `v:shapetype` 各 1 条、`fillcolor="#dddddd"`、`rotation="315"`、
> `<v:fill on="t" opacity="50%"/>`、`font-family:'黑体';font-size:1pt` + `fitshape="t"`、
> `string="内部资料"`，`w:hdr` 根上带 VML 三个命名空间；`document.xml` 里
> `<w:headerReference r:id="rIdHf1" w:type="default"/>`、rels 里
> `Target="header1.xml"` 那条关系在且 `[Content_Types].xml` 有对应 Override，**没有**凭空建出
> `word/footer1.xml`，正文两段一字不缺。用读文件的同一套代码回灌得到
> `{text:"内部资料", font:"黑体", sizePt:null, color:"#dddddd", semiTransparent:true, layout:"diagonal"}`，
> 再点一次保存 `header1.xml` **字节一致**、`headerReference` 与 `rIdHf1` 各 1 条；重开对话框显示的就是
> 刚写进去的值，取消勾选再保存后 `v:shape` 归零而页眉部件仍在、正文完好，回灌读成"没建模过"。
> 截图 `tmp/wm-dialog.png`（水印区排布与其余各区一致）。
> 已知不做：~~编辑画布**不预览**水印~~——**本条作废**（§8.13 的 `d079287` 把它做出来了：当时把它当成
> "又一套坐标计算"的取舍，代价是你勾完看不到任何东西、直接测成"不可用"，所以取舍让位于可用性）。
> 现在画布与打印共用 `editor/pageWatermark.ts` 一份几何；读侧仍只认 VML t136 形状，
> 若别的编辑器把水印写成 DrawingML，我们读不出来也就不会去动它。
> 用 Word / WPS 实际打开验证：`[!] BLOCKED`，形状与属性取自 Word 产出的样例加 ECMA-376 §14.1.2 /
> §9.5.1 推断，本机没能亲眼看到渲染结果（见 §8.11）。

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
  **（前半个已由 ④ 的 `d30939a` 补上：`rawPpr` 现在是段落的隐式属性，能活过"打开→保存"，
  live 用例里段内 `<w:pBdr>` 与整条文本框 `w:framePr` 都逐字留下了。但"整篇两节文档经桌面保存
  仍保持两节"这条还没有实测用例——需要一份真带段内 `w:sectPr` 的样例才能判定，留给下一片。）**
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
  曾受 `rawPpr` 注册缺口影响（与 §8.5 同一条），已由 ④ 的 `d30939a` 注册 `rawPpr` 解决。
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
  数据照样往返。~~真实行号由 Word 与打印/PDF 链路渲染~~——**此句作废**（§8.13 核对：`printHtml.ts`
  里没有行号代码，PDF 也不画，只有 Word/WPS 打开 `.docx` 时才看得到）。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。`w:pgBorders`/`w:lnNumType` 的形态、单位
  （`w:sz` 八分之一 pt、`w:space` pt、`w:distance` twips）与子元素顺序按 ECMA-376 §17.6.10 / §17.3.4
  推断，加上观测到的现存文档属性顺序，不是实测。
- 遗留（写清楚，不算本片做完）：① 编辑视图与 PDF 都无逐行行号（只有 `.docx` 数据里有，见上一条更正）；② 改任一条边会
  整段重建 `w:pgBorders`，没建模的 `w:display`/`w:shadow`/`w:zOrder` 在那一次会掉（与 ② 的 `w:sep`
  同一个取舍）；③ 对话框的颜色只能选 hex，没法把已改过的颜色重置回 `auto`；④ `w:restart` 不建模，
  只在整段未动时保住；⑤ 段内分节符曾受 `rawPpr` 注册缺口影响（与 §8.5、§8.6 同一条），
  已由 ④ 的 `d30939a` 注册 `rawPpr` 解决；
  ⑥ 桌面原地保存没有产出 `.bak`（只有 `apps/cli/src/fsutil.ts:38` 有），与 §8 安全红线第 4 条对不上，
  等你决定是补桌面侧还是改文档。
- 提交：`587a9c9` + `96afc34` 已推送，`git ls-remote origin main` = `96afc34`。

### 8.8 排名 7 ④（首字下沉）验收记录（`2cfe911` `d30939a`）

- 往返用例：`docx-io` 新增 6 条——`读出位置与行数，没改动的段落字节一致`、
  `不是首字下沉的框架原样带走：真文本框、none 与表外值都算没建模过`、
  `源文档没写 w:lines 时不猜行数，写回也就不会补上它`、
  `新建的下沉补出锚定：缺省的 hAnchor/vAnchor 是页面，会把首字钉到纸边上去`、
  `改行数只动 dropCap 与 lines，源文档那条框架的其余属性留着`、
  `关掉下沉：纯下沉框架整条删掉，真文本框只把 dropCap 写成 none`。
- live 用例（§5.3：`dropCap` 是段落属性，这次必须有真实 schema 条目）：`docBlocks.live.test.ts` 新增
  5 条——`有值 / 显式关掉 / 没建模过三种状态都活过 doc JSON 往返`（`null` 在 JSON 里仍是
  `"dropCap":null`，不与"没建模过"混）；`源文档的下沉进真实 schema，一个没动过的段落写回还是那一句`；
  `面板改成突出页边距外、四行：写回换掉 dropCap 与 lines，别的属性不牵动`（断言
  `<w:framePr w:dropCap="margin" w:lines="4" w:hSpace="180"/>` 与段落的 `<w:ind w:left="300"/>` 并存）；
  `选「无」删掉整条纯下沉框架：留下空的 w:framePr 会把正文变成匿名文本框`；
  `没建模的段落属性跟着 rawPpr 走过编辑器：真文本框的框架一个字都不动`（整条 ISO 文本框 `w:framePr`
  与段内 `<w:pBdr>` 逐字留下）。段落默认值那条用例补了一句 `dropCap` 必须是 `undefined`。
- 属性 DOM 读写（粘贴 / `getHTML→setContent` 路径）：`ParagraphFormatting.test.ts` +7 条——三种状态的
  `data-drop-cap` 进出、`--oc-drop-cap` 变量、`margin`→`N 1`、缺省行数、读不懂的形状退回 `undefined`
  （行数为 0/小数/字符串时只保留位置，让面板显示"下沉"而不是"无"）、`null` 不挂类。
- 局部格式：不涉及——`dropCap` 是段落属性，不碰 `w:rPr`，没有"部分文字带不同格式被拍平"的风险（§2.3）。
- 撤销：不涉及——没有新增 op，仍走 `formatParagraph` 既有路径；保存仍 `.tmp` → `rename`（§2.4）。
- 类型 + 测试：`pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包 docx-io **271**（6 个文件，+6）/
  desktop **298**（27 个文件，+16：live 5 + `ParagraphFormatting.test.ts` 7 + `ParagraphPopover.test.ts` 3 +
  `printHtml.test.ts` 1）/ word-client 8 / word-tools 6 / server 2；`apps/cli`、`packages/shared`、
  `apps/task-pane` 仍无测试文件。`vitest.config.ts` 开了 `css: true`：默认配置会把 `?raw` 的样式表
  stub 成空串，打印那条用例会因此假绿。
- UI：真实 app 探针（`tmp/dropcap-app-probe.mjs`，独立 `--user-data-dir` + CDP）——面板初始
  `位置=无 / 行数框禁用 / 段落无类名`；选"下沉"自动补 3 行，段落挂上 `class="oc-drop-cap"`、
  `style="--oc-drop-cap: 3"`，计算样式 `initial-letter: 3`；行数改 4 → `initial-letter: 4`；
  位置改"突出到页边距外" → `initial-letter: 4 1`；保存读回
  `<w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="text" w:dropCap="drop" w:lines="4"/>`，
  切 margin 后是 `w:dropCap="margin" w:lines="4"`；再选"无"→ 整份文件 `w:framePr` 命中 **0** 次。
  样例第一段故意带一条编辑器没建模的 `<w:pBdr><w:bottom w:val="single" w:sz="6" …/></w:pBdr>`，
  四次保存后它仍在（`PBDR_KEPT_AT_END=true`）、正文一字不缺（`TEXT_KEPT=true`）、连点两次保存
  字节一致（`STABLE_ACROSS_SAVES=true`）。
- 已知差距（画布保真）：单行段落的大行数下沉字会压到下一段，Word 按 `w:lines` 留空而 Chromium 128
  不撑行盒（数字见上面 ④ 的判据）。纸上由 Word 自己排，交付物不受影响。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。`w:dropCap`/`w:lines` 的合法值、`w:lines` 缺省 1
  与新建框架的锚定属性顺序按 ECMA-376 §17.3.1.11（微软 XML-Schema 参考页）推断，不是本机实测。
- 遗留（写清楚，不算本片做完）：① 画布单行段落的下沉字溢出；② 面板不做 Word 对话框里的"字体"与
  "距正文"两项——`w:framePr` 不管字体（首字字体来自首字符的字符格式，动它就是 §2.3 的红线），
  "距正文"是 `w:hSpace`，源文档里的值现在靠 `rawPpr` 逐字保住，但编辑器还不写它；
  ③ 桌面原地保存仍没有 `.bak`（与 AGENTS §8 安全红线第 4 条不符，仍等你决定）。
- 提交：`2cfe911` + `d30939a` 已推送，`git ls-remote origin main` = `d30939a2606545f5e9d298e2e11d0b67c4857cbe`。

### 8.9 排名 4 ⑤（批注侧栏 UI）验收记录（`3a0a771` `9dafc61` `1588561` `aa589be`）

- 往返用例：`packages/docx-io/src/docx.test.ts` 新增 1 条——`最后一条批注也删掉时，部件连同
  Override 与关系一起摘掉`（断言 `word/comments.xml` 为 `null`、`[Content_Types].xml` 里既没有
  `/word/comments.xml` 也没有 `comments+xml`、`word/_rels/document.xml.rels` 不再含
  `relationships/comments`）。⑤ 之前的部件读回、逐字节透传、二次保存稳定、删一条不牵动另一条等
  用例是 ① 那一片铺的，这次全部保持绿。
- live 用例（§5.3：三个批注锚点节点在 ② 已注册进 `EDITOR_EXTENSIONS`，⑤ 补的是命令层）：
  `docBlocks.live.test.ts` 新增 8 条——真实 schema 下 `insertCommentRange` 按 `id` 把
  `commentRangeStart/End/Reference` 包住选区、跨段选区、选中整段、光标压到块边界时钳到上一段末尾、
  `deleteComment` 一次清掉三个锚点、缺节点时的兜底。
- 引擎层单测：`apps/desktop/src/editor/comments.test.ts`（新）5 条——id 从 `["0","2"]` 分出 `"3"`、
  署名 `undefined`/`""`/`"   "` 三种都兜底成 `Office Copilot`、没选区时回
  `{ok:false,reason:"no-selection"}` 且**不动锚点**、删除按 id 过滤（锚点已先没了也照样过滤）、
  `commentStamp` 到秒带 `Z`。`contextMenu.test.ts` +1 条："新建批注"在有选区时可用、
  `hasSelection:false` 时置灰。
- 局部格式：不涉及——批注锚点是 inline 原子节点，不碰 `w:rPr`，没有"部分文字带不同格式被拍平"的风险
  （§2.3）。
- 撤销：不涉及——新增/删除批注是编辑器事务，走 PM 自己的 undo 栈；没有新增 CoreOp，
  保存仍走 `.tmp` → `rename`（§2.4）。**注意**：`userData/modlog` 不会记批注的新删（它只记 AI/格式 op）。
- 单位与颜色：不涉及（不产生长度、颜色属性，§2.1/§2.2 无新换算点）。
- 复用（§3.1/§3.2）：`emptyPmDoc` 从 `docx-io` 包口出去给两个窗格共用；`bodyPreview` 加在
  `docBlocks.ts` 里，`FootnotePane.tsx` 删掉自己那份 `previewOf` 与 `EMPTY_BODY`；
  增删判据落 `src/editor/comments.ts`，组件只做装配。无新增平行实现、无新增依赖。
- 类型 + 测试：`pnpm build` EXIT=0（`grep -c "error TS"` = 0）；`pnpm test` EXIT=0，逐包 docx-io **272**
  （6 个文件，+1）/ desktop **312**（28 个文件，+14：live 8 + `comments.test.ts` 5 + `contextMenu.test.ts` 1）/
  word-client 8 / word-tools 6 / server 2；`apps/cli`、`packages/shared`、`apps/task-pane` 仍无测试文件。
- UI：真实 app 探针（`tmp/comment-app-probe.mjs`，独立 `--user-data-dir` + CDP，一份**没有**批注部件的
  `.docx`）实跑读回的数——设置里写 `authorName:"测试署名"` 后回读一致；双击第一段选中"有"→
  右键"新建批注" `disabled:false` → 侧栏 1 行、序号 `1`、meta `测试署名 · 2026-10-10 22:18`、
  锚点芯片 `["start:0","end:0","reference:0"]`；嵌套编辑器里打"这句要改口径。"，行内预览同步显示；
  保存后包里凭空长出 `word/comments.xml`，条目是
  `<w:comment w:id="0" w:author="测试署名" w:date="2026-10-10T22:18:34Z">`，正文里
  `<w:commentRangeStart w:id="0"/><w:r><w:t xml:space="preserve">有</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/>`，
  Content_Types 与关系各登记 1 次。第二条用 Range 量出的坐标拖选"第一段里"→ 新建后侧栏
  `numbers:["1","2"]`、`ANCHOR_IDS` 是 `w:id="1",w:id="0"`（按正文区间顺序排）。点行"跳到锚点"
  选区回到"第一段里"。删一条 → 部件里只剩另一条；把最后一条也删掉 →
  `AFTER_DELETE_ALL={rows:0,chips:0}`、`PARTS4` 里没有 `word/comments.xml`、
  `CT_HAS_COMMENTS4=false`、`RELS_HAS_COMMENTS4=false`，正文
  `第一段里有一句可以拿来批注的文字。第二段。另一句可以批注的话。` 一字不缺；再存一次
  `STABLE_RELS=true`、`STILL_NO_PART=true`（不会反复改写 rels，也不会又长出部件）。
  两处**日志假阴性**要记下来免得下次误判：`TEXT_KEPT=false` 与 `DOC_AFTER_DELETE=false` 都不丢字——
  锚点芯片在 DOM 里就是字面的〔〕批，`textContent.includes(整句)` 必然匹配不上，
  去掉标签与空白后的 `TEXT_AFTER_ALL` 才是判据。
- 探针改出来的两个缺陷（都已修，写在这里防止回退）：① 第二次"新建批注"没反应，实测是双击没产生
  选区、菜单按判据正确置灰，探针改成 Range 拖选后正常；② 删光批注以后 `word/comments.xml` 仍带条目
  （`syncCommentIds` 已经把 model 清空，但 `serializeCommentsPart` 返回空串时旧字节原样留在包里），
  即 `1588561` 那片修复。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。`w:comments` 的部件注册、`w:author` 必填、
  `w:date` 形态按 ECMA-376 与 Open-XML-SDK 部件清单推断 + 与 Word 写出样例对比，不是本机实测。
- 遗留（写清楚，不算本片做完）：① 编辑器只画区间端点与"批"字芯片，没有 Word 那种区间高亮与编号；
  ② `w15:commentEx`（已解决状态）与 `w15:people` 不生成（读到的靠透传带走），侧栏也就没有"解决/恢复"；
  ③ 部件被标成 `isUnmodeled` 时（认不出的批注形状）`writeCommentPart` 直接短路，侧栏的改动不会落盘——
  这条在 ① 的判据里是刻意的，但没有 UI 提示；④ 批注的新删不进 `modlog`（见上面撤销那条）；
  ⑤ 桌面原地保存仍没有 `.bak`（与 AGENTS §8 安全红线第 4 条不符，仍等你决定）。
- 提交：`3a0a771` + `9dafc61` + `1588561` + `aa589be` 已推送，
  `git ls-remote origin main` = `aa589be6081797196d42189391e86bb4d3123213`。

### 8.10 排名 4 ⑥（修订开关）验收记录（`d281eb3` `27169f2`）

- 往返用例（§5.3）：**不新增** docx-io 用例——开关产出的就是 ③ 那条 `revision` 标记，
  写回走 `<w:ins>`/`<w:del>`/`w:delText` 的既有路径（③ 的往返用例覆盖）。⑥ 的落盘形状改在
  桌面侧断言：`TrackChanges.test.ts` 末条用真实 `EDITOR_EXTENSIONS` schema 建文档、
  `pmToOoxml(docJSON)` 里断言 `<w:ins`/`<w:del`/`<w:delText` 与 `w:author="张三"` 都在。
- 引擎层单测：`apps/desktop/src/editor/extensions/TrackChanges.test.ts`（新）14 条——开关关着时
  什么都不记；开着打字记 `ins` 并带上署名与时间；连打三个字符只有一条 run、一个 `w:id="1"`；
  删除走"还原原节点 + 挂 `del`"并把光标停在还原段之后；选区替换落成 `del(id=2)+ins(id=1)` 的
  Word 顺序；刚插入的字再删掉自我抵消（不留 `ins` 也不留 `del`）；别人的 `ins` 被我们删掉时
  外层套 `del`；撤销事务不记成新修订；改粗体不记；回车合并段落不记；关掉以后再打字不记；
  文档里已有 `w:id="5"` 时新 id 取 6；署名为空白时 `w:author` 留 `null`；以及上面那条 `pmToOoxml`。
  全部用 `getSchema(EDITOR_EXTENSIONS)` + `EditorState.create({plugins:[history(), trackChangesPlugin()]})`
  的真 schema harness，不是临时扩展数组（§5.3）。
- 局部格式：不涉及——开关只加/摘 `revision` 标记，不碰 `w:rPr`，也没有整段覆盖字符样式的路径（§2.3）。
  格式改动**故意不记**（`w:rPrChange` 未建模），所以不存在"把别人的红色字拍平"的可能。
- 撤销：一次 ⌘Z 撤销的是"编辑 + 标记"这一整组（`appendedTransaction` meta 让 history 不开新组），
  撤销事务被 `historyKey` 挡掉不会二次记账；没有新增 CoreOp，保存仍走 `.tmp` → `rename`（§2.4）。
  `userData/modlog` 与这条无关（它记的是 AI/格式 op，§上面 8.9 那条同样适用）。
- 单位与颜色：不涉及（不产生长度、颜色属性，§2.1/§2.2 无新换算点；`w:id` 保持字符串不做数字转换）。
- 复用（§3.1/§3.2）：`revision` 标记直接复用 ③/④a 那一份（含呈现与接受/拒绝命令，开关与它们天然闭环）；
  编号复用 docx-io 的 `nextAnnotationId`——为此先把四处并列的 max+1 计数器收成一处（`d281eb3`，
  按 §1.4 单独成片，不夹进功能提交）。无新增 schema 元素、无新增依赖、无第二套编号器。
- 类型 + 测试：`pnpm build` EXIT=0（`grep -c "error TS"` = 0）；`pnpm test` EXIT=0，逐包 docx-io **274**
  （7 个文件）/ desktop **326**（29 个文件，+14 全在 `TrackChanges.test.ts`）/ word-client 8 /
  word-tools 6 / server 2；`apps/cli`、`packages/shared`、`apps/task-pane` 仍无测试文件。
- UI：真实 app 探针（`tmp/track-app-probe.mjs`，独立 `--user-data-dir` + CDP，一份**没有**修订的
  `.docx`），实跑读回的数记在 §四 的"⑥ 的实测"那段：按钮 `active` 双向切换、`ins`/`del` 芯片的
  `title` 带署名与时间、存盘后 `<w:ins w:id="1" …>` 与 `<w:del w:id="2" …><w:delText>` 各一条、
  ⌘Z 以后只剩 `ins` 且文字回来、关掉开关再打字不记（二次存盘 `IDS2` 仍只有 `w:id="1"`，
  新字是裸 `<w:t>`）、右击落在插入文字上"接受修订/拒绝修订"都可用、接受以后三次存盘
  `<w:ins>`/`<w:del>` 归零而四句正文完整。截图 `tmp/track-changes.png`。
  探针第一次跑出来的一处**探针自身缺陷**记下来防误判：`selectInParagraph` 按节点序号取文字节点，
  芯片把段落拆成多个文字节点之后偏移越界（`IndexSizeError: offset 12 > 10`），改成按内容找节点。
- Word 实测：`[!] BLOCKED`——本机未安装 Word / WPS。`w:ins`/`w:del` 的包 run 形状、`CT_TrackChange`
  的 `w:author`/`w:date` 可选性按 ECMA-376 类页与 ③ 已核对过的写出样例推断，不是本机实测。
- 遗留（写清楚，不算本片做完）：① 格式修订（`w:rPrChange`/`w:pPrChange`）与段落标记修订（回车、
  合并段落）不记，见 §四 ⑥ 的判据；② 原子节点（图片、引用点）上的删除记不了；③ 跨段删除整批不记；
  ④ 移动类修订（`w:moveFrom`/`w:moveTo`）读侧仍没有分支（③ 那一片就写明了）；
  ⑤ 开关状态不进文档也不跨会话（Word 的"修订"是文档级状态，我们的是会话级，重开文档要重新按）；
  ⑥ 别的插件若在同一事务里 append `ReplaceStep` 会被当成用户编辑记账（当前无此插件）；
  ⑦ 桌面原地保存仍没有 `.bak`（与 AGENTS §8 安全红线第 4 条不符，仍等你决定，与 8.9 同一条）。
- 提交：`d281eb3`（前置整理）+ `27169f2`（本片）+ 本条 docs 记录一起推送；推送后 `git ls-remote origin main`
  与本仓 `git rev-parse HEAD` 一致才声称已推送。

### 8.11 排名 7 ⑤（水印：数据层 + 桌面入口）验收记录（`2044594` `1a44229`）

- 往返用例（§5.3）：`packages/docx-io/src/docx.test.ts` +6 条，全部走 `saveTwice` 真文件路径——
  读进 `section.watermark`；页眉没建模时字节原样穿过；没改过的水印二次保存**字节一致**且
  `v:shape`/`v:shapetype` 各 1 条；改文字/字体/颜色/版式后是**替换**而不是叠加（断言
  `string="作废&quot;&lt;草稿&gt;"` 的转义、`font-family:'黑体';font-size:48pt`、无 `rotation`、无 `opacity`）；
  `watermark: null` 删形状但留下页眉文字；正文文档原本没有页眉时**只**建 `header1.xml`
  （Content_Types Override + `Target="header1.xml"` 关系 + `w:headerReference r:id="rIdHf1"`，
  不建 `footer1.xml`），二次保存仍只有一条 `rIdHf1`。不新增编辑器节点/标记，所以
  `EDITOR_EXTENSIONS` 与 `docBlocks.live.test.ts` 这一片不涉及（§5.3 的注册项没变）。
- 引擎层单测：`packages/docx-io/src/watermark.test.ts`（新）13 条——读 Word 的对角半透明
  （含 `50000f` 定点数写法）、读固定字号与横向、实体解码；读不懂的一律当没建模过（无形状 /
  没有 `string` / `fillcolor` 不是 hex 都返回 `undefined`）；写侧建形状与命名空间、写完能读回、
  替换不叠加、剥水印时留下页眉正文、**绝不删** `w:txbxContent` 里的艺术字、引号注入被转义、
  `sameWatermark` 逐字段比较。
- 局部格式（§2.3）：不涉及——水印是文档级对象，不碰任何 run 属性，也没有整段覆盖字符样式的路径。
- 原子写与撤销（§2.4/§2.5）：不涉及新机制——仍走 `fileHandlers` 的 `.tmp` → `rename`；水印与页边距
  同属"文档级设置跟随 App state、不进修改历史"那一类（与 ②③④ 一致，套版路径的注释就写明了这条边界）。
- 单位与颜色（§2.1/§2.2）：VML 的 `rotation` 是度、`font-size` 是 pt、形状尺寸是 pt，全部按原生单位
  存，**没有**新增第二处换算；颜色只认 `#rrggbb`，`fillcolor` 读成非 hex 就当没建模（不往 `w:color`
  那种位置塞读不懂的值）。属性值转义分两条：`@style` 只转 `&<>` 与双引号（`font-family:'黑体'` 的单引号
  必须留着），`@string` 走通用 `escapeAttr`。
- 复用（§3.1/§3.2）：`ensureHeaderFooterPart` 取代 `ensureDefaultHeaderFooter`（旧实现删掉，§3.5）；
  `addRelationship` 从 `ensurePartReference` 里抽出共用，同时修掉空 rels 表被 `replace("</Relationships>")`
  静默丢关系的缺陷；`watermarkFor` 与 `pageBordersFor`/`lineNumbersFor` 同一条三态规则、
  `geometryFor` 的 `extras` 同一入口，没有第二套"文档级设置"通道；对话框控件复用 `.ps-grid`/`.ps-seg`，
  未新增样式文件。
- 类型 + 测试：`pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包 docx-io **293**（8 个文件，+19 全在水印）/
  desktop **330**（29 个文件，+4 在 `PageSetupDialog.test.ts`）/ word-client 8 / word-tools 6 / server 2；
  `apps/cli`、`packages/shared`、`apps/task-pane` 仍无测试文件。
- UI：真实 app 探针 `tmp/wm-app-probe.mjs`（独立 `--user-data-dir` + CDP），跑了两遍结果一致。
  看到的数记在 §七"⑤的判据与实测"那段：初始禁用态、勾上后解锁、保存后 `header1.xml` 的
  `fillcolor="#dddddd"` / `rotation="315"` / `opacity="50%"` / `string="内部资料"` /
  `fitshape="t"`、`headerReference` 与 rels 那条 `rIdHf1`、二次保存字节一致、取消勾选后形状归零而
  页眉与正文完好。截图 `tmp/wm-dialog.png`：水印区与分栏/边框/行号排布一致，颜色是色块、
  版式是"对角/横向"分段按钮。画布上不显示水印（本片有意不做，见上；后续被 §8.13 推翻并补上）。
- Word 实测：`[!] BLOCKED`——形状与属性取自 Word 自己产出的样例 XML 与 ECMA-376 §14.1.2 / §9.5.1
  **推断**，本机没有 Microsoft Word，装有 WPS 但它的窗口抓不到（原因见 §8.12）。
- 遗留（写清楚，不算本片做完）：① ~~编辑画布与打印/PDF 都不预览水印~~（已由 §8.13 的 `d079287` 补上，
  画布与 PDF 逐页都画）；② 只认 VML `#_x0000_t136`，别的编辑器写成 DrawingML 的水印读不出来（读不出来
  也就不动它）；③ 只管 `w:type="default"` 的页眉，`even`/`first` 页眉里的水印不建模；
  ④ 桌面原地保存仍没有 `.bak`（与 AGENTS §8 安全红线第 4 条不符，仍等你决定，与 8.9/8.10 同一条）。
- 提交：`2044594`（数据层）+ `1a44229`（桌面入口）+ 本条 docs 记录一起推送；推送后
  `git ls-remote origin main` 与本仓 `git rev-parse HEAD` 一致才声称已推送。

### 8.12 环境更正：WPS 装了，但渲染结果看不到

① ~ ⑤ 的验收记录里有 12 处写着"本机未安装 Word / WPS"，这句**不准确**，按本条统一更正：
本机装有 WPS（`/Applications/wpsoffice.app`），没有的是 Microsoft Word 与 LibreOffice。
**`[!] BLOCKED` 的结论不变**，但原因不是"没装"，而是"装了也看不到渲染结果"——⑤ 期间实测过三条路：
`osascript` 驱动 WPS 报 `-1728 不允许辅助访问`（运行探针的终端没有辅助功能权限）；
`python3 -c "import Quartz"` 无模块，拿不到窗口列表；整屏截图只有壁纸与 WPS 菜单栏，
文档窗口不在可捕获范围内。
所以此前所有"形状/单位/元素顺序按 ECMA-376 与 Word 产出的样例推断"的措辞仍然成立，
只是别再把它写成"没装"。要真验渲染，二选一：给终端开"辅助功能 + 屏幕录制"权限后由我重跑探针，
或你人工用 WPS 打开 `tmp/` 里那几份 `.docx` 回一句结论。

### 8.13 排名 7 可用性修复（水印看得见 + 导出不再只有一页）验收记录（`d079287` `177f21b`）

起因是你的原话："需要确保页面设置中的功能都是正常可用的，我现在测试水印和分栏不可用"。
复盘出来的事实是：**XML 一直是对的，屏幕上什么都不画**——② ⑤ 两片都把"画布不预览"当成了
取舍写进记录（见 8.6/8.11），于是勾选项在 app 里没有任何反馈，这就是你测到的"不可用"。

- 判据（什么叫"可用"）：① 在页面设置里勾上之后，**编辑画布当场能看到**；② 保存的 `.docx`
  里属性在（这两条此前已由 ②③⑤ 的往返单测覆盖）；③ 导出的 PDF **逐页都有**且**页数正确**。
  三条按此顺序验，第③条把下面那个更严重的老 bug 挖了出来。
- 水印（`d079287`）：新增 `apps/desktop/src/editor/pageWatermark.ts` 作为**唯一**的水印呈现几何
  来源（468pt × 117pt 形状盒、对角 -45°、半透明 0.5、`sizePt` 缺省时按文字宽度自适应字号），
  画布 `PageChrome` 与打印 `printHtml` 共用同一个函数，pt→px 仍只在既有的一处换算里（§2.1、§3.3）；
  字体名过白名单（含引号/尖括号/分号等一律退回 `serif`），颜色非 `#rrggbb` 退回黑（§2.2），
  文字走已有的 `escapeHtml`（导出复用，不再写第二个转义函数，§3.1）。画布上每页一份
  `.oc-page-watermark`（`pointer-events: none` + `z-index: 1`：压纸面、在正文下、不吃掉点击）。
- 导出链路的存量缺陷（`177f21b`，与分栏无关但被分栏暴露）：`.page { height: X; overflow: hidden }`
  让 Blink 把这一页当成**不可拆分的裁剪盒**，于是**任何**导出的 PDF 都被截成 1 页——单栏 60 段
  同样只有 1 页（实测：OLD / OLDSINGLE 两个变体都是 1 页）。改成 `min-height` + 去掉
  `overflow: hidden`；分栏的 `column-*` 不再带高度（带高度时溢出会往**右**长列：
  `pmScrollWidth 1242 > clientWidth 602`，最后一段落在 x=960）；页面边框与水印改用
  `position: fixed`，Chromium 打印时每页重画一次，效果与 Word 把水印放页眉里一致。
  修后实测：**2 页**，第 1 页左栏段 1–21 / 右栏段 22–42，红色上边框与对角"机密"水印都在，
  逐页重复（`sips -s format png` 渲染 `tmp/ps-cols-F.pdf` 第 1 页肉眼确认）。
- 分栏画布（本片**没**改成重排，原因记在这里）：`PagedLayout` 是在**单个** ProseMirror 内容元素上
  按块高度算分页、靠 `.oc-page-gap` 占位 widget 造出多页视觉；CSS multicol 会把 N 栏横向铺满整个
  盒子，"每页 N 栏"需要每页一个 multicol 容器，与这个架构不可调和。所以编辑视图仍只画栏边界导引线
  （`b93c1ef` 的取舍），而**保存出的 `.docx` 与导出的 PDF 是真分栏**。画布导引线的实测几何：
  415.767px = 120.2 + 295.6（第二条栏边界，与页边距 + 栏宽 + 栏间距一致）。
- 行号顺带更正一处**假记录**：8.7 写着"真实行号由 Word 与打印/PDF 链路渲染"——`printHtml.ts` 里
  没有任何行号代码，这句是错的。事实只有"数据完整写进 `.docx`，由 Word/WPS 渲染"，画布与 PDF 都不画。
- 类型 + 测试：`pnpm build` EXIT=0（`grep -c "error TS"` = 0）；`pnpm test` EXIT=0，逐包 docx-io 293
  （8 个文件）/ desktop **397 → 398**（34 个文件；新增 `pageWatermark.test.ts` 9 条，`printHtml.test.ts`
  水印 4 条 + 分栏重写 + 新增"不截断"与边框 `position: fixed` 各 1 条）/ word-client 8 / word-tools 6 /
  server 2；`apps/cli`、`packages/shared`、`apps/task-pane` 仍无测试文件。
- UI：真实 app 探针（`tmp/ps-usability-probe.mjs`、`tmp/ps-reload-probe.mjs`，独立 `--user-data-dir` + CDP，
  不碰你在跑的 app）实跑数字：重新打开 `tmp/ps-user.docx` 后画布出现 7 个水印层，
  计算样式 `font-size 156px` / `rgb(192,192,192)` / `opacity .5` / `rotate(-45deg)` / `z-index 1`，
  点水印落到的元素是 `P`（说明没挡住编辑）；对话框回显栏数 2、"启用水印"勾选、字体"黑体"；
  控制台 0 条 error。截图 `tmp/ps-reload-canvas.png`、`tmp/ps-cols-F-1.png`。
- Word 实测：`[!] BLOCKED`——同 §8.12，装了 WPS 但窗口不可捕获。上面的"画布可见 + PDF 逐页"
  是**本机 Chromium 实测**，`.docx` 里 Word 怎么排这些属性仍是推断。
- 遗留：① 行号画布与 PDF 都不显示（只有 `.docx` 有）；② 分栏在编辑视图只画导引线、不重排文字，
  与 PDF/Word 行为不一致，要改成真重排得换掉分页架构；③ 桌面原地保存仍没有 `.bak`（AGENTS §8
  安全红线第 4 条，仍等你决定）。
- 提交：`d079287` + `177f21b` 已推送，`git ls-remote origin main` 分别与两者一致。

### 8.14 排名 7 可用性修复 ②（页面设置其余档位实跑 + 横向宽高约定）验收记录（`57bcccd`）

§8.13 修完水印与导出后，把页面设置剩下的档位一次跑完（`tmp/ps-audit-probe.mjs`，独立
`--user-data-dir` + CDP 起真实 app，跑的是你那份文档的副本 `tmp/ps-audit.docx`），
顺带挖出一个**方向**上的真 bug。

- 实测数字（每一档都看三件事：改完画布动不动、重开对话框回显、存盘 XML）：
  纸张 Letter + 方向横向 + 四边距(上2/下2/左4/右3cm) + 页眉页脚距边 1.2cm + 行号(每 2 行/起始 7/距正文 0.8cm)
  → 画布 `.page-card-real` 从 `793.733 × 1122.53px` 变成 `1056 × 816px`，
  编辑器宽度 553.727 → 791.398px，`--oc-margin-top` 96 → 75.6px、`--oc-margin-left` 120 → 151.2px、
  `--oc-header/--oc-footer` 56.73/66.13 → 45.4/45.4px（全部等于 cm×96/2.54，没有第二处换算，§2.1）；
  重开对话框回显 `Letter`、"横向"是 active、六项边距与行号四项逐字对得上（2/2/4/3/1.2/1.2、勾选/2/7/0.8）；
  存盘 `<w:pgMar w:top="1134" w:right="1701" w:bottom="1134" w:left="2268" w:header="681" w:footer="681">`
  （2cm=1134twips、3cm=1701、4cm=2268、1.2cm=681 ✓）、`<w:lnNumType w:countBy="2" w:distance="453" w:start="7"/>`
  ✓；正文 195 段 5440 字一字不缺；控制台 0 条 error；截图 `tmp/ps-audit-canvas.png`。
- **行号仍然是唯一"勾了看不到"的一项**：画布上 `[class*=line-num]` 0 个，PDF 里也没有——
  与 §8.13 里水印修复前同一状态。这一项要不要做预览、做到哪一步（只画布 / 画布+PDF），等你定；
  §8.7 当初的取舍是"逐行编号会把 `PagedLayout` 的分页测量打乱"，理由仍然成立。
  **→ 本条已由 §8.15 的 `12216f0` 结掉**：你定的范围是"只加编辑画布行号"，PDF 仍不画行号。
- 方向的真 bug（`57bcccd`）：`buildSectPr` 与 `readSection` 按 `w:orient` **交换**宽高，
  也就是横向页写出 `w:w=12240 w:h=15840`（短边在 w）。微软官方
  [How to: Change the print orientation](https://learn.microsoft.com/office/open-xml/word/how-to-change-the-print-orientation-of-a-word-processing-document)
  明说"改方向不够，还必须交换宽高"——Word 自己存横向 Letter 是 `w:w="15840" w:h="12240" w:orient="landscape"`。
  两头都错：我们存的横向文档在 Word 里是"横向标记 + 竖版尺寸"（打印按纸张方向转一次，正文会歪/超出），
  而 **Word 存的横向文档在我们这里被读成竖版形状**。改法是进出口都不交换，`w:w` 就是屏幕上看到的宽，
  `w:orient` 只当标记（`pageWidth > pageHeight` 推出来）。
- 判据与实测：`tmp/ps-orient-probe.mjs` 真实 app 跑两遍（第二遍是**关掉重开**同一份文件）——
  选 Letter + 横向 → 画布 `1056 × 816px`（编辑器 815.6px）→ 存盘
  `<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>` → 重开画布仍是 `1056 × 816px`。
  读写约定自此自洽，且与 Word 一致。
- 为什么以前没暴露：方向此前**零测试覆盖**（`grep -l landscape packages/docx-io/src/*.test.ts` 空）。
  补 3 条：读 Word 的横向页必须得到宽页、存横向时长边写进 `w:w`、横向文档二次保存字节一致。
- 类型 + 测试：`pnpm build` EXIT=0（`grep -c "error TS"` = 0）；`pnpm test` EXIT=0，逐包
  docx-io **293 → 296**（8 个文件）/ desktop 398（34 个文件）/ word-client 8 / word-tools 6 / server 2；
  `apps/cli`、`packages/shared`、`apps/task-pane` 仍无测试文件。
- Word 实测：`[!] BLOCKED`（同 §8.12）。宽高约定是**官方文档 + 与 Word 产出样例的一致比对**，
  不是亲眼看到 Word 渲染；`w:pgMar`/`w:lnNumType` 的 twips 数值是浏览器实跑 + 单位换算核对。
- 提交：`57bcccd` 已推送，`git ls-remote origin main` = `57bcccd`。

### 8.15 排名 7 可用性修复 ③（行号画进编辑画布 + 分栏/行号生效范围说明）验收记录（`12216f0` `6856b9a`）

你定的范围（AskUserQuestion 各二选一）：行号 →**只加编辑画布行号**，打印 / PDF 不画；
分栏 →**保持导引线**，画布不重排正文，只在对话框里加一句说明，免得下次又被当成 bug。

- 判据（Word 语义，`w:lnNumType`，ISO/IEC 29500-1 §17.6.9）：`w:countBy` = 每几行显一个号，
  显示的值仍是**这一行本身的号**（所以 `countBy=2 start=7` 是 7、9、11…，不是 2、4、6）；
  `w:distance` = 号与正文的间距；`w:start` = 起始号；`w:restart` 的缺省是 `newPage`，
  规范例子原话 "each line, restarting at one at the top of each new page" → **每页从 `start` 重新数**。
- 实现：`apps/desktop/src/editor/lineNumbers.ts` 只做几何与计数（不碰 DOM，可单测），
  `PageChrome` 负责量布局——顶层块（列表按 `w:numPr` 的每一条单独算）摊成视觉行，
  坐标一律 `.page-stage` 系（屏幕差值除以缩放），**按行中心落在哪一页的纸面**归页。
- 踩到并修掉的两处错位：① 号浮在页顶之上 96px——我拿"页高 × 页码 + 页边距"自己算了坐标系，
  而 `.ProseMirror` 的 rect 本来就在正文顶部；② 段中分页时号不复位——第一版按
  `pageNumberOf(块首)` 给整块定页，跨页段落的后半截把上一页的大号码（49、51、53）写进了下一页顶部。
- 实跑证据（`tmp/ps-ln-page.mjs`：真实 app + 独立 `--user-data-dir` + CDP，文档 = 你那份简历的副本；
  每 2 行一个号、起始 7）：7 张 `.page-card-real` **每一张**首个号都是 7、序列 `7,9,11,13`、
  `stepOk` 7/7 true、落在所有页范围之外的号 0 个。`tmp/ps-linenum-probe.mjs` 另测：
  `countBy=1` 时 267 个号，`countBy=2/start=7` 时 135 个，打字后 135 → 153（跟着重排）；
  号宽 `96px` = 左边距 120px − `w:distance` 24px（0.635cm，Word 的 0.25 英寸缺省）；
  颜色 `rgb(138,138,138)`、`pointer-events:none` 不挡正文；首行对齐：号顶 120px vs 正文首行
  Range 顶 122px（差 2px 是 half-leading，号高 = 行高 18.9px 居中）；控制台 0 条 error；
  截图 `tmp/ps-ln-page.png`（对话框那两句说明见 `tmp/ps-hint.png`：栏数 >1 才出现分栏那句、
  勾了行号才出现行号那句，单栏未勾时都收掉）。
- **环境坑（本轮最费时间的一条，已同步进 AGENTS §8）**：探针 `npx electron apps/desktop`
  不带 `VITE_DEV_SERVER_URL` 时走的是 `win.loadFile(dist/index.html)`（`apps/desktop/electron/main.ts:109`），
  量的是**上一次 build 的包**。所以"改完 src 重跑探针数字一模一样"不是缓存错觉；
  `curl http://127.0.0.1:5174/src/...` 看到新源码只能证明 dev server 是新的，证明不了 app 跑的是它。
  改完必须先 `pnpm build`（或 `pnpm --filter @office-copilot/desktop exec vite build`）再跑探针。
  顺带一条：`packages/sheet-io/dist` 落后于它的 src 会让桌面 renderer 的 rollup 构建直接失败
  （`"readOnlyReason" is not exported`），先 `pnpm --filter @office-copilot/sheet-io build` 再 build。
- 类型 + 测试：`pnpm build` EXIT=0（日志里 `error TS` 0 处）；`pnpm test` EXIT=0，逐包
  desktop **416**（36 个文件，含新 `lineNumbers.test.ts` 6 条；§8.14 时是 398/34，中间还进了
  另一位作者的 sheet 用例）/ docx-io 296（8 个文件）/
  sheet-io 45（7 个文件，另一位作者的包）/ word-client 8 / word-tools 6 / server 2；
  `apps/cli`、`packages/shared`、`apps/task-pane` 仍无测试文件。
- Word 实测：`[!] BLOCKED`（原因见 §8.12，装了 WPS 但看不到渲染画面）。"每页从 `start` 重新数"
  是按 ISO 原文 + Word 帮助页**推断**并照此实现，Word 里真实渲染没亲眼看到。
- 遗留（写清楚，不算本片做完）：① 打印 / PDF 不画行号（你定的范围，对话框里已写明）；
  ② Word 在多栏里是**每栏各画一组**行号（规范原话 before each column of text），我们按整页单栏画；
  ③ 表格与列表按块自身 computed `line-height` 摊行，跨页表格的行号不会与 Word 逐行对齐；
  ④ `w:restart` 仍不建模（只靠"值等价就不重建"保住文件里的原值），文件写了 `restart` 的非缺省值时
  画布预览照旧每页复位；⑤ 桌面原地保存仍没有 `.bak`（AGENTS §8 红线 4，仍等你决定）。
- 提交：`12216f0`（行号）+ `6856b9a`（两句说明）+ 本条 docs 与 AGENTS 记录一起推送；推送后
  `git ls-remote origin main` 与本仓 `git rev-parse HEAD` 一致才声称已推送。

### 8.16 排名 7 收尾总检（分栏 / 页面边框 / 水印逐字段实跑，含"版式=横向"的水印居中复核）

§8.14 那轮总检只跑到纸张 / 方向 / 边距 / 行号，**页面边框与水印的逐字段、以及分栏的存盘数值**
没在同一轮里看过。这一轮补齐（`tmp/ps-final-probe.mjs` → `tmp/ps-final.log`，真实 app + 独立
`--user-data-dir`，文档 = `tmp/ps-final.docx`）。

- 填进去的值：栏数 2 / 栏间距 1cm；边框 上=双线 3pt 间距4 #cc0000、左=虚线 1.5pt 2 #0070c0、
  下=单线 1pt 1 #00b050、右=无；水印 启用 + 半透明 + 文字"机密" + 字号 60pt + 颜色 #dddddd +
  版式"横向"；行号 启用（每 1 行 / 起始 1）。
- 存盘 XML（`word/document.xml` 与 `word/header1.xml` 都是读回真文件，不是断言推断）：
  `<w:pgBorders w:offsetFrom="page">` 里 `<w:top w:val="double" w:sz="24" w:space="4" w:color="CC0000"/>`、
  `<w:left w:val="dashed" w:sz="12" w:space="2" w:color="0070C0"/>`、
  `<w:bottom w:val="single" w:sz="8" w:space="1" w:color="00B050"/>`，**右边没线就不写** ✓
  （`w:sz` 是 1/8 pt：3pt→24、1.5pt→12、1pt→8 ✓；子元素顺序 top→left→bottom 合 `CT_PgBorders`）；
  `<w:cols w:space="567" w:num="2"/>`（1cm=567twips ✓）、`<w:lnNumType w:countBy="1" w:distance="363"/>`
  （起始 1 是缺省所以不写 `w:start` ✓）、`<w:pgSz w:w="11906" w:h="16838"/>` 与 `w:pgMar` 未被牵连 ✓；
  水印写进 `word/header1.xml` 的 VML：`type="#_x0000_t136"`、`fillcolor="#dddddd"`、
  `<v:fill on="t" opacity="50%"/>`（半透明 ✓）、`v:textpath style="font-family:'宋体';font-size:60pt"
  string="机密"` ✓，版式选"横向"时不写 rotation ✓。
- 画布：`.oc-page-border` 21 条（3 条有线 × 7 页，线宽/线型/颜色经 CSS 变量交给
  `pageBorders.css`，其几何由 `pageBorders.test.ts` 6 条单测机检）、`.oc-column-guide` 7 条、
  `.oc-line-number` 262 个、`.oc-page-watermark` 7 份；重开对话框四条边的线型/线宽/间距/颜色与
  水印六个字段逐字回显（含"右=none"）✓；正文 195 段 / 5440 字一字不缺 ✓；控制台 0 条 error ✓；
  截图 `tmp/ps-final-canvas.png`（红双线在上、蓝虚线在左、中线是导引线、左边距里一列行号、
  右下那团淡灰是"机密"）。
- **顺手复核的一条疑点（结论：不是 bug）**：截图里水印看着偏右下，怀疑"版式=横向"时盒子没居中，
  于是量了一次（`tmp/ps-wmpos-probe.mjs`）：对角与横向两种版式下 `.oc-page-watermark` 的
  **盒子中心都正好落在页面中心** `[397, 561]`（页 794×1123），`transform` 分别是
  `matrix(0.707,-0.707,0.707,0.707,0,0)`（-45°）与单位矩阵。看着偏是因为旋转盒的**外接矩形**
  变大了（对角时 `wmRectCss` 是 552×552），而文字本身仍居中。
- 一个探针侧的坑（记下来免得再误判）：`.ps-border-grid` 的行序是 **上 / 左 / 下 / 右**
  （`BORDER_SIDES` 按 Word 那个方框的阅读顺序排），不是上/下/左/右；按位置索引写控件前先读回
  行名 span，否则会把"左"的虚线当成"下"来核对。
- 门禁：本轮无代码改动（只补文档与探针），沿用 §8.15 末尾那次 `pnpm build` EXIT=0 /
  `pnpm test` EXIT=0 的逐包数字（desktop 416 / docx-io 296 / sheet-io 45 / word-client 8 /
  word-tools 6 / server 2）。
- Word 实测：`[!] BLOCKED`（§8.12）。边框 `w:sz` 的 1/8pt 口径、`v:fill/@opacity` 的写法是
  ECMA-376 §17.6.1 + Word 产出的 VML 样例比对，不是亲眼看到 Word 渲染。

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
