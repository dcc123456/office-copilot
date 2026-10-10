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
