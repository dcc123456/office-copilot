# 全格式文档编辑规格（Office / WPS parity spec：pptx + pdf 两条能力线）

> 立档日期：2026-10-14。触发指令（dcc 原话）："你应该要去 github 上搜索 office 或 wps 相关的 excel 文档编辑相关的项目，
> 然后将它们的文档编辑功能全部抄到这个项目中。我希望这个项目不仅支持 doc excel，还要支持 ppt pdf 等所有类型的文档编辑功能，
> 然后 AI 能智能操作。"
>
> 本文档只管 **pptx（演示）** 与 **pdf** 两条新线。`.docx` 线的规则见 `docs/word-parity-spec.md`，
> `.xlsx` 线见 `docs/excel-support-spec.md`（D9 拆列/合列仍未落码，不在本文范围）。

---

## 零、全局约束（每一节都受它管）

1. **既有红线一条不减**：`AGENTS.md` §2 全部适用。pptx/pdf 不是新类别，只是把"未建模部件原样带过 + op 自带逆操作 +
   内容寻址撤销 + 单位换算只在 IO 层进出口 + 数字一律引擎算 + 模型只出提案"这六条搬到新的部件树上去。
2. **"抄"的边界先说死（许可红线，逐条一手核对过，见 §二）**：
   - **AGPL-3.0 一行代码都不搬**：`PPTist`、`ONLYOFFICE/DocumentServer`、`MuPDF`。网络 copyleft 会把整个衍生品要求开源，
     本仓是私有产品。这三家只能抄**功能清单与交互词汇**。
   - **Apache-2.0 / MIT / MPL-2.0 可以搬码**，但必须保留 `LICENSE`/`NOTICE` 并在文件头标注改动点与来源 commit。
     能搬的只有 `Univer`（OSS 部分）、`genoffice`、`OfficeCLI`、`pdf-lib`、`jsPDF`、`qpdf`、`pikepdf`、`pdf.js`、`pdf2docx`。
   - **`Univer` 的 Pro 包不是 OSS，不得参考其实现**（详见 P1）。
3. **一次只推进一个能力**（§0 的既有纪律）：pptx 与 pdf 虽然"并行立项"（dcc 裁决），但**落地串行**——
   任何时刻只有一条线在写产品代码，另一条线只允许停在取证与夹具。理由是这两条线都要动同一批公共层
   （OPC、原子写、modlog、工具 schema），并行改公共层会互相踩。
4. **交付物必须能被办公软件打开**：pptx → Keynote / WPS 演示；pdf → 任意阅读器 + pdf.js 自己读回。
   判据见 §六 A 组，**本机实测通道已经找到（见 §二 P9/P10 与附录 A），pptx 线不再是"完全没视觉通道"**。
5. **AI 不许自己算数、不许交原始 XML**。pptx 的图表数值走 `sheet-io` 现成的 `resolveChartSpec` 那道闸（§2.11）；
   pdf 的页码/坐标一律由引擎从内容流算。参考项目里 OfficeCLI 的 `raw`/`raw-set` XPath 逃生口**明确不抄**（违反 §2.8）。

---

## 一、总目标

把本仓从"文档 + 表格"扩成"**文档 + 表格 + 演示 + PDF**"四类可编辑格式，三类客户端（桌面 app / CLI / 加载项）
共用同一批引擎包，AI 在四类格式上都能出**结构化提案**并一键落地、一键撤销。

拆成三个可独立验收的子目标：

- **G1（pptx）**：桌面 app 能打开 `.pptx`，看幻灯片、改文字/形状/图片/表格/图表/段落格式，存回去 Keynote 与 WPS
  照原样打开且**没被改的东西一个字节都不动**；每次改动能撤销、能进修改历史。
- **G2（pdf）**：桌面 app 能打开 `.pdf`，读得出页/文本块/字体/图片，**真·正文改写**（dcc 裁决的那一档：不是标注层，
  是改到正文里的字），改完仍是合法 PDF，能撤销。
- **G3（AI）**：四类格式共用一套"工具面 → 提案 → 确认 → 落地 → 撤销"的装配，`sheetAgent`/`agent` 的模式不再各长一份。

**范围界定（"全部抄过来"必须落成一张有边界的表，否则无法验收）**——按四档切，档位来自 §二 的功能清单：

| 档 | 含义 | 包含 |
| --- | --- | --- |
| **T1 必做** | 没有它就不叫"能编辑" | pptx：读+保真回写、母版/版式/占位符继承链、文本与段落、形状与填充、图片、表格、节与页眉页脚页码、备注页、主题、切换效果（结构级）、撤销与 modlog、AI 提案面。pdf：读、结构审计、文本层改写（我们自产的那一档）、撤销 |
| **T2 次做** | T1 站稳后按序补 | pptx：图表（含内嵌工作簿）、连接线与成组/层级、批注（legacy + 现代）、图形对齐分布。pdf：任意第三方 PDF 的 CID/ToUnicode 反查改写 |
| **T3 只保真带过** | 读得懂、原样带回，**不提供编辑入口** | pptx：`ppt/tags/*`、动画 timing 树、SmartArt(`dgm`)、OLE、3D `.glb`、嵌入字体、`handoutMasters`、`printerSettings`、缩略图。pdf：表单控件、数字签名、注释层 |
| **T4 明确不做** | 写进文档免得以后反复问 | pptx：母版可视化重建器、自定义 `custGeom` 路径编辑器。pdf：扫描件 OCR、版面重排回流、加密文件破解 |

**完成判定（总目标）**：G1/G2/G3 各自的 §六 验收表全绿，且 `README.md` 的"项目是什么"一句话里四类格式都成立。

---

## 二、选型研究结论（先取证，后写码 —— `AGENTS.md` §6.1/§6.2）

许可与星数是 2026-10-14 用 GitHub / PyPI / npm API 一手核对的，不是转述。

| # | 来源 | 许可（实测） | 读了什么、结论 |
| --- | --- | --- | --- |
| **P1** | `dream-num/univer` 22 589★ | Apache-2.0（默认分支 `dev`，OSS 60 个包） | **不能当内核**。README「🔓 Open Source and Pro」表（`dev` 分支 301-323 行）逐条写明 **Sheets/Docs/Slides 的 import/export、charts、pivot、print、collaboration 全在 Pro/商业层**，OSS 里只有数据模型 + 编辑 UI + 公式引擎 + 渲染引擎 + Facade + Node headless/RPC。我们要的恰恰是"保真读写 OOXML"那一侧，而它闭源。**可抄的**：分层方式（`core`/`engine-render`/`engine-formula`/`ui`/`server`）、`protocol`+`rpc`+`action-recorder` 的"命令即记录"思路、`presets` 的装配方式、生态里的 `dream-num/univer-sdk-skills`（AI skills 的包装形态）。 |
| **P2** | `genspark-ai/genoffice` 9 606★ | Apache-2.0 | 与本仓目标最像的一个：apps 含 docs/sheets/slides/pdf/markdown/html/shell，packages 含 `pptx-engine`/`pptx-ops`/`pptx-render`/`xlsx-gateway`(Rust)/`pdf2docx`/`agent-core`。README 自称 "byte-preserving. Only what you edit is rewritten"——**与我们的"未建模部件原样带过"同一条路线，可当作可行性佐证**。**可抄的**：`pptx-ops` 的 op 词汇表（6 组约 55 条，每条一个 `### opName` + 签名 + JSON 样例，且有测试断言"文档覆盖注册表且每个样例都真跑过 dry-run"）、`target:{slide, el}` 的**持久元素 id**（`a16:creationId` GUID → `e_<guid8>`，回退 `cNvPr id`）、`srcPara`/`srcRun` 携带源索引以免把格式拍平、`slides audit` 的几何 lint（越界/溢出/重叠/图片形变>5% 且各带一条现成的 `setTransform` 建议）。**不抄的**：它没有撤销历史（我们必须有，§2.5）。 |
| **P3** | `iOfficeAI/OfficeCLI` 31 967★ | Apache-2.0 | .NET 单文件 + 自有 OOXML 引擎，覆盖 docx/xlsx/pptx（**无 PDF**）。命令面 `create/view/get/query/set/add/remove/move/swap/validate/batch/dump/merge/save/raw-set`，路径寻址 `/slide[1]/shape[1]`（**1 起始、元素本地名、非 XPath**，支持 `@name=`/`@id=`/`[last()]`）。**可抄的**：QA 闸门顺序（`validate` schema → `view issues` → 残留占位符扫描 → 强制出 PNG 做视觉审计）、以及"位置索引在结构操作后会漂移，所以按名字寻址才稳"这条约束。**明确不抄**：`raw`/`raw-set` 让模型直接改任意 XML part——违反 §2.8 与本仓"引擎只出结构化 op"的底线。 |
| **P4** | `pipipi-pikachu/PPTist` 9 384★ | **AGPL-3.0 → 只抄功能清单** | 9 384★ 的 Vue 幻灯片编辑器，是最好的"一个演示编辑器该有哪些功能"清单来源（文本/形状/图片/表格/图表/连线/成组/层级/对齐分布/切换/动画/备注/节/主题/页眉页脚）。**它的架构恰好是我们的反例**：`pptxtojson` 导入（自述约 85% 保真）+ `pptxgenjs` 导出（约 95%）= **重新生成型**模型，导出时 `defineSlideMaster` 造一个合成的 `PPTIST_MASTER`，占位符链接被拍平成绝对文本框，动画/切换/节/批注/`tags`/母版全丢。它的 `doc/Q&A.md` 自己承认"实现不了的部分没有办法"。**所以：抄它的功能面，架构反面教材。** |
| **P5** | `ONLYOFFICE/DocumentServer` 7 012★ | **AGPL-3.0** + 商业双授权 | 功能天花板参照物（四类格式全有 + 协同）。许可不允许搬码。 |
| **P6** | `Hopding/pdf-lib` 8 656★ | MIT | 纯 JS，能改现有 PDF 的对象树并写回，是 pdf 线唯一能在 Electron/Node 里直接用的写侧库。**能力上限要实测**（见 D8）。 |
| **P7** | `parallax/jsPDF` 31 310★ / `pdfdora/pdf2docx` | MIT / MIT | jsPDF 只能**生成**不能改现有文件；`pdf2docx` 是"PDF→可编辑 docx"的转换路线（许可 MIT，可搬），值得作为 D8 乙档的一条备选通路记录。 |
| **P8** | `ArtifexSoftware/mupdf` 2 998★（npm `mupdf` = AGPL-3.0-or-later） | **AGPL-3.0 → 不搬** | 能力最强（渲染 + 内容流 + 字体）但许可是红线，**连 WASM 版都不行**。这条要先立住，否则很容易在"pdf 渲染怎么办"时顺手引进来。 |
| **P9** | `qpdf/qpdf` 5 480★（Apache-2.0）/ `pikepdf/pikepdf` 2 814★（MPL-2.0）/ `mozilla/pdfjs`（`pdfjs-dist` = Apache-2.0） | 见左 | qpdf/pikepdf 是 C++/CPython 绑定，本机 `command -v qpdf pdftotext mutool gs soffice` **全部 not found**，没有 brew 通道就不能假设 CI 上有 → 只能当"可选外部工具"，不能当依赖。`pdfjs-dist` 是 Apache-2.0 的**渲染 + 文本层提取**，读侧可用。 |
| **P10** | **本机实测（2026-10-14，附录 A 有原命令与输出）** | — | ① `python3 3.14.2` + 新装 `python-pptx 1.0.2`（已 vendor 进 `tmp/pylibs`，与 openpyxl 3.1.5 同址）→ pptx 的**结构交叉读取器**成立。② **`Keynote.app` 14.2 能开我们生成的 `.pptx`**：`count of slides` = 2，`object text of text items` 读回"标题甲/第一行\n第二行/标题乙/文本框内容"。③ **Keynote 能 `export ... as slide images`**：拿到 `tmp/kn-export/kn-export.001.png`、`.002.png`，**PNG 里中文字形真的渲染出来了**（已用肉眼看过图）。④ Keynote 还能 `export ... as Microsoft PowerPoint`（26 909 字节，python-pptx 读回形状名与文字都对）与 `export ... as PDF`（13 915 字节）。→ **pptx 线第一次有了"第三方办公软件真读过这份文件 + 画面看得到"的双层证据通道**，不再是 xlsx/docx 那种目视 BLOCKED。⑤ 本机**没有** Microsoft Office、**没有** LibreOffice，WPS 无可视通道（不变）。 |
| **P11** | **本机实测：PDF 正文为什么不是字符串替换**（附录 A） | — | Keynote 导出的那份 PDF：`%PDF-1.3`、12 个流、字体只有 `AAAAAB+PingFangSC-Regular` 与 `AAAAAC+ArialMT`（**子集字体**，前缀 `AAAAAB+` 是子集标记），编码 `Identity-H` + `MacRomanEncoding`，`/ToUnicode` 只有 **1** 个而嵌入字体有 **2** 个，`/FontFile` 2 个。第一页内容流里文字是 `q 1 0 0 1 294 457.735 cm BT 44 0 0 44 0 0 Tm /C1 1 Tf <0ebd1dcc12a8> Tj ET`——**CID 十六进制串，不是字符**，而且部分文字带 `/Span <</ActualText (...)>> BDC … EMC` 标记内容。结论：**"把正文里的 A 改成 B"在任意第三方 PDF 上 = 反查 CID→Unicode + 往子集字体里合并新字形**，子集里没有那个字形就根本写不出来。这条决定了 D8 必须分两档，不能对 dcc 承诺"任何 PDF 都能改正文"。 |
| **P12** | `xuri/excelize`（Go, BSD-3）/ `yb2460/harness-anything`（AI 代理控 WPS/Office） | — | 前者语言不通（只作 xlsx 侧思路参考，已完成）；后者是"用自动化驱动真办公软件"路线，本机 WPS 无可视通道（§8 实测 `-1728`），**不作为产品依赖**，只作探针手法参考。 |

---

## 三、架构决策

### D1 — 内核自研，不引入外部运行时；`Univer` 只作分层与 AI 包装的参考

**问题**：dcc 点名"去看 Univer"，隐含选项是"拿它当内核省一半工"。
**核查结论**（P1）：Univer 的 OSS 部分**没有 OOXML 读写**，import/export 在闭源 Pro 层；把它当内核意味着
我们的核心交付物（"能被 Excel/PowerPoint 正常打开的文件"）依赖一个买来的黑盒，且 OSS 分支里根本没有那条路。
**决策**：pptx/pdf 的保真引擎**自研**，复用本仓已经验证过的路线（原始字节回写 + 未建模部件原样带过 + op 自带逆操作）。
从 Univer 只借三件：包分层命名、`action-recorder` 那种"命令流水即审计"的思路、以及它把 AI 能力包成 SDK skills 的形态。
**代价（如实写）**：pptx 的部件树比 xlsx 大（一个 WPS 默认模板 35 个 part，最富的那个 125 个 part 含 65 个 `tags`），
自研的工作量比 xlsx 线大，必须靠 §四 的 T1/T2/T3 分档控制范围。

### D2 — 先抽 `packages/opc`，不允许出现第三套 OPC 实现

**问题（实测到的现状，不是假想）**：本仓**已经有两套** OPC 处理——`packages/docx-io/src/docx.ts` 自带一套
rel/content-type 逻辑，`packages/sheet-io/src/opc.ts` 另一套（`loadPackage`/`parseRelationships`/`resolveTarget`/
`unmodelledParts` 全是通用的，但**包私有**，`sheet-io/src/index.ts` 没导出）。pptx-io 一进来就是第三套，直接撞 §3.2/§3.3。
**决策**：pptx 阶段 0 之前先把 OPC 提到 **`packages/opc`**（新包需说明理由，§3.4：**它只做一件事——按 ECMA-376 Part 2
读写 zip 包、部件枚举、关系图、内容类型登记、原始字节回写**；docx/sheet/pptx 三形都需要，且 pdf 线的 xref 模型完全不同，
不塞进去）。`sheet-io`/`docx-io` 改为依赖它的公开入口，删掉各自那套平行实现。
**顺序红线**：**先抽公共层并让 xlsx/docx 两线测试全绿，再动 pptx**。不允许"先复制一份 opc.ts 到 pptx-io，以后再抽"——
那正是 §3.2 禁止的那一步。

### D3 — `pptx-io` 的模型：母版/版式/占位符的继承链**只读解析、不写回拍平**

**问题（P2 实测到的最高保真风险）**：一个 `ctrTitle` 形状在 `slide1.xml` 里 `<p:spPr/>` 与 `<a:bodyPr/>` 都是**空的**，
真实几何在 `slideLayout1.xml`、字号在 `slideMaster` 的 `p:txStyles`、字体颜色在 `theme1.xml`。
**决策**：读模型必须给出 **`effective`（算完继承的有效值）+ `src`（这个值来自哪一层）** 两个字段，
沿用 genoffice 的 `srcPara`/`srcRun` 思路；写侧只允许写"这一层自己显式带的属性"，
**禁止把继承来的几何/字号烤成绝对值写进 slide**（一烤就永久断掉版式与主题的联动，等价于 §2.3 的"拍平局部格式"）。
匹配键是 `p:ph/@type` + `@idx`。
**验收**：一份夹具里"只改标题文字、不碰版式"之后，`slide1.xml` 里那个 `<p:spPr/>` 仍必须是空的。

### D4 — 单位与颜色：EMU / 百分点 只在 `pptx-io` 进出口换算（§2.1 的 pptx 版）

实测到的换算表（写进代码注释，别处再出现即违规）：`1 inch = 914400 EMU`、`1 px@96dpi = 9525 EMU`；
字号是**百分点**（`sz="1800"` = 18pt）；行距 `spcPct val="90000"` = 90%、`spcPts val="1000"` = 10pt；
`lIns="91440"` = 0.1 英寸；16:9 页面 `sldSz cx="12192000" cy="6858000"`。颜色 `a:srgbClr val=` 或
`a:schemeClr val="tx1"` + `tint`，要经 `p:clrMap`（`bg1="lt1" tx1="dk1"`）解析到 `theme` 的 `clrScheme`——
**schemeClr 必须解析成 hex 再交给 UI**，但不许把解析结果烤回 slide（同 D3）。

### D5 — op 与撤销：形状 id 用 `a16:creationId` GUID，回退 `cNvPr id`

**决策**：pptx 的改动块以"形状"为粒度寻址，id 取 PowerPoint 已经写进 `p:extLst` 的 `a16:creationId`（实测 master 与
layout 里都有 `p14:creationId`），拿不到才回退 `cNvPr id`——因为**位置索引在插入/成组/z 序变化后会漂移**（P3 的同一坑）。
op 形状与 xlsx 一样自带逆操作，落进**同一份** modlog（`userData/modlog/<文档名>.json`，`apps/desktop/electron/modlog.ts`
已经是通用 JSON，直接复用，不新建第二套历史栈）。
**新 op 种类（T1）**：`setText`、`setParagraphFormat`、`setRunFormat`、`setShapeProperty`(xfrm/fill/stroke)、
`insertSlide`/`deleteSlide`、`setSection`、`setNotes`。**T3 的部件不产生 op**（D9 的"判定不了角色就不出 op"同一条）。

### D6 — 工具面：抽一份通用 schema 核心，各格式仍各自声明工具

**问题**：`parseToolArguments` + zod 校验这套已经存在**两份**（`packages/sheet-tools`、`packages/word-tools`），
pptx 会是第三份（§3.2 说第二次就该抽）。
**决策**：抽 `packages/agent-tools`（或落到 `packages/shared`，阶段 0 定）：只放"工具声明 + zod 入参解析 +
永不抛异常地回中文理由"这个通用核心；`SHEET_TOOLS`/`WORD_TOOLS`/`PPT_TOOLS` 各自保留自己的工具名与 schema。
装配层（`sheetAgent.ts` 的 8 轮工具循环、`ai:tool-turn` 传输）已经通用，直接复用。

### D7 — 幻灯片视图：`App.tsx` 的 `isSheet` 布尔必须先换成模式枚举

`apps/desktop/src/App.tsx:1118` 现在是 `isSheet = sheetWorkbook !== null`，所有文档面板用 `!isSheet` 门控。
第三个格式进来时这个布尔就不够用了。**决策**：加 `type EditorMode = "doc" | "sheet" | "slide" | "pdf"`，
一次性把 `!isSheet` 的门控改成按模式判定；扩展点同时落在扩展名路由（`:374,447` 的 `docx|json|xlsx` 正则）与
拖拽门控（`:450`）。样式仍走 `src/styles/*.css` 那一套（§5.4），幻灯片编辑区自己写视图，
**不引入第二套样式体系、也不把 Tiptap 的 `EDITOR_EXTENSIONS` 拉去渲染幻灯片**（幻灯片是绝对定位画布，不是流式文档）。

### D8 — PDF 的"真·正文改写"分两档，并对做不到的一档如实拒绝

dcc 选的是最贵的一档（真·正文改写，不是标注/表单层）。按 P11 的实测，必须落成两档：

- **甲档（T1，承诺做）**：**本仓自己产出的 PDF**（现有 `print:pdf` 通路，`apps/electron/printHandlers.ts:15`）。
  我们在生成侧保留一份旁路模型（文本块 ↔ 内容流偏移 ↔ 字体与编码），改正文时按这个映射定位到 `Tj`/`TJ`，
  重排该文本块的显示算子并写回，**不动其他页的字节**。撤销照 §2.5 记前像/后像。
- **乙档（T2，先取证再承诺）**：**任意第三方 PDF**。要改必须先解决三件事：① CID→Unicode 反查（`/ToUnicode` CMap，
  实测 Keynote 那份 2 个嵌入字体只有 1 个 ToUnicode，**缺就没有可靠反查**）；② 目标字形在子集字体里不存在时
  要合并字体程序（子集只有用到的那几个字形）；③ 改完 xref/偏移要重建。
  **准入判据写成硬拒绝**：无文本层（扫描件）、缺 ToUnicode、有数字签名、加密——**整条拒绝并给中文理由**，
  不许"看起来改了其实丢了字"。备选通路：搬 `pdf2docx`（MIT）先转可编辑 docx 再重排，作为乙档的一条实现候选记录在案。
- **明确不做**：OCR、版面回流、破解加密（T4）。

### D9 — pptx 里的图表：复用 `sheet-io`，内嵌工作簿是真 `.xlsx`

实测结构（P2 附录）：`slide3.xml` 的 `p:graphicFrame` → `a:graphicData uri="…/drawingml/chart"` →
`c:chart r:id="rId2"` → `charts/chart1.xml` → 它自己的 `.rels` → `ppt/embeddings/Microsoft_Excel_Sheet1.xlsx`，
**而 openpyxl 能直接读那个内嵌 part**（`Sheet1 dims A1:C4`，值与 `c:numCache` 一致）。
**决策**：图表写侧复用 `packages/sheet-io` 的公开入口与现成的 `resolveChartSpec`（它已经会拒绝模型自填数字和二维引用），
由引擎同时导出 **cache 与内嵌工作簿的单元格**，保证 `c:f` 区间、`ptCount`、缓存值、工作簿值四者互相自洽
——不一致就是 Excel/PowerPoint 眼里的"陈旧图表"。登记件数比 §2.10 的五件更宽：图表部件 + 图表自己的 rels +
slide 的 `graphicFrame` `r:id` + slide 的 rels + `embeddings/*.xlsx`（`<Default>` 类型），**真 PowerPoint 还会带
`colors1.xml`/`style1.xml`**。注意 pptx 的图表在 `ppt/charts/`，不是 `xl/charts/`。

### D10 — 新建部件的登记规则与结构审计：`oc ppt check`

pptx 里**每个 XML 部件都要自己的 `<Override>`**（`presentationml.slide+xml`/`slideLayout+xml`/`slideMaster+xml`/
`notesSlide+xml`/`notesMaster+xml`/`tags+xml`/`theme+xml`/`drawingml.chart+xml`/`presProps`/`viewProps`/`tableStyles`），
二进制按扩展名走 `<Default>`（`bin`/`jpeg`/`xlsx`/`rels`/`xml`）。`p:spTree` 的**子元素顺序是规范的**
（先 `p:nvGrpSpPr` 再 `p:grpSpPr`，然后 `p:sp`/`p:pic`/`p:graphicFrame`/`p:cxnSp`/`p:grpSp`）。
**决策**：照 `packages/sheet-io/src/compat-audit.ts` 的形状写一份 `pptx-io` 的审计（内容类型覆盖、关系可解析、
`r:id` 落点、`spTree` 子序），并接成 CLI `oc ppt check <file>`（有问题退出码 1）。**这是 pptx 线的 §7.4⑥ 机检证据**，
配合 Keynote 的 PNG 导出（P10）就是"结构 + 画面"双证。

---

## 四、阶段目标（每段一句完成判定）

### 阶段 0 · 取证与夹具（不写产品代码）— `[ ]`
把 WPS 自带模板（`/Applications/wpsoffice.app/Contents/Resources/office6/…/newfile.pptx` 35 part /
`new_wpp.pptx` 125 part 含 65 个 `tags`）与 python-pptx 生成的 4 页夹具（标题/项目符号/图表/表格/备注）收进
`packages/pptx-io/scripts/` 的生成脚本，**并确认本机 `tmp/pylibs` 里有 python-pptx 1.0.2**（已 vendor）。
**完成判定**：三类夹具（空模板 / 富模板 / 带图表备注）都能被 python-pptx 与 Keynote 各自读通，Keynote 能导出 PNG 且肉眼可见文字。

### 阶段 1 · 抽 `packages/opc`，两线不改行为 — `[ ]`
按 D2 把 OPC 提到公共包，`sheet-io`/`docx-io` 改依赖它。
**完成判定**：`pnpm test` 逐包通过数与抽取前**完全一致**（不许靠改测试过），xlsx/docx 产物字节级往返夹具全绿。

### 阶段 2 · `packages/pptx-io`：读 + 保真回写 — `[ ]`
部件树（presentation/master/layout/slide/notes/theme）、D3 的 `effective`+`src` 继承解析、D4 换算、未建模部件原样带过。
**完成判定**：三份夹具"打开→原样存回"后 `unzip -l` 条目集与**未改动部件的字节**逐字一致，`oc ppt check` 0 条，
python-pptx 读回形状名/占位符类型/文字全等，Keynote 打开后 `count of slides` 与文字一致。

### 阶段 3 · 幻灯片视图可用（人手动编辑）— `[ ]`
D7 的模式枚举 + 幻灯片列表 + 画布编辑（选中形状、改文字、改字号颜色、拖拽尺寸）。
**完成判定**：起真实 app 走"改标题→存盘→Keynote 导出 PNG→肉眼比对文字与位置"，并把 PNG 里的实际文字写进收尾报告（§7.2）。

### 阶段 4 · 结构化 op + 撤销/modlog — `[ ]`
D5 的 op 集合与内容寻址撤销，接同一份 modlog。
**完成判定**：每类 op 有逆操作单测；"改完→用户手工再改别处→撤销"只回到自己那一格，不牵连后面的手工编辑。

### 阶段 5 · AI 工具面 + `slideAgent` — `[ ]`
D6 的通用 schema 核心 + `PPT_TOOLS`（分析/整理/生成三类），提案卡复用现有装配。
**完成判定**：单测（mock fetch）覆盖"模型给数字/给原始 XML/给越界索引都要回中文理由"；真端点探针实测模型会不会点对工具、
落地后 Keynote PNG 里文字真的变了。

### 阶段 6 · 图表与表格（D9）— `[ ]`
**完成判定**：新建图表的 `c:numCache` 与内嵌 `.xlsx` 的值由 openpyxl 独立读回且一致；改数据后两者同步；撤销逐格回原样。

### 阶段 7 · pdf 线：读 + 审计 + 甲档正文改写（D8 甲）— `[ ]`
`packages/pdf-io` 读页/文本块/字体/内容流偏移，`print:pdf` 通路带旁路模型。
**完成判定**：本仓自产的 PDF 改一处正文 → pdf.js 读回新文字、其他页字节不动、能撤销；对无 ToUnicode/扫描件/签名文件
**给出中文拒绝理由而不是静默改坏**（逐条有单测）。

### 阶段 8 · 收口：CLI、文档、规则落地 — `[ ]`
**完成判定**：`oc ppt check`/`oc pdf check` 可用；`README.md` 与 `AGENTS.md` 里四类格式的既有事实更新；
本文档 §六 全绿或如实标 `[!] BLOCKED`。

---

## 五、执行计划（任务 → 落点）

| 阶段 | 主要落点 | 备注 |
| --- | --- | --- |
| 0 | `packages/pptx-io/scripts/make-sample-deck.ts`、`tmp/pylibs`（不入库） | 夹具由脚本现生成，不提交二进制 |
| 1 | 新包 `packages/opc`；`packages/sheet-io/src/opc.ts`（删）、`packages/docx-io/src/docx.ts`（改依赖） | 见 D2 的顺序红线 |
| 2 | `packages/pptx-io/src/{pptx-file,pptx-model,placeholder-inherit,units,color,compat-audit}.ts` | 单位换算单点：D4 |
| 3 | `apps/desktop/src/App.tsx`（模式枚举）、`apps/desktop/src/slide/**`（新视图）、`electron/slideHandlers.ts`（照 `sheetHandlers.ts` 的 `.bak` + 存前审计） | §5.2 preload 白名单同步 |
| 4 | `packages/pptx-io/src/slide-edit.ts`（op + 逆操作）、`apps/desktop/src/slide/slideHistory.ts` | 复用 modlog，不新建历史栈 |
| 5 | `packages/agent-tools`（新，D6）、`packages/ppt-tools`、`apps/desktop/src/ai/slideAgent.ts` | 工具循环复用 `ai:tool-turn` |
| 6 | `packages/pptx-io/src/chart-spec.ts` + 复用 `packages/sheet-io` 公开入口 | D9 |
| 7 | `packages/pdf-io/src/{pdf-file,content-stream,xref,text-map}.ts`、`apps/electron/printHandlers.ts` | D8 甲档 |
| 8 | `apps/cli/src/{ppt,pdf}.ts`、`README.md`、`AGENTS.md`（由 dcc 改） | — |

---

## 六、验收标准列表（逐条打分，任一不过不得声称完成）

### A. 文档正确性（最高优先级，对应 §2）
- `[ ]` A1 未建模部件原样带过：富模板（125 part / 65 `tags`）打开→存回，`unzip -l` 条目集一致且未改部件字节逐字一致。
- `[ ]` A2 母版/版式继承不被拍平：只改标题文字后 `slide1.xml` 的 `<p:spPr/>` 仍为空（D3）。
- `[ ]` A3 单位换算单点：全仓 grep 不到第二处 `914400`/`9525`/百分点换算（D4）。
- `[ ]` A4 颜色：schemeClr 解析成 hex 交 UI，且不烤回 slide；`#rrggbb` 一条不减（§2.2）。
- `[ ]` A5 原子写：幻灯片存盘走 `writeFileAtomically` + `.bak`（照 `sheetHandlers.ts:159`）。
- `[ ]` A6 新建部件登记齐全（D10 的宽版五件），审计 0 条 + `oc ppt check` 退出码 0。
- `[ ]` A7 图表 cache 与内嵌工作簿自洽（D9），openpyxl 独立读回为证。
- `[ ]` A8 PDF：改一处正文后其他页字节不动；pdf.js 读回新文字。
- `[ ]` A9 PDF 的硬拒绝成文可测：无文本层 / 缺 ToUnicode / 有签名 / 加密 → 中文理由，不静默改坏（D8）。

### B. AI 行为（对应 §2.8/§2.11）
- `[ ]` B1 模型只交引用与意图，数字由引擎算（图表、页码、坐标）。
- `[ ]` B2 模型交原始 XML / XPath 一律拒（不抄 OfficeCLI 的 `raw-set`）。
- `[ ]` B3 半自动必须点确认才落地；全自动仍受撤销约束。
- `[ ]` B4 送模型的幻灯片内容按 §2.9 截断（形状数、每形状文本 ≤120 字），不许整包上送。
- `[ ]` B5 真端点探针：模型会点对 `ppt` 工具名与形状 id；结论与原始 tool_calls 记进文档（探针在 `tmp/`，不进主干）。

### C. 可撤销性与历史（对应 §2.5）
- `[ ]` C1 每类 op 有逆操作单测；C2 内容寻址用 `creationId`，位置漂移不失效；
- `[ ]` C3 撤销不牵连用户后来的手工编辑；C4 做不到时如实报"部分撤销"。
- `[ ]` C5 复用同一份 modlog，无第二套历史栈。

### D. UI 与工程（对应 §3/§4/§5）
- `[ ]` D1 `App.tsx` 的 `isSheet` 已换成模式枚举，无新增布尔门控。
- `[ ]` D2 无第二套 OPC / 解析 / 撤销 / 工具 schema / 样式体系（§3 红线）。
- `[ ]` D3 新函数全部有中文 JSDoc（含单位与失败语义）；命名有含义。
- `[ ]` D4 渲染层不碰 Node，跨界只走 preload 白名单。

### E. 测试与门禁（对应 §7）
- `[ ]` E1 每条写入路径有单测（mock fetch），**逐包核对通过数**（`packages/shared` 无测试是已知假绿源）。
- `[ ]` E2 `pnpm build`（含 `tsc --noEmit`）+ `pnpm test` 重定向后读退出码，不许管道 tail。
- `[ ]` E3 视觉验收走 Keynote PNG 导出通道并报告看到的文字内容；跑不通就标 `[!] BLOCKED` 并写原因。
- `[ ]` E4 探针与夹具产物只进 `tmp/` 并清理；禁止 `git add -A`。

---

## 七、风险与未决

1. **工作量**：pptx 部件树规模与继承链是 xlsx 的数倍（D1 已如实写）。若 T1 也做不完，**先交付"读 + 保真回写 + 文本编辑"**，
   不要为了看起来完整而把 T3 的东西半建模——半建模比不建模更危险（会静默丢用户设置，富模板那 65 个 `tags` 就是例子）。
2. **Keynote 不是 PowerPoint**：它能开、能导出，不等于 PowerPoint 也认。凡涉及 `p:extLst` 扩展、`p18c` 现代批注，
   本机无法取证 → 相关条目一律 `[!] BLOCKED`，不许推断。
3. **动画/切换的播放效果**：无通道（genoffice 与 OfficeCLI 自己都标注"需在别的查看器里验证"）→ T3 只保真带过。
4. **PDF 乙档（任意第三方文件正文改写）可行性未定**：D8 的三条前置要在阶段 7 前用真夹具逐个实测，
   测不过就回来向 dcc 重新裁决范围，不许自行降级成"只做标注层"。
5. **未决**：`packages/agent-tools` 与落到 `packages/shared` 二选一（D6），阶段 5 前定。

---

## 八、总验收表（随进度填写，键为提交哈希）

| 提交 | 阶段 | 命令与输出（逐包通过数） | 视觉/第三方读回结论 | 状态 |
| --- | --- | --- | --- | --- |
| — | 0 | 待填 | 待填 | `[ ]` |

---

## 附录 A · 本机取证记录（2026-10-14，命令与真实数字）

```
python3                                    → 3.14.2
PYTHONPATH=tmp/pylibs python3 -c "import pptx"      → 1.0.2（本日 pip install --target tmp/pylibs python-pptx）
PYTHONPATH=tmp/pylibs python3 -c "import openpyxl"  → 3.1.5（既有）
command -v qpdf pdftotext mutool gs soffice         → 全部 none
/Applications/Keynote.app                            → 14.2
/Applications/Numbers.app                            → 存在（xlsx 线既有通道）
Microsoft Office / LibreOffice                        → 未安装；WPS 无可视通道（AGENTS §8）
```

Keynote 通道（python-pptx 先造 `tmp/pptx-kn-check.pptx`：2 页，标题甲 + 两行项目符号 + 标题乙 + 文本框）：

```
open POSIX file …            → document id 22462B5D-…（开成功）
count of slides              → 2
object text of text items    → 标题甲 / 第一行\n第二行 / 标题乙 / 文本框内容
export … as slide images     → tmp/kn-export/kn-export.001.png、.002.png（PNG 里中文字形肉眼可见已渲染）
export … as Microsoft PowerPoint → tmp/kn-export.pptx（26 909 字节；python-pptx 读回 2 页、形状名 Title 1 / TextBox 2、文字一致）
export … as PDF              → tmp/kn-export.pdf（13 915 字节）
close … saving no            → 无输出（不写回用户文件）
```

PDF 内部结构（同一份 `tmp/kn-export.pdf`，纯 stdlib `zlib` + 正则，不依赖外部二进制）：

```
%PDF-1.3  streams=12
/BaseFont  = {AAAAAC+ArialMT, AAAAAB+PingFangSC-Regular}   ← 子集字体（前缀 AAAAAB+）
/Encoding  = {MacRomanEncoding, Identity-H}
/ToUnicode = 1 个，/FontFile = 2 个                        ← 至少一个字体没有反查表
第一页内容流：q 1 0 0 1 294 457.735 cm BT 44 0 0 44 0 0 Tm /C1 1 Tf <0ebd1dcc12a8> Tj ET
带标记内容：/Span <</ActualText (…)>> BDC … <04de> Tj … EMC
```

> 结论见 P11 与 D8：**PDF 正文是 CID 十六进制，不是字符**，所以"改正文"要先解决反查与子集字体合并。

探针产物全部在 `tmp/`（`pptx-kn-check.pptx`、`kn-export/`、`kn-export.pptx`、`kn-export.pdf`、`pptx-pip.log`），
按 §6.5/§7.3 由本次执行者清理，不入库。
