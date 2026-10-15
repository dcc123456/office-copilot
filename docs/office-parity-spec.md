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

**范围界定（dcc 2026-10-14 裁决：「不许收窄，全部要能编辑」）**——所以下面这张表**不是取舍表，是顺序表**：
每一"波"只表示**先做谁**，不表示"后波可以不做"。任何一项在轮到它之前都处于"原样带过"的临时状态
（这是不丢数据的唯一安全态），轮到它就建模、就出 op、就开放编辑。

| 波 | 先做什么（**没有"不做"档**） | 内容 |
| --- | --- | --- |
| **W1** | 主干跑通 | pptx：OPC 读 + 保真回写、母版/版式/占位符继承链、文本与段落、形状与填充、图片、表格、撤销与 modlog、AI 提案面。pdf：读、结构审计、**增量更新式正文改写**（同框同字体同长度的那一档，见 D8） |
| **W2** | 数据部件 | pptx：图表（含内嵌 `.xlsx` 的 cache 一致性，D9）、连接线与成组/层级、节与页眉页脚页码、备注页、主题编辑。pdf：宽度感知的框重写 + 局部回流（要重排子集字体） |
| **W3** | 协作与呈现 | pptx：切换效果、**动画 timing 树**、批注（legacy + 现代 `p18c` 线程）。pdf：表单控件、注释层、附件 |
| **W4** | 长尾建模 | pptx：SmartArt(`dgm`)、OLE、3D `.glb`、嵌入字体、母版可视化重建器、`custGeom` 路径编辑。pdf：扫描件（要 OCR 才能编辑文字，OCR 本身作为功能立项）、加密文件、数字签名场景下的编辑策略 |

**"本机没有验收通道"的项 ≠ 不做的项**，只是这些条目的验收只能停在"文件合法 + 结构正确 + 独立读取器读回一致"，
**画面效果标 `[!] BLOCKED`**：动画/切换的播放效果、Morph、3D 渲染、字体观感、PowerPoint 独有的 `p:extLst` 扩展行为
（本机没装 Microsoft Office；WPS 无可视通道）。这条如实写在这里，不是收窄，是验收判据的边界。

**完成判定（总目标）**：W1–W4 全部建模完成，G1/G2/G3 各自的 §六 验收表全绿，且 `README.md` 的"项目是什么"一句话里四类格式都成立。

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
| **P13** | **本机实测：Keynote 的 PNG 导出连图表与表格都能渲染**（附录 A2） | — | python-pptx 造的"1 页 + 柱状图(2 系列里的 1 条, Q1=1/Q2=2) + 2×2 表格"存成 `.pptx` → Keynote 打开（`count of slides`=1、文字"乙"读回）→ `export … as slide images` 出的 PNG **肉眼可见：柱子画出来了、坐标轴 0/1/2 与 Q1/Q2 在、表格首格 "x" 与表头底色在**。含义很大：**pptx 线是本仓第一条能拿到"图表画面"证据的线**（xlsx 线到死都是 BLOCKED，因为 Numbers 不画我们的图表部件）。三条使用约束（实测踩到的）：① `count of slides` 要先 `set d to front document` 再问，直接 `count of slides of front document` 在部分时机报 `-1708`；② `open` 会偶发 AppleEvent 超时 `-1712`（同一批夹具里 3 页带备注那份中过一次，拆开造都能开）→ 判据必须写"重试 + 如实报告"，不能写"必能开"；③ 每次问完 `close … saving no`，别把 Keynote 的重新解释写回用户文件。 |

---

## 三、架构决策

### D1 — 内核自研，不引入外部运行时；`Univer` 只作分层与 AI 包装的参考

**问题**：dcc 点名"去看 Univer"，隐含选项是"拿它当内核省一半工"。
**核查结论**（P1）：Univer 的 OSS 部分**没有 OOXML 读写**，import/export 在闭源 Pro 层；把它当内核意味着
我们的核心交付物（"能被 Excel/PowerPoint 正常打开的文件"）依赖一个买来的黑盒，且 OSS 分支里根本没有那条路。
**决策**：pptx/pdf 的保真引擎**自研**，复用本仓已经验证过的路线（原始字节回写 + 未建模部件原样带过 + op 自带逆操作）。
从 Univer 只借三件：包分层命名、`action-recorder` 那种"命令流水即审计"的思路、以及它把 AI 能力包成 SDK skills 的形态。
**代价（如实写）**：pptx 的部件树比 xlsx 大（一个 WPS 默认模板 35 个 part，最富的那个 125 个 part 含 65 个 `tags`），
自研的工作量比 xlsx 线大，所以靠 §四 的波次（W1–W4）控制**先后顺序**——**不靠砍范围**（dcc 裁决：不许收窄）。

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
**已落（几何那一档，2026-10-15）**：`inherit.ts` 给出 `effectiveBoxPx` + `boxPxSources` 两个字段（有效值 + 来源层，逐轴回退），
视图把继承来的四个数念出来但**不给输入框**；颜色（`schemeClr`→theme）已解成 hex（`theme.ts`，`8023501`）；
字号那一档也已沿六层链算出（`text-styles.ts`，`ab559e9`：`effectiveSizePt` + `effectiveSizeSource`，`sizePt` 仍只记显式值），
D3 那三档至此全落，且三档都是"只读、绝不烤回部件"。
实测与偏离 python-pptx 的两处见阶段 2 的"D3 几何继承链已落"。

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
**新 op 种类（W1 先落这些）**：`setText`、`setParagraphFormat`、`setRunFormat`、`setShapeProperty`(xfrm/fill/stroke)、
`insertSlide`/`deleteSlide`、`setSection`、`setNotes`。**还没轮到建模的部件不产生 op**（只原样带过；
"判定不了角色就不出 op"同一条），但**它们是 W3/W4 的待办，不是终态豁免**。

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
**已落（2026-10-14）**：`EditorMode` 就在 `App.tsx` 里（`"doc" | "sheet" | "slide" | "pdf"`，pdf 是这条决策留的扩展点，
现在还没有分支会赋到它），九处 `!isSheet` 门控与 `isSheet || isDeck` 都改成按 `mode` 判定；
视图落在 `apps/desktop/src/slides/**`（不是 `src/slide/**`）、出入口落在既有的 `electron/pptxHandlers.ts`（不新建 `slideHandlers.ts`，§3.4 扩展现有模块）。

### D8 — PDF 的"真·正文改写"：走**增量更新**，并且先修我们自产 PDF 已经坏掉的中文子集

dcc 选的是最贵的一档（真·正文改写，不是标注/表单层），并且裁决"不许收窄"。所以这里不写"做不到就不做"，
只写**按什么顺序做到**，以及每一档的硬判据。

**先记一条本日实测出来的既有缺陷（比 pptx 线更紧急）**：我们自己的 `printToPDF` 产物里字体名是
`/SourceHanSansCN-Regular-6087`，**违反 ECMA-376/ISO 32000 的子集命名规则**（`9.6.4` 要求 `AAAAAA+Name`：六个大写字母 + `+`），
实测 FreeType/MuPDF 直接拒绝该字体程序（`unknown file format`），Apple 的渲染器把那行中文画成了 `! " #$%&' &(` 原始字形。
**也就是说：本仓现在导出的 PDF 里中文本来就是坏的**，只是我们没有视觉通道所以一直没发现（QuickLook / MuPDF 一交叉就露）。
→ 立为 **W1 前置修复项**（验收 A10），排在 pdf 编辑之前。

**机制（ISO 32000-1 一手条款 + 本机实测）**：文字在内容流里（`9.4`：定位 `Td/TD/Tm/T*`、显示 `Tj/TJ/'/"`、
状态 `Tf/Tc/Tw/Tz/TL/Tr`），**PDF 里没有"回流"这回事**——几何全靠 `cm`/`Tm` 矩阵 + 每字形推进宽度（简单字体 `/Widths`、
CID 字体 `/W`，单位 1/1000 em）。编码分 `9.6` 简单字体（1 字节 + `/Differences`）与 `9.7` 复合字体：
Type 0 + `Identity-H` 时"要显示的串是成对字节表示的 CID"，所以**新字符能不能写出来取决于子集里有没有那个 CID**；
`9.10.2` 明确"没有办法确定字符码代表什么"，`9.10.3` 的 `/ToUnicode` 是可选的且经常缺或错（P11 实测：2 个嵌入字体只有 1 个）。
文件侧：`7.5.4` xref、`7.5.7` 对象流、`7.5.8` xref 流、**`7.5.6` 增量更新——"改动必须追加到文件末尾，原内容保持不动"**。

**增量更新是本仓 pdf 线的保真答案，且已在本机跑通**（纯 JS、MIT 的 `@cantoo/pdf-lib` 2.11.1）：
`PDFDocument.load(bytes,{forIncrementalUpdate:true})` → `takeSnapshot()` → 打补丁 → **`snapshot.markRefForSave(ref)`**
（**坑：只 `context.assign` 会被静默丢掉，必须显式标记**）→ `saveIncremental(snap)` 返回一段 delta 拼到文件尾。
真 12 页文件实测：**前 163 205 字节逐字不变 + 追加 20 160 字节**，`/Prev 162756`，24 个流被重写；
MuPDF、Apple PDFKit（`multI=12 Multi=0`）、QuickLook 三方渲染一致。对照：pdf-lib 的**全量** `save()` 会把 163 205 字节重写成 790 601 字节——那就是"改一个字动整份文件"，不可接受。

**W1 落地后的两条实测补充（2026-10-15，`packages/pdf-io` 13 条 + 两份真文件）**：

1. **`/FlateDecode` 是 zlib 框，不是 raw deflate**（`ISO 32000-1 7.4.4` + 本机实测）。第一版按 raw deflate 解，
   两份真实文件都读出 **0 条 run**，而未压缩夹具那 11 条测试照样全绿——只有压缩夹具能抓住这种"对真实文件全瞎"的缺陷，
   所以 `flateWinAnsiFixturePdf()`（literal + hex 两条串都在压缩流里）已固化为回归网。
2. **真实中文产物是一条 run 一个字**：WPS 那份简历 6 页 5196 条 run，**全部**长度为 1（每字形一次 `Tj`，为了字距调整），
   没有一条 ≥3 字。所以 W1 那档"整条同长度替换"在真实中文文件上只能**重排同一 `Tj` 内的几个字形**
   （`tmp/ps-cols-OLD.pdf` 实测：把"段：这是"改成"是这：段"，前缀 33 706 字节逐字不变 + delta 1 657 字节，
   Apple PDFKit 读回新文字、QuickLook 出的 PNG 肉眼核对只有那四个字变了、第二栏一字未动）。
   → **W2 的前置不是字体子集化，而是把逐字形的 run 拼成逻辑文本行**（按 `Tm`/`Td` 与字体缩放判行），
   否则任何"改一句话"的需求都会在 W1 这一层就退化成"改一个字"。

**A10 那条既有缺陷已于 2026-10-15 现跑判定为不再适用**：经与 `printHandlers.ts` 相同的 `printToPDF` 通路重产一份中文 PDF，
读回生效字体名是合规的 `/AAAAAA+PingFangSC-Regular`，`-6087` 那种违规名在当前 Chromium 通路复现不出（见 §六 A10 与 §八 那行）。
同一次实跑翻出的是**另一条**旁路保真缺陷（PingFang 的 ToUnicode 把长/高映成康熙部首码位，属上游字体数据），不是子集命名。

**分档（都要做，只是排队）**：
- **W1 · 同框同字体同长度替换**：按 hex/字面量两种串形态定位，**必须递归进 Form XObject**（我们的表格里就在那儿），
  改完走增量更新。判据：前缀字节逐字不变、体积只增 delta、MuPDF + PDFKit + pdf.js 文本层三方读回新文字、QuickLook 出 PNG 肉眼核对。
- **W1.5 · 整行等长替换（W2 的前置链，2026-10-15 已落）**：真实中文产物**一条 run 只有一个字**，所以"改一句话"在 W1 那一档要发几十条。
  前置链三刀都在**等长语义内**，不是 W2 本体：第一刀让内容流跟住每条 run 的原点与生效字号，第二刀按坐标层把逐字形的 run
  拼成**逻辑文本行**（`pdfTextLinesOf`），第三刀让模型只报一个行号、由 `pdfLineEditsOf` 按该行每条 run 原有的字数切片拆回逐 run 的
  等长替换。判据：一次提交拆出的条数 = 该行的 run 条数、行外一条 run 不写、落盘与撤销都仍是前缀逐字不变 + 只追加 delta（真文件已核对）。
- **W2 · 宽度感知的框重写 + 局部回流**：用 `/W` 推进宽度算重排，需要**重新子集化/合并字体**
  （候选：MIT 的 `@pdf-lib/fontkit`，或 fontTools 级的纯 JS CFF/TTF 子集器）。判据：新字形真的进了子集
  （独立交叉：PyMuPDF `Page.get_texttrace()` 的 CID 清单 + PDFKit 抽取）。
- **W3 · 表单控件 / 注释层 / 附件**、**W4 · 整页重排、扫描件（OCR 作为独立功能立项）、加密与数字签名文件的编辑策略**：
  排队但**不砍**。签名文件要单独设计（任何全量重写都会让签名失效，只能"另存为新文件 + 明示原签名已失效"）。
- **通路否决**：`pdf2docx` 名义 MIT 但依赖 PyMuPDF（AGPL-3.0-or-later）→ **许可污染，不走**；
  MuPDF / Ghostscript / Poppler / pdftk-java / iText 全部 AGPL/GPL，**只能当本机外部取证工具，不得成为依赖**；
  `qpdf-wasm@0.1.0`（Apache-2.0）能在 Electron 渲染进程里跑但**跑不了裸 Node 24**（要 `self`/`location`）→ 结构审计走渲染进程窗口或本机 CLI。
- **本机环境补充实测**：`pip3 install --target` 能拉 arm64 轮子（pikepdf 10.17 / PyMuPDF 1.28 可装，仅作交叉读取器），
  `brew` 存在但 `mutool`/`qpdf`/`pdftotext`/`gs`/`pdftk`/`rustc` **当前都没装**；**QuickLook 能出 PNG 且探针看得见** →
  pdf 线**有**视觉通道（与 WPS 不同）。

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

**已落（2026-10-14，`auditPptxPackage` + `oc ppt check <file>`）**，与上面的原决策有一处**有意的偏差**：
- 与格式无关的那五条（内容类型覆盖、Override 指向存在、关系 Target 可解析、关系 Id 不重复、`r:id` 落点）不再各写一份，
  而是从 `sheet-io/compat-audit.ts` **上抽到 `@office-copilot/opc` 的 `auditOpcPackage`**（`AGENTS.md` §3.2：`.xlsx`/`.pptx` 第二次用到同一条判定就必须抽公共层），
  `auditXlsxPackage` 只留 `<worksheet>` 顺序与 `<c:chart>` 两条专属判定。sheet-io 原有 245 条单测一条未改、一条未减且全绿，是这次抽取行为等价的证据。
- `.pptx` 专属两条落成了**幻灯片必选元素**而非 `spTree` 的**子序**：`<p:sld>` 里必须有 `cSld > spTree`（cSld 是 Part 1 §19.3.1.16）、
  幻灯片必须有一条指向 slideLayout 的内部关系。`spTree` 子序留作后续：我们自己的写入从不新建形状，
  只有"增删形状"那批结构 op 才可能写错顺序，等那批 op 落地时一起补，避免现在写一条永远命中不到的判定。
- **这两条的一手依据按 §6.2 复核过一遍**（初稿凭记忆写的 §19.3.1.9/§19.3.1.10/§19.3.1.2 是错的，已改）：Learn 的 `Slide` 类页给出
  `<p:sld>` 子元素的条款号（cSld §19.3.1.16、clrMapOvr §19.3.1.7、extLst §19.3.1.20、timing §19.3.1.48、transition §19.3.1.50），
  而"最小一张幻灯片"要 `CommonSlideData` 里套 `ShapeTree` 是 "Working with presentation slides" 的原话
  （"All of these class instances ... are required to create the minimum number of XML elements necessary to represent a new slide"）。
  **slideLayout 那条在 ISO 正文里没有对应的 MUST**，同一页的原话是"Each instance of the Slide class must also be associated with instances of
  the SlideLayout and SlideMaster classes"——属 PowerPoint 侧约束，所以代码注释与 CLI 文案都按"办公软件会修复或拒绝打开"来说，不冒充条款号。
  母版（slideMaster）由版式自己的关系链上去，没有单列判定。
- 问题码文案在 CLI 装配层（§5.7），`Record<PptxAuditIssue["code"], string>` 让"新增码忘了配文案"在编译期就报错。
- **第三条专属判定（`feat(pptx-io)` 0eebd25，阶段 4 那条缺陷的后续）**：几何元素 `<a:off>`/`<a:ext>`（组合形状那对是 `<a:chOff>`/`<a:chExt>`）的
  x、y 与 cx、cy 四条在 Part 1 §20.1.1.15（`CT_Point2D`）与 §20.1.1.14（`CT_PositiveSize2D`）里全是 required attribute，只写一半 PowerPoint 就判损坏。
  写侧那条守卫（`edit.ts` 的 `missingPairValues`）只能挡住"我们自己写的"，旧产物或别的工具写进来的这种 XML 之前审计回 **0 条**，等于没机检——现在审计与写侧各守一头，
  判定按**本地名**比对，所以写成 `<off y="…"/>` 这种换前缀/默认命名空间的合法写法也抓得到。同一刀还管 `<a:xfrm>` 少整块：`CT_Transform2D` 把 `off|chOff` 与 `ext|chExt` 两个 choice 各列为必选子元素，
  只写 off 不写 ext 的 `<a:xfrm>` 属性一个个都齐、照样是损坏（正是 a3bd03c 那条缺陷当时写出来的形状）。
**本机实测**：python-pptx 1.0.2 生成的真实 40 部件 deck → `oc ppt check` 0 条、退出码 0；`setShapeRunText` 改标题后的写回同样 0 条；
把 slide1 的关系表从包里删掉的那份 → 报 1 条"幻灯片没有指向版式（slideLayout）的关系"、退出码 1（闸门真的会拦）。
**第三条那条的实测**（`tmp/audit-check/inject.mjs`，真实 deck 注入半个 `<a:off y="571500"/>`）：`oc ppt check` 报
"几何元素只写了一半必填属性（x 与 y、cx 与 cy 得成对），打开时报损坏：a:off 缺 x"、退出码 1；同一份文件 python-pptx 抛
`InvalidXmlError: required 'x' attribute not present on element {…}off`——独立读取器与我们的判定同向；未注入的对照组 0 条、退出码 0（不误报）。
**对真实文件的误报面**（新加判定必须核一遍）：python-pptx 默认模板、我们的两份探针产物、WPS 自带的两份新演示模板（`new_wpp.pptx` 125 部件那份与 `mui/zh_HK/templates/newfile.pptx`）
共 6 份 `.pptx` 逐份 `oc ppt check` 全部 0 条、退出码 0。

---

## 四、阶段目标（每段一句完成判定）

### 阶段 0 · 取证与夹具（不写产品代码）— `[x] 已完成 2026-10-14（`13e62e2` 之后的实测，数字见附录 A2）`
把 WPS 自带模板（`/Applications/wpsoffice.app/Contents/Resources/office6/addons/kfindermenu/res/templates/new_wpp.pptx`
125 part 含 65 个 `tags`；`…/mui/zh_HK/templates/newfile.pptx` 35 part）与 python-pptx 生成的夹具收进
`packages/pptx-io/scripts/` 的生成脚本，**并确认本机 `tmp/pylibs` 里有 python-pptx 1.0.2**（已 vendor）。
**完成判定**：三类夹具（空模板 / 富模板 / 带图表备注）都能被 python-pptx 与 Keynote 各自读通，Keynote 能导出 PNG 且肉眼可见文字。
→ **过了，而且比预期强**：PNG 里**柱状图与表格都真的渲染出来了**（不只是文字），见 P13。
唯一没复现成功的那份（3 页 + 备注 + 图表 + 表格）是 AppleEvent 超时 `-1712`，同类夹具分开造都能开——
所以判据要写成"通道可用但不稳，超时须重试并如实报告"，不许当成"能开"的普遍结论。

### 阶段 1 · 抽 `packages/opc`，两线不改行为 — `[x] 已完成 2026-10-14（`7178c65` 读层 + `740c6e4` 写侧登记）`
按 D2 把 OPC 提到公共包，`sheet-io`/`docx-io` 改依赖它。
**完成判定**：`pnpm test` 逐包通过数与抽取前**完全一致**（不许靠改测试过），xlsx/docx 产物字节级往返夹具全绿。
**实测**：抽前抽后逐包通过数一致（docx-io 296 / sheet-io 245 / desktop 457 …），仅新增 `packages/opc` 12 条；未改任何既有测试。serialize 顺带修回注释与 CDATA 的往返（原 docx-io 版会静默丢掉）。

### 阶段 2 · `packages/pptx-io`：读 + 保真回写 — `[~] 读模型 + 扁平 DFS 形状列表（组内子形状带 `groupShapeId`、连接线 `p:cxnSp` 进列表）+ 保真回写 + `oc ppt check` 审计 + Keynote 复开与 PNG + D3 三档继承（几何/颜色/字号）+ notes 读侧 + 备注写侧第一档（只改已存在的备注部件，含 AI 提案面、落地通道与面板的显示/编辑入口）+ 表格那一格（读侧逐格给出「每段几截 run」、合并起点占几列、被并进来的续格标成只读；写侧第九个 op 只换那一格那一段那一截，多 run 与续格各给一句拒绝理由，每行格数恒等于 `a:gridCol` 数）已落；新建 notesSlide 那一档仍缺，组框↔子坐标系那对换算（`a:off`/`a:ext` 与 `a:chOff`/`a:chExt` 要连着组的缩放与旋转一起算）也仍缺——所以组里的形状在视图与画像里一律不给页面位置数字`
部件树（presentation/master/layout/slide/notes/theme）、D3 的 `effective`+`src` 继承解析、D4 换算、未建模部件原样带过。
**完成判定**：三份夹具"打开→原样存回"后 `unzip -l` 条目集与**未改动部件的字节**逐字一致，`oc ppt check` 0 条，
python-pptx 读回形状名/占位符类型/文字全等，Keynote 打开后 `count of slides` 与文字一致。
**实测（读+保真这条已达标）**：本机 python-pptx 1.0.2 交叉读一份真实 38 部件 `.pptx` 往返产物 → 形状文字 `['标题甲','副标题乙','正文丙']` 全等、38 个部件字节级 0 变化（含二进制）；读模型对占位符 `hasOwnGeometry=false`、文本框 `true`（D3 未烘焙）。

**Keynote 复开证据（2026-10-14 实测，P10 那条通道第一次用在"我们自己的写回产物"上）**：拿 `setShapeRunText` 改过标题的那份 `.pptx`
（`tmp/audit_real_edited.pptx`，29 330 字节）→ `tell application "Keynote" to open POSIX file …` 直接开成功 →
`count of slides of front document` = **2** → `object text of text item j of slide i` 读回
`1.1=改过的标题 / 1.2=正文一行 / 2.1=第二页标题 / 2.2=正文一行`（**改的那格变了，没改的那页原样**）→
`export … as slide images with properties {image format:PNG}` 出两页 PNG，**肉眼看过**：第一页居中黑字"改过的标题"、
下方灰色"正文一行"，第二页"第二页标题"，占位符的继承样式（居中/灰色副标题）在 Keynote 里真的生效了 → `close … saving no`。
这条同时是 §7.4⑥ 的"第三方办公软件真读过"与 §7.2 的"画面看得到"两层证据，pptx 线到这里不再有目视 BLOCKED。
**已落齐**：D3/D4 同一条红线下的三档全落——几何逐轴继承见下段，颜色（`packages/pptx-io/src/theme.ts`：`a:schemeClr` 沿母版 `p:clrMap` 与主题 `a:clrScheme` 解成 `#rrggbb` 交 UI，带 `lumMod`/`tint` 这类变换子元素与 `phClr` 一律不解、回 `undefined`，本机没有能核对 PowerPoint 画面的通道，算错的颜色比没有颜色更坏），字号（`packages/pptx-io/src/text-styles.ts`：run 没写 `sz` 时沿 段落 `a:pPr/a:defRPr` → 形状 `a:lstStyle` → 版式那一格 → 母版那一格 → 母版 `p:txStyles` 逐层问，交 `effectiveSizePt` + `effectiveSizeSource`，`sizePt` 仍只表示"自己写了 `sz`"，链上查不到就不给数、不拿 18pt 这类常见默认凑）。notes 部件的解析已落（`162cb38`，只读侧：`Slide.notes` 逐段文字 + 画像里的 `notes`/`notesPartPath`，只取 `p:ph type="body"` 那一格，页码 `<a:fld>` 与缩略图占位符不算备注；写侧那条备注 op 的第一档已落（`fccd4f9`，见下面"备注写侧的前置取证"那段的收尾），AI 那一头也已接上（`ad4d995`：`ppt_submit_proposal` 第八类改动 `set_notes`，页号→`notes.partPath` 由 `slideAgent` 翻，压平与否跟引擎共用 `notesFlattenNote`），新建 `notesSlide` 那一档仍缺）。**面板那一格已落**（`2dc83f4`：备注直接显示在幻灯片面板里、可整篇换掉，见下面「面板的备注那一格」那段）。审计与 CLI 那条已落，见 D10 的"已落"段。

**备注写侧的前置取证（2026-10-15，`zipfile` 直读 WPS 自带模板与 python-pptx 产物，不拿我们的读侧自证）**：结论是**写侧几乎一定要新建部件**，而不是"改已有备注那一格"。
本机 WPS 的 106 份 `.pptx` 模板里，抽样能打开的 8 份真文件（`normal.pptx`、`normal_mac.pptx`、`normal_gray_mac.pptx`、`new_wpp.pptx`、`newfile*.pptx`、`default.pptx`）**notesSlide 部件数全为 0**——模板只给版式与母版，不给备注页；
其中**有 notesMaster 的只有 4 份**（`normal_mac` / `normal_gray_mac` / `normal_black_mac` / `new_wpp`），另外 4 份连备注母版都没有。python-pptx 侧同样：默认模板产的 deck `notesMaster=0`，从 WPS `new_wpp.pptx` 生成的那份 `notesMaster=1`、`notesSlide=0`。
所以那条 op 的**第一档只能是"这页已有备注页时改它"**（改动只在已建模的 `p:txBody` 里，走现成的 `toOp` 前像通道）；**第二档是新建 `notesSlide`**——按 §2.10 要三处一起登记（部件本体 + 幻灯片侧 `<Relationship>` + `[Content_Types].xml` 的 Override，另加备注页自己那两份 `.rels` 指向幻灯片与 notesMaster），而**指向 notesMaster 这一条在 4/8 的真文件里根本无从指向**。
据此定的最小范围：**没有 notesMaster 就整条拒绝并给一句中文理由**，不许自己造一份备注母版——那等于凭空发明版式，与 D3"读不懂宁可不写"同一条线。
**第一档已落（2026-10-15，`fccd4f9`，`setSlideNotes`）**：只改这一页**已经存在**的备注部件，op 仍是部件粒度的前/后像，坐标就用读模型的 `notes.partPath`，落地与撤销完全复用现成那条通道（§3.4，没发明第二种 op 形状）。三条红线落在代码里：现有备注里有**多截 run** 或**带 `<a:rPr>` 的 run** 时整条拒绝（整篇覆盖会拍平，§2.3）；新段从旧第一段带过 `<a:pPr>` 与 `<a:endParaRPr>`（只换文字，不动这一页已有的版式意图，§2.7）；读侧那句"找正文那一格"抽成 `findNotesBodyShape` 给读写共用（§3.2）。**第二档（新建 `notesSlide`）没做**；AI 那一头已接上（`ad4d995`：`pptx-tools` 多第八类改动 `set_notes`，`slideAgent` 把页号翻成读模型那份 `notes.partPath`，"会不会拍平"与引擎共用 `notesFlattenNote` 那一份判定），面板的显示与编辑入口也已接上（`2dc83f4`，见下面「面板的备注那一格」那段）；按上面那份取证，真模板里 8/8 都没有备注页，所以"这页本来没备注"仍是眼下最常见的情形，那一档（新建 `notesSlide`）一定要等 §2.10 那三处登记落齐才走得通。

**面板的备注那一格已落（2026-10-15，`2dc83f4`，`SlideDeck` + `slideDeckModel`）**：上一条留的那格空（"引擎与 AI 都改得了备注，用户手工改不了"）补的是装配面，不是第三条编辑链——判定仍旧只有一份：`notesSaveBlock(slide, draftText)` 把"这一页没有备注页""备注没有改动。"与 `notesFlattenNote` 那三类理由拼成按钮的 `title` 与格子下面那句说明（§3.2，`slideAgent` 用的是同一个 `notesFlattenNote`），落地走 `SlideDeck` 现成的 `commitNow`（与"加文本框/删形状/调层级"那一条立即落盘的通道同一个口），所以存前审计（`auditPptxPackage`）、modlog 与「撤销上一次保存」一个字都没改（§2.4/2.5）。**两处刻意的取舍**：① 备注**不进**「N 处待保存」计数——那一栏记的是 run 级草稿（`drafts`/`pendingEdits`）的账，备注是整篇换、当场落盘的另一类粒度，混进同一个计数会让用户以为顶栏那个「保存」也能存备注；② "会拍平"那句理由**只在用户已经动手打字之后**才显示（没改过时按钮只以"备注没有改动。"禁用），免得三页面板一进来有两页顶着一句警告。**真 app 已实跑**（`tmp/probe/slide-notes-panel-ui.mjs`，不起端点、不发 Key，只验装配面；夹具沿用 `tmp/notes-ai/ui/probe.pptx` 那份三页 49 部件）：细节与逐部件字节证据见 §八 那行。**这条 op 至此三面齐**（引擎 / AI 提案面 / 用户手工入口）；仍缺的只剩新建 `notesSlide` 那一档（真模板 8/8 都没有备注页，"这一页没有备注页"因此是面板上最常出现的那一句）与阶段 5 那格真端点行为探针。

**新建 `notesSlide` 那一档的前置取证（2026-10-15，已取证未实现）**：这一档缺的不是 OOXML 知识而是**一份可信的参照产物**，所以按 §6.2 让 python-pptx 1.0.2 在**真模板**上现建一个备注页再拆包看它写了什么（`tmp/notes-new/`，模板 `/Applications/wpsoffice.app/.../kfindermenu/res/templates/new_wpp.pptx`，它带 1 条 notesMaster 关系）。**它一共动四处**（新增 2 个部件 + 改 2 个原有部件，`消失部件: []`）：① `ppt/notesSlides/notesSlide1.xml`——`<p:notes>` 下 `p:cSld/p:spTree` 三格 `p:sp`：`ph type="sldImg" idx="2"`（只 `p:spPr`，无 `p:txBody`）、`ph type="body" idx="3" sz="quarter"`（**正文那一格，`p:txBody` 就 `<a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>…`，没有 `<a:pPr>`、没有 `lang`**）、`ph type="sldNum" idx="5" sz="quarter"`（同样无 `p:txBody`），树根带 `<p:nvGrpSpPr>` 的 `id="1" name=""` 与一个全零 `<a:xfrm>`，末尾 `<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>`；② `ppt/notesSlides/_rels/notesSlide1.xml.rels` 两条，**`rId1` 指向 `../notesMasters/notesMaster1.xml`、`rId2` 指向 `../slides/slideN.xml`**（顺序是 notesMaster 在前）；③ `ppt/slides/_rels/slideN.xml.rels` 追加一条 `Type=".../notesSlide"`；④ `[Content_Types].xml` 加 `/ppt/notesSlides/notesSlide1.xml` 的 Override（notesMaster 那条模板里本来就有）。**这三格占位符恰好印证读侧那条判据**：只有 `type="body"` 那一格是备注正文，`sldImg`/`sldNum` 念进来就是编内容。**交叉验证没拿我们自证**：这份产物过我们的 `oc ppt check` **EXIT=0**，`oc ppt info` 如实念出「备注（ppt/notesSlides/notesSlide1.xml）：取证第一句」——读侧已经能吃自己写侧将要产的那种形状，这一档不需要动读模型。**顺手再一次坐实 D1 那条反面**：python-pptx 保存时把整包重序列化，`[Content_Types].xml` 里模板自带的 `Default Extension="JPG"` 被它丢掉、`docProps/core.xml` 与全部关系表都重排过，这正是"重新生成型"模型的代价，所以这一档我们自己写时**只动那四处**。**实现要跨的一处门槛（留给下一批，别顺手做完）**：`SlideEditOp` 现在是 `{partPath, beforeXml, afterXml}` 且 `beforeXml` 是必填字符串，"新建部件"要么让它能表达"部件本来不存在"、要么另立一种 op，而撤销就得从"写回前像"变成"删部件 + 撤那三处登记"（§2.5 内容寻址在这里意味着撤销必须能识别"这份后像仍是我建的那种形状"而不是按位置猜），落地/审计/modlog 那条链（`pptxHandlers.ts` → `apply`/`revert`）要一起跟着改，所以它是一个**独立批次**，不是把 `setSlideNotes` 的守卫删掉就完事。**判据按已定的最小范围**：模板没有 notesMaster 就整条拒绝并给一句中文理由（取证过 4/8 真模板根本没有备注母版），绝不自己造一份备注母版（D3 那条"读不懂宁可不写"同线）。
**D3 几何继承链已落（2026-10-15，`packages/pptx-io/src/inherit.ts`）**：一手依据是本机读到的 python-pptx 1.0.2
`shapes/placeholder.py` —— `_InheritsDimensions._effective_value` 是**逐轴**回退（自己有就用，没有才问上一层），
`_BaseSlidePlaceholder._base_placeholder` 用 `layout.placeholders.get(idx=)` 按 **idx** 匹配幻灯片→版式，
`LayoutPlaceholder._base_placeholder` 再按**版式那一格自己的类型**映射到母版类型（`ctrTitle`→`title`，
`body`/`subTitle`/`obj`/`chart`/`tbl`/`pic`/`clipArt`/`dgm`/`media`→`body`，`title`/`dt`/`ftr`/`sldNum` 原样）。
两处偏离是有意的：python-pptx 遇到表里没有的类型直接 `KeyError`，我们按"原样类型去比对，查不到就停在上一层"；
母版那一格也没写时**不**往下用 `defaultTextStyle`/`p:sldSz` 猜。
读模型据此多给两个字段：`SlideShape.effectiveBoxPx`（有效值，逐轴 CSS px）与 `boxPxSources`（每一轴来自哪一层），
`boxPx` 与 `hasOwnGeometry` 这两个"自己写没写"的事实**保留不动**——写侧只认它们，有效值纯只读。
`read.ts` 里那个"四个都齐才给"的 `boxPx` 现在由 `readExplicitBoxAxes` + `completeBox` 派生（同一份 `<a:xfrm>` 轴读取只留一处，§3.2）。
**实测（三层链 + 真文件交叉）**：新夹具 `inheritanceDeckParts()`（幻灯片/版式/母版三层各写一格）上，
本机 python-pptx 独立读回 `914400/457200/7315200/1143000`（标题四轴全来自母版）、`952500/1143000/2857500/762000`（正文自己写的赢）、
`476250/476250/1905000/952500`（无 `<p:ph>` 的文本框根本不继承）——与我们 `effectiveBoxPx × 9525` **三格逐轴全等**；
一份真实 38 部件 `.pptx`（python-pptx 现生成）同样全等，只有 `top` 差 1 个亚像素（`274638 EMU` → 我们按 D4 四舍五入成 29px，是既定的显示单位损失，不是解析分歧）。
真 app 实跑（`tmp/probe/slide-inherit-ui.mjs`，`--user-data-dir=tmp/slide-inherit-ui/profile`，CDP 驱动 `dist`）：
标题那行**没有几何输入框**，只有一句"位置与尺寸继承自母版：左 96 · 上 48 · 宽 768 · 高 120 px（写死就切断版式联动，本视图只读，spec D3）"，
另两行各自是 `100/120/300/80`、`50/50/200/100` 四个输入框；改一次标题文字→「已保存，动了 1 张幻灯片」→
磁盘 `slide1.xml` 里标题那格**仍是 `<p:spPr/>`**（母版那个 `7315200` 没被烤进来，D3 验收），模型重读后那句继承说明原样还在。
**画面位置仍 BLOCKED**：Keynote 的 AppleScript 字典拿不到对象的 `position`/`size`（`-1700`/`-1728` 三次都撞在这），
所以"我们显示的那四个数就是 PowerPoint 画的位置"这一条只有 python-pptx 这层独立读取器证据；样式与文字的目视证据沿用上批那份 PNG。

### 阶段 3 · 幻灯片视图可用（人手动编辑）— `[~] 打开分派 + 列表式文字/字号/颜色/几何编辑 + 备注那一格（显示与整篇换掉）已落（改→保存→撤销，真 app 与 python-pptx / Keynote 双向实测）+ 组里的形状与连接线各占一行并按大类点名（「组合里的文本框」「连接线」），几何那一格只给一句为什么不给数；表格里那一格按网格摊开——单 run 那段给输入框，多 run 那格与合并续格只读并各带一句为什么；画布式拖拽仍缺（几何继承链解析已落到视图上，见阶段 2）`
D7 的模式枚举 + 幻灯片列表 + 画布编辑（选中形状、改文字、改字号颜色、拖拽尺寸）。
**完成判定**：起真实 app 走"改标题→存盘→Keynote 导出 PNG→肉眼比对文字与位置"，并把 PNG 里的实际文字写进收尾报告（§7.2）。
**已落（2026-10-14，`apps/desktop/electron/pptxHandlers.ts` + preload 四条通道）**：跨界只送 `PresentationModel` 与
**定位坐标**（部件路径 + 形状号 + 段落/run 序号），op 一律由主进程现做——渲染层拿不到幻灯片原文，也就拿不到撤销所需的逆操作，
所以成功结果里把 `ops` 送回，`ppt:revert` 原样送回即可回滚（§2.5 内容寻址）。落地流程与表格侧同源（`applyBatch` 一份，
正向与撤销共用，§3.2）：重读底包 → 逐条现做并应用（**后面的那条看得到前面已落地的结果**）→ 出字节 → 结构审计只拦
"本次新引入的"问题 → 留 `.bak` → `写 .tmp → rename`（§2.4）。4 条出入口单测：`.ppt` 不进闸门、
改文字只动 `ppt/slides/slide1.xml` 且磁盘重读是新文字、撤销回到"标题甲"、同批两条按序生效 +
定位失败时整批不写盘（字节比对一致）。两处格式无关的公共层这次一并上抽：
"只拦本次新引入的"那道判定（`newlyIntroducedIssues`）从 `sheetHandlers`/`pptxHandlers` 各一份合成
`@office-copilot/opc/audit.ts` 里的一份（3 条用例随之移进 `opc.test.ts`，原 `sheetHandlers.test.ts` 删除——
它对桌面的价值只有那 3 条）；夹具从 `pptx-io` 的审计单测抽成 `buildDeckBytes`/`minimalDeckParts`，
经 `@office-copilot/pptx-io/testing` 这个 `exports` 子路径共用一份，**不进 `index.ts`**——否则一份假幻灯片
会被主进程打包带进 `dist-electron/main.js`（改前 vite 确实报过这条，改后产物里已查不到 `test-deck`），
pptx-io 原 26 条未改未减。

**已落（2026-10-14 第二批：D7 的模式枚举 + 打开分派 + 列表式文字视图）**：
- **D7 落地**：`apps/desktop/src/App.tsx` 加 `export type EditorMode = "doc" | "sheet" | "slide" | "pdf"`，
  原来那 9 处 `!isSheet` 门控与 `isSheet || isDeck` 全部改成按模式判定（`mode` 一处算出来）。pdf 线进来时只多一个分支，
  不再叠 `!isSheet && !isDeck && !isPdf` 这种布尔组合。
- **打开分派**：`file:open` 对话框的 filter 与扩展名闸门带上 `pptx`（`electron/fileHandlers.ts`）、`main.ts` 的
  `OPENABLE_EXT` 认 `.pptx`（argv / 拖到 Dock / 访达"打开方式"那条路）、`App.tsx` 的 `openPath` 加 `.pptx` 分支
  （`readPptx` → 清掉表格状态 → `setDeck`），拖拽正则与两处提示文案同步。
- **视图**：`src/slides/SlideDeck.tsx` + `src/styles/slides.css`（只用 `tokens.css` 那一套，§5.4）。按页列形状
  （名字 + 占位符角色 / 大类），**单条 run 的段落给输入框，多条 run 的段落只读并写明理由**（§2.3）；
  保存走 `ppt:edit`、撤销把上一次那批 op 逆序送回 `ppt:revert`，都只经 preload 白名单（§5.2）。
- **落地范围抽成模型层**：`src/slides/slideDeckModel.ts`（`draftKeyOf` + `pendingTextEdits`）决定"哪些段落会被写进文件"，
  视图只显示条数和送批——与表格侧 `sheet/sheetGridModel.ts` 同一个分工（§3.3）。4 条单测钉住：只有真改过的单 run 段落成一条、
  多 run 段落一次都不交出、文字来自字段（0 条 run）的不交出、顺序跟着"页—形状—段"的显示顺序。
- **读模型补一条事实**：`SlideShape.paragraphRunCounts`（每段 `<a:r>` 条数）——视图判断"能不能整段覆盖"要靠它，
  不靠猜（pptx-io 26→27 条）。跨界载荷类型（`PptxEditRequest`/`PptxEditPayload`/`PptxRevertPayload`）收进
  `pptx-io/edit.ts` 一份，主进程、preload、渲染层三处共用（§3.2）。
- **真 app 实测**（`pnpm build` 后跑 `dist`，CDP 连上去看 DOM，§7.2）：argv 打开 → 视图给出"2 页 · 0 处待保存"与
  4 个形状（`Title 1 占位符 ctrTitle` / `Subtitle 2 占位符 subTitle` ×2）→ 改第一页标题 → "1 处待保存" → 点保存 →
  toast"已保存，动了 1 张幻灯片" → `oc ppt info` 读磁盘：第一页是"Keynote 要看到的第一页标题"、**第二页原样** →
  `.bak` 在 → `oc ppt check` 退出码 0（"没有发现会让办公软件判损坏的结构问题。"）→ 点"撤销上一次保存" →
  视图与磁盘都回到"第一页标题" → **Keynote 导出两页 PNG 肉眼看过**：第一页居中黑字"Keynote 要看到的第一页标题"、
  下方灰色"正文一行"，第二页"第二页标题"（文字与位置都对，§7.4⑥ 的第三方真读过 + §7.2 的画面证据一并拿到）。
- **§2.3 的目视证据**：`tmp/multi-run.pptx`（一段里两条不同格式 run）在视图里 **0 个输入框**，显示
  "这段有 2 截不同格式的文字，整段覆盖会拍平粗体与字号，本视图只读"。
- **三种模式回归**（同一份 `dist`，同一实例依次换文件）：`.docx` → 工具栏/大纲/状态栏/185 段都在；`.xlsx` → 文档面板全隐、
  表格面板在；`.pptx` → deck 视图在、文档面板全隐。证明这次换枚举没改坏前两条路。
**已落（2026-10-14 第三批：字号 / 颜色 / 形状几何的编辑入口 + D4 的换算收到一处）**：
- **D4 的换算单点变成结构**：新建 `packages/pptx-io/src/units.ts`（`emuToPx`/`pxToEmu`/`halfPointToPt`/`ptToHalfPoint` +
  `ShapeBoxPx`/`ShapeGeometryPatch` 两个盒子类型），`edit.ts` 里那条私有的 `EMU_PER_PX` 与 `ShapeBoxPx` 删掉、改为引用它，
  `read.ts` 也走同一份。**读侧从此交给视图的是 px 与 pt**（原来只给 EMU 的话，渲染层就得自己乘 9525 / ÷100，那是第二处换算）。
- **读模型再补两条事实**（视图不猜，D3）：`SlideShape.paragraphRunFormats`（每段每条 run 自己显式写了的 `sizePt`/`bold`/`hex`，
  没写的留 `undefined`＝继承，绝不填猜测值）与 `SlideShape.boxPx`（形状自己写的 `<a:xfrm>` 换成 px；没写就是 `undefined`，
  与 `hasOwnGeometry` 一道构成"能不能改几何"的判据）。
- **模型层扩到四类改动**：`slideDeckModel.ts` 的 `pendingTextEdits` → **`pendingEdits`**（返回 `PptxEditRequest[]`，
  几何先于该形状的各段），草稿值从"一串文字"变成 `SlideDraft { text?, sizePt?, hex?, box? }`，几何另用键 `partPath@shapeId`。
  四条守卫都在这一层：多 run 段落不交出（§2.3）、**与现场值相同的字段不交出**（空点"保存"不写盘）、
  字号只把改过的那一项塞进 `RunFontFormat`（不连带写 `bold`，§2.3）、`boxPx` 为 `undefined` 时几何一个字段都不写（D3）。
  单测随之 4 → 9 条。
- **视图**：每个形状一行四个数（左/上/宽/高，只有自带几何时才出现，否则显示一句"继承自版式，本视图不写死"）；
  每条单 run 段落在文字框右边加字号（`step=0.5`，PowerPoint 允许半磅）与颜色（`input[type=color]`，值一律 `#rrggbb`，§2.2），
  继承来的字号占位显示"继承"、颜色 `title` 提示"当前颜色继承自版式"。样式只加进 `src/styles/slides.css`（§5.4）。
- **单位/颜色的往返有机检**：`pptx-io` 新增 `units.test.ts` 2 条（9144000 EMU ↔ 960 px、负偏移、24.5pt ↔ sz 2450）；
  `pptx.test.ts` 读模型两条新断言（`boxPx` = 96/384/288/96、多 run 段的 `[{bold:true},{sizePt:18}]`）+
  一条"写进去再读回来"（同一形状连做 font/color/geometry 三条 op 后重新解析：`24.5pt`、`#3cb371`、四个 px 原样对上）。
  **这条测试顺手抓到一处真缺陷写法**：把三个 op 生产者写成 `[await A, await B, await C]` 再依次 apply，会因为三条后像都基于
  同一份原文而互相覆盖（只剩最后一条）——主进程 `editPptxDoc` 一直是"逐条现做逐条落地"，所以线上没这个洞，
  但探针/视图必须照同一个顺序（已在测试注释里写明）。pptx-io 27 → 30 条。
- **真 app 实测**（`pnpm build` 后跑 `dist`，CDP 连 DOM，§7.2）：夹具是 python-pptx 现生成的两页 deck
  （第 1 页：继承几何的标题占位符 + 一个自带 `<a:xfrm>` 的三 run 段落文本框；第 2 页：20pt、`#aabbcc` 的一行）。
  视图如实显示：第 1 页标题那格**没有**几何输入框、只有"继承"占位，那个三 run 段落 **0 个输入框**并写明"这段有 3 截不同格式的文字…本视图只读"；
  文本框四个数是 `96/384/288/96`（= 1″/4″/3″/1″ @96DPI，逐 px 对上）；第 2 页那行回显 `20` 与 `#aabbcc`。
  在第 2 页那个形状上改文字＋字号 32＋颜色 `#00cc00`＋左 150＋宽 300 → 顶栏"**4 处待保存**" → 点保存 →
  toast"已保存，动了 1 张幻灯片"、顶栏归零、四个输入框从磁盘重读回来。**python-pptx 独立读回**（不是我们的 parser 自证）：
  `left/top/width/height` = `150/192/300/96` px（**只有改过的左与宽变了，上与高一个 EMU 没动**）、`size=32.0pt`、
  `rgb=00CC00`、`bold=None`（没连带写粗体，继承链保住）、第 1 页三 run 段落仍是"正文丙/粗(粗体)/细(18pt)"原样。
  `oc ppt check` 退出码 0 → 点"撤销上一次保存" → 视图回到 `192/192/384/96`、`20`、`#aabbcc`、原文字；
  **逐部件字节比对**：40 个部件里内容不同的是 0 个（撤销后整包回到原样）。
- **发现的既有缺陷（不在本批修，另开）**：`.pptx`/`.xlsx`/`.docx` 的写回会给 zip **新增目录项**（`ppt/`、`ppt/slides/`），
  而 python-pptx 产的原文件里没有目录项。内容与结构审计都不受影响（办公软件照开，`oc ppt check` 0 条），
  但"部件清单逐条一致"这条保真标准不该靠它打折：根因是 `zip.file(path, content)` 会顺手合成父目录对象，
  三个 IO 包各写一次（`docx.ts:823`、`sheet-file.ts:414`、`pptx-file.ts:67`）→ 该把序列化收进 `@office-copilot/opc` 一份并带
  `createFolders: false`。
**仍缺**：画布式选中与拖拽尺寸（本批是列表里的数字输入框，不是 WYSIWYG 拖拽）。三档继承现在都落到了读模型：颜色解成 hex 后 `SlideDeck.tsx` 那格颜色输入框与 title 提示自动有基线色（一行 UI 代码没改）；字号那档另加了一句 hover 说明（`slideDeckModel.sizeFieldHint`：把 `effectiveSizePt` + 来源层念成人话），**输入框的 value 仍只放显式 `sizePt`**——填进继承值会让用户一个字没敲也落出一条 `sz`，那就违反 D3；notes 读侧同样已落。这四样（读不懂提示 / 「加文本框」禁用态 / 颜色框的 hex / 字号那句 hover）已在真 app 里逐条看到，见 §八 最后两行。

### 阶段 4 · 结构化 op + 撤销/modlog — `[~] run 级四类编辑 op + 叠放次序结构 op（含视图上那四个按钮）+ 增删形状那批结构 op（含视图上的「加文本框 / 删形状」）+ 内容寻址撤销 + modlog 落盘已落（真 app 实测跨进程撤销）+ 结构那三条一律在形状所在的那一层上动（组里的孩子挂在 `p:grpSp` 下，拿页 spTree 摘会静默摘不掉却报成功），叠放的两端判定只比同层邻居；第九个 op 是表格那一格的那条 run（`table_text`，坐标多带行与列，落地/撤销/modlog 全走现成那条通道，没有第二条写路径）；历史面板的幻灯片渲染与画布式选中仍缺`
D5 的 op 集合与内容寻址撤销，接同一份 modlog。
**完成判定**：每类 op 有逆操作单测；"改完→用户手工再改别处→撤销"只回到自己那一格，不牵连后面的手工编辑。
**实测（文字 + 字体 + 颜色 + 几何这四条已达标）**：`pptx-io/edit.ts` 的 `setShapeRunText` 只改目标 `<a:r>` 的 `<a:t>`、`setRunFont` 只改目标 run 自己的 `<a:rPr>` 的 `b`/`sz`（pt×100，字号唯一换算点落 IO 边，§2.1；rPr 缺失时新建插到 `<a:t>` 前）、`setRunColor` 只替换该 run 的填充组为 `<a:solidFill><a:srgbClr val>`（`#rrggbb` 归一，§2.2；同 run 重复上色只留一个填充）、`setShapeGeometry` 只写目标形状的 `<p:spPr><a:xfrm>` 的 off/ext（px×9525→EMU，几何唯一换算点，§2.1；组合形状拒绝；占位符只在用户明示覆盖时才写 xfrm，不在读侧烘焙，D3）；四者共用一套定位（`locateShape`→`locateRun`）与 get-or-create 助手与前/后像 op 契约，`apply`/`revert` 内容寻址（§2.5），邻居 run 与别的形状逐字不动（§2.3）。本机 python-pptx 交叉读真实 deck：`setShapeRunText` 编辑后标题变、粗体 run 保住，`apply→revert` 后 **38 个部件全部字节级回到原样（0 个变化）**；`setRunFont` run0 改非粗体 20pt、run1 加粗，读回 run0 `bold=False size=20.0`、run1 原 11.0pt 未受牵连；`setRunColor` run0 上 `#c00000`，读回 `rgb=C00000` 且其 `bold=True` 留住、run1 无颜色且 11pt 不动；`setShapeGeometry` 把文本框移到 0.5″,1.0″ 缩放为 6″×2″，python-pptx 读回 `left=0.5 top=1.0 width=6.0 height=2.0`（inches），文字保住。
**几何那条查出的一处红线缺陷（`fix(pptx-io)` a3bd03c，本批）**：`setShapeGeometry` 从前"给了哪个写哪个"，在**自己没写 `<a:xfrm>` 的形状**（标题/正文占位符，`boxPx` 为 `undefined`）上只给 `topPx` 时，get-or-create 就写出一半：`<p:spPr><a:xfrm><a:off y="571500"/></a:xfrm></p:spPr>`——既缺 `@x` 也缺 `<a:ext>`。
一手依据：ECMA-376 Part 1 §20.1.1.15 `CT_Point2D` 把 `@x`/`@y` 都定成 required attribute、§20.1.1.14 `CT_PositiveSize2D` 同理，`CT_Transform2D` 又把 `a:off` 与 `a:ext` 都列为必选子元素。
**独立读取器实测**（`tmp/probe/geo-probe.mjs`，`PYTHONPATH=tmp/pylibs`）：python-pptx 读这份产物直接抛 `InvalidXmlError: required 'x' attribute not present on element {…}off`，也就是 §2 那条"产物必须能被办公软件打开"被踩穿；而我们的 `auditPptxPackage` 当时回 **0 条**（它只判 OPC 五件与 `cSld>spTree`、版式关系，没判 DrawingML 的属性），所以这条当时只能靠写侧自己拦。
**这一类现在是机检了**（`feat(pptx-io)` 0eebd25）：审计补 `slide-incomplete-transform` 一条，落盘前看得见半个 `<a:off>`/`<a:ext>`（含组合形状的 `chOff`/`chExt`），判定与写侧那刀同一套条款号，细节与实测记在 D10 的"已落"段第三条。
修法两层：**引擎动手前先按必填属性核一刀**（`missingPairValues`，缺哪个就在理由里报哪个字段名；形状已有的那一半原样留住，所以正常形状单改宽/单改高照旧合法），**代理再拦在提案面**（`slideRequestsOfProposal` 见 `shape.boxPx` 缺失且四个值没给齐时逐条不交出、给中文理由，免得把注定失败的一条递给主进程换来整批不落地，§2.4）；工具面描述与 `SLIDE_SYSTEM` 同步这个约束，模型看得到。
**为什么不补另一个轴的数**：那是把用户没要求的那一轴也钉死、顺手切断版式继承链（D3 的反面）；真要补继承值该走"母版/版式继承链解析"那条（下面阶段 2 的"仍缺"），不是在写侧猜。
真 app 实跑（`tmp/probe/slide-geo-placeholder-ui.mjs`，桩端点 + 假 Key，`--user-data-dir=tmp/slide-geo/profile`，CDP 驱动 dist）：一条只给 `topPx` 的占位符提案进卡后**当场标成"改不了"**（卡上 `2 处改动` + 那句跳过理由，另两条照样落地），磁盘上两个形状各有一对完整的 off/ext（`952500/571500 + 4762500×1143000` 与新宽 `2857500`、原 `cy` 留住），点「撤销上一次保存」后 `slide1.xml` **逐字回原**，python-pptx 读回的仍是继承值（标题 `457200/274638/8229600/1143000` 来自版式，形状自己没写）。
**叠放次序这批（`moveShapeLayer`）**：一手依据是 ECMA-376 Part 1 §19.3.5 的 `CT_GroupShape`——spTree 的子元素序列是 `p:nvGrpSpPr`、`p:grpSpPr`，其后才是 `p:sp`/`p:grpSp`/`p:graphicFrame`/`p:cxnSp` 的 choice；标准正文**没有**写遮挡语义，那句由 python-pptx 的一手源码补上（`src/pptx/shapes/shapetree.py:541`："The first shape in the sequence is the backmost in z-order and the last shape is topmost."）。所以引擎只重排形状子元素、不写任何属性：`rewriteShapeOrder` 只填形状那些槽位，`p:nvGrpSpPr`/`p:grpSpPr` 与节点间空白原地不动；`up`/`top` 已在最顶、`down`/`bottom` 已在最底时回一句理由，不交 `before==after` 的空改动，装配端与视图共用 `slideDeckModel.ts` 的 `layerEdgeNote`（`slide.shapes` 本身就是文档顺序，不必再解析一次），
先到端的条目逐条不交出，免得主进程整批拒（§2.4）。实测 `tmp/slide-layer-probe.mjs`：一页两形状的夹具（`twoShapeDeckParts`）里 `top` 之后自己的读侧与 python-pptx **都**读到 `3:TextBox 2 -> 2:Title 1`，`revert` 后两边都回到 `2:Title 1 -> 3:TextBox 2`，三份文件的 `auditPptxPackage` 各 **0 条**。

**视图上那四个层级按钮这批**：`SlideDeck.tsx` 每个形状行的头部一排 `上移一层 / 下移一层 / 置于顶层 / 置于底层`，标签仍取 `LAYER_DIRECTION_LABEL`（与提案卡同一份，§3.2 / §5.7）。
**这一条不进草稿、点一次落一次盘**：`up`/`down` 相对的是当前顺序，攒进草稿再点第二次会退化成只挪一层，而这四个按钮的语义就是点一次动一次（PowerPoint 功能区同样即点即动）；
到头的那两头由 `layerEdgeNote` 直接置灰并在 `title` 里说理由，所以点出去的那条不会被引擎整批拒。
落地后主进程回包带重读磁盘的模型，形状列表按新的文档顺序**当场重排**——这就是按钮的唯一可视反馈，视图不自己画缩略图。
同一批改动顺带把**撤销不再清草稿**定死：撤销只回已落盘的那批，用户还没保存的输入是另一件事（草稿存的是绝对值，回滚后 `pendingEdits` 对着新现场重算，等于现场值的那条自己就不交了）。
真 app 实跑（`tmp/probe/slide-layer-ui.mjs`，`--user-data-dir=tmp/slide-layer-ui/profile`，CDP 驱动 dist，一页三形状夹具）：
初始禁用态如实——最底的 `Title 1` 是 `下移一层(禁用),置于底层(禁用)`、最顶的 `TextBox 3` 是 `上移一层(禁用),置于顶层(禁用)`、中间的 `TextBox 2` 四个都可点；
在 `TextBox 2` 上留一条未保存的文字草稿后点 `Title 1` 的「上移一层」→ toast「已把「Title 1」上移一层」，磁盘当场变 `3:TextBox 2 -> 2:Title 1 -> 4:TextBox 3`，
列表行序跟着变、那个形状的四个按钮同时全部解禁，草稿仍是 `1 处待保存` 且输入框值仍是「正文乙改过」；再点「置于底层」→ 磁盘回到 `2:Title 1 -> 3:TextBox 2 -> 4:TextBox 3`（草稿仍未保存）；
点「撤销上一次保存」只回最后那一次点击（磁盘回到 `3:TextBox 2 -> 2:Title 1 -> 4:TextBox 3`），**草稿没被吞掉**（仍 1 处待保存、值仍是「正文乙改过」）；
最后点「保存」→ 磁盘上 `<a:t>` 依次是 `正文乙改过 | 标题甲 | 正文丙`，与新的叠放顺序一致。事后 `auditPptxPackage` **0 条**，python-pptx 独立读到叠放 `[3, 2, 4]`、文字 `['正文乙改过', '标题甲', '正文丙']`。
顺手实测出一条**读侧的宽容**：夹具的幻灯片 XML 少一个 `</a:p>` 时，`readSlideShapes` 不报错，面板如实显示「这张幻灯片没有可读出的形状」——解析失败现在被折成"没有形状"，见下面「仍缺」。
**仍缺**：历史面板按格式渲染幻灯片记录（`components/HistoryPanel.tsx` 正在 dcc 手上改，本批没动它，见下）、画布式选中与拖拽尺寸（现在能改几何的是数字输入框）。
上面那类"DrawingML 必填属性缺一半"已经从"只有写侧拦"变成机检（`slide-incomplete-transform`，见 D10 第三条），但审计目前只覆盖 `xfrm` 这一族属性/子元素；
版式与母版部件（`slideLayout`/`slideMaster`）里同一族元素还没纳入判定，母版的 `cSld > spTree` 那两条也只对幻灯片生效。
**增删形状这批（`addTextBox` / `removeShape`，本批新落）**：一手依据纠正了本文件前面那句"ECMA 要求 `p:cNvPr @id` 在**整份文档**内唯一"——**它是错的**。
python-pptx 分配形状号的那段源码（`tmp/pylibs/pptx/oxml/shapes/groupshape.py:151-163` 的 `max_shape_id`）取的是**这一页 spTree 上的** `xpath("//@id")` 最大值，注释写 "XML id-values have document scope"，那份 document 就是这一个幻灯片部件；
本机实测也对得上：`tmp/pptx-edit/probe.pptx` 两页的 id 分别是 `[1,2,3]` 与 `[1,2]`，**跨页本来就重复**。
所以引擎按"这一页部件内 `max(@id)+1`"分配（`maxShapeIdInPart` 扫这一页 XML 里所有数字 `@id`，`p:section/@id` 那种 GUID 非数字、正好被过滤掉），新建的 `<p:sp>` 骨架逐字照 python-pptx 的 `CT_Shape._textbox_sp_tmpl()`（`oxml/shapes/autoshape.py:388`，也就是 PowerPoint 自己插入文本框写的那一段：`p:nvSpPr > cNvPr/cNvSpPr txBox="1"/nvPr`、`p:spPr > a:xfrm + a:prstGeom prst="rect" + a:noFill`、`p:txBody > a:bodyPr + a:lstStyle + a:p`），
名字沿用它的 `"TextBox %d" % (id-1)`；空文字时保留模板那个 `<a:p/>`，不造空 run；文字来自用户输入，写进 `<a:t>` 前过 `opc` 的 `escapeText`（§3.6 只在系统边界校验）。
新形状挂在形状子序列末尾 = 最顶层（叠放语义同上），删形状就是把那一个元素从 spTree 里摘掉——**不动 `.rels`、不动 `[Content_Types].xml`**：OPC 只要求被引用的关系存在，删元素只会少引用不会多引用，所以 §2.10 那"三处一起登记"没有要回滚的东西，撤销与前像核对一句没改（结构 op 共用整页前/后像这一条通道）。
单测 7 条：形状号是本页最大值 +1 且已有形状逐字不动、px→EMU 只在 `units.ts` 那一处（96px→914400）、空串保留 `<a:p/>`、文字里的 `<` `&` 转义后读回原样、宽高非正与部件缺失各回理由且不落盘、删形状后其余形状逐字不动 + 撤销逐字回原、删到最后一个形状是合法操作（读侧回空数组不抛异常）。
**顺带抽出 `loadSlidePart`**：`locateShape` 里"读部件 + 解析 + 失败给中文理由"那一段现在与整页粒度的 `addTextBox` 共用（§3.2 第二次出现就抽）。
另有一条本批实测撞到的**读侧宽容**：幻灯片部件解析失败时 `readSlideShapes` 回空数组，面板于是只显示「这张幻灯片没有可读出的形状」，把"这一页读不懂"说成了"这一页是空的"——该如实分成两句（阶段 2 的读侧要补一条失败理由通道）。**已 closed（`e5fdb16`）**：读侧现在回三种失败码 `missing-part` / `invalid-xml` / `no-sp-tree`（`Slide.readFailure`），中文句子只在 `packages/shared` 存一份（`describeSlideReadFailure`），面板、`oc ppt info`、模型收到的 `reason` 三处共用同一句；`readSlideText` 对读不动的页直接回"这一页不是'空的'，别对它提改动"，面板在该页禁用「加文本框」。取证见 §八 对应行。
**实测（modlog 落盘 + 内容寻址撤销这条已达标，本批新落）**：撤销的凭据不再只活在渲染层的内存批次里。
`apps/desktop/src/slides/slideHistory.ts` 把每次落地那批 op（整页前/后像）连同 label/summary 写成一条 `幻灯片改动` 记录，
存进**同一份** `userData/modlog/<文档名>.json`（复用 `electron/modlog.ts` 那四个通道，无第二套历史栈 = C5）；
撤销走 `revertSlideEdits`→**从历史里取最新一条没撤销的记录**，`pickLatestSlideEdit` 与表格那条 `pickLatestSheetEdit` 同构。
记录形状靠"每条 op 都带 `beforeXml`/`afterXml`"这个签名识别——文档的 `CoreOp` 与表格的 `SheetCoreOp` 都没有这两个字段，
所以三类记录混在同一个文件里也不会互相认错。
**现场核对下沉到引擎**：`pptx-io/edit.ts` 新增 `revertSlideEdits(file, ops)`，倒序逐条问"这一页现在是不是正是改完后的那个样子"
（与 `op.afterXml` 逐字一致才写回前像），对不上就跳过并给中文理由；`pptxHandlers` 的成功回包因此多出 `skipped`，
渲染层据此报"部分撤销（x/y）"而不是假装全回去了（§2.5）。
**真 app 实跑**（`--user-data-dir=tmp/slide-modlog/profile`、这一批不涉及 LLM、CDP 驱动 dist）：
第一轮在幻灯片面板改标题→点保存，toast「已保存，动了 1 张幻灯片」，磁盘标题变「重启后还能撤」，
modlog 里出现 `{label:"幻灯片改动", summary:"手工改了 1 处", opCount:1, beforeXml 含"标题甲", undone:false}`；
**把整个进程杀掉再重开同一份文件**，「撤销上一次保存」不再是禁用态（凭据在磁盘上，不在内存里），
点下去 toast「已撤销上一次保存的那批改动」、磁盘回到「标题甲」、那条记录变 `undone:true`，按钮随即回到禁用（一条记录不许撤第二次）。
单测：4 条引擎用例钉现场核对（一致才写回 / 同批两条改同一页能一路链回原样 / 那一页后来又被手工改过则整条撤不动且"用户自己改的"一个字不丢 /
部件已不在包里给理由不抛异常），1 条出入口用例钉"拿旧批次撤不动、最新那批照旧能撤"，6 条记录层用例钉三类记录互不认错与按 ts 取最新。
**没有新增第二条撤销通道**：手工保存与 AI 提案仍共用 `App.tsx` 那一个落地口，历史里也只有那一个来源。
**真 app 实跑（`tmp/probe/slide-add-remove-ui.mjs`，这一批不涉及 LLM、CDP 驱动 dist）**：打开 `tmp/slide-add-remove/probe.pptx`（第 1 页两个形状 `ids=[2,3]`）→
点第 1 页标题行的「加文本框」→ 磁盘形状子序列当场变 `[2,3,4]`，列表多出一行 `TextBox 3`，类型标签是"文本框"、那四个几何框里的值是 `96 / 96 / 480 / 144`（px 交出什么读回什么，视图里没有第二次换算），
段输入框是可编辑的"新文本框"而不是那行只读提示；把它改成 `Q4 目标` 点「保存」→ 磁盘 `<a:t>` 变 `Q4 目标`、脏点不出现；点那行的「删形状」→ 磁盘回到 `[2,3]`、
`txBox="1"` 由 2 变回 1，且整页 XML **去掉 XML 声明后与原文件逐字一致**（声明那行的差异是本仓既有的整页重写行为：原文件是 `<?xml version='1.0'…?>` 单引号 + 换行，我们一律写双引号且不接换行，语义等价）；
点「撤销上一次保存」（那批是"删形状"）→ 形状 4 连同 `Q4 目标` 一起回来。事后 `auditPptxPackage` **0 条**、`CHANGED_PARTS` 只有 `ppt/slides/slide1.xml`（其余部件逐字节没动）；
python-pptx 独立读到 3 个形状：`id=4 name='TextBox 3' type=TEXT_BOX (17) text='Q4 目标' left/top/w/h = 1.0/1.0/5.0/1.5 英寸`（= 96/96/480/144 px @96DPI），
且 `s.shapes[-1]` 就是新那个（最顶）。历史里是 3 条 `幻灯片改动`（`在第 1 页加一个文本框` / `手工改了 1 处` / `删掉「TextBox 3」`），撤完最后那条转 `undone:true`——**没有新开第二条历史栈**。
视图侧的一条新守卫：删形状落地后把**那个形状**的未保存草稿一并丢掉（`slideDeckModel.dropShapeDrafts`）。
不丢会出一条静默写错地方的路：形状号是"本页最大值 +1"，删掉的正好是最大号时下一次新建会把这个号发给新形状，用户给旧框打的字就落进新框；
3 条单测钉住"只掉那一个形状的各段与它的几何"、"号 4 的清理不带走号 41 的草稿（分隔符算在键里）"、"没命中时原对象返回"。
**历史面板这一批故意没动**：`HistoryPanel` 现在把所有记录都按文档记录渲染（`describeCoreOp` 对幻灯片 op 返回 undefined，
逐条改动列表是空的，且那个「撤销这次修改」按钮接的是 Tiptap 那条撤销），幻灯片记录要进面板得先按格式分派——
而该文件正带着 dcc 未提交的改动，改它会撞车，所以本批只把记录与撤销做扎实，AI 提案卡上的幻灯片那张仍不给「查看修改历史」入口。
**"改完的 `.pptx` 会不会被办公软件判损坏"现在有机检面**（与 sheet 侧 `oc sheet check` 同源）：与格式无关的五条判定（内容类型覆盖、Override 指向存在、关系 Target 可解析、关系 Id 不重复、部件里 `r:id` 查得到）从 `sheet-io/compat-audit.ts` 上抽到 `@office-copilot/opc` 的 `auditOpcPackage`（`AGENTS.md` §3.2：`.xlsx`/`.pptx` 共用，不许长第二套；`localName` 同时抽进 `opc/xml.ts`），`auditXlsxPackage` 只留 `<worksheet>` 顺序与 `<c:chart>` 两条专属判定——**sheet-io 245 条单测一条没改、一条没减**，证明这次抽取是行为等价的。`.pptx` 专属两条落 `pptx-io/audit.ts` 的 `auditPptxPackage`：`<p:sld>` 里必须有 `cSld > spTree`（cSld 是 Part 1 §19.3.1.16，"最小一张幻灯片"要 `CommonSlideData` 里套 `ShapeTree` 出自 Learn "Working with presentation slides"）、幻灯片必须有一条指向 slideLayout 的内部关系（同页原话把 SlideLayout 关联列为必须，**ISO 正文里没有对应 MUST**；Target 是否存在由通用那条覆盖）。条款号初稿凭记忆写过 §19.3.1.9/.10/.2，已按官方一手页复核改过（见 D10 的复核记录）。7 条新单测（合规包 0 条 / 未登记部件 / 空 `cSld` / 无关系表 / 版式关系指向缺失部件 / 夹具与改后产物各自 0 条）。**本机实测**：python-pptx 1.0.2 生成的真实 **40 部件** deck 审计返回 `[]`（对第三方产物零误报），`setShapeRunText` 改标题后的写回同样 `[]` 且 `changedParts` 只有 `ppt/slides/slide1.xml`，python-pptx 读回标题为"改过的标题"、页数 2 未变。

### 阶段 5 · AI 工具面 + `slideAgent` — `[~] 工具面 pptx-tools（九类改动，与引擎那九个 op 齐平；表格那一格那条必须同时带行与列，边界只在 zod 那一份里说）+ slideAgent（含 23 条单测：页号/形状号/行/列/段/run 五层对不上现场模型的一律不交出，逐条给中文理由）+ 桌面提案卡装配已落；真端点模型行为探针仍缺`
D6 的通用 schema 核心 + `PPT_TOOLS`（分析/整理/生成三类），提案卡复用现有装配。
**D6 落定**：不新建 `packages/agent-tools`、也不塞进 `packages/shared`，而是**照 `word-tools`/`sheet-tools` 的先例每格式一个工具面包** → 新建 `packages/pptx-tools`（纯 `zod`，不 import `pptx-io`，与既有两包同构）。工具面含 `ppt_describe`/`ppt_read_slide_text`（读，页/形状/段/run 由引擎从现场文件列）与 `ppt_submit_proposal`（写，初批四类：`set_run_text`/`set_run_font`/`set_run_color`/`set_shape_geometry`，一一对应 `pptx-io/edit.ts` 当时那四个带逆操作 op；引擎把叠放与增删建成七个 op 后，工具面在下面"第二批"一段里同步成七类）。坐标是模型视角：页号 1 基、shapeId 取自 describe、段/run 0 基；单位说在描述里（字号 pt、几何 px，换算留在引擎边 §2.1）；跨字段"至少给一项"约束放外层 `superRefine`（`discriminatedUnion` 不接受被 `.refine` 包的成员）。
**完成判定**：单测（mock fetch）覆盖"模型给数字/给原始 XML/给越界索引都要回中文理由"；真端点探针实测模型会不会点对工具、
落地后 Keynote PNG 里文字真的变了。
**实测（工具面校验这条已达标）**：`packages/pptx-tools` 15 条单测覆盖坏 JSON / 缺字段 / 多字段(strict) / 越界(页号 0、字号 401、宽 0、文 2001 字) / 未知 type / font-与-geometry 不给任何属性 → 全部回中文理由、绝不抛异常；并断言每个 `PPTX_TOOLS` 定义都能在同名 schema 表里找到（不给模型一个校验不了的口子）。
**实测（`slideAgent` 装配这条已达标，AGENTS §2.4/2.8/2.11）**：`apps/desktop/src/ai/slideAgent.ts` 是**唯一**把模型视角（页号 1 基）换成引擎视角（`ppt/slides/slideN.xml` + `shapeId` + 段/run 0 基）的地方（`slideRequestsOfProposal`），12 条单测钉住四类改动各自落到正确的部件路径与序号、多 run 段落按模型指的 run 改（不整段覆盖，§2.3）、
页号/形状号/段号/run 号对不上现场模型的**逐条不交出并给中文理由**（"没有第 9 页""没有 shapeId=99""越界（只有 1 段）""没有文字体""文字来自字段"），而不是整批静默不落；上送首条消息带结构画像并说清单位与截断规则（25 页只描述前 20 页，§2.9）；暂停留在两轮之间、`messages` 原样攒着供「继续」复用；8 轮封顶。
**真 app 实跑（`tmp/slide-ai-app-probe.mjs`，端点是本机 8899 桩 + 假 Key，验装配面不是模型行为）**：打开 `tmp/slide-ai/probe.pptx` → 幻灯片视图 + 那条幻灯片引导语 + 幻灯片专属 placeholder，文档的快捷指令入口不出现（`chips:0`）；发一句"把第 1 页标题改成 Q3 复盘，正文里『粗』那一截标红" → 卡片头写 **2 处改动**（只算能落的两条）、清单里三条都念出来、越界那条后面跟 `· 没有第 9 页（这份文稿共 2 页），这条跳过。`；点「应用到幻灯片」→ 卡上出现"已写入 1 张幻灯片。要回退请在幻灯片面板点「撤销上一次保存」。"，磁盘 `slide1.xml` 的标题 run 变 `Q3 复盘`、`srgbClr val="C00000"` 只出现在第 1 截 run 上（`正文丙`/`细` 两截原样，§2.3 没被拍平），标题栏脏点**不出现**（落地当场写盘，内存没有待存改动），幻灯片面板的「撤销上一次保存」由禁用变可点；点它 → toast"已撤销上一次保存的那批改动"，磁盘回到 `标题甲` 且 `colors: []`。
**表格那半边没被这次泛化改坏（`tmp/sheet-ai-app-probe.mjs`，同一只桩）**：表格引导语/placeholder/无快捷指令都在；`sheet_create_chart` 回来的图表卡显示"1 张图表 + 引用与数值都由引擎从现场表读"、`ChartPreview` 的 canvas 在、点「插入图表」只排进待写清单并置脏（dirty=true），文案如实说"撤销栈管不到图表"；`sheet_submit_proposal` 回来的改动卡"2 处改动 · C2 ← 999 · B5 ← 华南"，点「应用到表格」后网格当场是 `999`/`华南`，且 `D3` 的 `SUM(C2:C5)` 缓存值跟着重算成 2269（§2.11），卡上给「查看修改历史」。
**第二批（把引擎已有的两条结构 op 开进提案面）**：阶段 4 建完 `addTextBox`/`removeShape` 后，工具面还停在五类——那是一处没人说过的收窄（模型能提的比引擎能做的少，用户说"加个致谢框"时模型只能回"做不了"）。现在 `PptxEdit` 的判别式联合是七条，与 `edit.ts` 那七个 op 生产者齐平：
`add_text_box` 要页号 + 四个几何值 + 文字（全必填，`strict` 连 `shapeId` 都拒——新形状号由引擎在落地时按该页现有最大值 +1 分配，模型在这一份提案里引用不到自己刚要建的那个框，这条写进字段描述与系统提示词），`remove_shape` 要页号 + 形状号。
给模型的 JSON `oneOf` 与 zod 联合由一条新单测钉住必须逐项相等（只加校验忘了改声明 = 模型根本提不了那条；反向也一样）。装配面加一条**同批守卫**：这一批里已经删掉的形状再被指到，那条只进 `notes`——按数组顺序落地时它必然定位失败，而失败语义是整批不写（§2.4），等于让用户白点一次确认；现场模型是提案前那一份，所以光查"这一页有没有它"拦不住重复删。
**真 app 实跑（`tmp/probe/slide-ai-structure-ui.mjs`，本机 8899 桩 + 假 Key，验装配面不是模型行为）**：桩交四条提案（加文本框、删形状 3、改刚被删的那条、删根本不存在的第 9 页）→ 卡片头 **2 处改动**，新文本框那行是「第 1 页 加文本框「谢谢观看」（96,420px 起 480×120px）」、删形状那行是「第 1 页 · 形状 3 删掉整个形状（连同它里面的文字）」，两条拦下的各跟一句理由（「在同一批里已经被删掉」「没有第 9 页（这份文稿共 2 页）」）；点「应用到幻灯片」→ 磁盘形状序列 `[2:Title 1, 3:TextBox 2]` 变 `[2:Title 1, 4:TextBox 3]`（新框拿到 4 号，因为它落地时页内最大是 3，随后 3 号被删），`<a:t>` 只剩 `标题甲 | 谢谢观看`，脏点不出现，幻灯片面板那行跟着列出新框；点「撤销上一次保存」→ `slide1.xml` 与原文件**连声明都逐字一致**（`byte:true`，前像本来就是从磁盘读的那一份）。落地仍走 `ppt:edit` 那一个口，所以存前审计（`auditPptxPackage`）与 modlog 一句没改。
**第三批（几何单轴提案的预检，与阶段 4 那条缺陷同一批）**：`slideRequestsOfProposal` 对 `set_shape_geometry` 多加一刀——`shape.boxPx` 缺失（占位符那种从版式继承）且四个值没给齐时只进 `notes`，中文理由直说"要改就得四个一起给"；`SLIDE_SYSTEM` 与 `ppt_submit_proposal` 的描述同步这句，让模型第一次就提对。完整细节与实测记在阶段 4 的「几何那条查出的一处红线缺陷」。
**第四批（备注那条 op 开进提案面，`ad4d995`）**：`fccd4f9` 落的 `setSlideNotes` 当时是"引擎有 op 而模型够不着"，这一批补的就是这一环——`PptxEdit` 加第八条 `SetNotesEdit`（`type`/`slideNumber`/`paragraphTexts`，一段一个字符串、1~200 段、单段 ≤2000 字，`strict`），给模型的 JSON `oneOf` 同步加分支，那条"zod 与 `oneOf` 逐项相等"的单测把清单从七项钉到八项。**模型只交页号**：`notesSlideN.xml` 那个路径由 `slideRequestsOfProposal` 按现场模型翻，与 `add_text_box` 同一条理由（引擎坐标不上送）。**预检用的是读模型新带出的两个事实**：`SlideNotes.paragraphRunCounts`（每段几截 run）与 `styledRunCount`（带 `<a:rPr>` 的截数），压平与否那份判定从 `setSlideNotes` 抽成 `notesFlattenNote` 给两边共用（§3.2），于是"那一页根本没有备注部件"和"现有备注带多种格式"这两类都在提案阶段就只进 `notes`、不进批次——不然用户点了确认才发现整批不写（§2.4）。落地通道一个字没动（`pptxHandlers.ts` 那条 `set_notes` case 早在）。**真 app 已实跑**（`tmp/probe/slide-notes-ui.mjs`，8901 桩 + 假 Key，细节与逐部件字节证据见 §八 那行）：三条提案只算「1 处改动」，两条拦下的各带一句中文理由，落盘后 49 个部件里只有 `notesSlide1.xml` 变了，python-pptx 独立读回新备注三句、第 3 页那截加粗原样还在，撤完 49 个部件逐字回原样、`oc ppt check` EXIT=0。**仍没跑的**：真端点一次都没发，所以"模型自己会不会主动提 `set_notes`、会不会把段数与空行给对"没有证据（与 B5 同一格欠账）；面板的备注显示与编辑入口随后由 `2dc83f4` 接上（见阶段 2「面板的备注那一格已落」那段），这一条 op 至此引擎 / AI / 手工三面齐。
**仍缺**：真端点探针看模型会不会点对 `ppt_*` 工具名与形状 id（B5），也包括它会不会主动用 `add_text_box`/`remove_shape`/`set_notes`。pptx 的 modlog 落盘已在阶段 4 收掉。
### 阶段 6 · 图表与表格（D9）— `[~] 幻灯片里的表：逐格读侧 + 第九个 op（改某一格的那条 run）+ AI 提案面 + 视图网格式编辑已落；表格里的字号/颜色**四面齐**（`fffda8a` 引擎那两条 op 的坐标带行列就走同一套 `<a:rPr>` 改动，`4e351ef` 补上 `pptx-tools` 的可选行列与视图上那两个输入框）；还缺的是表格里的**继承字号解析**（那一格没写显式值时框子为空）与图表部件的读写、`c:numCache` 同步`
**完成判定**：新建图表的 `c:numCache` 与内嵌 `.xlsx` 的值由 openpyxl 独立读回且一致；改数据后两者同步；撤销逐格回原样。

### 阶段 7 · pdf 线：读 + 审计 + 甲档正文改写（D8 甲）— `[~] 读模型（两种串形态 / Form XObject 递归 / CID+ToUnicode / 压缩流）+ W1 同长度替换 + 增量更新 + 引擎侧按内容寻址撤销 + 桌面装配（pdf:read/edit/revert/save-as 四条通道、PdfView 面板、modlog 落盘）+ AI 工具面（`packages/pdf-tools` 三条工具 + `src/ai/pdfAgent.ts` 提案面，落地复用 `commitPdfEdits` 那一个口）+ W2 第一刀（内容流跟住文字状态，每条 run 带出字形原点与**生效字号**）+ W2 第二刀（逐字形的 run 拼成**逻辑文本行**，AI 概要与视图按行读）+ W2 第三刀（逻辑行进**编辑**面：`set_line_text` 由 `pdfLineEditsOf` 按行内各 run 原字数拆回逐 run 的等长替换，真文件改完能撤）+ **结构体检**（引擎 `auditPdfStructure` 七码 + CLI `oc pdf check`，三份真文件实测、无误报）+ **打开时把结论念给用户的那条提示**（面板顶部红/黄一档，读模型都建不起来时那两句进 toast；文案抽进 `packages/shared`，两个装配层只剩一份）已落（两份真文件 + QuickLook 肉眼核对 + 真 app 实跑截图，四份形状各一轮）；`print:pdf` 旁路已具备（`readPdfModel` 读得动自家 `printToPDF` 产物，A10 按现通路关闭），残留的是字体 ToUnicode 保真（部首码位与拉丁行残缺）；W2 本体（宽度感知重排 + 字体重新子集化，见 D8 分档那句澄清）仍缺`
`packages/pdf-io` 读页/文本块/字体/内容流偏移，`print:pdf` 通路带旁路模型。
**完成判定**：本仓自产的 PDF 改一处正文 → pdf.js 读回新文字、其他页字节不动、能撤销；对无 ToUnicode/扫描件/签名文件
**给出中文拒绝理由而不是静默改坏**（逐条有单测）。
**现状**：拒绝理由已是引擎错误码（`length-mismatch` / `char-outside-subset` / `run-not-found` /
`no-encoding`（看不出这段文字念作什么，含缺 `/ToUnicode` 那格，有专门的夹具钉住）/ 撤销时的 `text-drifted`），
失败时整批不产出字节，符合 §2.4；中文文案要等装配层。撤销走 `revertPdfTextEdits`，与写入共用同一份实现，
先按**内容**比对再动手，所以撤那一条不会把用户后来的手工编辑一起冲掉（§2.5）。
装配那一半已落（2026-10-15）：`electron/pdfHandlers.ts` 开四条通道，`src/pdf/{pdfRuns,PdfView}` 是判定层与视图（草稿只交定位键，
字节预算不在渲染层判——编码规则属于 `pdf-io`，再算一遍就是第二套换算，§2.1），`src/pdf/pdfHistory.ts` 复用同一份 modlog
（记录标签「PDF 改动」，签名是 `applied`，与文档 / 表格 / 幻灯片那三类同存一个 docKey）。
**AI 工具面已落（2026-10-15，阶段 5 的 pdf 那一半）**：`packages/pdf-tools` 给三条工具（`pdf_describe` / `pdf_read_page_text` /
`pdf_submit_proposal`），zod 入参解析与 OpenAI 形态声明逐项相等（沿用 `ppt-tools` 那条单测口径）；**模型只看 pageNumber（1 基）与
runIndex（该页第 0 基条），引擎坐标 `streamRef`/`withinForm`/`start` 绝不上送**，换算只发生一次，在 `apps/desktop/src/ai/pdfAgent.ts`
的 `pdfRequestsOfProposal` 里——页码越界、runIndex 越界、`runEditBlock` 判不可改（内联流 / 认不出念作什么）、与原文相同、字数不等
这五种各给一句中文 `notes` 而不是把坏条目送去写（§2.4 是整批不写，混进去等于让用户白点一次确认；字数那条只是**预筛注定失败的项**，
字节权威仍在引擎）。落地仍是 `App.commitPdfEdits` → `pdf:edit` 那一个口，所以 modlog 记录、撤销凭据、失败文案（`pdfFailureText`）
都只有一份；卡片不新开第三张，`TOOL_CARD_COPY.pdf` 加进 docx/表格/幻灯片共用的那一张工具回合卡。pdf 模式下 docx 那套
快捷指令与选区改写入口一并关掉（它们改的是 Tiptap 编辑器，对着 PDF 点了会误导），引导语换成 PDF 那条。
**这一批顺手抓出并修掉一条 §5.2 红线缺陷**：渲染层 import 了 `pdf-io` 的**主入口**，Vite 于是把 `require("zlib")` 与 `@cantoo/pdf-lib`
打进浏览器包，真 app 里 PDF 视图挂载前就抛 `ReferenceError: require is not defined`，整页空白——`pnpm build` 与 `pnpm test` 全都能过，
只有实跑 CDP 才看得见。修法是把包分两半：无 Node 的那一半（`pdf-model.ts` 的类型与判定 + `analysis.ts` 的 AI 读面）经新子入口
`@office-copilot/pdf-io/pure` 给渲染层，带 `node:zlib`/pdf-lib 的 `pdf-file.ts` 只留主进程；并补 `apps/desktop/src/pdf/rendererIoEntry.test.ts`
作机检（`src/` 下任何文件 import 主入口即失败）。
**真端点实跑（2026-10-15，用户本人配的服务商，Key 只在 `tmp/`）**：改署名与改职位行两次都对——第二次目标是页内**第 6 条**（不是第一条，靠位置猜不出来），
落地仍是"前缀逐字不变 + 只追加 delta"；那句"改成 32 磅红色"没有出提案卡，模型直说做不了并点名 `set_run_text` 只能等长整条替换。
**观测不到的一条如实记**：想在渲染层数工具回合（包 `window.officeCopilot.aiToolTurn`）包不上——contextBridge 交出的对象在页面主世界里只读，
赋值静默无效；而 `pdf_describe` 的每页概要本身是按 run 拼起来的，所以**不能**据"坐标对"断言它一定调过 `pdf_read_page_text`，要看回合数得在主进程侧接计数。**`print:pdf` 旁路现状（2026-10-15 现跑，见 §八 A10 那行）**：`readPdfModel` 能把 `printToPDF` 的产物读回并定位（1 页 59 run，旁路是通的），A10 的子集命名缺陷按现通路不再适用；残留的是字体侧 ToUnicode 把长/高映成部首码位、拉丁行读回残缺（属上游数据，非本仓解析器）。`oc pdf check`、逻辑行拼接、桌面打开时的提示条都已落。

### 阶段 8 · 收口：CLI、文档、规则落地 — `[~] oc ppt info / oc ppt check / oc pdf check、README 的 PPT 与 PDF 事实、AGENTS.md 的四类格式事实都已落；只剩 §六 里那几条既有项（A1/A3/A4/A6/A7/A9 的签名与加密两格/A10、B1–B3/B5）`
**完成判定**：`oc ppt check`/`oc pdf check` 可用（两条均已落，pdf 那条有问题时退出码 1）；`README.md` 与 `AGENTS.md` 里四类格式的既有事实更新；
本文档 §六 全绿或如实标 `[!] BLOCKED`。

---

## 五、执行计划（任务 → 落点）

| 阶段 | 主要落点 | 备注 |
| --- | --- | --- |
| 0 | `packages/pptx-io/scripts/make-sample-deck.ts`、`tmp/pylibs`（不入库） | 夹具由脚本现生成，不提交二进制 |
| 1 | 新包 `packages/opc`；`packages/sheet-io/src/opc.ts`（删）、`packages/docx-io/src/docx.ts`（改依赖） | 见 D2 的顺序红线 |
| 2 | `packages/pptx-io/src/{pptx-file,read,inherit,units,analysis,audit}.ts`（初稿的 `pptx-model/placeholder-inherit/compat-audit` 落地时按本包既有短名改成 `read/inherit/audit`，与 `sheet-io` 的 `compat-audit.ts` 只差这一处） | 单位换算单点：D4；几何继承链：D3 |
| 3 | `apps/desktop/src/App.tsx`（模式枚举，已落）、`apps/desktop/src/slides/**`（视图 + `slideDeckModel.ts`，已落）、`electron/pptxHandlers.ts`（照 `sheetHandlers.ts` 的 `.bak` + 存前审计，已落） | §5.2 preload 白名单同步；落点与初稿写的 `src/slide/**`+`slideHandlers.ts` 的差别见 D7 的"已落" |
| 4 | `packages/pptx-io/src/edit.ts`（op + 逆操作 + 现场核对）、`apps/desktop/src/slides/slideHistory.ts` | 复用 modlog，不新建历史栈；落点与初稿写的 `slide-edit.ts`/`src/slide/` 的差别同 D7 的"已落" |
| 5 | `packages/agent-tools`（新，D6）、`packages/ppt-tools`、`apps/desktop/src/ai/slideAgent.ts` | 工具循环复用 `ai:tool-turn` |
| 6 | `packages/pptx-io/src/chart-spec.ts` + 复用 `packages/sheet-io` 公开入口 | D9 |
| 7 | `packages/pdf-io/src/{pdf-file,pdf-model,pure,content-stream,xref,text-map}.ts`、`packages/pdf-tools/src/index.ts`、`apps/desktop/src/ai/pdfAgent.ts`、`apps/electron/printHandlers.ts` | D8 甲档；渲染层只准走 `@office-copilot/pdf-io/pure`（§5.2，见阶段 7 那条白屏缺陷） |
| 8 | `apps/cli/src/{ppt,pdf}.ts`、`README.md`、`AGENTS.md`（由 dcc 改） | — |

---

## 六、验收标准列表（逐条打分，任一不过不得声称完成）

### A. 文档正确性（最高优先级，对应 §2）
- `[x]` A1 未建模部件原样带过：富模板打开→存回，**文件部件清单逐条一致、且每个部件字节逐字一致**。**实测（2026-10-15，随 zip 序列化收口那批一起跑通）**：WPS 自带的 `normal_mac.pptx`（124 条目 / **108 个文件部件**，其中 **65 个 `ppt/tags/*`** 我们完全不建模）与 `normal.pptx`（114 / 102 / 同样 65 个 tags）各走一遍 `PptxFile.open` → `toBytes`（一条都不改）——**内容对不上的部件 0 个**，108/102 的文件部件清单与原文件逐条对齐；往返产物再过 `oc ppt check`（EXIT=0）与 python-pptx（1 页 / 1 母版 / 11 版式全读得到）。**判据原文那句"`unzip -l` 条目集一致"按 OPC 收窄成"文件部件清单一致"**：OPC（ECMA-376 Part 2 §3）的目录项本来就是可选的，而 `serializePackage` 统一不写它们，所以条目数 124→108、114→102 是那条收口的**预期结果**，不是丢部件——独立读取器与结构审计在没有目录项的产物上给出同一结论。
- `[x]` A2 母版/版式继承不被拍平：只改标题文字后 `slide1.xml` 的 `<p:spPr/>` 仍为空（D3）。**实测**：真 app 里第 1 页那个继承几何的标题占位符**根本没有几何输入框**，保存只记账第 2 页那一个部件（toast"动了 1 张幻灯片"），python-pptx 读回第 1 页三 run 段落"正文丙/粗/细"连 18pt 都原样。
- `[ ]` A3 单位换算单点：**pptx 侧已收到一处**（`pptx-io/units.ts`，`edit.ts` 里那条私有 `EMU_PER_PX` 已删）；但"全仓只有一处"这条判据写错了——每个 IO 包各有一份才是正解（sheet 侧 `chart-spec.ts`、docx 侧 `docx.ts:810`/`ooxml-pm.ts`/`textbox.ts:31` 各自一处），**docx-io 里那三处重复是既有缺陷**，与本线的 `pptx-io/units.ts` 不同源，待另开一批收敛。
- `[x]` A4 颜色：schemeClr 解析成 hex 交 UI，且不烤回 slide；`#rrggbb` 一条不减（§2.2）。**已落（2026-10-15，`8023501`）**：`pptx-io/theme.ts` 沿 幻灯片→版式→母版取 `p:clrMap`、再沿母版关系取 `a:clrScheme`（`srgbClr @val` 或 `sysClr @lastClr`），`read.ts` 只在 `schemeClr` **没有变换子元素**时解色，带 `lumMod`/`lumOff`/`tint`/`shade`/`alpha` 与 `phClr` 一律留 `undefined`（§2.2 那条"读不懂宁可不写"，本机无力核对 PowerPoint 的渲染，算错的颜色比没有颜色更坏）；解出来的 hex **从不写回部件**（单测钉住：`open → toBytes` 后 `slide1.xml` 逐字不变）。**实测**：python-pptx 现产的三条 run deck，配一份不经我们 parser 的独立 python 解色（`tmp/theme_cross.py`）逐条全等——`accent1 → #4f81bd`、带变换那条 `NONE`、直接 RGB `#ff0000`；两份 WPS 真模板 `oc ppt check` 仍 EXIT=0。**真 app 那一半已实跑（2026-10-15，`91cc3b9` 那批的 CDP 截图）**：第 1 页那条主题色 `accent1` 的文字，颜色框 `value` 就是 `#4f81bd`（框里渲染成那块蓝）、`title` 也是同一个 hex，同页另一条显式 `#00cc00` 显示成绿——不再是"当前颜色继承自版式"那句兜底（见 §八 最后两行）。
- `[x]` A5 原子写：幻灯片存盘走 `writeFileAtomically` + `.bak`（照 `sheetHandlers.ts:159`）。**实测**：真 app 点保存后 `tmp/app-deck2.pptx.bak` 出现、磁盘重读是新文字，点撤销后磁盘回到原字。
- `[ ]` A6 新建部件登记齐全（D10 的宽版五件），审计 0 条 + `oc ppt check` 退出码 0。
- `[ ]` A7 图表 cache 与内嵌工作簿自洽（D9），openpyxl 独立读回为证。
- `[x]` A8 PDF：改一处正文后其他页字节不动；**读回用 Apple PDFKit + 我们的读模型**（本仓不引 pdf.js，判据等价——第三方独立读取器抽得到新文字才算数）。**实测**：`ps-cols-OLD` 改完前缀 33 706 字节逐字不变、PDFKit 抽到"是这：段"且第二栏一字未动；那份 WPS 简历整行改完前缀 334 523 不变、行外 605 条 run 一字未动、总 run 5 196 不变（QuickLook 出的 PNG 肉眼核对只有那四个 CID 字形换位、无豆腐块）。
- `[~]` A9 PDF 的硬拒绝成文可测：无文本层 / 缺 ToUnicode / 有签名 / 加密 → 中文理由，不静默改坏（D8）。**已落的三格**：缺 `/ToUnicode` 的 run 读回如实 `null`、改写判 `no-encoding`（有专门夹具 `cidNoToUnicodeFixturePdf()`），体检另出一档 `text-not-editable`（"1 113 条里 181 条认不出内容"）；纯图片页/扫描件出"整份文件一条显示文字都没读到"；字节装不回去出 `length-mismatch`、子集缺字出 `char-outside-subset`、定位键失配出 `run-not-found`、撤销对不上出 `text-drifted`，全部整批不产出字节。**没落的那两格**：有签名 / 加密——本机 `qpdf`/`mutool`/`gs` 实测全 none，产不出这两类文件就等于写不出用例，与其留一条永不亮的分支不如整档记进 W4（D8 的加密与签名策略），所以这条仍是 `[~]` 而不是 `[x]`。
- `[x]` A10 **既有缺陷（本条已按现通路复现不出，判为不再适用）**：原记本仓 `printToPDF` 的中文字体子集命名不合 `9.6.4`
  （实测 `/SourceHanSansCN-Regular-6087`，缺 `AAAAAA+` 前缀），FreeType/MuPDF 拒绝该字体程序、Apple 渲染成原始字形。
  **2026-10-15 现跑判定**（spec 上文 line 196 要求"下次动 `print:pdf` 通路时现跑一遍再判"）：用与 `printHandlers.ts` 逐字相同的
  `printToPDF({pageSize:'A4',printBackground:true,margins:0})` 产一份中文 PDF，经**我们自己的 pdf-io 读模型**（即那条旁路）读回，
  生效字体名是 `/AAAAAA+PingFangSC-Regular`——**子集命名合规 §9.6.4**，`-6087` 那种违规名在当前 Chromium 打印通路上不再出现
  （当初那份 `-6087` 产物早不在 `tmp/`，无法再造）。判据"MuPDF/PDFKit/QuickLook 里中文正常显示"依赖的正是这个合规名（名字不合规模块才被拒），
  故本条按现通路关闭。⚠️ 但同一次实跑翻出**另一条更该记的旁路保真缺陷**（见 §八 那条新行，属上游字体数据、非本仓解析器）：PingFang SC 的
  ToUnicode 反查表把 长/高 等少数汉字映成了康熙部首码位（`<1CE9> <2ED3>`、`<1E5B> <2FBC>`），**字形画面对**（走 CID→字形，与 ToUnicode 无关）、
  只有文本抽取/复制/无障碍那一层拿到部首字；拉丁行也读回残缺。
- `[x]` A11 PDF 改动必须走**增量更新**：改一处正文后**原文件前缀字节逐字不变**、只追加 delta（对照：全量 `save()` 把 163 205 字节重写成 790 601 字节，判为不合格）。**实测（四轮，含撤销那一次）**：W1 的 1 657 / 1 644 字节、W2 第三刀的 3 344、真端点那轮的 3 342 与撤销再追加 3 343（累计 6 685，该行读回原句）——**撤销同样只追加，不做全量重写**。

### B. AI 行为（对应 §2.8/§2.11）
- `[ ]` B1 模型只交引用与意图，数字由引擎算（图表、页码、坐标）。
- `[ ]` B2 模型交原始 XML / XPath 一律拒（不抄 OfficeCLI 的 `raw-set`）。
- `[ ]` B3 半自动必须点确认才落地；全自动仍受撤销约束。
- `[x]` B4 送模型的幻灯片内容按 §2.9 截断（形状数、每形状文本 ≤120 字），不许整包上送。
  **实测**：`pptx-io/analysis.ts` 的 `describeDeck`/`readSlideText` 各带页数、每页形状数、每段 run 数与单条 run 文字 120 字的上限，
  被截掉的部分置 `truncated` 让模型接着调 `readSlideText`；`slideAgent.test.ts` 用 25 页的模型断言上送里写的是"只描述了前 20 页"。
- `[ ]` B5 真端点探针：模型会点对 `ppt` 工具名与形状 id；结论与原始 tool_calls 记进文档（探针在 `tmp/`，不进主干）。

### C. 可撤销性与历史（对应 §2.5）
- `[x]` C1 每类 op 有逆操作单测（文字/字体/颜色/几何四条：`pptx-io` 的 `apply`/`revert` 用例 + `pptxHandlers.test.ts` 的「撤销回到标题甲」）；
- `[ ]` C2 内容寻址用 `creationId`，位置漂移不失效（现在仍按 `cNvPr @id` 定位，D5 的 GUID 那条未做）；
- `[x]` C3 撤销不牵连用户后来的手工编辑；C4 做不到时如实报"部分撤销"。**实测（本批）**：`revertSlideEdits` 逐条比现场与 `op.afterXml` 逐字一致才写回前像，
  那一页后来又被改过就整条跳过并给理由（引擎单测钉住"用户自己改的"一个字不丢），出入口用例钉"拿旧批次撤不动、磁盘一个字不动、最新那批照旧能撤"；
  `pptxHandlers` 成功回包带 `skipped`，`App.tsx` 据此 toast「部分撤销（x/y）：理由」。撤不动时宁可整条不落，绝不猜位置。
- `[x]` C5 复用同一份 modlog，无第二套历史栈。**实测（本批）**：幻灯片记录经 `slideHistory.ts` 写进 `electron/modlog.ts` 那四个通道（同一份 `userData/modlog/<文档名>.json`），
  记录判定靠 op 的 `beforeXml`/`afterXml` 签名，与文档、表格两类记录混放互不认错（`slideHistory.test.ts` 6 条）；
  真 app 里杀掉进程重开后撤销按钮仍可点、撤销成功、记录转 `undone:true`（阶段 4 的实测段）。

### D. UI 与工程（对应 §3/§4/§5）
- `[x]` D1 `App.tsx` 的 `isSheet` 已换成模式枚举（`EditorMode`：九处文档门控与 `isSheet || isDeck` 全改），无新增布尔门控；三种模式在真 app 里各跑一遍确认没改坏表格与文档视图（阶段 3 的实测段）。
- `[ ]` D2 无第二套 OPC / 解析 / 撤销 / 工具 schema / 样式体系（§3 红线）。本批新增的那条**没长第二套**：
  幻灯片提案卡与表格那张收成同一张 `ToolPlanCard`（落地出口是 `ToolLand` 的三个分支，卡片装配、暂停续跑、聊天模式不动文件都只剩一份），
  幻灯片落地与手工「保存」共用 `App.tsx` 的同一个 `commitSlideEdits` 口、撤销凭据也只有一份。仍开着的是 zip 目录项那份重复（阶段 3 的"发现的既有缺陷"）与 `docx-io` 的三处单位换算（A3）。
- `[x]` D3 新函数全部有中文 JSDoc（含单位与失败语义）；命名有含义（前两批：`pendingEdits`/`draftKeyOf`/`lockedParagraphHint`/`paragraphRunCounts`；本批：`emuToPx`/`pxToEmu`/`halfPointToPt`/`ptToHalfPoint`/`boxDraftKeyOf`/`pushGeometryEdits`/`changedBoxFields`/`readRunFormat`/`readBoxPx`/`patchBox`/`numberOrUndefined`）。
- `[x]` D4 渲染层不碰 Node，跨界只走 preload 白名单（`SlideDeck` 只调 `window.officeCopilot.editPptx/revertPptx`；幻灯片原文从不过进程边界）。

### E. 测试与门禁（对应 §7）
- `[x]` E1 每条写入路径有单测（本批：`slideDeckModel.test.ts` 4 条钉住「哪些段落会被写进文件」，视图本身不装配落地逻辑），**逐包核对通过数**（`packages/shared` 无测试是已知假绿源）。
- `[x]` E2 `pnpm build`（含 `tsc --noEmit`）+ `pnpm test` 重定向后读退出码，不许管道 tail。本批两次都是 `EXIT=0`，且 build 真抓到一处联合类型窄化缺陷（`slideDeckModel.test.ts` 里 `edit.target.paraIndex`）。
  阶段 5 这批的门禁：`pnpm build` 与 `pnpm test` 都 `EXIT=0`，逐包通过数 `opc 15 / sheet-tools 12 / pptx-tools 15 / pptx-io 37 / word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / cli 5 / desktop 478`（合计 1119；`packages/shared` 仍无测试文件，是已知假绿源）。
  阶段 4 的 modlog 落盘这批：`pnpm build`（含 `tsc --noEmit`）与 `pnpm test` 都 `EXIT=0`，逐包 `opc 15 / sheet-tools 12 / pptx-tools 15 / **pptx-io 41**（+4 条现场核对）/ word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / cli 5 / **desktop 485**（+6 条记录层 +1 条出入口）`，合计 **1130**。
  叠放次序这批：`pnpm build` 与 `pnpm test` 都 `EXIT=0`，逐包 `opc 15 / sheet-tools 12 / **pptx-tools 16**（+1 条方向枚举）/ **pptx-io 45**（+4 条重排与逐字回原）/ word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / cli 5 / **desktop 487**（+1 条装配端判端 +1 条出入口）`，合计 **1137**。
  视图上那四个层级按钮这批：`pnpm build` 与 `pnpm test` 都 `EXIT=0`，逐包 `opc 15 / sheet-tools 12 / pptx-tools 16 / pptx-io 45 / word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / cli 5 / **desktop 491**（`layerEdgeNote` 从 `slideAgent` 抽进 `slideDeckModel` 后补 4 条判端用例：中间四向都能挪、最顶只给向下两头、最底反向、单形状那页四向全到头）`，合计 **1141**。
- `[x]` E3 视觉验收走 Keynote PNG 导出通道并报告看到的文字内容：`keynote-out.001.png` 第一页居中黑字「Keynote 要看到的第一页标题」+ 下方灰色「正文一行」，`002.png` 第二页「第二页标题」原样（阶段 3 实测段）。
- `[ ]` E4 探针与夹具产物只进 `tmp/` 并清理；禁止 `git add -A`。

---

## 七、风险与未决

1. **工作量**：pptx 部件树规模与继承链是 xlsx 的数倍（D1 已如实写）。dcc 已裁"不许收窄"，所以**只能靠排队消化，不能砍项**。
   排队期间的纪律：**未建模的部件一律原样带过，不许半建模**——半建模比不建模更危险（会静默丢用户设置，
   富模板那 65 个 `ppt/tags/tagN.xml` 就是例子：重新生成型的编辑器会把它们全删掉）。
2. **Keynote 不是 PowerPoint**：它能开、能导出，不等于 PowerPoint 也认。凡涉及 `p:extLst` 扩展、`p18c` 现代批注，
   本机无法取证 → 相关条目一律 `[!] BLOCKED`，不许推断。
3. **动画/切换的播放效果**：建模要做（W3），但**画面验收无通道**（genoffice 与 OfficeCLI 自己都标注"需在别的查看器里验证"）
   → 判据停在"结构合法 + 独立读取器读回 timing 树一致 + 增量存回不丢部件"。
4. **PDF 正文改写的三处硬约束已取证但没消解**（D8）：① CID 子集里可能没有目标字形（要重子集化）；
   ② 增量更新的盲区（加密 `7.6`、线性化 Annex F、PDF/A 与数字签名——任何全量重写都会让签名失效）；
   ③ `@cantoo/pdf-lib` 的 `markRefForSave` 是**未写进文档的必需步骤**，漏了就静默不生效。
   这三条各配一条单测与一条"如实拒绝"的文案，不许用"看起来改了"糊过去。
5. **我们自产的 PDF 中文已经是坏的**（D8 首段，实测 FreeType 拒绝字体程序、Apple 渲染成原始字形）——
   这是既有缺陷，不是新线引入的，**排在 pdf 编辑之前修**（验收 A10）。
6. **已决**：D6 工具面每格式一个包（照 `word-tools`/`sheet-tools` 先例）→ 新建 `packages/pptx-tools`，不合并成 `packages/agent-tools`、不塞进 `packages/shared`（见阶段 5）。

---

## 八、总验收表（随进度填写，键为提交哈希）

| 提交 | 阶段 | 命令与输出（逐包通过数） | 视觉/第三方读回结论 | 状态 |
| --- | --- | --- | --- | --- |
| `13e62e2` + 本次实测 | 0 | 无产品代码改动（纯取证）；`pnpm build`/`pnpm test` 未跑（未动 `src`） | 三类夹具 python-pptx 1.0.2 全读通（35/125/49 part）；Keynote 14.2 打开并导出 PNG，**柱状图与表格肉眼可见**（附录 A2） | `[x]` |
| `7e271ff` | 3（视图 + 打开分派） | `pnpm build` EXIT=0（含 `tsc --noEmit`，抓到一处联合类型窄化缺陷）；`pnpm test` EXIT=0，逐包：pptx-io 27 / opc 15 / desktop 462 / sheet-io 245 / docx-io 296 / pptx-tools 15 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1093 条 / 76 个文件**（`packages/shared` 仍无测试文件，已知假绿源） | `oc ppt info` 读磁盘：第一页是新标题、第二页原样；`oc ppt check` 退出码 0；`.bak` 生成；撤销后磁盘回原字。**Keynote 导出两页 PNG 肉眼看过**：第一页居中黑字「Keynote 要看到的第一页标题」+ 灰色「正文一行」，第二页「第二页标题」；`tmp/multi-run.pptx` 在视图里 0 个输入框（§2.3） | `[~]` 文字那条闭环，字号/颜色/几何的入口仍缺 |
| `fb5b4f9` | 3（字号 / 颜色 / 几何入口 + D4 换算单点） | `pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包：pptx-io **30** / opc 15 / desktop **466** / sheet-io 245 / docx-io 296 / pptx-tools 15 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1300 条**（`packages/shared` 仍无测试文件） | 真 app（`dist` + CDP）打开 python-pptx 现生成的两页 deck：继承几何的占位符**没有**几何输入框、三 run 段落 0 个输入框、文本框四数 `96/384/288/96` 逐 px 对上；改文字+字号 32+`#00cc00`+左 150+宽 300 → 「4 处待保存」→ 保存 toast「动了 1 张幻灯片」→ **python-pptx 独立读回** `150/192/300/96`px、`32.0pt`、`rgb=00CC00`、`bold=None`（没连带写粗体），第 1 页三 run「正文丙/粗/细 18pt」原样 → `oc ppt check` 退出码 0 → 撤销后 40 个部件**内容不同的 0 个**。**画面效果仍 BLOCKED**（Keynote 只给文字与位置的 PNG，字体渲染未比对） | `[x]` 四类改动（文字/字号/颜色/几何）在真 app 里闭环；画布拖拽与继承链解析归入下批 |
| 本次（D3 几何继承链） | 4（`inherit.ts` + 读模型两字段 + 视图那句只读说明） | `pnpm build` EXIT=0（含 `tsc --noEmit`，抓到一处 `key` 窄化与三处夹具字面量缺字段）；`pnpm test` EXIT=0，逐包：pptx-io **74**（+16 条：逐轴读取 4 + 读模型的继承链 7 + 解析器的类型映射与容忍 3 + 落盘面 2）/ opc 15 / desktop **503**（+5 条 `inheritedBoxNote`）/ sheet-io 245 / docx-io 296 / pptx-tools 20 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1186 条**（`packages/shared` 仍无测试文件） | 新夹具三层链上 **python-pptx 独立读回三格逐轴与我们 `effectiveBoxPx×9525` 全等**；真实 38 部件 deck 同样全等（仅 `top` 差一个亚像素，是 D4 的显示取整）；真 app（`dist`+CDP）标题行念出"继承自母版：左 96 · 上 48 · 宽 768 · 高 120 px"且**零输入框**，改文字保存后磁盘仍是 `<p:spPr/>`、模型重读后说明不变。**画面位置 BLOCKED**：Keynote 的 AppleScript 拿不到对象 `position`/`size`（`-1700`/`-1728`） | `[x]` D3 几何那一档落到读模型与视图；字号/颜色两档与 notes 归入下批 |
| 本次（pdf 线第一档：`packages/pdf-io` W1） | 7（读模型 + 同长度替换 + 增量更新） | `pnpm build` EXIT=0（`tsc` 抓到 `test-pdf.ts` 漏 `deflateSync` 导入——vitest 不做类型检查，所以只有 build 能抓到）；`pnpm test` EXIT=0，逐包：**pdf-io 13（新包）** / pptx-io 74 / opc 15 / desktop 503 / sheet-io 245 / docx-io 296 / pptx-tools 20 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1199 条**（`packages/shared` 仍无测试文件） | 两份真实文件读回：WPS 简历 6 页 **5196 条 run 全解出**、`printToPDF` 产物 1 页 1113 条（932 条有文字，缺 ToUnicode 那 181 条如实给 `null` 而不是猜）。改→存→读回：`ps-cols-OLD.edited.pdf` **前缀 33 706 字节逐字不变 + delta 1 657 字节**，Apple **PDFKit** 抽取到"第 1 是这：段"且第二栏文字一字未动，**QuickLook 出的 PNG 肉眼核对**只有那四个 CID 字形换位、无豆腐块；自造 Helvetica 两页文件同样只改第一页。**依赖只有 `@cantoo/pdf-lib`（MIT），无原生模块** | `[~]` W1 那一档闭环（同长度、同框、同字体）；真实中文产物一条 run 一个字，"改一句话"要先拼逻辑行 → 归 W2 前置 |
| 本次（pdf 线 W1 的撤销那一半） | 7（`revertPdfTextEdits` + `text-drifted` + 无 `/ToUnicode` 夹具） | `pnpm build` EXIT=0（`tsc` 抓到一处 `run.text` 可能为 null 就写进 `fromText` 的缺陷——不是类型洁癖：那条 run 我们**不知道原文念作什么**，现在改成 `no-encoding` 整批拒绝）；`pnpm test` EXIT=0，逐包：**pdf-io 17（+4：撤销回原文且那次也只追加 delta 1、别处后来的编辑一字不动、对不上内容给 `text-drifted` 且不产出字节、缺 ToUnicode 的读+拒）** / pptx-io 74 / opc 15 / desktop 503 / sheet-io 245 / docx-io 296 / pptx-tools 20 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1 203 条** | 真文件（`tmp/ps-cols-OLD.pdf`）改→撤→读回：撤销那一次同样**前缀 35 363 字节逐字不变 + 再追加 1 644 字节**，新读模型里旧句"段：这是"回来了、改过的"是这：段"一条不剩，**Apple PDFKit 抽取第一句就是原文"第 1 段：这是一段足够长的中文正文"**；再拿一条对不上内容的 `applied` 去撤，只回 `text-drifted` 且不产出字节 | `[~]` §2.5 那一半在引擎侧闭环；modlog 落盘与真 app 通道归下一批 |
| 本次（pdf 线的桌面装配那一半） | 13（`electron/pdfHandlers` 6 + `pdf/pdfRuns` 4 + `pdf/pdfHistory` 3）+ 四条通道进 preload 白名单 | `pnpm build` EXIT=0（`tsc` 抓到两处：测试里拿 `PdfTextRun` 当定位键用（`streamRef` 可为 null）、以及一处 JSX 少括号）；`pnpm test` EXIT=0，逐包：pdf-io 17 / pptx-io 74 / opc 15 / desktop **516**（+13）/ sheet-io 245 / docx-io 296 / pptx-tools 20 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1 216 条**（`packages/shared` 仍无测试文件） | **真 app（`dist` + CDP）跑第三方文件**（`~/Downloads/Modern Resume-edited.pdf`，一份 `printToPDF` 简历）：面板挂上后念"显示文字 14 段，可改 14 段"、每行带 `/BaseFont` 与"原位只有 N 字节"→ 把 "Jane Q. Doe" 改成同长度换位 "ane Q. DoeJ"（**新文字只用这条 run 自己出现过的字符**，所以字体子集必然收得下）→ 保存 toast「追加 402 字节增量，原来的字节没重排」→ 磁盘实测**前缀逐字节不变**、新读模型里新文字在 / 旧文字一条不剩 → `userData/modlog/<那份>.json` 出现一条「PDF 改动」（`applied` 带旧新文字）→ 撤销按钮随之可用，点下去旧文字回到文件里、那条记录 `undone:true`、文件仍是只追加。**截图留证** `tmp/pdf-ui/panel.png`。**AI 面板在 pdf 模式下仍是 docx 引导语**（工具面未开工，等 `pdf-tools` 一起改） | `[x]` 阶段 7 的"打开→改→存→撤"真 app 通道闭环；`oc pdf check`、AI 工具面、逻辑行拼接（W2 前置）归入后续 |
| 本次（pdf 线的 AI 工具面，含顺手修掉的 §5.2 白屏缺陷） | 32 新条（`pdf-io/src/analysis.ts` 的 AI 读面 8 + 新包 `packages/pdf-tools` 10 + `src/ai/pdfAgent.test.ts` 10 + `src/pdf/rendererIoEntry.test.ts` 1；`pdf-io` 现 28 = 读写 20 + 读面 8） | `pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包：**pdf-io 28** / **pdf-tools 10（新包）** / pptx-io 74 / pptx-tools 20 / opc 15 / desktop **527** / sheet-io 245 / docx-io 296 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1 248 条**（`packages/shared` 仍无测试文件）。提交分六刀：`6a7e403` 读面 → `823fc8f` 工具包 → `26dca36` 入参判定去重 → `f99a4a6` 代理 → `1f7ace9` 装配 → `1bbb4fa` 白屏修复 | **真 app（`dist` + CDP，同一份第三方 `printToPDF` 简历）**：① 面板挂上（14 段）＋ PDF 引导语 ＋ PDF 那句输入提示；② docx 才有的两个入口（快捷指令开关、自动排版）在 pdf 模式下回 `false`——它们改的是 Tiptap 编辑器，对着 PDF 点了会误导；③ 模式说明是"…点「应用到 PDF」才写文件，之后可在 PDF 面板点「撤销上一次保存」回退"；④ 发一句话后卡头 **AI PDF 提案**、正文「还没有配置模型。点右上角⚙️…」、状态 **未写入**。**探针 profile 没配 Key，所以这轮证的是装配面**（工具回合卡、needsSetup 分支、落地共用 `commitPdfEdits`→`pdf:edit`）；模型会不会点对 `pdf_*` 工具名与页内 runIndex **仍待真端点探针**，与 B5 同口径。**缺陷实跑记录**：第一次跑连①都过不去（"等不到 PDF 面板挂载"、`#root` 空子），`Runtime.exceptionThrown` 抓到 `ReferenceError: require is not defined`、bundle 里有 `require("zlib")`——渲染层 import 了 `pdf-io` 主入口，而 `pnpm build` 与 `pnpm test` 一路全绿。拆出 `@office-copilot/pdf-io/pure`（只含类型、判定与 AI 读面）后 bundle 0 命中、面板挂上，并补 `rendererIoEntry.test.ts` 作机检 | `[x]` 阶段 5 的 pdf 那一半闭环：模型只交引用与页号/runIndex、引擎坐标不上送、只提案不落地（§2.8）、失败条目逐条给中文理由（§2.4 整批语义不被拖坏）；`[!]` 模型行为未验（本机无端点），`oc pdf check`、`print:pdf` 旁路、逻辑行拼接（W2 前置）仍缺 |
| 本次（pdf 线 AI 工具面的**真端点**实跑，把上一行那个 `[!]` 消掉一半） | 0（只是取证，不改代码） | 不动主干，`pnpm build`/`pnpm test` 与上一行同一份数字 | 用**用户本人的服务商**（`userData/@office-copilot/desktop/config.json` 里那两条，Key 只复制进 `tmp/pdf-ai-real/profile*/`，绝不入库，§7.3）在真 app（`dist` + CDP）上跑两份：**① 改署名**「把 Jane Q. Doe 的名和姓调个顺序」→ 卡片 **1 处改动**、正文自己写出"保持 11 个字符长度不变"、落地 toast「追加 402 字节增量」→ 磁盘**前缀逐字不变**、`Doe Q. Jane` 在、`Jane Q. Doe` 一条不剩、其余 13 条没动、modlog 一条「PDF 改动」。**② 改职位行**（目标那条在页内第 **6** 条，不是第一条，靠位置猜不出来）「把 Corp 改成同长度的 Comp」→ 卡片**第 6 条**改对、`Senior Engineer – Example Comp` 落盘、**前缀未动 + 追加 2 439 字节**、14 条齐全。**③ 该拒绝的**「SUMMARY 改 32 磅红色」→ 没有提案卡，回答直说"做不了"，并点名 `set_run_text` 只收等长整条替换、还引用了引擎给的 charCount。**观测不到的那半如实说**：想数"模型到底调了几轮 `pdf_describe`/`pdf_read_page_text`"，办法是在渲染层包一层 `window.officeCopilot.aiToolTurn` 计数器——**包不上**（contextBridge 交出去的对象在页面主世界里只读，赋值静默无效，`__turns` 恒为 `[]`），这恰好是 §5.2 那层保护在起作用；而 `pdf_describe` 的每页概要是**按 run 用 `⏎` 拼起来**的，所以模型数得出第 6 条，**不能**据坐标对就断言它一定调过 `pdf_read_page_text` | `[~]` 模型行为这一半有了实测：坐标对、等长约束守住、该拒的拒；`[!]` 工具调用轮次本机不可观测（不是缺陷，是没有通道），要看得接一条主进程侧的回合计数 |
| 本次（pdf 线 W2 第一刀：内容流跟住文字原点与生效字号） | 15（新文件 `packages/pdf-io/src/content-stream.test.ts`；`TextRun` 加 `origin`/`fontSizePt`，`PdfTextRun` 同步带出） | `pnpm build` EXIT=0（`tsc` 抓到 5 处 `noUncheckedIndexedAccess`：操作数数组按下标取值要判 `undefined`，不许拿 0 顶）；`pnpm test` EXIT=0，逐包：pdf-io **43**（+15）/ desktop 527 / docx-io 296 / sheet-io 245 / pptx-io 74 / pptx-tools 20 / opc 15 / pdf-tools 10 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1 263 条** | **两份真文件实测**（`tmp/pdf-w2/probe.mjs`，跑 `dist`）：**① Word/Google 那份** 14 条 run，x 恒 40.0、y 740→420 递减，字号 28/14/12/11/10/9，`origin=(0,0)` 的组 0 个。**② WPS 那份 6 页 5196 条**：解压后的原文是 `q 1 0 0 -1 0 841.9 cm … BT /FT8 439 Tf … 0.05 0 0 -0.05 253.559 114.89 Tm <1D2C>Tj 441.594 -0 TD <1DBC>Tj …`——**逐字形发一条 `TD`**，所以每个原点都不同，手工核对第一字形 (253.559, 841.9−114.89=727.01) 与读模型 (253.6, 727.0) 一致；**字号必须乘过 Tm 的 0.05** 才是 21.95/15.95/10.45/12（照抄字面 439 就是给下一个读者挖坑，这条已写进字段注释与用例）。**第二刀的判据也是这次量出来的**：同一页里**有 19 行在同一个 y 上 x 断开 >25pt**（最大 171pt；左栏 x=90…142、右栏 x=289…341 都在 y=615.0），所以**只按 y 分组必然串栏**，拼行必须再按 x 断句。算子语义逐条对过 ISO 32000-1 §9.4.4 与 pdf.js `canvas.js`：`T*` 是 `0 -Tl Td`（第一版写成 `+leading` 会让行序倒过来）、`TD` 把 leading 设成 `-ty`、`"` 那两个操作数是**词/字间距而不是 leading**（第一版误当 leading 用了） | `[x]` W2 第一刀闭环：原点与生效字号进了读模型、过了两份真文件；`[ ]` 逻辑文本行拼接（#21）仍缺，`oc pdf check` 与 `print:pdf` 旁路模型仍缺 |
| 本次（pdf 线 W2 第二刀：把逐字形的 run 拼成逻辑文本行） | 12（新文件 `packages/pdf-model.test.ts` 10 条 + `analysis.test.ts` +2） | `pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包：pdf-io **55**（+12：pdf-file 20 / analysis 10 / content-stream 15 / pdf-model 10）/ desktop 527 / docx-io 296 / sheet-io 245 / pptx-io 74 / pptx-tools 20 / opc 15 / pdf-tools 10 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1 275 条** | **阈值是量出来的，不是拍的**（`tmp/pdf-w2/gaps.mjs`，同一坐标层内先按 y 归行再量 x 间隔）：WPS 那份 5 196 条 run 的**间隔/字号** p50=**1.00em**、p90=**1.01em**，而 p99 已经是 **15.04em**（最大那几个 94/136/147/157pt 全是右栏起点）；那份两栏合成夹具同一量法 p90=1.33em、p99=17.0em。**行内 ≤1.4em、断栏 ≥10em，中间是一条空的带**，所以取 4 倍生效字号断句。**为什么不是固定 pt**：这份文件里行内最大间隔是 21pt、栏间最小 94pt，25pt 这种常数恰好夹在中间、蒙得住这两份，但它按字号说话就不成立——标题「个人简历」每字进 22pt，正文只有 10.5pt，同一份文件里两种尺度混在一起，字号再大一号的标题就会被固定阈值从中间劈开；按 em 归一才与字号无关（用例照这三档各造了一条：22pt 标题里 22pt 的推进是 1em 仍是一行、10.45pt 正文里 30pt 仍合并、94pt 才断，`pdf-model.test.ts`）。**y 容差 0.6pt 同样是量出来的**：WPS 那份里同一条视觉行内相邻字形的 y 最多差 **0.119pt**（`TD` 累加的抖动，326.331 与 326.45 那对确实属于同一行、40 段拼出来是一整句），而真是两条行时最近的两行差 **15.12pt**，0.6pt 正好落在这两条之间。**真文件核对**（`tmp/pdf-w2/{lines,outline}.mjs`，跑 `dist`）：① WPS 简历 5 196 条 → **206 行**（平均 25.2 run/行），206 行**全部**由多条 run 拼成、**0 行**只有一条 run（= 没有"本该同行却被拆开"的单字行），第 1 页 611 条 → 29 行，左栏 x 90…142.4 与右栏 289.2…388.5 在同一个 y=615.0 上**没被粘成一行**；模型可见的那一面（`describePdf`）现在上送「姓名：丁长长」「出生日期：1994.01.16」「毕业院校：安庆师范大学」这样的整行，不再是 611 个单字。② Word/`printToPDF` 那份 14 条 → 14 行，一条不多合并（那本来就是"一行一条 run"的产物）。③ 两栏合成夹具 1 112 条 → 164 行，栏间 178.5pt 处照规则断开了。**规矩与限制如实记**：分组**只改怎么断行、不改文字本身**（行拼接必须等于 run 拼接，用例里钉死），段间**不插空格**（中文会多出一个看不见的字）；旋转文字没有 `Tm` 的旋转角可用，只能按 y 判，会**过度切分**（如实记，未修）；没有字形宽度，`lastOriginX` 是"最后一个字形的起点"而不是行的右边界；那份合成夹具里**认不出原文的那条 run**（数字用的是没 `/ToUnicode` 的子集）起点比邻居跳了 10em，于是它把一行切成了两段——这是 W2 第三刀之前唯一观察到的过度切分，只在缺字那类 run 上出现 | `[x]` #21 闭环：逻辑文本行进了引擎、AI 概要按行读、两份真文件过；`[!]` 真端点那半这次**没重跑**（改的只是概要与视图读到的文字怎么断行，工具入参、等长判定、落地通道一行没动，上一行那三份实测仍然成立）；`[ ]` 把逻辑行接进 AI 的**编辑**面（提交仍是页号 + run 序号，没被行号替换，这是刻意的）；`oc pdf check` 与 `print:pdf` 旁路模型仍缺 |
| 本次（pdf 线 W2 第三刀：把逻辑行接进 AI 的编辑面） | 14（`analysis.test.ts` +6、`pdf-file.test.ts` +1、新夹具 `cidPerGlyphFixturePdf()`、`pdf-tools/src/index.test.ts` +3、`pdfAgent.test.ts` +4） | `pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包：**pdf-io 62**（+7：pdf-file 21 / analysis 16 / content-stream 15 / pdf-model 10）/ **pdf-tools 13**（+3）/ **desktop 531**（+4）/ docx-io 296 / sheet-io 245 / pptx-io 74 / pptx-tools 20 / opc 15 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / server 2 = **1 289 条**（`packages/shared` 仍无测试文件） | **展开住在引擎里**（`pdfLineEditsOf`），代理只翻坐标：新文字按该行**每条 run 原有的字数**依次切片，中文那行六个字回六条各一字、Word 那行回一条整行，`streamRef`/`withinForm`/`start` 一律由引擎填（§2.1 换算单点、§3.2 不许有第二份字数账）。拦在代理层的四种情形各回一句中文理由、一条都不送去写：行号越界（回"这一页只有 M 行"）、字数不等（回"原文是 T 个字，新文字是 C 个字"）、与原文一样、行内有一条改不了（**整行不动**，缺一条就拼不回这句话）。卡片头部条数按**模型提交的那一批**算（`units`，与表格侧 D8 同一条判据），一条整行显示"1 处改动 · 第 N 页 · 第 K 行改为「…」"。**为什么新开一份夹具**：四个 CID 挤在一条 `Tj` 里时"一行 = 一条 run"，切片没有对象可切，整行落盘能不能保住前缀字节就测不到——`cidPerGlyphFixturePdf()` 把同样的四个 CID 写成**四条 `Tj` + 每字一条 `8 0 Td`**（WPS 的真实形状），那条用例读回的是"引擎拆的六条 == 落盘的四条 == 读回的顺序"。**WPS 那份简历实测**（`tmp/pdf-w3/probe.mjs`，跑 `dist`）：概要里「姓名：丁长长」= 第 1 页第 2 行、由 run 10…15 共 6 条拼成、`editableLineCount=29/29` → 一次 `set_line_text` → 引擎回 6 条等长替换 → **前缀 334 523 字节逐字不变、只追加 3 344 字节** → 读回该行是新文字、行外 605 条 run 一字未动、总 run 数 5196 不变 → 再按内容寻址撤销这一次整行改动，**撤销那一次同样只追加 delta**，读回原句「姓名：丁长长」 | `[x]` #25 闭环：整行那一档进了引擎、工具面、代理与卡片，真文件改完能撤；`[!]` **真端点那半本轮未跑**——`set_line_text` 是新条目、系统提示词刚改（"改一句话一律报行号"），模型会不会选行号而不是发几十条 `set_run_text` 属于模型行为，须按 B5 口径实跑才算数，本轮只有单元/夹具/离线探针三级证据；`oc pdf check` 与 `print:pdf` 旁路模型仍缺 |
| 本次（W1.5 的**真端点**实跑，把上一行那个 `[!]` 消掉） | 0（只是取证，不改代码） | 不动主干，`pnpm build`/`pnpm test` 与上一行同一份数字 | 真 app（`dist` + CDP，`tmp/pdf-w3/app-probe.mjs`；Key 只复制进 `tmp/pdf-w3/profile/config.json`，绝不入库 §7.3）对着**那份 WPS 简历**（6 页 5 196 run、第 1 页 611 条）问：「第 1 页那行「姓名：丁长长」，把名字那三个字倒过来写成「长长丁」，整行其余不动」。**模型提交的是 1 条 `set_line_text`、`lineIndex=2`**（卡片头部「1 处改动」、清单只有一行「第 1 页 · 第 2 行改为「姓名：长长丁」」），**没有**退化成 6 条 `set_run_text`，更没有对着 611 个单字发几百条 → 落地 toast「已写入 6 处文字：追加 3 342 字节增量，原来的字节没重排」→ 磁盘**前缀逐字不变**、该行读回「姓名：长长丁」且仍由 6 条 run 拼成、总 run 数 5 196 不变 → `userData/modlog` 一条「PDF 改动」（`applied` 6 条带旧文字）→ 面板「撤销上一次保存」再**追加 3 343 字节**（累计 6 685），该行读回原句「姓名：丁长长」。截图留证 `tmp/pdf-w3/app.png`。**两处如实记**：① 模型正文把它念成"第 1 页第 3 行"（它按 1 基数数了行号），但**送进工具的 `lineIndex` 是 2**、与引擎行序一致——落点以入参为准，那句散文不进引擎；② 追加字节与上一行离线探针的 3 344 不是同一次产物，判据只看"前缀逐字不变 + 只追加"，两轮都过。**顺带一条环境事实**：`--user-data-dir` 会带着 `app.getPath("userData")` 一起走（modlog 落在 `tmp/pdf-w3/profile/modlog/`），AGENTS.md §8 那句"只是 Chromium 开关"只对**单实例锁**成立 | `[x]` 模型行为那一半实测闭环：改一句话就报行号、只提案不落地、撤销同样只追加；`oc pdf check` 与 `print:pdf` 旁路模型仍缺 |
| 本次（pdf 线的结构体检：引擎 `auditPdfStructure` + CLI `oc pdf check`） | 9（新文件 `packages/pdf-io/src/check.test.ts` 8 + `apps/cli/src/pdf.test.ts` 1；另有新文件 `check.ts`、`streams.ts`（从 `pdf-file.ts` 抽出，不是新逻辑）与 `apps/cli/src/pdf.ts`） | `pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包：**pdf-io 70**（+8）/ **cli 6**（+1）/ desktop 531 / docx-io 296 / sheet-io 245 / pptx-io 74 / pptx-tools 20 / opc 15 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / server 2 = **1 298 条**（`packages/shared` 仍无测试文件） | **判据按 ISO 32000-1 的显式硬约束**：§7.5.2（文件头 `%PDF-`、文件尾 `%%EOF`）、§7.7.3.4（`/Pages` 的 `/Count` 必须等于 `/Kids` 项数、也必须等于解析出的页数）、§7.3.5（`/Filter` 解不出来就是坏页），再加一条 D8 自己的"改不了"档（`text-not-editable`：字体没带 `/ToUnicode` 的那些 run）。**引擎只出码与数字、中文文案全在 CLI**（§5.7），键类型是 `Record<PdfAuditIssue["code"], string>`，新增码忘配文案会编译不过。**真文件三份实测**（`tmp/pdf-check/`，跑 `dist`）：① 那份 WPS 简历（6 页 5 196 run）**0 条**、退出码 0——**不误报是这条体检的第一判据**；② `printToPDF` 产物如实报一条黄色"1 113 条显示文字里 181 条认不出内容（字体没带 `/ToUnicode`）"、退出码 1；③ 砍掉尾巴的文件报红色 `eof-missing` + 红色 `catalog-missing`。**坏夹具靠等长字节手术造**（`patchSameLength`：改 `/Count` 那位数字、把第一段 `/FlateDecode` 流正文填 0），因为 PDF 的交叉引用按**字节偏移**寻址，任何长度变化都会连带一堆假失败——这与 `.docx`/`.pptx`/`.xlsx` 那三份审计完全不同（它们是 zip，改 XML 不用管偏移）。**抓出并如实记一条实测事实**：截断到只剩一节的文件 `PDFDocument.load` **照样成功**，但 `doc.catalog` 是 `undefined`（`@cantoo/pdf-lib` 2.11.1 的只读属性，且它没有 `getContext()`），第一版体检因此抛 `Cannot read properties of undefined`——这就是"载得动不等于目录在"，于是新增 `catalog-missing` 码并在它之后**早退**（页树无从定位，后面的检查全是瞎猜）。**§3.2 的那一刀**：流解码逻辑（`decodeStream`/`streamOf`/`plainName`/`contentKeys`）本来长在 `pdf-file.ts`，体检要判同一件事就是第二次出现，抽进 `streams.ts` 供读侧与体检侧共用，**没有第二份 inflate**；`streams.ts` 刻意不从 `index.ts` 导出（内部件，不是公开面）。**没查的两条如实说**：`/Encrypt` 与 `/Sig` ——本机既没有 qpdf/mutool/gs（附录 A 实测全 none）也没有签过名的 PDF，**造不出夹具就等于写不出用例**，与其留一条永远不会亮的分支，不如把这两档记进 W4（D8 的加密与签名策略）再一起做。**长任务里的一处旧账**：`oc pdf` 之前 CLI 只有 `sheet`/`ppt` 两条只读命令，`package.json` 的 workspace 依赖要补 `@office-copilot/pdf-io`（已补并 `pnpm install`，锁文件随之变动） | `[x]` 阶段 8 的 `oc pdf check` 闭环：七码两档（`broken`/`limited`）、退出码可当闸门、三份真文件（含 0 误报）；`[ ]` `print:pdf` 旁路模型、把体检结论接进桌面打开时的提示仍缺 |
| 本次（AGENTS.md 收口：把四类格式的既有事实同步进规则文件） | 0（只改文档，不动代码） | 不动 `src`，`pnpm build`/`pnpm test` 与上一行同一份数字（1 298） | 逐条**先核再写**，没有一条是凭印象补的：§0 改成四种模式（`.docx`/`.xlsx`/`.pptx`/`.pdf`，并写明 PDF 没有回流）；§1 的 scope 枚举补上 `pptx-io`/`pptx-tools`/`pdf-io`/`pdf-tools`/`opc`/`sheet-io`/`sheet-tools`（这几天实际用过的 scope 之前一个都不在表里）；§2 新增 **2.12**（pptx 的 px↔EMU / pt↔半磅只在 IO 边，占位符几何必须按 D3 逐轴继承、不许烘焙）与 **2.13**（PDF 只许等长整条替换 + 只追加增量、缺 `/ToUnicode` 不许猜字、撤销与写入共用一份实现）；§2.4 改成"四类产物共用 `fsutil.writeFileAtomically`"（`grep -l` 实测 `fileHandlers`/`sheetHandlers`/`pptxHandlers`/`pdfHandlers` 四个入口都在）；§2.5 补四个 modlog 标签（「AI 修改」「表格改动」「幻灯片改动」「PDF 改动」）与 `slideHistory.test.ts`/`pdfHistory.test.ts` 两条机检；§2.9 补 PDF 那半（单页、默认 50 条、上限 200、超出标 `runsTruncated`——数字直接读 `pdf-tools` 的 zod，不是我拍的）；§3 的"禁止第二套"点名 `pdf-io/streams.ts` 是内容流解压唯一一份；§5.1 补三条依赖禁令 + 那条**只有实跑能看见**的白屏坑（渲染层 import 主入口 → `require("zlib")` 进浏览器包），并把 `rendererIoEntry.test.ts` 标成 `[机检]`；§7.4 的⑥ 补一句"`.pptx`/`.pdf` 也不是 BLOCKED"（`oc ppt check`/`oc pdf check` 当闸门、PDF 有 QuickLook + PDFKit 视觉通道、pptx 只有 Keynote 的文字与位置）；§8 新增三条环境事实（PDF 视觉通道 / 坏 PDF 只能等长改写 / `@cantoo/pdf-lib` 的 `doc.context` 与 `doc.catalog` 可为 undefined）并订正 `apps/cli` 的通过数（3 → 6）；§9 加 §2.12、§2.13 两行、把 §5.1/§5.2 那行升级成"已落地（含静态 import 检查）"。**同时把 §六 的 A8/A11 从 `[ ]` 改成 `[x]`、A9 改成 `[~]`**：A8 那句原文写的是"pdf.js 读回"，本仓从不引 pdf.js，如实改成"Apple PDFKit + 我们的读模型"并把两轮字节数写进去；A9 保持不满钩的原因是**签名与加密那两格本机产不出文件**（`qpdf`/`mutool`/`gs` 全 none），已落的是缺 ToUnicode / 纯图片页 / 装不回去 / 子集缺字 / 定位失配 / `text-drifted` 这六格 | `[x]` 阶段 8 的"AGENTS.md 四类格式事实"那一半闭环（本行只动文档，代码与上一行同一份）；`[ ]` §六 剩余项与 `print:pdf` 旁路另起一批 |
| 本次（把体检结论接进桌面打开时那句提示，顺手把两份装配层的中文抽成一份） | 4（新文件 `packages/shared/src/index.test.ts` 3 条 + `apps/desktop/electron/pdfHandlers.test.ts` +3；另有 `apps/cli/src/pdf.test.ts` 的 1 条跟着文案搬走，所以 cli 6 → 5） | `pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包：**shared 3**（第一次有测试文件，根命令 `passWithNoTests` 那个假绿源少一格）/ **desktop 534**（+3）/ **cli 5**（−1）/ docx-io 296 / sheet-io 245 / pptx-io 74 / pdf-io 70 / pptx-tools 20 / opc 15 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / server 2 = **1 303 条** | **文案抽进 `@office-copilot/shared`**：判定住在引擎（§5.7），可装配层有两个（`oc pdf check` 与桌面面板），"这一页的内容流解不出来"念两遍就是第二套（§3.2）。`pdf-io/src/check.ts` 改成从 shared 引类型、再按公开面原样导出，两边 import 路径都没动；依赖方向仍是 `pdf-io → shared`、`apps → packages`（§5.1），桌面为此新增一条 `@office-copilot/shared` workspace 依赖（`pnpm install`，锁文件 +3 行）。**`pptx` 那条没一起抽**：`describePptIssue` 今天只有 CLI 一个装配层，第二次出现才抽（§3.7）。**主进程里的顺序是反的**：先跑体检、再建读模型——实测截断的真文件 `readPdfModel` 抛 `Cannot read properties of undefined (reading 'Pages')`，那时只剩体检答得出"为什么打不开"，所以 `PdfOpenResult` 的失败分支也带 `audit`。真 app 四份形状各一轮（`tmp/pdf-w4/ui-probe2.mjs`，跑 `dist`，**没复制真 Key**：这轮一条 AI 请求都不发）：① WPS 那份 6 页简历——面板挂载、「显示文字 611 段，可改 611 段」、**没有提示条**；② `printToPDF` 产物——黄色一条「1 处能打开但有一部分改不了」+「1113 条里 181 条认不出内容（字体没带 /ToUnicode）」；③ 自家夹具砍一半（读模型 2 页 0 run，面板仍挂得上）——红色四条（文件尾缺 `%%EOF`、第 1/2 页内容流解不出来、整份读不到显示文字）；④ 只剩 `%PDF-1.7` 的文件——面板挂不上，toast 念「文件尾没有 %%EOF…；目录（/Root 或 /Pages）在文件里查不到…」，第二参 `withDetail=false` 把「结尾 16 字节读到的是 …」那串原文挡掉（一条 toast 放不下）。截图 `tmp/pdf-w4/{clean,partial,fixture-half,headeronly}.png`。**顺带如实记一条旧账**：`electron apps/desktop <文件>` 时 macOS 的 `open-file` 与冷启动 argv 两条路都送同一个路径，所以失败 toast 出现两遍——这是 `main.ts:36` 那条双分派本来的行为（`.docx` 一样），与本次改动无关，另起一批再收 | `[x]` D8 阶段 7 那句"把体检结论接进桌面打开时的提示"闭环：引擎结论 → 主进程回包 → 面板提示条 / 失败 toast 三处都看得见，全仓只剩一份中文；`[ ]` `print:pdf` 旁路模型仍缺 |
| 本次（`print:pdf` 通道的 A10 现跑判定 + 旁路读回保真取证） | 0（纯取证探针，不动产品代码；只在 `tmp/pdf-w5/`） | 不动 `src`，`pnpm build`/`pnpm test` 与上一行同一份数字（1 303） | **spec line 196 那条要求兑现**："下次动 `print:pdf` 通路时要现跑一遍再判 A10，不要沿用旧结论"。两步探针（`tmp/pdf-w5/print-main.cjs` 用与 `printHandlers.ts:renderPdf` 逐字相同的 `printToPDF` 选项产一份含中英的 A4 PDF；`judge.mjs` 经**我们自己的 `readPdfModel`** 读回 = 那条"旁路"；`inspect.mjs` 用 `@cantoo` 把 `/ToUnicode` CMap 原样解出）①**A10 判为不再适用**：生效字体名 `/AAAAAA+PingFangSC-Regular`，子集命名**合规 §9.6.4**，当年 `/SourceHanSansCN-Regular-6087` 那种违规名在当前 Chromium 打印通路复现不出（旧产物早不在 `tmp/`，无法再造）。名字合规正是 MuPDF/PDFKit 接受字体程序、中文正常渲染的前提（A10 原判据依赖的就是这一步），故 §六 A10 由 `[ ]` 改 `[x]`。②**翻出更该记的一条**（属上游字体数据、非本仓解析器）：字体结构是一份 `/Type0 /Identity-H` + 其 `/CIDFontType0` 后代，ToUnicode 只在 Type0 上（正常）；可 PingFang SC 的 ToUnicode 反查表把少数字映进了康熙部首码位——CMap 里字面写着 `<1CE9> <2ED3>`（长→⻓）、`<1E5B> <2FBC>`（高→⾼）、`<0AE1> <2F2F>`，我们的读模型**如实**回吐这三个部首字（不是解析 bug，两边证据都对得上）。**字形画面对**（CID→字形与 ToUnicode 无关），只有文本抽取/复制/无障碍那层拿到部首；同一份里拉丁整行也读回残缺（`The quick…` → `T … .`）。③**旁路本身是通的**：`printToPDF` 产物 1 页 59 run 能被 `pdf-io` 读出并定位，即阶段 7 那句"`packages/pdf-io` 读页/文本块/字体，`print:pdf` 通路带旁路模型"已具备，剩下的是 ToUnicode 保真。**如实标注**：字体名合规=实测；"渲染正常"是**推断**（本机没有 Word/能 OCR 的渲染通道，且 A10 破坏渲染的根因就是违规名，现已不存在）；部首映射=实测（CMap 原文 + 读模型回包双证）。 | `[x]` A10 按现通路关闭 + 把"旁路读回"的保真缺陷成文；`[ ]` 若要修旁路的部首/拉丁残缺，得改我们产 PDF 那侧的字体来源或 ToUnicode，属新立一档，另起一批 |
| 本次（三份 zip 序列化收进 `@office-copilot/opc` 一份，产物不再写目录项） | 2（`packages/opc/src/opc.test.ts` +2：新增部件后序列化**目录项 0 条**且文件部件清单与原文件逐条一致；没新增部件时清单一字不变，只少目录项） | `pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / desktop 534 / pptx-io 74 / pdf-io 70 / pptx-tools 20 / **opc 17（+2）** / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 3 / server 2 = **1 305 条**。文档同步订正三处旧账：README 表头 1 303 → 1 305（opc 15 → 17）、`desktop` 明细那句写的 527 → 实测 534、"packages/shared 目前没有测试文件"改成"当前唯一没用例的包是 `apps/task-pane`"（实测日志里它打印「No test files found, exiting with code 0」，这正是 `passWithNoTests` 那条假绿现在唯一的落点） | **任务原文那句"加 `createFolders:false`"被实测推翻**（三步取证，每次都看 `unzip -Z1` / `zipfile.namelist()` 而不是我们自己的读侧——JSZip 的读侧分不清"真目录项"与"列出来的目录"）：① `JSZip.loadAsync` **不合成**目录对象，一份没有目录项的输入读回再写照样没有；② `zip.file("ppt/slides/slide2.xml", …)` 会**顺手造出 `dir:true` 的父目录对象**，而 `generateAsync({createFolders:false})` **挡不住**——同一份加了这个开关仍写出 `ppt/`、`ppt/slides/` 两条，照原文写只是空转；③ 清目录项**不能**用 `zip.remove(name)`，它是递归删（实测把 `ppt/` 下所有真部件一起删光），只能 `delete zip.files[name]`。`serializePackage` 就是那一句 + `generateAsync`。为什么值得抽：三处 `toBytes` 各写一份 `generateAsync`（§3.2 第二次出现就抽），而 `.pptx` 新建幻灯片、`.xlsx` 新建图表、`.docx` 新建部件都会带出目录项；OPC（ECMA-376 Part 2 §3）的目录项**本来就是可选的**，办公软件按名字查部件，所以三种格式统一不写。用例落在 `opc` 而不是 `pptx-io`：夹具 `buildDeckBytes` 自己就带目录项，在格式包里断言"产物无目录项"会与夹具纠缠。**没有拿自证**：三份真文件各走 open→toBytes（一条都不编辑）——`tmp/inherit-chain.pptx`（python-pptx 产的，原带 **10** 条目录项）、`tmp/xlsx-probe/sample.xlsx`（**6** 条）、`tmp/ps-audit.docx`（**6** 条），往返产物目录项全为 **0**、文件部件清单与原文件**逐条一致**（11 / 10 / 15 个）、每个部件**内容逐字一致**；再让第三方读取器各读一遍：python-pptx 读出 1 页、首形文字"正文乙"，openpyxl 读出表名「销售」与 `A1=月份`，Apple `textutil` 转出正文首行"个人简历"；`oc ppt check` 与 `oc sheet check` 两份都 EXIT=0（"没有发现会让办公软件判损坏的结构问题。"）。依赖方向：`docx-io` **第一次** import `@office-copilot/opc`（此前它完全没引这层），`sheet-io`/`pptx-io` 沿用已有那条，仍是 `apps → packages`、格式包 → 格式无关层（§5.1） | `[x]` §3.2/§3.5 那一格：全仓只剩一份 zip 序列化，产物少掉 22 条与内容无关的目录项，三种格式的部件清单往返逐条对齐；`[ ]` 真 Word / WPS 的画面验证仍 BLOCKED（AGENTS §8），本机由 python-pptx / openpyxl / `textutil` 三个独立读取器 + 两条结构审计代证 |
| 本次（A1 那条"未建模部件原样带过"的实测与判据收窄） | 0（取证，不动代码，`tmp/a1-check/` 用完即清） | 不动 `src`，`pnpm build`/`pnpm test` 与上一行同一份数字（1 305） | **拿最富的真文件证，而不是自家夹具**：WPS 自带模板 `normal_mac.pptx`（`unzip -Z1` 124 条目 = **108 个文件部件 + 16 条目录项**，其中 **65 个是 `ppt/tags/*`**——数据标签，我们一件都没建模）与 `normal.pptx`（114 = 102 + 12，同样 65 个 tags），各走 `PptxFile.open` → `toBytes`（**一条都不改**）。用 `zipfile` 逐部件比字节（不拿我们的读侧自证）：**内容对不上的部件 0 个**，108 / 102 的文件部件清单与原文件**逐条对齐**；往返产物再过 `oc ppt check` → EXIT=0（"没有发现会让办公软件判损坏的结构问题。"），python-pptx 独立读出 **1 页 / 1 母版 / 11 版式**。**判据原文那句"`unzip -l` 条目集一致"因此按 OPC 收窄成"文件部件清单一致"**：目录项在 Part 2 §3 里本来就是可选的，而上一行那条收口统一不写它们，所以条目数 124→108、114→102 是**预期结果**而不是丢部件——判据若不收窄就等于把"少写 16 条无关目录项"记成保真缺陷。A1 从 `[ ]` 改 `[x]` | `[x]` §六 A 组那条最硬的保真判据闭环（未建模的 65 个 tags 部件逐字带过）；`[ ]` A3/A4/A6/A7 与 B/C 组剩余项仍开着 |
| 本次（line 439 那条读侧宽容：一页读不懂回三种失败码，不再报成"这页是空的"） | 3（`packages/pptx-io/src/pptx.test.ts` +2 新增一个 describe「一页读不懂时的失败码」、`packages/pptx-io/src/analysis.test.ts` +2、`packages/shared/src/index.test.ts` +1；夹具助手 `minimalDeck` 加了"给字符串就覆盖 slide1.xml / 给 null 就整个部件不给"两个档） | `pnpm build` EXIT=0（含 `tsc --noEmit`，0 条 `error TS`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / desktop 534 / **pptx-io 78**（+4）/ pdf-io 70 / pptx-tools 20 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / **shared 4**（+1）/ server 2 = **1 310 条** | **先测分支可达再写代码**（§3.6 不许写不会亮的分支）：拿 `packages/opc/dist/xml.js` 直接喂一段少了 `</a:p>` 的 XML，抛的是 `XmlError: Missing end tag for element a:p`——所以 `invalid-xml` 那一档不是想象出来的，`parseDocument` 真会炸。**失败码住在引擎、中文句子只存一份**（§5.7 + §3.2）：`read.ts` 只回 `missing-part`/`invalid-xml`/`no-sp-tree` 三个码并原样导出类型，句子在 `@office-copilot/shared` 的 `describeSlideReadFailure`，面板（`SlideDeck.tsx`）、`oc ppt info`（`apps/cli/src/ppt.ts`）、模型看到的 `reason`（`analysis.ts` 的 `readSlideText`）三个装配层念同一句；`ppt_describe` 那条走 `JSON.stringify(describeDeck(...))`，`readFailure` 原样上送，不用另写文案。**判据是真文件而不是夹具**：把 WPS 自带模板（108 个文件部件那份）的 `ppt/slides/slide1.xml` 抽掉再 `oc ppt info`，打印「⚠     这一页的部件不在文件里（关系表指向的幻灯片部件缺失）。」；同一份模板不抽时该行**一条不印**、照常列形状（`id=2 标题 1（shape 占位符:ctrTitle）`），即**零误报**——这条和 §六 A1 用同一份富文件，才证得了"缺失"与"只是没建模"分得开。**多 run 段的整段改禁令不受影响**：`analysis.test.ts` 里新加的那条只验"读不动的页回理由"，没碰 `singleRunParagraphs`（§2.3）。**没实跑的如实标**：§7.2 的真 app 那一格本批**没做**——面板上那句新提示与「加文本框」按钮的禁用态是照代码改的，没起 CDP 探针看画面；下批接画布式编辑时一并跑 | `[x]` line 439 关闭：读不懂的页在 CLI / 面板 / 模型三处都说成"读不懂"而不是"是空的"；`[!]` §7.2 的幻灯片面板实跑本批未做（改了 UI 但没起 app 看） |
| 本次（notes 部件的读侧：备注进读模型与画像，"没有备注页"与"有备注页但没写字"分开说） | 5（`pptx.test.ts` +3 新增「演讲者备注的读侧」那组、`analysis.test.ts` +2；夹具助手 `minimalDeck` 加第二参 `notesText`，给定时按 §2.10 把部件本体 + `ppt/slides/_rels/slide1.xml.rels` + `[Content_Types].xml` 的 Override 三处一起登记） | `pnpm build` EXIT=0（含 `tsc --noEmit`，0 条 `error TS`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / desktop 534 / **pptx-io 83**（+5）/ pdf-io 70 / pptx-tools 20 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 315 条** | **只认 `p:ph type="body"` 那一格**：备注页上还有 `sldImg`（幻灯片缩略图）与 `sldNum`（正文是 `<a:fld type="slidenum">`，页码是算出来的），把它们念成备注就是编内容——夹具里故意给缩略图那格写上"缩略图里不该被读到的字"、给页码那格写 `1`，断言只回正文两句。**失败通道不外溢**：备注部件缺失 / 解析失败一律回 `undefined`，`readSlideShapes` 那三个失败码只描述**这一页本身**读不读得动，一页备注读不动不许把整页判成读不动。**没拿自证**：python-pptx 现产一份带备注的 deck（`notes_text_frame` 两段"先讲背景/再讲结论"），它自己读回 `'先讲背景\n再讲结论'`，我们 `oc ppt info` 打的是同一句「备注（ppt/notesSlides/notesSlide1.xml）：先讲背景 / 再讲结论」；再拿 WPS 自带模板 `prometheus/newfile_mac.pptx`（**真文件里就有一格空备注**，python-pptx 读到 `''`）——我们如实念「有备注页但没写字」，同一份的图形/图片部件 0 误报；`oc ppt check` 对 python-pptx 那份 EXIT=0。**回写那一半没动**：备注部件属未建模部件，`toBytes` 原样带过（`pptx.test.ts` 里钉了一条：存回后该部件字节逐字一致、重读段落还在）。**视图与 AI 的写面本批都没接**：`SlideDeck` 不显示备注、`ppt_submit_proposal` 提不了备注，读得到但改不了，如实记在阶段 2 那句里 | `[x]` 阶段 2 那条「notes 部件的解析」关闭（读侧）；`[ ]` 备注的写侧 op、面板显示、`pptx-tools` 那条工具各归下一批 |
| 本次（D3 同构的颜色那一档：`a:schemeClr` 沿母版 `p:clrMap` + 主题 `a:clrScheme` 解成 hex） | 5（`pptx.test.ts` 新增一个 describe「主题色 a:schemeClr 解成 hex」5 条：映射查槽位、带变换与 `phClr` 不给色、主题链缺失不给色、夹具过结构审计、解出来的色不许烤回部件；新文件 `theme.ts`、夹具新函数 `colorDeckParts`） | `pnpm build` EXIT=0（含 `tsc --noEmit`，0 条 `error TS`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / desktop 534 / **pptx-io 88**（+5）/ pdf-io 70 / pptx-tools 20 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 320 条** | **先确认没有现成算法可抄**（§6.1/§6.2）：本机 python-pptx 1.0.2 的 `font.color.theme_color` 只回槽位名（`ACCENT_1`），它不做 `clrMap`→theme 的解色，所以这一档只能自己算；依据是 ECMA-376 Part 1 §19.2.1.9（`p:clrMap`：应用角色→主题槽位）与 §7.3.2.28（`a:schemeClr @val` 就是那些角色名）。**有意的非目标，写进代码注释而不是留给后人猜**：带 `a:lumMod`/`a:lumOff`/`a:tint`/`a:shade`/`a:alpha` 的与 `val="phClr"` 一律回 `undefined`——变换的 HSL 语义本机算得出，但没有能核对 PowerPoint 画面的通道（Keynote 只给文字与位置，WPS 无可视通道），**算错的颜色比没有颜色更坏**，会被当成事实烤回文件。**链路的复用而不是第二套**（§3.2）：`幻灯片 → /slideLayout → /slideMaster → /theme` 这一路走 `inherit.ts` 里那份 `relatedPartPath`（把它的可见性从文件私有改成导出），色板按幻灯片缓存，一份 deck 几十页只算一次母版与主题。**夹具专门防"蒙对"**：`colorDeckParts` 把母版 `p:clrMap` 的 `accent1` 故意映射到 `accent3`（`#ffffff`），按同名回退会解成 `#ff0000`，一眼可辨；`dk1` 用 `<a:sysClr lastClr>` 写，`sysClr` 那条路也被钉住。**没拿自证**：python-pptx 现产一份三条 run 的真 deck（`ACCENT_1` / `TEXT_1` 手工塞 `lumMod+lumOff` / 直接 RGB），另写一份**不经我们 parser** 的 python 解色（`tmp/theme_cross.py`：自己沿关系取 `clrMap` 与 `a:clrScheme`）——独立算法给 `accent1 → #4f81bd`、带变换那条给 `NONE`，我们的读模型回 `{sizePt:20,hex:"#4f81bd"}`、`{sizePt:20}`、`{hex:"#ff0000"}` **逐条全等**。两份 WPS 真模板（125 部件那份与 0 页那份）读侧正常、`oc ppt check` 仍 EXIT=0 零误报。**UI 侧一行没改**：`SlideDeck.tsx` 的颜色输入框与 title 提示、`slideDeckModel.ts` 的"与现值不同才出 op"、`analysis.ts` 上送给模型的 `format` 都本来就吃 `RunFormat.hex`，所以真实文件里那些写主题色的文字从此有基线色（不再是"继承自版式"）。**只读不烤**：解出来的 hex 不进部件——`open → toBytes` 后 `slide1.xml` 逐字不变（单测盯住）。**字号那一档（`p:txStyles`）本批没动**，仍是 D3 下的最后一档 | `[x]` §六 A4 的颜色解析那一半（引擎 + 交 UI + 不烤回）；`[!]` 真 app 里亲眼看见某条主题色文字显示成解出来的 hex 本批未跑（与 line 439 那格同一批 CDP 欠账，见阶段 3「仍缺」）；`[ ]` `p:txStyles` 有效字号仍缺 |
| 本次（D3 最后一档：run 没写 `sz` 时的有效字号沿六层链算出，`sizePt` 仍只记显式值） | 9（`pptx.test.ts` 新增一个 describe「有效字号的继承链」5 条：五档各在自己那一格赢、显式与继承那一对字段的语义、`lvl="1"` 那一层没写就不给数、夹具过结构审计 + 母版没 `txStyles` 时不给数、读完原样回写部件逐字不变；新文件 `text-styles.ts`、夹具新函数 `textStyleDeckParts`；`desktop` 侧 `sizeFieldHint` 4 条 + 一条"继承值一个字段都不交出"） | `pnpm build` EXIT=0（含 `tsc --noEmit`，0 条 `error TS`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / **desktop 538**（+4）/ **pptx-io 93**（+5）/ pdf-io 70 / pptx-tools 20 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 329 条**（提交 `ab559e9`，9 个文件 497 行插入） | **先确认没有现成算法可抄**（§6.1/§6.2）：本机 python-pptx 1.0.2 的 `font.size` 对没写 `sz` 的 run 直接回 `None`——它压根不算继承字号，所以这一档没有 OSS 实现可参照，只能按 ECMA-376 Part 1 §19.2.5（`p:txStyles`）与 §21.1.x 的占位符匹配自己搭。**两处只实测不轻信**（lxml 直读真文件）：① `<p:sldMaster>` 的子元素顺序实测是 `cSld, clrMap, sldLayoutIdLst, txStyles`——**`txStyles` 在最后**（先前"排第二位"的印象是错的，夹具按实测写）；② python-pptx 默认模板与全部 WPS 模板里那三个 `p:txStyles` 子元素**都没有 `@phType`/`@styleId`**，所以角色只能按元素名选（`titleStyle`/`bodyStyle`/`otherStyle`），照属性查会一条都查不到；WPS 模板与默认模板的 `otherStyle/lvl1pPr` 都写 `sz="1800"`——这就是"纯文本框为什么画 18pt"的来处，也是夹具里那条 `id=5` 的期望值。**六层的顺序**：run `@sz` → 段落 `<a:pPr><a:defRPr @sz>` → 形状自己的 `a:lstStyle` → 版式那一格的 lstStyle → 母版那一格的 lstStyle → 母版 `p:txStyles`；占位符匹配**复用几何那条规则**（slide→layout 按 `@idx`、layout→master 按 `MASTER_TYPE_OF_LAYOUT_TYPE`，版式那格缺失时从形状自己的 ph type 推 masterType），不写第二份匹配逻辑（§3.2）。**夹具防"蒙对"**：五档各给一个互不相同的值（段落 20 / 版式 40 / 母版格 24 / 形状 lstStyle 14 / otherStyle 18），任何一个层被跳过都会露出来。**没拿自证**：另写一份**不经我们 parser** 的 lxml 链式实现（`tmp/size_chain.py`），对**真文件**产的 deck（`tmp/size-real.pptx`：从 WPS `new_wpp.pptx` 经 python-pptx 生成，11 版式 × 占位符 + 文本框、段落 `level` 交替、34 条 run）逐条比 `sizePt`/`effectiveSizePt`/`effectiveSizeSource`——**40 条 0 不一致**，且档位分布非空（`master-text-styles` 20、`layout` 9、`run` 3、链上无值 8），不是"全都 undefined 所以全都对"；`oc ppt check` 对夹具与真文件都 EXIT=0。**只读不烤**：`open → toBytes` 后**幻灯片部件与母版部件都逐字不变**，夹具里断言 slide 内不含 `sz="4400"`（母版那个值）。**UI 那格的取舍**：`sizeFieldHint` 只进 `<input title>`，`value` 仍只放显式 `sizePt`——把 24pt 填进输入框会让用户一个字没敲也落出一条 `sz`，等于 D3 禁止的烘焙（那条用例直接断言 `pendingEdits` 回 `[]`） | `[x]` D3 那三档至此落齐（几何 `1ee3158`、颜色 `8023501`、字号 `ab559e9`），且三档都是"只读、绝不烤回部件"；`[!]` 真 app 里亲眼看见那句 hover 文案本批未跑（与 line 439 / A4 同一批 CDP 欠账，见阶段 3「仍缺」）；`[ ]` 画布式选中与拖拽、备注写侧仍缺 |
| 本次（把 §7.2 欠的那格 CDP 实跑补上：读不懂提示、「加文本框」禁用态、主题色色框、字号 hover 文案——顺带修掉实跑翻出的一句重复文案） | 0（只改一句 UI 文案，没新增用例：本仓没有 React 渲染 harness，画面证据由真 app 出） | `pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包与上一行同一份数字：docx-io 296 / sheet-io 245 / desktop 538 / pptx-io 93 / pdf-io 70 / pptx-tools 20 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 329 条**（`apps/task-pane` 仍是"没有测试文件"那格）。文案修复单独一刀 `91cc3b9` | **一份文件同时喂出四种形状**（`tmp/slide-ui-32/`）：python-pptx 现产两页 deck（第 1 页标题 run 只写 `schemeClr accent1` 不写字号 → 继承 44pt + 主题色；正文 run 显式 `Pt(32)` + `#00cc00`），再用 `zipfile` 把 `ppt/slides/slide2.xml` 抽掉做成第二份（关系留在，正好走 `missing-part` 那一档）。跑的是 `dist`（先 `pnpm build`，§8 那条"探针跑 dist 不跑 src"），`--user-data-dir=tmp/slide-ui-32/profile` + `--remote-debugging-port`，一条 AI 请求都不发（不需要 Key）。**屏幕上看到的四样**：① 第 2 页一句「这一页的部件不在文件里（关系表指向的幻灯片部件缺失）。」，同页**没有**形状列表；② 第 2 页的「加文本框」是灰的（`disabled=true`），第 1 页的是可点的——同一屏里对照，禁用不是全局样式；③ 第 1 页那条主题色文字的颜色框是**蓝的**，`value` 与 `title` 都是 `#4f81bd`（不再是"当前颜色继承自版式"那句兜底），另一行绿框 `#00cc00`；④ 字号框：继承那行 `value` 是**空**、`placeholder` 是"继承"、`title` 是「当前 44pt，继承自母版的文本样式（本形状没写字号，写死就切断母版联动）」，显式那行 `value=32`、`title` 是「当前 32pt（这条 run 自己写的）」。**D3 的 UI 侧硬证据**：一屏都没动过时顶栏是「2 页 · 0 处待保存」、保存与撤销都禁用、脏点没有——继承来的 44pt 没有变成待写改动。备注没回归：面板 `innerText` 里 0 次出现"备注"。**实跑翻出的缺陷**：第 2 页那句印成「这一页读不动：这一页的部件不在文件里…」——面板的前缀与 `describeSlideReadFailure` 那三句**各自都以"这一页"开头**，拼起来就重复；三句本身已经点名是哪一种失败，所以删掉面板那句前缀（`91cc3b9`），重跑同一份探针后该行只剩一句。**如实标注看不到的那部分**：④ 那句是读 `input` 的 `title` 属性拿到的，鼠标按八小步 glide 到了框上并截了图（`tmp/slide-ui-32/hover.png`，截图里只有面板），**原生 tooltip 飘窗由 macOS 绘制、CDP 截不到**——所以"文案内容"是实跑，"飘窗长什么样"不是。**没碰用户进程**：跑前 `pgrep -f "apps/desktop"` 只有一个 vite dev server、0 个带 `tmp/` 的 Electron，所以没有杀任何进程 | `[x]` line 648 与 line 650 留的那两格 `[!]`（真 app 里的读不懂提示 / 按钮禁用态 / 主题色色框 / 字号 hover 文案）一并消掉，§六 A4 由 `[~]` 改 `[x]`；`[!]` tooltip 像素本机无通道；`[ ]` 画布式选中与拖拽、备注写侧仍缺 |
| 本次（备注写侧第一档：`setSlideNotes` 只改这一页**已经存在**的备注部件，现有备注带多截格式就整条拒绝） | 7（`pptx.test.ts` 新增一个 describe「演讲者备注的写侧」7 条：三段整篇换且幻灯片部件逐字不动、`revertSlideEdits` 把备注部件字节级还原、`<a:pPr>`/`<a:endParaRPr>` 从旧第一段带过来、多截 run 那档回精确理由、带 `<a:rPr>` 那档回精确理由、备注部件不存在回「备注部件不存在：…」、落地产物经 `newlyIntroducedIssues` 零条新增） | `pnpm build` EXIT=0（含 `tsc --noEmit`，0 条 `error TS`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / desktop 538 / **pptx-io 100**（+7）/ pdf-io 70 / pptx-tools 20 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 336 条**（`fccd4f9`，4 个文件 186 行插入） | **拿真产物而不是夹具**（§6.3）：`tmp/notes-write/` 里 python-pptx 现产一份 deck（第 1 页三个形状 + 旧备注两句，43 个文件部件）→ 我们的读模型念 `["旧备注一句","旧备注二句"]`，`notes.partPath` 直接当 op 坐标 → `setSlideNotes` 换成三段（中间一句空行）存盘 → **python-pptx 独立读回 `['新备注一句', '', '新备注三句']`**，同一页正文三格仍是 `['', '标题甲', '正文乙']` 一字未动；空段落成 `<a:p><a:r><a:t/></a:r></a:p>`，"一段一个字符串"的对应关系不靠吞空行维持。**字节面用 `zipfile` 逐部件比，不拿我们的读侧自证**：`applied.pptx` 对 `in.pptx` **只有 `ppt/notesSlides/notesSlide1.xml` 一个部件变了**，其余 42 个部件字节逐字一致、文件部件清单 43 条逐条对齐（没多出目录项、没丢部件）；`revertSlideEdits` 回 `applied=1 / skipped=[]`，撤销后的备注部件与输入**1 288 字节逐字相同**（§2.5 内容寻址，不是按位置猜）。`oc ppt check` 对两份产物都 EXIT=0（"没有发现会让办公软件判损坏的结构问题。"），单测里 `newlyIntroducedIssues(base, out)` 同样 `[]`——**这一档不新增部件，§2.10 那三处登记一个字都不动**，这也是它跟"新建 notesSlide"那一档的本质区别。**第二条红线写成精确文案**：多截 run 与带 `<a:rPr>` 的 run 各回一句（「备注里有 N 段带多截不同格式的文字，整篇覆盖会把它们压平，这一档不改。」/「…N 截带字符格式的文字…」），拒绝而不是静默压平（§2.3）；新段的 `<a:pPr>`（项目符号、层级）与 `<a:endParaRPr>`（`lang`）从旧第一段带过来，只换文字不动版式意图（§2.7）。**跨界只多一种 kind**：`PptxEditRequest` 加 `{ kind:"set_notes", target:{notesPartPath}, paragraphTexts }`，`pptxHandlers.ts` 的 `makeEdit` 多一条 case，落地/撤销/modlog 全都走现成那条通道（§3.4 扩展现有接口，不新建第二条编辑链）。**如实记一条本机不通的通道**：想再让 Keynote 问一次"备注正文"没问成——`note of slide 1` 在 Keynote 14.2 的 AppleScript 字典里报 `-1723 不允许访问`，而这份 `.pptx` 的导入还撞上 `AppleEvent 已超时 -1712` / `document 1 无效的索引 -1719`（查询排在导入完成之前）。所以"第三方办公软件真读过这份带新备注的文件"本批**只有 python-pptx 一个独立读取器**；查完 Keynote 的文档数为 **0**，没留下窗口，也没碰用户自己的进程。**装配面这格还是空的**：`SlideDeck` 不显示也不编辑备注、`ppt_submit_proposal` 里没有 `set_notes`，模型改不了备注——引擎有 op 而用户够不着，如实挂在阶段 2 那句里 | `[x]` 阶段 2「备注的写侧」第一档（改已存在的备注部件）闭环，且三处红线（§2.3 压平、§2.5 撤销、§2.10 登记面）各有实测；`[ ]` 第二档新建 `notesSlide`（要 §2.10 三处登记 + 没有 notesMaster 就拒绝）、备注的面板显示与编辑入口、`pptx-tools` 那条工具、`slideAgent` 的页号→`notes.partPath` 映射各归下一批；`[!]` Keynote 那条第三方复开通道本机问不出备注 |
| 本次（把备注那条 op 开进 AI 的提案面：第八类改动 `set_notes`，页号→`notes.partPath` 由装配层翻，"会不会拍平"与引擎共用一份判定） | 5（`pptx-tools` +1：`set_notes` 的缺字段/空数组/单段超 2000 字三条拒绝，加上那条一致性用例把清单从七项钉到八项；`pptx-io` +1：读侧把「每段几截 run」与「带 `<a:rPr>` 的 run 截数」念出来；`desktop` +3：正常那条翻成 `{kind:"set_notes",target:{notesPartPath}}`、第 2 页没备注部件时不交出、备注带多截格式时不交出） | `pnpm build` EXIT=0（含 `tsc --noEmit`，0 条 `error TS`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / **desktop 541**（+3）/ **pptx-io 101**（+1）/ pdf-io 70 / **pptx-tools 21**（+1）/ opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 341 条**（`ad4d995`，8 个文件 190 行插入） | **引擎有 op 而模型够不着就是半成品**：`fccd4f9` 落的那条 `setSlideNotes` 上一批谁都用不上——面板没有备注输入，`ppt_submit_proposal` 也没有那一类改动，这一批只补"够得着"这一环，落地通道一个字没动（`pptxHandlers.ts` 那条 case 早在了）。**判定只存一份**（§3.2）：`notesFlattenNote(paragraphRunCounts, styledRunCount)` 从 `setSlideNotes` 里抽出来，引擎按现场 XML 数、装配层按读模型数，两边念同一句中文理由；装配层必须能提前判，否则用户点了确认才发现"任一条失败就整批不写"（§2.4）。**读模型那两个数是半个新事实，不是第二套坐标**：`SlideNotes.paragraphRunCounts` / `styledRunCount` 由 `readSlideNotes` 现数，`pptx-tools` 仍然不 import 引擎（§三 的包边界没破）。**引擎坐标不上送**（与 `add_text_box` 同一条理由）：模型只交页号，`notesSlideN.xml` 那个路径由 `slideRequestsOfProposal` 一处翻（§3.3）。**zod 与给模型的 JSON `oneOf` 同批改**：只加校验忘了声明，模型就提不了那条，所以那条一致性用例把"两边逐项相等"钉到八项上。**真 app 实跑**（`tmp/probe/slide-notes-ui.mjs`，端点是本机 8901 桩 + 假 Key，验装配面不是模型行为；夹具 `tmp/notes-ai/ui/probe.pptx` 由 python-pptx 现产 3 页 49 部件：第 1 页备注两句同一种格式、第 2 页没有备注部件、第 3 页备注里那一截写了 `b="1"`）：卡片头是 **1 处改动**，三条提案逐条都念出来（「第 1 页 备注整篇换成 3 段「新备注一句 / （空行） / 新备注三句」」），两条拦下的各跟一句理由（「第 2 页没有备注页…这条跳过。」「第 3 页的备注改不了：备注里有 1 截带字符格式的文字，整篇覆盖会把它们压平，这一档不改。」）；点「应用到幻灯片」→ **49 个部件里只有 `ppt/notesSlides/notesSlide1.xml` 一个变了**，其余逐字节一致、部件清单没多没少；**python-pptx 独立读回** applied 那份三页的备注是 `['新备注一句\\n\\n新备注三句', '', '加粗的一截']`、三页正文 `['标题甲','正文乙','正文丙']` 一字未动，第 3 页那截加粗原样还在（被拦下的那条根本没碰磁盘）；`oc ppt check applied.pptx` EXIT=0；点「撤销上一次保存」后 **49 个部件全回原样**（`changed:[]`），python-pptx 读回第一页又是 `旧备注一句\n旧备注二句`。跑前 `pgrep -f "apps/desktop"` 是 1（vite dev server）、跑后 `pgrep -f tmp/notes-ai` 是 0，没留窗口也没碰用户自己的进程。**仍然没跑的只剩一件**：真端点一次都没发，所以"模型自己会不会主动提 `set_notes`、会不会把段数与空行给对"没有证据——这一条与阶段 5 那句 B5 同一格欠账 | `[x]` 阶段 2 那句里的"AI 的备注写面"这一格关闭（工具面 + 装配层映射 + 预检有用例，且真 app 里卡片文案、只动备注部件、独立读回、撤销逐字回原样四条都有实跑证据）；`[!]` 真端点模型行为探针（B5）仍欠；`[ ]` 新建 `notesSlide` 那一档仍缺 |
| 本次（备注的用户手工入口：幻灯片面板每页底下一格备注，显示 + 整篇换掉，没备注页与会拍平两种情形只说理由） | 4（`slideDeckModel.test.ts` 新增一个 describe「notesSaveBlock / notesRequestOf：备注那一格的落地范围」：没有备注页那句、内容没变那句"备注没有改动。"、带格式那句走 `notesFlattenNote`、干净那份回 `null` 且 `notesRequestOf` 按换行切成 `paragraphTexts` 并只带 `notes.partPath`） | `pnpm build` EXIT=0（含 `tsc --noEmit`，0 条 `error TS`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / **desktop 545**（+4）/ pptx-io 101 / pdf-io 70 / pptx-tools 21 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 345 条**（`2dc83f4`，4 个文件 224 行插入） | **补的是装配面，不是第三条编辑链**（§3.3/3.4）：`SlideNotesField` 只做显示与装配，判定全在 `slideDeckModel.ts` 那两个纯函数里，"会不会拍平"仍调引擎那一份 `notesFlattenNote`（`slideAgent` 也用同一份，§3.2），落地仍走 `SlideDeck` 现成的 `commitNow`→`ppt:edit`，所以**存前审计、modlog、撤销栈一个字都没改**。文案在渲染层组装（§5.7），样式只用 `src/styles/slides.css` 那一份与既有 token（§5.4）。**真 app 实跑**（`tmp/probe/slide-notes-panel-ui.mjs`，跑 `dist`、`--user-data-dir=tmp/notes-ai/panel/profile`、9359 端口，**不起端点也不发 Key**——这一趟验的是手工入口，不是模型行为）：三页一进来各自的状态是**第 1 页**输入框初值从磁盘读来 `旧备注一句\n旧备注二句`、按钮禁用且 `title` 是"备注没有改动。"；**第 2 页没有输入框**，只有一句「这一页没有备注页（新建备注页这一档还没做，只能改已有的备注）。」；**第 3 页**输入框是 `加粗的一截`，一打字就顶出「备注里有 1 截带字符格式的文字，整篇覆盖会把它们压平，这一档不改。」且按钮始终禁用（§2.3 拒绝而不是静默压平）；把第 1 页改成三段（中间一句空行）后按钮解禁，**顶栏仍是「3 页 · 0 处待保存」**（备注不进 run 级草稿那个计数）；点「保存备注」→ toast「已保存那一页的备注，要回退点「撤销上一次保存」」，**49 个部件里只有 `ppt/notesSlides/notesSlide1.xml` 一个变了**，输入框里的值即落盘后重读的那一份、按钮回到禁用；再点「撤销上一次保存」→ **`changed: []`**（49 个部件逐字回原样），输入框回到 `旧备注一句\n旧备注二句`。**独立读取器交叉验证**（不拿我们的读侧自证）：python-pptx 读 applied 那份三页备注 = `['新备注一句','','新备注三句'] / None / ['加粗的一截']`（第 3 页那截加粗原样，第 2 页仍没有备注部件），对照原文件 `['旧备注一句','旧备注二句'] / None / ['加粗的一截']`；`oc ppt check tmp/notes-ai/panel/saved.pptx` EXIT=0。**如实记两处不足**：① 探针读到撤销后的 `.toast` 文案仍是上一条"已保存那一页的备注…"（两条提示叠在同一容器里，`querySelector` 取到先插入的那条），撤销的判据用的是**逐部件字节** `changed: []` 而不是那句文案；② 本机没有能看备注**画面**的通道（Keynote 那条 AppleScript 问 `note of slide` 报 `-1723 不允许访问`，见上一行），所以"Word/WPS 里备注真的显示成新句子"仍是推断，证据止于独立读取器读回同样的三段。**跑前后进程干净**：跑前 `pgrep -f "apps/desktop"` 只有用户自己的 vite dev server 与一个 helper，跑后 `pgrep -f tmp/notes-ai` = 0，没杀任何不属于探针的进程。 | `[x]` 阶段 2 与阶段 3 各留的"面板的备注显示与编辑入口"这一格关闭，`set_notes` 这条 op 至此引擎 / AI 提案面 / 用户手工三面齐；`[!]` 备注的渲染画面本机无通道（同上，Keynote `-1723`）；`[ ]` 新建 `notesSlide` 那一档（§2.10 三处登记 + 没有 notesMaster 就整条拒绝）与真端点模型行为探针（B5）仍缺 |
| 本次（取证：让 python-pptx 在真模板上现建一个备注页，拆包看新建 `notesSlide` 那一档到底要写哪四处） | 0（纯取证，没写实现代码，也没新增用例） | 不涉构建：`oc ppt check tmp/notes-new/out.pptx` EXIT=0，`oc ppt info` 念出「备注（ppt/notesSlides/notesSlide1.xml）：取证第一句」 | **一手参照而不是文档转述**（§6.2）：模板取 `/Applications/wpsoffice.app/.../kfindermenu/res/templates/new_wpp.pptx`（真模板，且带 1 条 notesMaster 关系——取证过的那 8 份里另外 4 份连备注母版都没有，这份是能用的那种），`add_slide` 后取 `slide.notes_slide` 逼它建部件、存盘，再拿 `zipfile` 逐部件对比。**四处登记逐条抄下来**：新部件 `notesSlide1.xml`（三格 `p:sp`：`sldImg idx=2` 无 `txBody`、**`body idx=3 sz="quarter"` 才是正文**、`sldNum idx=5` 无 `txBody`，末尾 `<p:clrMapOvr><a:masterClrMapping/>`）+ `notesSlide1.xml.rels`（**`rId1`→notesMaster、`rId2`→slide**）+ 幻灯片侧 `.rels` 追加 `notesSlide` 那条 + `[Content_Types].xml` 的 Override，`消失部件: []`。**三格占位符恰好印证读侧那句"只取 body"**：`sldImg`/`sldNum` 若被念成备注就是编内容，与阶段 2 读侧那条判据互为正反面。**没拿自证**：产物先过**我们自己的结构审计**（EXIT=0）与 `oc ppt info`（读侧已经能吃写侧将要产的那种形状，这一档不动读模型）。**反面教材再坐实一次**（D1）：python-pptx 保存时整包重序列化，把模板自带的 `Default Extension="JPG"` 丢了、`docProps/core.xml` 与全部关系表重排——所以我们的这一档**只动那四处**，不跟着重排。**判据里的硬门槛**：`SlideEditOp` 的 `beforeXml` 现在是必填字符串，表达不了"部件本来不存在"，撤销要从"写回前像"变成"删部件 + 撤三处登记"，落地/审计/modlog 那条链要一起改，故这是一批而不是删掉 `setSlideNotes` 那句守卫就完事；没有 notesMaster 时整条拒绝、绝不自己造一份备注母版 | `[x]` 这一档的取证与判据（四处登记的逐条形状 + 参照产物）落进文档；`[ ]` 实现未动一行代码（op 形状与撤销语义那处跨门槛留给下一批）；`[!]` 没在 Word/WPS/PowerPoint 里看这份产物的画面，Keynote 那条问 `note of slide` 仍报 `-1723` |
| 本次（`e7b36c3` 引擎那一半：形状列表摊成扁平 DFS——组里的子形状与连接线终于进模型，结构那三条 op 改在形状所在的那一层） | 6（`pptx.test.ts` 新增一个 describe「组合与连接线（读侧展开 + 组内形状的写侧边界）」：五行 DFS 顺序与 `groupShapeId`、连接线无文字且自己带几何、组内形状 `hasOwnGeometry` 为真却不给 `boxPx`/`effectiveBoxPx`/`boxPxSources`、组内改文字落地并逐字节撤得回、删组里的孩子只摘那一格且组本体不动、组内换叠放只在那个组里换序而页级顺序不变、组内几何整条拒绝而连接线的几何照写） | `pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包：pptx-io **107**（+6）/ desktop 545 / docx-io 296 / sheet-io 245 / pdf-io 70 / pptx-tools 21 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 350 条** | **这一刀是因为"看不见"是量出来的而不是猜的**：`readSlideShapes` 只走 `p:spTree` 的直接子元素，所以一份带组的 deck 在面板上只显示 3 行中的 2 行、`oc ppt info` 少念一条、模型画像里没有那一格——更要紧的是 `removeShape`/`moveShapeLayer` 拿页 spTree 去摘组里的孩子时 `indexOf` 返回 `-1`，`splice(-1,1)` 摘掉的是**别的形状**或删除根本不发生，却照样回 `ok: true`（§2.4 的静默失败）。修法是把"所在的那一层"跟着形状一起往下传：`findShapeById` 递归进 `p:grpSp` 并回 `{ shape, container }`，`rewriteShapeOrder`/`removeShapeFromTree` 都收 `container`；`SHAPE_HOST_TAGS` 补 `p:cxnSp`（它原本连"是个形状"都不算）。**连接线进列表但不进文字**：`p:cxnSp` 按 schema 就没有 `p:txBody`（§19.5.3），所以 `paragraphTexts` 恒空，读侧不用造一条"它没有文字"的假段落，装配层那句越界理由改成"这个形状没有文字体（图片、图形框或连接线）"。**D3 那条线不许为了让组里能编辑而破**：组里的 `<a:xfrm>` 写的是**子坐标系**（`a:chOff`/`a:chExt` 那一档换算要连着组的缩放与旋转一起算，还没做），烘焙成页面数字就是给错数，所以 `readShape` 在 `groupShapeId !== undefined` 时把三个几何字段全部留空、`setShapeGeometry` 在 `located.container !== shapeTreeOf(root)` 时整条拒绝——**宁可不给数也不给错数**，与占位符那一档同判据；文字、字体、颜色、删形状、组内换叠放都不受影响（叠放是"同一层内"的性质，所以组内照做，只在它自己那个组里换序）。**没拿自证**：夹具先过 `oc ppt check`（EXIT=0），产物再交给 python-pptx 独立读回 | `[x]` 组内子形状与连接线进了读模型、三条结构 op 按所在层落地、几何那一条在组内整条拒绝；`[ ]` 组框↔子坐标系换算、取消组合、对齐与分布仍缺（所以组内的几何一格只给一句话）；`[!]` 没在 PowerPoint/WPS 画面里看这份产物（本机 §8：看不到渲染结果），判据以结构审计 + python-pptx 独立读回为准 |
| 本次（`6baef49` 装配那一半：叠放两端的判定改成只比同一层的邻居，几何那一条在装配层就按 `inGroup` 拦下，视图上一行一个名字） | 5（`slideDeckModel.test.ts` 新增一个 describe「组合那一层」4 条：组内形状拿整页上下端判会**两头都判错**、同层只有一个形状时给"组内没有可叠放的邻居"、组本体的几何说明与组内子形状那句必须不同（组写的是 `<p:grpSpPr>` 那对组框，不是"自己没写几何"）、`inheritedBoxNote` 不给组内形状猜数；`slideAgent.test.ts` +1 条：一份提案里组内改文字交出、组内改几何与组内已顶端的换层各回一句中文理由而不是整批不落地） | `pnpm build` EXIT=0；`pnpm test` EXIT=0，逐包：desktop **550**（+5）/ pptx-io 107 / docx-io 296 / sheet-io 245 / pdf-io 70 / pptx-tools 21 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 356 条** | **`layerEdgeNote` 是这一刀翻出来的真缺陷**：扁平列表上 `shapes.indexOf(shape)` 就是整页下标，于是组里第一个形状被判定"已经在最底层"（禁用「下移一层」而它的组内邻居明明在它下面）、组里最后一个被判定"已经在最顶层"（禁用「上移一层」），两端判反的方向各不同。改成 `shapes.filter((entry) => entry.groupShapeId === shape.groupShapeId)` 再取索引，并把理由文案分成"组内最顶层/最底层"与页面级两种——这一条与引擎同判据是**刻意**的：引擎只在所在层换序，视图禁用态若按整页算就会把引擎必拒的那条交出去，§2.4 是整批放弃而不是丢一条。**几何那条的拦法**：`isGroupShape` 换成 `geometryBlockedNote`，组本体（子偏移语义不同）与 `groupShapeId !== undefined` 的子形状（组内子坐标系）两类各给一句理由，都在提交前拦；`ppt_submit_proposal` 的描述同步补上"inGroup=true 的那一格几何改不了"，让模型不必试。**视图只做装配**（§5.7）：一行一个名字的口径来自 `shapeKindLabel` 的 `组合里的` 前缀与 `连接线` 大类，几何那一格的"为什么不给数"仍是 `slideDeckModel.inheritedBoxNote` 那一句，视图不自己拼数字。**实跑（§7.2，`tmp/probe/slide-group-ui.mjs` 跑 `dist`，CDP 连真 app）**：五行全进列表且类别点名正确（「组合里的文本框」×2、「连接线」），组内 TextBox 2 的「下移一层」禁用态理由是"它已经在组内最底层"、TextBox 3 的「上移一层」禁用态是"它已经在组内最顶层"（**不是**页面级的两端），点 TextBox 2 的「上移一层」后磁盘顺序 `2->9->4->3->10`（组本体 `9` 位置不动、组内换成 `4->3`）、撤销回 `3->4`、组内改文字存盘后 python-pptx 独立读回 `[[2,'Title 1'],[group 9,[[3,'组内甲改过'],[4,'组内乙']]],[10,'Connector 9']]]`，`oc ppt check` 对夹具与存盘产物各 EXIT=0；未保存的草稿没被重排牵连（`DRAFT_AFTER_UP` 仍在） | `[x]` 视图与装配层认得组与连接线、叠放禁用态与几何拦截都按所在层判；`[ ]` 取消组合/对齐分布/画布式拖拽仍缺；`[!]` 组合与连接线的**画面效果**没本机通道（§8：WPS 不可捕获、无 PowerPoint/LibreOffice），只到"结构审计 0 条 + python-pptx 读回一致"这一档 |
| `ae4056a`（引擎那一半：表逐格读出来 + 第九个 op）+ 本次（装配那一半：`pptx-tools` 那条工具、`slideAgent` 的映射与预拒、AIPanel 预览行、视图的表格网格） | 2 / 4 / 5 / 3（表格那一格那一档） | `pnpm build` EXIT=0（含 `tsc --noEmit`：判别式联合多一条 `table_text` 后，`pptxHandlers.ts` 那条 switch 少了 case 编译器当场指着，不靠人记）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / desktop **555**（+5：AI 映射与预拒 3 条、视图 `pendingEdits`/`dropShapeDrafts` 2 条）/ pptx-io **113**（引擎那批 +6）/ pdf-io 70 / pptx-tools **22**（+1：行列边界、缺字段与多字段都回中文理由而不是抛异常）/ opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 368 条** | **实跑（§7.2，真 app 跑 `dist` + CDP）**：python-pptx 现产一份 3×2 表——第 1 行第 0 列故意做成两截 run（「张」加粗 + 「三」24pt），第 2 行第 0~1 列横向合并。① 视图按网格摊开：`第 1 行第 0 列` 只读并写"这段有 2 截不同格式的文字，整段覆盖会拍平粗体与字号"（§2.3），`第 2 行第 1 列`（续格）只读并写"这一格被并到左边那一格里了"，合并起点标「横向占 2 列」；② 在第 0 行第 0 列把「姓名」改成「名字」→「1 处待保存」→ 保存 toast「已保存，动了 1 张幻灯片」→ python-pptx 独立读回磁盘那格是 `名字`，其余五格（含两截 run 那格与合并那格）逐字未动；③ 撤销后磁盘回 `姓名`，重摊一次视图两格状态不变；`oc ppt check` 对夹具与产物各 EXIT=0。**AI 面**：页号/形状号/行/列/段/run 任一层对不上现场模型就不交出这条、逐条给中文理由，且一条被拒不牵连同批兄弟（单测里有 mixed batch 那条）。**表格里的字号与颜色没给入口**：引擎那两条 op 至今仍只认形状的 `p:txBody`，视图不许自己发明第二条写路径 | `[x]` 阶段 6 的表格那一半（逐格读侧 + 第九个 op + AI 提案面 + 视图编辑）关闭；`[x]` 表格的字号/颜色由上面两行（`fffda8a` + `4e351ef`）接上；`[ ]` 图表部件、`c:numCache` 同步归后续；`[!]` 表格的**画面效果**本机无通道（同上一行那两条约束），只到"结构审计 0 条 + python-pptx 读回一致"这一档 |
| `fffda8a`（引擎那一半：字体与颜色那两条 op 认得表格那一格） | 4（op 面）+ 2（写侧边界） | `pnpm build` EXIT=0（含 `tsc --noEmit`：`RunTarget` 变宽后跨界契约一处过）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / desktop 555 / pptx-io **115**（+2）/ pdf-io 70 / pptx-tools 22 / opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 370 条** | **没新开第十、第十一条 op**：`setRunFont`/`setRunColor` 的坐标改成 `ShapeRunTarget | TableCellRunTarget`，带 `row`/`col` 就交给已有的 `locateTableCellRun`——`<a:rPr>` 的 `b`/`sz` 与"同一 run 只能一个填充"那两条约束仍只有一份（§3.2）。夹具实测：给 `第 0 行第 1 列` 那条两截 run 的第一截上 `b="0" sz="2000"`，序列化后只有那一截的 rPr 变了，同段另一截与隔壁格逐字节不动；上色到合并起点第二段后**整页 `<a:solidFill>` 计数为 1**（没给别的格子叠加第二个填充）；apply→revert 后部件回到逐字原样。三类拒绝沿用表格那条理由：续格（"它不画自己的字"）、非表形状（"不是一张表"）、行越界（"行号越界：7（这张表有 2 行）"） | `[x]` 引擎那一半关闭；`[x]` 工具面与视图那两格由下一行（`4e351ef`）接上；`[!]` 画面效果同前两行 |
| 本次（`4e351ef` 工具面与视图那一半：`set_run_font`/`set_run_color` 加可选行列，表格那一格补字号与颜色输入框） | 5（`pptx-tools` +2：行列成对那条 refine 与"省略行列的旧条目照旧合法"、JSON 两条分支的 row/col 是可选而 `set_table_cell_text` 是必填；`desktop` +3：`slideAgent` 带行列翻成 `font`/`color` 且 target 带 row/col（含同批兄弟不被牵连、四种越界各回一句）、`slideDeckModel` 那一格的字号/颜色只交真改过的那一项）+ 2 处读侧断言（`pptx-io` 与 `analysis` 各钉一份 `paragraphRunFormats`，用例数不变） | `pnpm build` EXIT=0（含 `tsc --noEmit`）；`pnpm test` EXIT=0，逐包：docx-io 296 / sheet-io 245 / desktop **558**（+3）/ pptx-io 115 / pdf-io 70 / pptx-tools **24**（+2）/ opc 17 / pdf-tools 13 / sheet-tools 12 / word-client 8 / word-tools 6 / cli 5 / shared 4 / server 2 = **1 375 条** | **读侧补的是半个新事实，不是第二套坐标**：`SlideTableCell.paragraphRunFormats`（逐段逐 run 的**显式** `b`/`sz`/解出的 hex）——视图那两个框要有初值、模型要对准 run 序号，两边都得看得见那一截自己写过什么；表格里那几档继承（表格样式 / cell 的 `a:lstStyle`）没有现成解析器，所以**不给 `effectiveSizePt`、框子留空**，与 D3"宁可不给数也不给错数"同一条线（用户明示要覆盖时给个数即算改动，不是烘焙继承值）。**判定只存一份**（§3.2）：形状那截与表格那格共用的"只交真改过的那一项"抽成 `pushFormatEdits(edits, target: RunTarget, draft, currentFormat)`，两条路径的差异全在 `target` 带不带行列；`pptx-tools` 的 `row`/`col` 边界从 `set_table_cell_text` 里提出 `TableRow`/`TableCol` 两个常量，`TABLE_CELL_POS_JSON` 同样只写一份。**成对强制的理由**：只给一个下标时连"这条指普通形状还是指那一格"都判不出来，而引擎按行与列两个下标才走表格那条链。**实跑（§7.2，`tmp/probe/slide-table-format-ui.mjs` 跑 `dist` + CDP，9372 端口）**：① 视图上单 run 那格给三个框（文字 / 字号 / 颜色），多 run 那格与合并续格一个都不给（`CELL_FIELDS` 里两格的 `size`/`color` 均 false）；② 只填字号 30 与颜色 `#00cc00` → 顶栏「1 页 · **2 处待保存**」（文字那条没凭空多出来）；③ 存盘后磁盘上**只有那一截**的 `<a:rPr>` 多了 `sz="3000"` 与一个 `<a:solidFill><a:srgbClr val="00CC00">`，**整页 `<a:solidFill>` 计数 1**，同格另一截的 `b="1"`/`sz="2400"` 与隔壁四格逐字未动；④ 点「撤销上一次保存」后**整页 XML 与保存前逐字节相同**（`REVERT_BYTE_EQUAL true`，不是"看起来一样"）。**独立读取器交叉验证**（不拿我们的读侧自证）：python-pptx 1.0.2 读 `saved.pptx` 那一格 = `size 381000`（=30pt）/ `rgb 00CC00`，同段另一截仍是 `304800`（24pt）、`张` 的 `bold=True` 原样在，其余格 `None`；`oc ppt check tmp/slide-table-format-ui/saved.pptx` EXIT=0。**画面证据是一截图**（`tmp/slide-table-shot/panel.png`）：六格各 446px、`scrollWidth > clientWidth` 全 false（三个框挤在一行不出横向溢出）。**跑前后进程干净**：跑前 `pgrep -f "apps/desktop"` 只有用户自己的 vite dev server 与一个 helper（未动），跑后探针实例为 0 | `[x]` 阶段 6 表格那一半的"字号/颜色"这一格关闭（引擎 / 工具面 / AI 提案面 / 视图四面齐）；`[ ]` 表格的**继承字号解析**（表格样式与 `a:lvl` 那条链）没做，所以那一格没写显式值时框子是空的；`[ ]` 图表部件与 `c:numCache` 同步仍缺；`[!]` 幻灯片**画面渲染**本机无通道（§8：无 PowerPoint/LibreOffice，WPS 窗口不可捕获），判据止于"结构审计 0 条 + python-pptx 读回一致 + 我们自己的面板截图" |

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

### A2 · 阶段 0 夹具与 Keynote 通道（2026-10-14 实测，命令可复现）

WPS 自带模板（只读取，不入库）：

```
…/office6/mui/zh_HK/templates/newfile.pptx        → 35 part，slides=0，tags=0，masters=1，layouts=11
…/addons/kfindermenu/res/templates/new_wpp.pptx   → 125 part，slides=1，tags=65，masters=1，layouts=11
```

> 第一行就是"只遍历 `ppt/slides/` 的编辑器会当场毁掉一份空模板"的证据（35 个 part 里 0 张幻灯片）。

python-pptx 1.0.2 读回（`PYTHONPATH=tmp/pylibs python3 -c "from pptx import Presentation; …"`）：

```
bare.pptx  slides=0 masters=1 layouts=11     ← WPS 空模板原样读通
rich.pptx  slides=1 masters=1 layouts=11     ← 125 part / 65 tags 那份也读通
deck.pptx  slides=3 masters=1 layouts=11     ← 在 WPS 模板上加 3 页（标题/图表/表格）+ 1 页备注
           parts=49  charts=1  embeddings=1  notesSlides=1
```

Keynote 14.2 通道（AppleScript，逐条实测）：

```
open tmp/p0/b-lib-chart.pptx（python-pptx 默认模板 + 柱状图 + 2×2 表格，1 页）
  → count of slides = 1，object text = "乙"
  → export … as slide images → tmp/p0/pngb/pngb.001.png
  → 肉眼看过这张 PNG：柱子（Q1=1、Q2=2）、纵轴 0/1/2、横轴 Q1/Q2、表格首格 "x" 与表头底色全部渲染
open tmp/p0/a-wps-text.pptx / b-lib-chart.pptx / c-wps-chart.pptx → 三份都返回文档引用（开成功）
open tmp/p0/deck.pptx（3 页 + 备注 + 图表 + 表格，建在 WPS 模板上）
  → 一次 AppleEvent 超时 -1712、documents=0（拆开造同类夹具都能开 → 判据写"重试 + 如实报告"）
坑：`count of slides of front document` 在某些时机报 -1708，必须先 `set d to front document` 再 `count of slides of d`
收尾：每次问完 `close front document saving no`；探针结束 `quit`
```

探针产物全部在 `tmp/`（`pptx-kn-check.pptx`、`kn-export/`、`kn-export.pptx`、`kn-export.pdf`、`pptx-pip.log`），
按 §6.5/§7.3 由本次执行者清理，不入库。

  增删形状这批：`pnpm build` 与 `pnpm test` 都 `EXIT=0`，逐包 `opc 15 / sheet-tools 12 / pptx-tools 16 / **pptx-io 52**（+7 条新建/删除：形状号在本页内取、骨架与单位、空段、转义、拒绝不落、逐字回原、删到空）/ word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / cli 5 / **desktop 495**（+3 条 `dropShapeDrafts`）`，合计 **1152**。`packages/opc` 另多导出一个 `XmlNode` 类型（自己走子树时要按它接 `isElement` 收窄），不是新实现。
  把增删形状开进 AI 提案面这批：`pnpm build`（含 `tsc --noEmit`）与 `pnpm test` 都 `EXIT=0`，逐包 `opc 15 / sheet-tools 12 / **pptx-tools 20**（+4 条：`add_text_box` 五字段必填与非正尺寸与空文字、strict 拒 `shapeId`、`remove_shape` 缺号与 0 号、七类判别式在 zod 与给模型的 JSON `oneOf` 里逐项相等）/ pptx-io 52 / word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / cli 5 / **desktop 497**（+2 条装配：加文本框只定位到整页 + 页号越界，同批已删的形状后面那条不交出）`，合计 **1158**。
  把几何单轴的必填属性核上这批：`pnpm build`（含 `tsc --noEmit`）与 `pnpm test` 都 `EXIT=0`，逐包 `opc 15 / sheet-tools 12 / pptx-tools 20 / **pptx-io 53**（+1 条：占位符上只给一轴带原因拒绝、且拒了不许留下半个 `<a:off>`）/ word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / cli 5 / **desktop 498**（+1 条装配：继承几何的形状只给一轴时不交出）`，合计 **1160**。
  把几何那一类变成落盘前机检这批：`pnpm build`（含 `tsc --noEmit`）与 `pnpm test` 都 `EXIT=0`，逐包 `opc 15 / sheet-tools 12 / pptx-tools 20 / **pptx-io 58**（+5 条审计：半个 `<a:off>` 报出并念出缺的属性、四值写全不误报、`chOff`/`chExt` 与前缀换写也抓得到、`<a:xfrm>` 少整块 `ext/chExt` 也报、写全几何的产物自己过审计）/ word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / **cli 5**（那条文案用例加了几何那码的断言）/ desktop 498`，合计 **1165**。
  把 D3 的几何继承链落到读模型与视图这批：`pnpm build`（含 `tsc --noEmit`）与 `pnpm test` 都 `EXIT=0`，逐包 `opc 15 / sheet-tools 12 / pptx-tools 20 / **pptx-io 74**（+16 条继承链：逐轴读 `<a:xfrm>` 4、读模型的 版式→母版 逐轴回退 7、`createGeometryInheritance` 的类型映射与缺失部件容忍 3、带继承链的夹具过审计与"改文字后 `<p:spPr/>` 仍空"2）/ word-client 8 / docx-io 296 / word-tools 6 / sheet-io 245 / server 2 / cli 5 / **desktop 503**（+5 条 `inheritedBoxNote`：念出来源层、混合来源点两层、三层都没写不给数、非占位符不凑数、报了继承数仍一个字段都不交）`，合计 **1186**。类型上 `SlideShape` 多了两个必填字段，`slideAgent.test.ts` 那三份手写形状快照与 `slideDeckModel.test.ts` 的 `shapeOf` 一起补齐（`tsc` 抓出来的，不是改测试迁就实现）。
