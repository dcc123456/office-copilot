# Excel 表格支持规格（Excel / spreadsheet parity spec）

本文是**实现契约**，不是愿望清单。与 `docs/word-parity-spec.md` 同级：每一节的判据都必须能被一条命令、
一个单测、或一次真实 `.xlsx` 产物的打开结果打分。所有条款继承 `AGENTS.md`，冲突时以 `AGENTS.md` 为准，
冲突点在动手前先提出。

约定：`[机检]` 有测试/构建拦截；`[验收]` 逐条打分；`[纪律]` 只能靠收尾自检自觉回答。
状态记号：`[x]` 已达成、`[ ]` 未做、`[!] BLOCKED` 本机环境无法验证（写明卡在哪，不许用推测冒充验证）。

---

## 零、全局约束（每一节都受它管）

1. **交付物仍是"能被 Excel / WPS / Numbers 正常打开的 `.xlsx`"**。任何"表格里看得懂但文件坏了"的实现算缺陷，不算妥协。
2. **IO 单一入口**：单位/颜色/日期序列号的换算只允许发生在 `packages/sheet-io` 的进出口
   （类比 `AGENTS.md` §2.1 对 docx 的规定）。渲染层、AI 工具层、CLI 禁止出现第二处换算。
   新增的第二处换算一经发现即为缺陷。
3. **不许长出第二套同类基础设施**（`AGENTS.md` §3）：
   - 撤销：复用桌面既有的 op + `inverses` + 内容寻址 + `userData/modlog` 机制（`apps/desktop/src/ai/docEdit.ts:946 planEdit`、`:1156 revert`、`electron/modlog.ts`），表格只是换一种地址形态，**禁止**新建撤销栈或新建历史存储。
   - LLM 客户端：桌面继续用 `electron/aiHandlers.ts`（`callLLM:258` / `runStream:177` / `ai:stream-cancel:167`），**禁止**为表格另开一套。
   - 设置：继续用 `electron/settingsStore.ts` 的 `userData/config.json`。
   - 样式：继续只用 `apps/desktop/src/styles/*.css`（`tokens.css`），**禁止**引入 CSS-in-JS（这是下面否决 Univer / glide-data-grid 的主因之一）。
   - zip/XML 解析：复用仓库已有的 `jszip` + `@rgrove/parse-xml`（`packages/docx-io/package.json`），**不新增解析依赖**。
4. **模型只产生提案，不直接写文档**（`AGENTS.md` §2.9→§2.8）：表格工具入参必须过 zod `.strict()` 校验；
   半自动模式必须用户点确认才落地。**并且：数值计算不得交给模型心算**——模型只允许指定"对哪个区域做什么聚合"，
   实际数字由引擎算出（见 §五 决策 D4）。
5. **发给 LLM 的表格内容默认截断**（`AGENTS.md` §2.9）：默认只送表头 + 前 50 行样本 + 列类型推断结果 + 行列总数；
   禁止整表上送。区域取数走"模型请求 → 引擎返回该片数据"的工具回路，不靠加大默认窗口解决。
6. **未知结构必须原样保留，且保留粒度是"元素级"**：`.xlsx` 里我们建模不了的东西（条件格式、数据验证、透视表、外部链接、
   宏、图表、`x14ac` 扩展属性…）在读写往返中逐字节保留，类比 docx 的 `unmanagedBlock`
   （`packages/docx-io/src/ooxml-pm.ts:49`）。**实测已证明单靠"整部件 pass-through"不够**（见 §二 P5 派生约束）：
   必须持有原始 XML 并定点替换，禁止"解析成对象再整篇重新序列化"。
   这条是"不损坏用户文件"的技术底线，也是本仓与 Univer/SheetJS 这类"重序列化"库的根本分歧点。
7. **写文件必须原子**（`AGENTS.md` §2.4）：`write path + ".tmp" → fs.rename`，复用 `electron/fileHandlers.ts` 的落地方式；原地写回默认保留 `.bak`（`AGENTS.md` §8 安全红线 4）。
8. **一次推进一个阶段**，每阶段独立可提交、可回滚；提交信息中文、`<type>(<scope>)` 格式（`AGENTS.md` §1）。
9. 阶段 0/1 的**实测结论必须写进本文 §二末尾的"本机实测"表**，未实测的假设不得进入代码。

---

## 一、总目标

在**现有的 Electron 桌面 app**（`apps/desktop`）里增加第二种文档形态：`.xlsx` 工作簿。用户打开工作簿后，
可以像现在对 Word 文档那样用自然语言让 AI 干活：

- **分析数据**：这个表有什么、每列什么类型、缺失多少、分组汇总、找异常值、回答"哪个月销量最高"这类问题。
- **数据整理**：清洗表头、去重、拆列/合列、统一日期与数字格式、按条件排序、把散列转成规范的一列一条记录。
- **数据计算**：新增计算列/汇总行，写入的是**Excel 公式**（打开文件后 Excel 能自己重算），本地同时写缓存值。
- **生成图表**：AI 选定数据区域 + 图型 → 出图；图在 app 里能看，写进 `.xlsx` 后 Excel 里**是可编辑的原生图表**（不是贴图）。

同时保留用户手动操作：表格视图能直接改单元格、插入/删除行列，改动可撤销、可查修改历史。

**非目标**（明确不做，防止跑偏）：

1. 不做协作 / 冲突合并；单实例本地文件。
2. 不做完整公式引擎（400+ 函数、依赖图重算、循环迭代）。本地只求"显示所需的值"，重算交给 Excel/WPS。
3. 不做透视表（PivotTable）建模——按全局约束 6 原样保留。
4. 不做 VBA / 宏执行。含宏的 `.xlsm` 第一阶段直接拒绝打开并提示另存为 `.xlsx`。
5. 不做跨工作簿连接、Power Query、外部数据源刷新。
6. 不做实时渲染保真的**打印**排版（分页/缩放/页眉页脚）；`.xlsx` 的 PDF 导出留到最后的可选阶段。
7. 不做任务窗格 Excel 加载项（`apps/task-pane` 保持 Word）。它技术上是可复用的后端工具回路，但属于独立交付物。

---

## 二、选型研究结论（先取证，后写码 —— `AGENTS.md` §6.1/§6.2）

以下版本号 / 许可证为本机 `npm view` 输出（2026-10-12，`registry.npmjs.org`），不是文档转述。

| 候选 | 事实 | 结论 |
| --- | --- | --- |
| **SheetJS `xlsx`** | npm 上停在 `0.18.5`（Apache-2.0，依赖 cfb/ssf/wmf/crc-32/codepage），社区版早已迁移到自己的 CDN（`xlsx@0.20.2` 需从 `cdn.sheetjs.com` 装，`node-xlsx@0.24.0` 就是这么依赖它的） | **否决为主 IO 层**：0.20.2 不在 npm 意味着我们的 `pnpm install` 要么走第三方 URL 依赖（不可可重现构建），要么停在 0.18.5 的旧解析器。更关键的是它**重序列化整个包**，读不懂的结构会被丢掉——与全局约束 6 直接冲突 |
| **`exceljs@4.4.0`**（MIT） | 依赖 jszip/saxes/archiver/fast-csv 等 9 个包；不写原生图表；维护节奏慢 | **否决为运行时依赖，采用为夹具生成器**：用它（只在本机 `tmp/`）产出"真实世界写法"的 `.xlsx`，作为我们 reader 的对照样本（见 §三 阶段 0） |
| **Hyperformula `3.5.0`** | **`GPL-3.0-only`**（依赖 chevrotain） | **否决**：要分发的 Electron app 引入 GPLv3 依赖会把整个 app 的许可证拖进传染范围（`AGENTS.md` §8 安全）。这是本次调研里最硬的一条否决 |
| **Univer `@univerjs/*@1.0.4`**（Apache-2.0） | `@univerjs/charts`、`@univerjs/screens` 在 npm 上 **404 不存在**；社区报道与官方定位都把图表/透视表留在企业版 | **否决**：① 图表（我们的核心需求之一）不在开源包内；② 它自带渲染层与样式体系，违反全局约束 3；③ 它有自己的工作簿模型和序列化，等于把"未知部件保留"这条红线交给第三方 |
| **`@visactor/vtable@1.26.8`** / **`ag-grid-community@36.2.0`** / **`@glideapps/glide-data-grid@6.0.3`**（均 MIT） | 只是网格/表格可视化组件，不含 OOXML IO；vtable 自带主题系统与 canvas 绘制 | **否决作为文档模型宿主**；`ag-grid-community` 可在阶段 2 之后作为**可选**网格替换候选（MIT、纯 DOM、可被我们的 CSS 覆盖）。第一版网格自己写，理由见决策 D2 |
| **SpreadJS（葡萄城）/ Fortune-sheet** | SpreadJS 商业授权（npm 上取不到公开包）；`fortune-sheet` **已于 2022-06-19 从 npm 撤包** | **否决**：一个要钱，一个已死 |
| **ECharts `6.1.0`**（Apache-2.0，依赖 zrender/tslib） | 成熟、支持 SSR 到 canvas、按需引入可控体积 | **采纳为图表渲染/交互层**（app 内查看与出图），但 `.xlsx` 里的图表**不用它**——见决策 D5 |

一手来源（规格与 API 形态以这些为准）：ECMA-376 Part 1 的 `18.2 SpreadsheetML / Sheet Data`、`18.6`（sharedStrings）、`18.5`（styles/numFmtId）、`20.3 Drawing`、`21.2 Chartparts`；Microsoft Learn 的 *Workbook part and structure* 与 *Introduction to the Office Open XML file format*；`@rgrove/parse-xml` 与 `jszip` 的仓库 README。

**本机实测（2026-10-12，`tmp/xlsx-probe/`：exceljs 4.4.0 写出 7174 字节的样本工作簿 → jszip 解包 → `@rgrove/parse-xml@4.2.3`（与 `packages/docx-io` 同版本）解析）**

| # | 待验证问题 | 实测结论 | 证据 |
| --- | --- | --- | --- |
| P1 | 真实 `.xlsx` 的 OPC 部件清单与关系链 | 一个"最小但功能齐全"的工作簿 = **10 个文件部件**：`[Content_Types].xml`、`_rels/.rels`、`docProps/app.xml`、`docProps/core.xml`、`xl/workbook.xml`、`xl/_rels/workbook.xml.rels`、`xl/sharedStrings.xml`、`xl/styles.xml`、`xl/theme/theme1.xml`、`xl/worksheets/sheet1.xml`（另有 4 个目录项）。`xl/worksheets/_rels/sheetN.xml.rels` **只在有超链接/图片/图表时存在**；`sharedStrings.xml` 在无字符串时可缺 | `tmp/xlsx-probe/run.log` "P1 parts" 段 |
| P2 | `sharedStrings` 的富文本 `<r>` run 结构 | `<si>` 有两种形态：`<si><t>月份</t></si>` 与 `<si><r><rPr><b/></rPr><t>粗</t></r><r><rPr><color rgb="FF0000FF"/></rPr><t>蓝</t></r></si>`。**拼接所有 `<t>` 得到无损纯文本**，`rPr` 是 run 内格式。`<sst>` 上 `count="12" uniqueCount="11"`：`count` 是**引用次数**、`uniqueCount` 才是 `<si>` 条目数——我们重建 string 表时只能用后者做基数校验 | 同上 "P2 sharedStrings.xml" 行 |
| P3 | 日期/数字/公式在 `sheet1.xml` 里的真实形态 | 日期 = **纯序列号 + 样式决定显示**：`<c r="D2" s="2"><v>46053</v></c>`（2026-01-31 的序列号，与 `1899-12-30` 起算的推算值逐位相符；**没有 `t` 属性**，与数字无法从 `<c>` 本身区分）。`s` 是 **cellXfs 的位置索引**：`s="2"`→`numFmtId="164"`（自定义 `yyyy-mm-dd`），`s="5"`→`numFmtId="14"`（内建日期），`s="3"`→`numFmtId="10"`（百分比），`s="1"`→`numFmtId="3"`。自定义格式在 `<numFmts>` 里（`numFmtId>=164`）。公式 = `<c r="G2"><f>SUM(C2:C4)</f><v>360</v></c>`，**与决策 D4"写公式同时写缓存值"完全一致**。空但带样式的格会保留成 `<c r="H2"/>`，所以稀疏模型必须能存"只有格式没有值"的格。颜色一律 **ARGB 8 位十六进制**（`<color rgb="FFFF0000"/>`、`<color theme="1"/>` 两种形态并存） | 同上 "P3 sheet1.xml"、"P3 cellXfs" 行；`node -e` 取 `styles.xml` 的 `numFmts/cellXfs/fonts` |
| P4 | 我们的 writer 产出的 `.xlsx` 能否被 WPS 正常打开并保留未建模部件 | `[!] BLOCKED`——本机 WPS 无可视验证通道（`AGENTS.md` §8、`docs/word-parity-spec.md` §8.12）。**替代证据形式**：读回解析一致 + OOXML 结构断言 + 用第三个独立库（exceljs `readXlsxFile`）读我们的产物作为交叉验证 | 阶段 1 的单测 |
| P5 | `@rgrove/parse-xml` 处理带前缀命名空间的部件 | 三个部件全部解析通过：`sheet1.xml`（`xmlns` + `xmlns:r` + `xmlns:mc` + `mc:Ignorable="x14ac"` + `xmlns:x14ac`）、`sharedStrings.xml`、`styles.xml`。属性里的 `x14ac:dyDescent` 这类扩展属性**会原样出现在 attributes 里**，所以"重排/重建元素"就会丢它们 → 见下面 P5 派生出的硬约束 | 同上 "P5 parse OK …" 三行 |
| P6 | ECharts 离屏 canvas → PNG 在 Electron 里的可行性 | 留到阶段 5 用最小探针验证；app 内跑不了 CDP harness（`AGENTS.md` §8），届时按 `css-verification-without-electron` 的先例用独立 Chromium 探针出证据，主流程仍标 BLOCKED | `[ ]` |

**P5 派生出的硬约束（比原全局约束 6 更强，写在下面替代它）**：未建模结构不只存在于"整个部件"层面，
**同一部件内部**也有——实测样本的 `sheet1.xml` 里就有 `<conditionalFormatting>`、`<dataValidations>`、`<pageMargins>`、`<pageSetup>`、
`<sheetFormatPr x14ac:dyDescent>`、`<row spans x14ac:dyDescent>`。因此"解析成对象再重新序列化整篇"**不成立**，
必须像 `docx-io` 那样持有 `sourceXml` 并做**定点替换**（`packages/docx-io/src/docx.ts:455 replaceProseMirror` 保留根属性的先例）：
`sheet-io` 只重写 `<sheetData>` 区间与 string/style 表，其余元素与属性原样搬运；`cellXfs`/`<numFmts>`/`<fonts>` 只**追加不重排**
（`s="2"` 这类位置索引一旦重排，用户既有格式会整体错位——这是本文件里代价最高的一类错误）。

---

## 三、架构决策

### D1 — 新增 `packages/sheet-io`，不扩写 `packages/docx-io`

理由（`AGENTS.md` §3.4 要求新增包必须先说明）：`docx-io` 的职责是 **WordprocessingML ↔ ProseMirror**
（`packages/docx-io/src/index.ts` 导出 `ooxmlToPm/pmToOoxml/DocxFile` 与全部单位换算），它的数据模型是 PM 节点树。
工作簿的模型是**稀疏单元格网格 + 字符串表 + 样式表 + 分节维度**，与 PM 树没有任何结构共享点。
硬塞进 `docx-io` 会让一个包同时服务两种文档、两套不变量，反而制造"第二套换算藏在同一个包里"的风险。
真正**复用**的部分是依赖（`jszip`、`@rgrove/parse-xml`）与写法纪律（进出口换算、快照/还原、原子写）。

包边界：`sheet-io` 不依赖 `desktop`，不依赖 `word-tools`；`sheet-tools` 只依赖 `shared` + `zod`（与 `word-tools` 同构）。
依赖方向保持 `apps → packages`（`AGENTS.md` §5.1）。

### D2 — 网格 UI 自己写一版，文档状态放 store，不放在组件里

网格是**虚拟滚动的 DOM 表格**（固定表头 + 行高列宽 + 选区矩形 + 就地编辑一个 `<input>`），复用既有
`hooks/usePointerDrag.ts`（列宽拖拽，`AGENTS.md` §5.6）与 `src/styles/*.css`（§5.4）。
理由：本项目的一切能力都建立在"文档 = 结构化数据 + op 流"上（`src/editor/docBlocks.ts:548 executeOps`），
第三方网格库把文档状态握在自己肚子里，AI 提案→模拟→撤销这条链就接不上。
体积/复杂度可控（一个可见区 ≤ 2 千格），且是唯一能满足全局约束 3 的路径。

### D3 — 单元格地址用"列字母+行号"，但**撤销地址走行身份指纹**

`setCell("B7", ...)` 这类 op 在"插入一行"之后会指向别的内容，直接违反 `AGENTS.md` §2.5
（撤销按内容寻址，不许牵连用户后来的手工编辑）。因此沿用 `docEdit.ts:540 addressOf` 的思路：
每行算一个内容指纹（行内值序列 hash）作为身份，op 记录 `address`（给用户看）与 `key`（给撤销用），
撤销时按 `key` 重新定位；定位不到就明确报"部分撤销"，不许猜。

### D4 — 公式：写公式 + 写缓存值，不实现重算引擎

AI 被要求"加一列同比增长率"时，落到 `.xlsx` 的是 `<c><f>(C2-C1)/C1</f><v>0.12</v></c>`：
`<f>` 让 Excel/WPS 打开后能重算（用户拿到的是活公式，不是一堆死数字），`<v>` 让 app 里能显示。
本地只对**用到的少数函数**（`SUM/AVERAGE/MIN/MAX/COUNT/COUNTIF/SUMIF/ROUND/ABS/IF/CONCAT`，以及算术表达式）
实现一个求值子集，用于校验公式可解析 + 计算缓存值；其余函数写 `<f>` 但 `<v>` 留空并在 UI 明示"该格需 Excel 重算"。
这条决策把"GPL 依赖"和"重算引擎工程量"两个坑同时绕开，代价是本地看到的可能不是最新重算值——UI 必须如实标注。

### D5 — 图表分两层，`.xlsx` 内必须是原生 chart part

- **app 内**：AI 产出 `ChartSpec`（数据区域 + 图型 + 编码），用 ECharts 渲染/交互；导出 docx/PDF 时 canvas→PNG 插入（复用 `src/ai/diagramCanvas.ts:136 renderDiagram` 的现成思路）。
- **写进 `.xlsx`**：生成 OOXML 原生图表部件（`xl/charts/chart1.xml` + `xl/drawings/*` + 关系），这样在 Excel 里图表是**可编辑对象**并随数据更新。
- **风险与门槛**：chart part 是 ECMA-376 里最繁琐的部分之一。阶段 0 必须先做一个最小 spike（一张柱状图，手工拼 chart XML + 关系 + drawing，读回断言，`tmp/` 里，不进主干 —— `AGENTS.md` §6.5），
  spike 失败则图表降级为"贴图插入工作表"，并**在本文档与 README 明确记录降级**，不许静默把贴图说成图表。

### D6 — AI 侧：新增 `sheet-tools` 工具面 + 桌面 `sheetAgent`，复用面板状态机

工具分三类，全部 zod `.strict()`：
- **读（只回数据，不改文档）**：`sheet_describe`（表结构+列类型+缺失统计）、`sheet_read_range`（区域取数，受 §2.9 截断）、`sheet_aggregate`（分组聚合，值由引擎算）、`sheet_find`（按条件定位行）。
- **写（产出提案 op）**：`sheet_set_cells`、`sheet_insert_rows`、`sheet_delete_rows`、`sheet_sort_rows`、`sheet_add_column`、`sheet_set_number_format`、`sheet_dedupe`、`sheet_create_chart`。
- **提交**：`sheet_submit_proposal`（与 `word-tools` 的 `word_submit_proposal` 同构，`packages/word-tools/src/index.ts:162 editsFromProposal` 的对应物）。

桌面侧新增 `apps/desktop/src/ai/sheetAgent.ts`（提示词 + 区域截断 + 提案装配），与既有 `agent.ts:436 requestProposal` 平行；
`AIPanel.tsx` 按文档种类分支装配，卡片状态机（`PlanCard{summary,ops,before,after,inverses,status}`，`AIPanel.tsx:67`）、
确认/暂停/继续（`applyPlan:414`、`ai:stream-cancel`）全部复用，不复制一份。

---

## 四、阶段目标

> 每阶段一句"完成判定"：能演示什么、能提交什么。阶段之间串行，阶段内任务可并行。

### 阶段 0 · 取证与夹具（不写产品代码）— `[x] 已完成 2026-10-12`

**目标**：把 §二 P1–P6 全部填成实测结论；产出可重复生成的 `.xlsx` 夹具。
**完成判定**：`tmp/` 里有真实世界写法的样本工作簿；`packages/sheet-io/scripts/make-sample-sheet.ts`（对齐
`packages/docx-io/scripts/make-sample.ts`）能产夹具；本文 §二 表格里 P1–P5 从 `[ ]` 变成带证据的结论；P4 明确标 BLOCKED 及替代证据形式。
**实际结果**：P1/P2/P3/P5 已有带路径的实测结论（§二）；P4 按预期 BLOCKED 并定下替代证据形式；P6 顺延到阶段 5。
夹具用 `tmp/xlsx-probe/probe.mjs` 现场生成（`AGENTS.md` §6.5：探针不进主干），主干侧的 `make-sample-sheet.ts` 随阶段 1 首个提交落地。
额外收获两条：① 未建模结构在**部件内部**（强化了 §零.6）；② `exceljs@4.4.0` 自己写 `dataBar` 条件格式会抛
`cfvo undefined`（实测），说明它的写路径不可当作"OOXML 正确性"的裁判，只能当夹具。

### 阶段 1 · `packages/sheet-io`：读 + 保真回写 — `[x] 已完成 2026-10-12`

**目标**：`SheetFile.fromBytes()` → `toSheetDoc()` → 不改任何东西再 `toBytes()`，**输出与原文件在所有已建模与未建模部件上等价**。
**完成判定**：往返单测绿（未建模部件逐字节、sharedStrings 条目数、维度、合并区域数不变）；类型/日期/数字读回正确；
`.xlsm` 与含外部链接的包给出明确拒绝或"原样保留"结论。**此时还没有 UI。**

**实际结果**：`packages/sheet-io` = `src` 下 10 个源文件 + 4 个测试文件，25 条单测全绿（`address` 5 / `datetime` 4 / `color` 3 / `sheet-file` 13）；
根 `pnpm build` / `pnpm test` 通过（sheet-io 24、docx-io 296、desktop 398、word-client 8、word-tools 6、server 2；shared/cli 仍无测试文件）。

落地的部件（与原计划文件的对应）：`opc.ts`（解 zip + 关系图 + `.xlsm` 拒绝）、`model.ts`/`address.ts`/`datetime.ts`/`color.ts`
（唯一换算入口）、`xlsx-strings.ts`/`xlsx-styles.ts`（**只追加不重排**的 string 表与 `cellXfs`）、`xlsx-sheet.ts`
（持有 `sourceXml`，写回时用正则**定点替换 `<sheetData>` 与 `<dimension>`**，其余元素与扩展属性原样搬运）、`sheet-file.ts`
（惰性解析 + 快照/还原 + 只覆写变化过的部件）。夹具生成器 `scripts/make-sample-sheet.ts` 的 XML 逐字抄自 §二 的实测样本。

**交叉验证（P4 的替代证据，`tmp/xlsx-probe/cross-read.mjs`）**：我们的 writer 产出 5508 字节 → 第三方独立库 exceljs 4.4.0 读回：
工作表名 `销售/空表`、`A6="新增项目"`（新增共享字符串）、`C6=300`、`G6` 公式 `SUM(C2:C6)` 缓存值 `660`、
`D2` 日期 `2026-01-31` 且 numFmt `yyyy-mm-dd`、我们**新追加**的 `E6` 自定义格式 `yyyy"年"m"月"`（numFmtId 165）被正确识别、
被删的 `C4=null`、合并 `["H1:I2"]`、冻结 pane 完整、数据验证 `list "华东,华北"`、富文本两 run（粗体 + 蓝色）全部读回。
含跨簿引用 `'[1]Sheet1'!$A$2` 的格与 `xl/externalLinks/externalLink1.xml`（及其 `TargetMode="External"` 关系）也被第三方读回，包部件 13 个一个不少。
`<conditionalFormatting>` 在产物 XML 里逐字保留（`unzip -p cross.xlsx xl/worksheets/sheet1.xml` 断言），但 exceljs 的
`getConditionalFormattings` 返回 null —— 这是**该库读路径的局限**，不是产物缺陷（它自己写 `dataBar` 会抛异常，见 §二 阶段 0 额外收获 ②）。

**与执行计划的偏差（记录，不算违背 §1.4）**：§五 把阶段 1 拆成 5 个提交，实际合成 1 个 `feat(sheet-io)`。
原因是按文件切分会让中间提交处于"引用了还不存在的模块"的状态、`pnpm build` 在该提交上红——
§1.6 要求每个提交都可推送且远端复核，所以选择"一个包 = 一个可编译提交"。spec 回填单独一个 `docs` 提交。

### 阶段 2 · 表格视图可用（人手动编辑）

**目标**：桌面 app 打开 `.xlsx` 能看能改：滚动、选区、改单元格、插入/删除行列、撤销重做、保存回 `.xlsx`。
**完成判定**：`pnpm dev:desktop` 手工走通"开→改→撤销→存→重新打开内容一致"；打开/保存/文件关联（`fileAssociations` 加 `xlsx`）
与"最近文件"打通；保存后读回断言一致；类型检查 + 全量测试绿。

### 阶段 3 · 结构化 op + 撤销/modlog 接入

**目标**：把"改单元格/改行列/改格式"收敛成 CoreOp 集合，每个 op 有逆操作，走 `planEdit/runOps/revert` 同一条链路，落 `userData/modlog`。
**完成判定**：`sheetEdit.test.ts` 覆盖：op→模拟→落地→撤销内容一致；插入行后再撤销不牵连用户后续手工编辑；无法安全回滚时明确报"部分撤销"。

### 阶段 4 · AI 工具层：分析 + 整理 + 计算

**目标**：自然语言 → 工具回路 → 提案 → 用户确认 → 落地。覆盖分析（描述/聚合/找异常）、整理（去重/拆合列/统一格式/排序）、计算（新增计算列写公式）。
**完成判定**：`sheet-tools` 全部入参有 zod 单测；桌面 `sheetAgent` 用 mock `fetch` 测过工具环与截断；
用真实端点探针跑 ≥3 类任务（分析/整理/计算）并记录 token 与失败样本；产出的 `.xlsx` 读回一致且 Excel 能重算公式。

### 阶段 5 · 图表

**目标**：AI 选区出图；app 内可看（ECharts），`.xlsx` 内为原生可编辑图表（D5 spike 成功）或明确记录的贴图降级。
**完成判定**：`ChartSpec` 校验与渲染单测；写出的 chart part 读回结构断言；导出 PDF/docx 报告时图以 PNG 出现；WPS/Excel 可视验证标 BLOCKED 并给出结构级替代证据。

### 阶段 6 · 收口：CLI、文档、规则落地

**目标**：`apps/cli` 增加 `ocs`（或 `oc sheet`）子命令复用 `sheet-io`+分析引擎；`README.md` 增补 Excel 小节；把 §一~§三 新增的不变量补进 `AGENTS.md`；
可选：`sheet:*` 的 docx 报告嵌入（把表格区域作为 Word 表写入，复用 `DocxFile.insertTable`）。
**完成判定**：三端（desktop/cli/docs）对同一份 `.xlsx` 行为一致；`AGENTS.md` §9 表格里新增项标清"已落地/待落地"。

---

## 五、执行计划（任务 → 落点）

| 阶段 | 任务 | 文件落点 | 粒度 |
| --- | --- | --- | --- |
| 0 | P1/P5 部件清单与解析试验：用 exceljs 生成样本 → 解包 → 打印 parts/rels | `tmp/`（不进主干，`AGENTS.md` §6.5） | 1 提交（仅结论进本文） |
| 0 | 夹具生成器 | `packages/sheet-io/scripts/make-sample-sheet.ts` | 并入阶段 1 首个提交 |
| 1 | 新包骨架：`package.json`/`tsconfig`（**不建 `vitest.config`**：同类包 `packages/docx-io` 就没有，无 CSS 的包再建一套是平行基建，`AGENTS.md` §3.4；将来若引 CSS 必须带 `css.postcss.plugins: []`，§8） | `packages/sheet-io/**` | 提交 1（实际与 2–5 合为 1 个提交，见阶段 1 偏差说明） |
| 1 | OPC 打开：解 zip、`[Content_Types].xml`/`workbook.xml`/`_rels` 关系图、拒绝 `.xlsm` | `src/opc.ts` | 提交 2 |
| 1 | 模型：`SheetDoc`/`Cell`/`NumberFormat`/`styleToken` 与日期序列号换算（唯一入口）；ARGB(`FFFF0000`) ↔ CSS(`#rrggbb`+alpha) 只在此处（实测 §二 P3：`.xlsx` 里是 8 位 ARGB 或 `theme` 引用） | `src/model.ts` `src/datetime.ts` | 提交 3 |
| 1 | reader：`sharedStrings`（含 `<r>` run）、`sheetData`、`dimension`、`mergeCells`、`cols/rows`、`styles.xml` numFmt 定位 | `src/xlsx-sheet.ts` `src/xlsx-styles.ts` | 提交 4 |
| 1 | writer：`replaceSheetDoc` + **未建模部件原样回写** + `getSnapshot/restoreSnapshot` | `src/sheet-file.ts` | 提交 5 |
| 1 | 往返单测（真实夹具 + 手工最小样本） | `src/*.test.ts` | 随各提交 |
| 2 | 主进程按扩展名分派；新增 `sheet:*` handler；原子写 + `.bak` | `electron/fileHandlers.ts` `electron/main.ts:32`（`DOCX_EXT` 旁路）、`package.json fileAssociations` | 提交 6 |
| 2 | preload 白名单加 `sheet:*`（同时改类型，禁万能 invoke） | `electron/preload.ts:222/224` | 提交 6 |
| 2 | store 与网格组件 | `src/sheet/sheetStore.ts` `SheetGrid.tsx` `CellEditor.tsx` | 提交 7 |
| 2 | 样式（唯一 CSS 体系，不新增入口） | `src/styles/sheet.css`（在既有入口引入） | 提交 7 |
| 2 | 列宽拖拽走 pointer capture；浮层走 fixed + `useAnchoredMenu` | 复用 `hooks/usePointerDrag.ts` `hooks/useAnchoredMenu.ts` | 提交 8 |
| 3 | `SheetCoreOp` 类型、`planEdit/runOps/revert` 的表格实现、行内容指纹 | `src/ai/sheetEdit.ts` | 提交 9 |
| 3 | 撤销与修改历史接入（复用 modlog，不新建存储） | `electron/modlog.ts` 现有通道 | 提交 9 |
| 4 | 分析引擎（列类型推断、缺失统计、分组聚合、去重键、异常值） | `packages/sheet-io/src/analysis/**` | 提交 10 |
| 4 | 公式子集求值（D4 的函数白名单）+ 缓存值写入 | `packages/sheet-io/src/formula/**` | 提交 11 |
| 4 | 工具 schema 与 JSON 工具定义 | `packages/sheet-tools/src/index.ts` | 提交 12 |
| 4 | 桌面 agent：提示词、区域截断、提案装配、卡片渲染分支 | `src/ai/sheetAgent.ts` `src/ai/AIPanel.tsx` | 提交 13 |
| 5 | `ChartSpec` + ECharts 渲染 + canvas→PNG | `packages/sheet-io/src/chart-spec.ts` `src/sheet/ChartView.tsx` | 提交 14 |
| 5 | chart part 写入（`xl/charts` + drawing + rels + content-types 更新） | `packages/sheet-io/src/xlsx-chart.ts` | 提交 15 |
| 6 | CLI 子命令、README、AGENTS 更新 | `apps/cli/src/**` `README.md` `AGENTS.md` | 提交 16 |

推进纪律：每阶段结束跑一次 `pnpm build > tmp/b.log 2>&1; echo EXIT=$?` 与 `pnpm test > tmp/t.log 2>&1; echo EXIT=$?`
（**重定向后再读退出码**，`AGENTS.md` §8 的假绿），并逐包核对通过数（`apps/cli`、`packages/shared` 本就无测试）。
每次改动完成立即提交并推送 `origin`，推送后 `git status -sb` + `git ls-remote origin main` 双重复核（`AGENTS.md` §1.6）。

---

## 六、验收标准列表（逐条打分，任一不过不得声称完成）

### A. 文档正确性（最高优先级，对应 `AGENTS.md` §2）

1. `[x]` **空改动往返**：读入 → 立刻写出，与原文件在所有部件上语义等价；未建模部件（条件格式、数据验证、透视表、外部链接、图表、主题）逐字节保留。
   **实测**：`sheet-file.test.ts`「打开后不改任何东西，写出的每个部件与原文件一致」把 13 个部件（含 `xl/externalLinks/externalLink1.xml` 与它的 rels）逐条 `toEqual` 原文。
2. `[x]` 打开真实世界工作簿不丢文本：含富文本 `<r>` run 的 sharedStrings 读出为纯文本时，写回后**内容不丢**（格式丢失需给出 warning，不许静默）。
   **实测**：夹具里那条两 run 的 `<si>`（粗体 + `FF0000FF` 蓝）读成纯文本 `粗蓝`，写回时整条 `rawSi` 原样搬（string 表只追加、不改已有条目）；第三方 exceljs 读回 `richText:[{bold:粗},{蓝}]`。
3. `[x]` 日期/数字读回正确：序列号 ↔ 日期只在 `sheet-io` 进出口换算一次；同一文件反复打开得到同一显示值。
   **实测**：`46053 → 2026-01-31`（`datetime.test.ts` 锚定真实样本序列号，含 1900 闰年 bug 的 `dayParts<=60` 分支）；自定义 `numFmtId=164` 与内建 `14` 都判成日期而值仍是序列号；`grep -rn "\bserial\b|serialTo|ToSerial|25569|epochDays" apps packages`（排除 `dist/`）在 `packages/sheet-io` **之外零命中**（本机实跑 exit 1），§零.2 的单点成立。
4. `[~]` 公式写回后 **Excel/WPS 打开能重算**（读回断言 `<f>` 存在且合法，`<v>` 为缓存值）；无法本地求值的函数在 UI 明示"需 Excel 重算"，不许显示过期数字当作结果。
   **阶段 1 到位的部分**：`<f>`+`<v>` 成对读写，跨簿引用 `'[1]Sheet1'!$A$2` 不被改写或多余转义；exceljs 读回公式 `SUM(C2:C6)` 与缓存值 `660`。
   **未到位**：“UI 明示需 Excel 重算”属阶段 2/4，本阶段没有 UI。
5. `[ ]` 保存走 `.tmp → rename`，原地写回保留 `.bak`；保存中断不会留下半个文件。
6. `[x]` `.xlsm` / 含外部连接的包：按阶段 1 的结论明确拒绝或原样保留，二者之一并有测试，不许"看起来成功"。
   **实测结论**：含 VBA 宏（`xl/vbaProject.bin` 或宏 content-type）**明确拒绝**并给出可读消息；外部链接/透视表/图表/绘图**原样保留**并记进 `unsupported` 交给 UI 提示。各有单测（宏拒绝、`externalLinks` 部件逐字节 + `[1]` 引用活着）。
7. `[x]` 单位/颜色换算单点：全仓 grep 不出现第二处 serial/numFmt/颜色换算实现（§零.2）。
   **实测**：serial/numFmt 判定只在 `src/datetime.ts`，ARGB ↔ `#rrggbb` 只在 `src/color.ts`，列宽只存 `widthChars` 原样搬运（有单测钉住不做第二处换算）。
8. `[!] BLOCKED` 用 Excel 或 WPS 实际打开一遍：本机 WPS 无可视验证通道（`AGENTS.md` §8、`docs/word-parity-spec.md` §8.12）。**已做到的替代证据** = 自读一致 + OOXML 结构断言 + **第三方独立库 exceljs 4.4.0 交叉读**我们的产物（`tmp/xlsx-probe/cross-read.mjs`，结论见阶段 1 实际结果）；此条**必须**如实标 BLOCKED，直到有人在装了 Excel 的机器上打开一遍。

### B. AI 行为（对应 §零.4/§零.5）

9. `[ ]` 所有表格工具入参过 zod `.strict()`，非法入参返回给模型的是错误串而不是崩溃（对齐 `apps/server/src/agent.ts:117` 的做法）。
10. `[ ]` **数字来自引擎**：单测固定住"聚合值由 `sheet-io` 计算"，模型给出的数字若与引擎不符，以引擎为准并记 warning。
11. `[ ]` 上送内容默认截断（表头 + ≤50 行样本 + 结构摘要），有单测锁死这个上限。
12. `[ ]` 半自动模式：提案必须用户点确认才落地；全自动模式的每步仍受 §2.5 撤销约束。
13. `[ ]` 真实端点探针（一次性、`tmp/`）跑通"分析 / 整理 / 计算"各 ≥1 个任务，记录：模型实际调用的工具序列、失败样本、未转义引号这类解析坑（本仓已有 ark 回包先例）。
14. `[ ]` 暂停/继续行为与 Word 一致：中断后半截提案留在卡上，点"继续"能接着写。

### C. 可撤销性与历史（对应 `AGENTS.md` §2.5）

15. `[ ]` 每个 `SheetCoreOp` 都发出自己的逆操作；撤销按行内容指纹定位。
16. `[ ]` 插入/删除行列后再撤销早期改动，**不牵连**用户在这期间的手工编辑（用一条专门单测钉死）。
17. `[ ]` 无法安全回滚时明确报"部分撤销"；modlog 仍每文档 ≤60 条、路径不变。

### D. UI 与工程（对应 `AGENTS.md` §3/§4/§5）

18. `[ ]` 没有复制出第二套撤销栈 / LLM 客户端 / 设置存储 / 样式体系 / SQLite 连接。
19. `[ ]` 网格滚动与编辑在 5 万行 × 20 列的工作簿上不卡死（虚拟滚动的验收口径：初始打开 ≤ 2s，滚动不重建全部 DOM；具体阈值在阶段 2 实测后写回本节）。
20. `[ ]` 列宽拖拽用 `usePointerDrag` 的 pointer capture，热区带 `-webkit-app-region: no-drag`；浮层用 fixed + `useAnchoredMenu` 不出屏。
21. `[ ]` 渲染层只经 `window.officeCopilot` 的白名单通道访问主进程；无 `nodeIntegration` 变化。
22. `[ ]` 新函数有中文 JSDoc（作用/入参含单位语义/失败语义）；命名有含义，无 `data/res/temp/flag`。
23. `[ ]` 无孤儿文件、无注释掉的代码块、无未被调用的导出。

### E. 测试与门禁（对应 `AGENTS.md` §7）

24. `[x]` `pnpm build` 全绿（类型检查藏在 build 里，`pnpm test` 不能替代）。
   **实测**：`pnpm build > tmp/xlsx-probe/root-build.log 2>&1` → `BUILD_EXIT=0`，日志含 `packages/sheet-io build: Done`。
25. `[x]` `pnpm test` 全绿，且**逐包核对通过数**（新包 `sheet-io`/`sheet-tools` 必须有测试文件，不许靠 `passWithNoTests` 蒙过）。
   **实测**：`TEST_EXIT=0`；逐包核对 —— sheet-io **25**、docx-io 296、desktop 398、word-tools 6、word-client 8、server 2；`packages/shared`/`apps/cli` 仍 “No test files found”（本仓既有状态，不是本次引入的假绿）。
26. `[ ]` `pnpm dev:desktop` 起真实 app 手工走通主流程，收尾报告看到什么；跑不起来就如实说明卡点，不许用单测冒充可视验收。
27. `[x]` 测试产物只落 `tmp/` 且清理；不新增任何二进制夹具入库（`AGENTS.md` §7.3）。
   **实测**：夹具由 `scripts/make-sample-sheet.ts` 在测试里现生成，仓库内无 `.xlsx`；探针与日志全在 `tmp/xlsx-probe/`（未跟踪）。
28. `[ ]` 提交信息中文、`<type>(<scope>)`、一事一提交，功能与格式化不夹带；已推送并复核远端。

---

## 七、风险与未决

| 风险 | 影响 | 缓解 |
| --- | --- | --- |
| 原生 chart part 手工生成失败率高 | 阶段 5 的核心卖点 | 阶段 0 就 spike，失败即降级为贴图并**显式记录**；不把它伪装成图表 |
| 未建模部件"逐字节保留"在 sharedStrings/styles 上做不到（它们互相引用，改一个值可能动 string 表） | 阶段 1 红线 | **已被实测消解**（§二 P2/P3/P5）：策略是定点替换 `<sheetData>` + string 表**只追加** + `cellXfs/numFmts/fonts` **只追加不重排**（位置索引重排会让用户格式错位）。单测必须钉住"改一个格不移动任何索引" |
| 大表全量载入内存 | 阶段 2 可用性 | 阶段 1 就定"稀疏 + 分 sheet 惰性解析"；超阈值（如 >20 万格）明确提示只读 |
| 本地公式求值与 Excel 结果不一致 | AI"数据计算"可信度 | 只支持白名单函数；UI 与 modlog 里标注缓存值来源；不一致的场景（循环引用、volatile 函数如 `TODAY()`）一律留空 `<v>` |
| 无法用 Word/WPS 目视验证 | 验收永远缺最后一环 | 沿用 Word 线的做法：结构级替代证据 + 明确 BLOCKED（`docs/word-parity-spec.md` §8.12） |
| 引入 Univer/SheetJS 的诱惑（"省几周"） | 破坏全局约束 3/6 | 本文 §二 已留否决理由；未来若要改选，必须先推翻 §零.6 再讨论 |

---

## 八、总验收表（随进度填写，键为提交哈希）

| 阶段 | 提交 | 判据编号 | 状态 | 实测/BLOCKED 说明 |
| --- | --- | --- | --- | --- |
| 0 | 本提交 | P1–P6 | `[x]`（P4 BLOCKED 已定替代证据，P6 顺延阶段 5） | 证据在 `tmp/xlsx-probe/run.log`（不入主干，`AGENTS.md` §7.3） |
| 1 | `e5d5ba4` | A1–A3、A6、A7、E24–E25、E27 | `[x]` | 25 条单测绿；第三方 exceljs 交叉读回值/公式缓存/日期 numFmt/合并/冻结/数据验证/富文本/外部链接部件全部一致 |
| 1 | `e5d5ba4` | A4 | `[~]` | `<f>`+`<v>` 读写已测；“UI 明示需 Excel 重算”待阶段 2/4 |
| 1 | — | A5、A8 | `[ ]` / `[!] BLOCKED` | A5 的 `.tmp→rename` + `.bak` 在主进程侧（阶段 2）；A8 本机 WPS 无可视通道（`AGENTS.md` §8），替代证据即上一条 |
| 2 | — | D19–D22 | `[ ]` | |
| 3 | — | C15–C17 | `[ ]` | |
| 4 | — | B9–B14 | `[ ]` | |
| 5 | — | D5 落地 | `[ ]` | |
| 6 | — | D18/E24–E28 | `[ ]` | |

---

## 附录 A · 现状盘点（本仓与表格相关的全部既有能力，带 `path:line` 证据）

- 无 xlsx/csv/excel 相关实现：全仓 grep 仅命中 `README.md:366` 路线图（"Excel/PPT：新增 `packages/excel-client`… 复用服务端工具环"）与 `package.json` 描述里的"Word/Excel/PDF (Word first)"。
- 无图表库依赖；唯一相邻能力是流程图渲染：`apps/desktop/src/ai/diagramSpec.ts` + `diagramCanvas.ts:136 renderDiagram`（canvas 2D → PNG dataURL）。
- Word 表格（非网格引擎）已打通，可作概念对齐：`shared` 的 `insert_table` 提案（`packages/shared/src/index.ts:36`）、
  `word-tools` 的 `InsertTableArgs`（`packages/word-tools/src/index.ts:35`）、`docx-io` 的 `DocxFile.insertTable`（`packages/docx-io/src/docx.ts:753`）、
  Tiptap Table 注册在 `apps/desktop/src/editor/extensions/editorExtensions.ts:62-65`。
- 复用的既有链路：`electron/fileHandlers.ts:245 openDocx` / `:311 saveDocx`（原子写）、`electron/preload.ts:222`（contextBridge 白名单）、
  `electron/aiHandlers.ts:177 runStream` / `:258 callLLM` / `:167 cancel`、`electron/settingsStore.ts:31`（atomic config）、
  `electron/modlog.ts:14 CAP=60`、`src/ai/docEdit.ts:540 addressOf` / `:946 planEdit` / `:1156 revert`、
  `src/ai/AIPanel.tsx:67 PlanCard` / `:414 applyPlan`、`src/ai/agent.ts:92 extractJson` / `:641 simulateProposal` / `:657 applyProposal`。
- 测试基建：vitest 全包；`live` 测试用真实 schema 的先例是 `apps/desktop/src/editor/docBlocks.live.test.ts:26`（新包应照此建立"用真实模型而非临时结构"的习惯）。
