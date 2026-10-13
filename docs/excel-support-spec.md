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
| P7 | 手工拼原生 chart part 是否可行（D5 的前提，原计划阶段 0 做 spike） | **可行，不降级**。本机用 openpyxl 3.1.5 往同样的数据上插一张柱状图（`tmp/chart-ref.xlsx`）拆出最小集：`[Content_Types].xml` 两条 Override（`drawingml.chart+xml` / `drawing+xml`）+ `xl/charts/chart1.xml`（`c:chartSpace>chart>title/plotArea/<kind>Chart+catAx/valAx`）+ `xl/drawings/drawing1.xml`（`xdr:oneCellAnchor>from/ext>graphicFrame>…>c:chart r:id`）+ `xl/drawings/_rels/drawing1.xml.rels` + `xl/worksheets/_rels/sheet1.xml.rels`，**五件少一件即判损坏**。两条由此定死的约束：① `CT_Worksheet` 只允许一个 `<drawing>`，已有画图部件（图片/别的图表）的表**明确拒绝**而不是追加；② `@Target` 必须相对源部件目录（`../drawings/drawing1.xml`），少一层就断链。落地后由**第三方读取器交叉验证**（替代 P4 的 BLOCKED）：`tmp/chart-out-{bar,line,pie}.xlsx` 用 openpyxl 读回，`charts=1`、`val='销售'!$C$2:$C$4` 缓存 `120/150/90`、`cat` 缓存 `粗蓝/2月/3月`、锚点 `(col 4,row 1)`、含 `<&"` 的标题原样取回 | `tmp/chart-ref.xlsx`、`tmp/chart-check.py` 输出、`packages/sheet-io/src/xlsx-chart.test.ts`（`294ce88`） |
| P8 | 只有 `xl/sharedStrings.xml` 部件本身够不够（阶段 5 由独立读取器抓出的缺陷，P1 的补强） | **不够**。来源表若全程用 `<c t="inlineStr">` 存文本（openpyxl 默认就这么写），包里本来没有 `xl/sharedStrings.xml`；我们第一次写文本会新建它，但只塞部件不够——必须在 `xl/_rels/workbook.xml.rels` 补一条 `Type="…/sharedStrings"` 的关系，并在 `[Content_Types].xml` 补 `/xl/sharedStrings.xml` 的 Override，两者少任一个：openpyxl 读回直接 `IndexError: list index out of range`（`shared_strings[int(value)]`），Excel 判文件损坏。原文件已有这张表时结构不动，只在"新建"这条分支补登记 | `tmp/chart-app.xlsx`（openpyxl 生成的 inlineStr 样本）、`tmp/sst-fix-probe.mjs`、`sheet-file.test.ts`「原文件没有字符串表时把新表登记齐全」（`fa693d2`） |
| P9 | "能不能被办公软件直接打开"这句话在本机有什么可用通道（用户明确要求：导出要能在 office Excel 里直接打开） | **三条，都不如桌面 Excel，但合起来能机检**。① **openpyxl 3.1.5**（`tmp/pylibs`，`PYTHONPATH=tmp/pylibs python3`）：独立读取器，只读值与结构，**不做合规判定**，我们写歪了它照样返回（P8 那次是它读回直接 `IndexError` 才暴露问题）。② **Numbers.app 走 AppleScript**：`tell application "Numbers" to open POSIX file "…"` 之后能问 `count of tables/charts of sheet 1`、`value of cell "A1" of table 1`，**这是本机唯一能拿到"某个办公软件真的把你的文件读成了什么"的通道**；实测 `tmp/chart-out-line.xlsx` → `tables=1 charts=1 A1=月份`。要拆成两次 `osascript` 调用（open 与查询合在一段会报 `-2741` 语法错误），查完 `close d saving no`。③ **自己按 ECMA-376 写结构审计**（`compat-audit.ts`，`6b770f5`）：Excel 打开时真正会查的那几项硬约束（内容类型覆盖、关系可解析、`r:id` 有落点、`<worksheet>` 子元素顺序）翻成 8 个问题码。**判据的对照实验**：故意删掉 `xl/drawings/_rels/drawing1.xml.rels` 并把 `pageMargins` 插到 `dimension` 之前，审计报 2 条（`unresolved-relationship-id` + `worksheet-child-order`），**同一个文件交给 Numbers → 打得开，但 `charts=0`**（图表被静默丢掉），而修补前的产物是 `charts=1`。所以两条各自都不够：Numbers 宽容（能开 ≠ 内容都在），审计只能保证"规范明写的硬约束没违反" | `tmp/make-broken-xlsx.mjs`、`tmp/check-ours.log`、`tmp/check-broken.log`、上面两组 Numbers 读数、`packages/sheet-io/src/compat-audit.test.ts`（`6b770f5`） |
| P10 | 我们的图表部件与"已知能在 Excel 里打开的实现"差在哪（P4/P9 的参照对拍） | **不比人家缺**。同一份数据让 openpyxl 生成 `tmp/xlsx-compat/ref-openpyxl.xlsx`，把两边的 `chart1.xml`/`drawing1.xml` 逐元素摊平对（`tmp/xlsx-compat/chart-xml-dump.txt`，291 行）：`chartSpace>chart>title/plotArea/<kind>Chart/ser(idx,order,cat,val)/axId×2/catAx/valAx/legend/plotVisOnly/dispBlanksAs` 两边都有，我们**多**写 `<c:marker val="1"/>`、`<c:delete/>`、`strCache`/`numCache`（带 `formatCode`）、`<c:spPr>`、`<c:varyColors>`，openpyxl 一个都没写也能开。由此确定：差异不在元素缺失，**风险在"引用与登记是否齐全"**（P7 的五件、P8 的三处），所以把力气花在 P9 ③ 的机检上而不是照抄 openpyxl | `tmp/xlsx-compat/chart-xml-dump.txt` |
| P11 | "打开→保存"这一趟有没有把用户没碰的东西弄瘦（P9 的审计判不出这个） | **拿 openpyxl 自己写的文件逐特性对拍**。`tmp/compat-sweep/gen.py` 造 10 份"每份只压一类结构"的夹具：合并+居中+边框、冻结+列宽+换行、日期与自定义 numFmt、跨行/跨表公式、条件格式+数据验证、超链接+批注+定义名、表格对象、inline 富文本+空表+只有格式的格、原生图表（顺带测"一张表只能一个 `<drawing>`"的拒绝分支）、行高+隐藏列+分组+打印设置。`roundtrip.mjs` 走**桌面的真实保存路径**（`fromBytes` → `sheetPayloadOf` → `applySheetPayload` → `writeSheet` → `toBytes`），`compare.py` 再让 openpyxl 分别读原文件与我们的产物逐字段比。结论：10 份的 P9 ③ 审计写前写后都是 0 条；**9 份逐字段一致，1 份抓到真缺陷**——`<c t="inlineStr"><is><r>` 的 run（粗体 + `FF0000FF`）被拍平成纯文本 `粗普通蓝`，还凭空新建了一张 `xl/sharedStrings.xml`，`bbc7eba` 按"未建模内容挂 `rawXml` 原样回写"这条既有口径修掉。教训如实记：**审计全绿不等于内容没变瘦**，合规判据与保真判据缺一不可 | `tmp/compat-sweep/{gen.py,roundtrip.mjs,compare.py,roundtrip2.log,compare2.log}`（§6.5 不入主干）、`packages/sheet-io/src/sheet-file.test.ts`「inline 富文本原样回写」两条 |

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
- **spike 结论（2026-10-13，`294ce88`）**：**成功，不降级**。写侧已落地成引擎能力 `resolveChartSpec` + `SheetFile.addChart`，
  产物由 openpyxl 3.1.5 读回交叉验证（数字见 §二 P7）。两条实测出来的边界进了实现：`CT_Worksheet` 只能有一个 `<drawing>`
  （已有图片/图表的表明确拒绝，返回 `drawing-exists`，不去重排别人的关系），关系 `@Target` 必须相对源部件目录（`opc.ts:relativeTarget`）。
  app 内的 ECharts 查看与 `sheet_create_chart` 工具面**还没接**，这是阶段 5 剩下的部分。

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
**实际结果**：P1/P2/P3/P5 已有带路径的实测结论（§二）；P4 按预期 BLOCKED 并定下替代证据形式；P6 顺延到阶段 5；
P7（原生 chart part 的 spike，本阶段最大的风险项）在阶段 5 的写侧提交里补做并**通过**，不降级（§二 P7、决策 D5）。
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

**实际结果**（`bbf23c5` 主进程按扩展名分派 + `sheet:*` 通道 → `ab27d0e` 几何与输入的纯函数层 → `c8c3f3b` 网格组件 + CSS + App 装配；
前置是 `dcd6a1f`/`e060525`/`7a0f606` 三个 `sheet-io` 提交）：

- **打开**：单个"打开"对话框同时接受 `.docx` 与 `.xlsx`，主进程只回**路径**，渲染层按扩展名分派（`App.tsx:314`）。
  改这个形态的原因：先前让 `file:open` 返回"docx 载荷 | sheet 载荷"联合，而 TS 的 `"workbook" in r` **排除不掉属性可选的那个成员**，
  11 处 `TS2339`；只回路径 + 按扩展名分派后类型自然收敛。拖拽入窗、`addRecent` 最近文件、`fileAssociations`
  （`apps/desktop/package.json:66` 新增 xlsx 一条）、`electron/main.ts:37` 的 `OPENABLE_EXT` 都走同一条路。
- **查看**：虚拟网格 + 列头/行头**双轴**同步（transform 挂在内层 `*-track` 上：`overflow:hidden` 的容器一被 transform，裁剪盒就跟着走，表头会飘）
  + 合并区覆盖层（带锚点文本，数字靠右）+ 表页签 + 名称框 + 只读编辑栏 + 公式重算提示。列宽/行高换算只在 `sheet-io`（`7a0f606`），渲染层拿到的已是 px（`AGENTS.md` §2.1）。
- **编辑**：打字/Enter/Tab/F2/Delete/Backspace/Home/Ctrl+Home/PageUp/PageDown、拖拽选区、列宽拖拽
  （`startPointerDrag` + 热区 `-webkit-app-region: no-drag`，§5.6）。几何与键盘步进全收在纯函数层 `src/sheet/sheetGridModel.ts`
  （`ab27d0e`）——桌面 vitest 是 `environment:"node"`，不纯就根本测不到。
- **保存**：`sheet:save` / `sheet:save-as` 在主进程**重开底包**、按 `partPath` 配对，配不上就整批不写（§2.4），先出 `.bak` 再 `.tmp → rename`；
  `sourceXml` 不过 IPC（`sheet-dto.test.ts`）。
- **门禁**：`sheet-io` 41 + desktop 416 单测绿，根 `pnpm build` 全包过。类型检查由 build 覆盖（§8）。
- **可视验证做到了哪一步**：本机没有 electron 的 CDP harness，所以用真实 Chromium 探针（连到用户已在跑的 vite dev，**没有** kill 它，§8）
  把这个组件的**实际计算样式与几何**量了出来，并因此发现一个真缺陷——画布 885px < 视口 1154px，右边露一条空白；
  于是 `TRAILING_COLUMNS` 8 → 24（表头 12 → 28 列，画布 1909px）。**这替代不了 E26 那条"起真实 app 手工走一遍"**，那条仍是 `[ ]`。

**与执行计划的偏差（记录，不算违背 §1.4）**：

1. 落点名换了：计划里的 `sheetStore.ts` + `SheetGrid.tsx` + `CellEditor.tsx` 三件套，实际是 `sheetGridModel.ts`（纯函数）
   + `SheetView.tsx`（一个组件，编辑框内联）。三件套是按"store / 视图 / 编辑器各一"预判的切法；真写起来编辑态只有"一格的文本 + 光标"，
   单开组件反而要把状态再传回 store，遂合并。
2. **插入/删除行列与撤销重做顺延到阶段 3**。理由：`AGENTS.md` §2.5 要求每个 op 发自己的逆操作、撤销按内容寻址，
   这套机制属于阶段 3 的 `SheetCoreOp`；在本阶段手搓一条撤销栈就是造第二个撤销栈（§3 禁止平行基建）。
   所以"完成判定"里的"开→改→**撤销**→存"只做到"改→存→重开内容一致"，**撤销那一步未实现**，不是实现失败。（阶段 3 的 `3d1c357` 补上了这一步，见下一节。）
3. 不做新建空白工作簿：本阶段只读写**已存在**的 `.xlsx`。从零造一个合法包（`[Content_Types].xml`/`workbook.xml`/rels/`styles.xml`）
   与"读回一致"是两件事，混进来会让本阶段的红线（不损用户文件）变得不可测。
4. `sheet:open` 通道与 `openSheetFile`（preload + `global.d.ts`）在偏差 1 的对话框改造后成为死代码，随即删除（§3.5）。

### 阶段 3 · 结构化 op + 撤销/modlog 接入（`3d1c357`）

**目标**：把"改单元格/改行列/改格式"收敛成 CoreOp 集合，每个 op 有逆操作，走 `planEdit/runOps/revert` 同一条链路，落 `userData/modlog`。
**完成判定**：`sheetEdit.test.ts` 覆盖：op→模拟→落地→撤销内容一致；插入行后再撤销不牵连用户后续手工编辑；无法安全回滚时明确报"部分撤销"。

**实际结果**（与计划的差别都在这儿，不是"按计划做完"）：

- op 层落在 **`packages/sheet-io/src/sheet-edit.ts`**，本阶段先有两种：`setCellValue`、`setColumnWidth`（插删行列见偏差 2）。
  形状与 Word 的 `CoreOp` 同构：`before`/`after` 都存进 op，逆操作就是两者交换；`runSheetOps` 是唯一落地入口，
  `revertSheetOps` 就是把逆队列倒序再走一遍同一个入口。
- **现场核对**（对不上就整条跳过并报理由，绝不猜位置）：
  ① 行/列按**内容指纹**寻址——`rowFingerprints` 是 `列号=值|…`，新增的 `colFingerprints` 是 `行号=值|…`。
  为什么不用 A1 地址：地址会随插删漂移，拿它当身份等于宣布"挪过位置的那一行不是原来那一行"，撤销就再也落不回去；
  单测「插入行把内容整体下移后，撤销仍命中原来那一行」把这条钉死。
  ② 空白行/列没有可认的内容，记伪指纹 `∅` 只按位置回，真正拦住写错地方的是"这格现在装的是什么"或"`<col>` 分段形状"的比较。
  ③ 逆操作的落点按**落地之后**的样子重新寻址——这次改动本身就改了行指纹，沿用旧指纹会认不出自己刚写的那一行。
  ④ 富文本 run 与 `rawXml` 只比"在不在"：把红字拍平成纯文本是 §2.3 禁止的事，现场多出或少掉这类结构就不许覆盖。
- 改列宽只动被拖的那一列：`<col min max>` 是一段范围，落地时劈成左段/本列/右段三段，撤销把三段整体换回；
  输出的分段按列号排序，与 Excel 自己写出的 `<cols>` 顺序一致（否则同一份文件两次保存会产出不同字节序）。
- 桌面装配：`SheetView` 不再自己改载荷，编辑框与列宽拖拽各交出**一条 op**（一次拖拽只发一条，拖动过程中的宽度是本地预览——
  按鼠标移动记的话一次拖拽会在历史里留几十条，撤销只退一像素），由 `App.commitSheetOps` 落地并把逆操作写进 modlog。
  存储与通道**全部复用**文档那四条（`modlog:append/patch/list/remove`，`electron/modlog.ts` 一字未改，CAP 仍是 60），
  没有新建第二份历史；`sheetHistory.ts` 只做表格专属的事：认形状、挑最近一条没撤销的、把逆操作倒回去。
  Cmd/Ctrl+Z → `pickLatestSheetEdit` + `undoSheetEdit`，认不出的一律不碰并 toast「部分撤销（x/y）」——文案与 Word 那条一致。
- **落盘证据**（§6.3 要求的"真实产物"）：`sheet-file.test.ts`「op 改完再撤销：写出去的文件读回来与原来一致」——
  改 C2 + 加宽 B 列 → 写盘 → 重开读到 9999 → 撤销 → **从改过的那份**再写盘 → 重开时值与 `<cols>` 都回到原样。

**偏差**（都是主动选择，不是遗漏）：

1. 计划的落点是 `apps/desktop/src/ai/sheetEdit.ts`，实际落在 `packages/sheet-io`：§5.1 的依赖方向（apps → packages）
   决定了阶段 4 的 `packages/sheet-tools` 要产出**同一批** op，落在渲染层就必然被复制第二套（§3.3/§3.4）。
2. 行列插入/删除的 op 不在本提交：C16 的字面要求是"插删之后撤销早期改动不牵连手工编辑"，本阶段先把 op 层与内容寻址做对、
   并用手工构造的位移表把这条语义测住；插删作为**同类 op** 已在 `ec4bda4`（op 层）+ `e662914`（视图入口）补齐，C16 升 `[x]`。
3. 撤销入口只有键盘，没有工具栏按钮，也没进「修改历史」面板：那个面板目前仍归文档那条（表签视图里不渲染），
   所以**表格改动在面板里看不到**，只在按 Cmd/Ctrl+Z 时被读出来。等阶段 4 的 AI 提案也记进同一份历史时一起做，
   否则会变成两条 UI 各挑一次记录。

**阶段 3 收尾（`ec4bda4` + `e662914` + `c099727`，C16 与 E26 一起结）**：

- **入口**：`SheetView.tsx` 在行/列表头上挂 `onContextMenu`，菜单四项插入（上/下、左/右）+ 分隔线 + 删除本行/本列；
  定位沿用正文那套 `editor/contextMenu.menuRect`（§5.5），样式沿用 `.ctx-menu`/`.ctx-item`（§5.4 没新增 CSS 入口）。
  菜单项拦 `mousedown`：否则编辑框的 `onBlur` 会抢在同一次落地里多提交一条改动。表格一滚动就关菜单——它按视口坐标贴着那条表头。
  只做单行/单列：选区→"带"的语义要引入第二套位移规则，超出本阶段（§3.6）。
- **真实 app 的整链实跑**（临时 CDP harness，探针不入库，`tmp/track-app-probe.mjs`）：开 `sample.xlsx` →
  右键 H 列表头删列 → toast 拒绝并给出理由（`H1:I2 这个合并区正好被这一带劈开…`），文件未动；
  右键第 3 行表头在上方插入 → 网格 `A3` 空、`2月→A4`、`3月→A5` → 命令面板"保存"落盘（`<dimension>` 到 `A1:I5`、
  写出 `.bak`）→ Cmd+Z → 网格回到 `A1:I4` → 再保存，**`sheet1.xml` 与 `sharedStrings.xml` 与插入前逐字节一致**。
  这条把 E26（"起真实 app 走一遍开→改→存"）从 `[ ]` 结掉，也第一次在真实进程里跑到了拒绝路径。
- **实跑顺带钉出一个 §2.3 缺陷**（`c099727`，不是插删引入的，是保存路径一直就有的）：桌面保存会逐张表 `writeSheet`，
  富文本格过去一律重新 `intern`——写出的 `<si>` 只剩能重新序列化出来的那部分格式，**`<color rgb="FF0000FF"/>` 被拍平成黑字**，
  还追加一条重复条目（`uniqueCount` 11→12）。修法是"按内容复用"：`intern` 给带 run 的条目按 run 逐字段指纹复用，
  新增 `reuseOrAppend(值, run, 读入时的索引)` 让值与格式都没动的格沿用原条目，`writeEntry` 补回 run 颜色。
  三条单测钉住（全部表写一遍仍指原条目 / 走桌面载荷那条路同样复用 / 真改过时重序列化必须带回颜色）。
- **探针环境事实（这次踩到的）**：`--user-data-dir` 只是 Chromium 开关，`main.ts:73` 的
  `app.requestSingleInstanceLock()` 仍按默认 `userData` 算——**上一个探针实例没退干净，后一个就会静默拿不到窗口**
  （`/json` 返回空数组），而连上去的调试端口又可能是旧实例的，于是"改了代码、跑的还是旧进程"。
  与 §8 的"探针跑的是 dist"同一条族：跑 app 级验收前先 `pgrep -f "apps/desktop"` 数一遍，清零再起。

### 阶段 4 · AI 工具层：分析 + 整理 + 计算 — `[x] 已完成 2026-10-13`（整理按 §七 的定案只做四类 op，排序/拆合列/统一格式不在写侧）

**目标**：自然语言 → 工具回路 → 提案 → 用户确认 → 落地。覆盖分析（描述/聚合/找异常）、整理（去重/拆合列/统一格式/排序）、计算（新增计算列写公式）。
**完成判定**：`sheet-tools` 全部入参有 zod 单测；桌面 `sheetAgent` 用 mock `fetch` 测过工具环与截断；
用真实端点探针跑 ≥3 类任务（分析/整理/计算）并记录 token 与失败样本；产出的 `.xlsx` 读回一致且 Excel 能重算公式。

**提交 10 的实际结果**（`470d370`，分析引擎）：列画像 / 分组聚合 / 去重键落在 `packages/sheet-io/src/analysis/**`，
模型只指定"按哪几列分组、对哪列做什么聚合"，数字一律由引擎算（§零.4）；列引用解析不出来时返回 `ok:false` + 可读理由，
工具层把理由回灌给模型而不是抛异常。

**提交 11 的实际结果**（`0e7cb74`，公式子集求值 = D4 的写侧）：

- 落点是 **`packages/sheet-io/src/formula-eval.ts`（单文件）**，不是计划里的 `src/formula/**`：
  求值器是一个自足的递归下降解析器（切词 → 比较/连接/加减/乘除/幂/一元/后缀 → 最小单元），拆成目录会把一条
  token 流拆到三个文件里读；与包内 `cell-text.ts`、`formula-refs.ts` 同一族命名（§4.5 kebab-case）。
- 白名单按 D4：`SUM/AVERAGE/MIN/MAX/COUNT/COUNTIF/SUMIF/ROUND/ABS/IF/CONCAT(ENATE)` + 算术、比较、`&`、`%`、
  一元正负、`^`（右结合、高于乘法，与 Excel 一致）。**口径照 Excel**：区间里的空格/文本/布尔跳过，
  "看着像数字的文本"仍按文本处理（`SUM(A2:A4,D2)` 不把 `D2` 的 `"15"` 计进来），文本与数字比较时文本更大。
- **失败不猜**：volatile（`TODAY/NOW/RAND/RANDBETWEEN/OFFSET/INDIRECT/CELL/INFO`）、跨表 `!`、外部簿 `[1]`、
  数组常量 `{}`、错误值 `#`、白名单外的函数、循环引用 → `{ok:false, reason}`，`setCellFormulaOp` 照常写 `<f>`
  但**不写 `<v>`**，并把 `recalcNeeded` 回报给调用方（验收 A4）。求值器**永不抛异常**，垃圾输入也是错误串（验收 B9）。
- 依赖格的值**现算不信缓存**：带 `<f>` 的格被引用时递归求值；引用链上的格子集合往下传，
  `A1↔B1` 这种互相引用在耗尽调用栈之前判成循环。依赖格自己算不出时整条判为算不出——
  中间层把它咽下来就会把"算不出"变成"算出个 0"，那是 D4 红线的反面。
- 桌面入口：`SheetView` 编辑框以 `=` 开头改走 `setCellFormulaOp`（其余文字仍走 `setCellTextOp`）；
  编辑栏那条 `.sheet-recalc-hint` 分成两句话——"这条公式本地算不出来，格子里没有值"与
  "公式显示的是缓存值，改了数据要在 Excel / WPS 里重算"，不再用一句覆盖两种情况。
- **测试**：`formula-eval.test.ts` 18 条（口径、优先级、拒绝结构、永不抛异常）；`sheet-edit.test.ts` 新增 4 条
  （算出缓存值 + 撤销后整格消失、算不出时值为空且 `recalcNeeded`、样式/`attrs` 原样带走、空公式返回 null）；
  `sheet-file.test.ts` 新增 3 条落盘往返（`<c r="F2"><f>AVERAGE(C2:C4)</f><v>120</v></c>`、
  `<c r="F3"><f>TODAY()</f></c>` **只有 `<f>`**、带格式的格改成公式后 `@s` 不动且显示仍按原格式）。
  逐包核对：sheet-io **152**（12 个文件）/ desktop **416**（37 个文件）/ docx-io 296 / word-client 8 / word-tools 6 / server 2，
  `pnpm build` 与 `pnpm test` 均 EXIT=0（重定向后读退出码，§8）。
- **第三方交叉读**（`tmp/xlsx-probe/formula-cross-read.mjs`，不进主干，§6.5/§7.3）：把上面三条 op 落到真实 `.xlsx` 后
  用 exceljs 读回 —— `F2={formula:"AVERAGE(C2:C4)",result:120}`、`F3={formula:"TODAY()"}`（**没有 result，即没有过期数字**）、
  `C2=120`（这条 op 被同批跳过，见下条），`xl/workbook.xml` 里 `<calcPr fullCalcOnLoad="1"/>` 在。
  这是"不是我读我自己写"的证据；"Excel/WPS 真的重算"那一半本机仍无可视通道，A4 保持 `[~]`、A8 保持 BLOCKED。
- **实跑钉出一个 op 层的既有约束，并已在 `7de845f` 解决**（不是公式引入的，是阶段 3 内容寻址在批内的死角）：一条批里的 op 全部照**批前的快照**
  锚定行指纹，所以**同一行在一批里改两格时，第二条会被"认不出这一行"跳过**（上面 `C2` 那条就是这个现象，修完后 `op ok: true`、
  exceljs 读到 `C2={formula:"C3+C4",result:240}`）。修法是批内记一张"旧指纹 → 改后的指纹"的锚点表，**只在这批 op 内部生效**：
  同一格在一批里改两次时第二条仍报"装的不是记录时的内容"，落地后用户再改这一行时撤销仍认不出并报部分撤销——
  C15 的现场核对一点没放松。四条单测钉住（同行两格、整列五条、同格两次、批外改动仍拦）。
- **同一族的列宽问题已在 `512fb0f` 解决**（阶段 3 遗留，与公式无关）：`setColumnWidthOp` 存的是"覆盖这一列的那几段
  `<col>`"的整段前后像，一批里改两列时第一段被劈开、第二段期望的分段形状对不上，第二条 op 以"宽度段与记录时的
  分段不一致"被跳过。它不是内容指纹问题，所以修法和 `7de845f` 不同：**先走精确路径**（现场分段与记录逐段一致时
  整段换回，撤销能还原成原来那一段 `<col>` 写法），不一致时**只核对目标列自己的宽度没变**，再按现场重算劈法，
  相邻列宽度一律按现场保留。落点坐标也改记成落地时的 `target.index`（前像与逆操作同坐标系，撤销不再二次平移）。
  七条单测钉住：批内两列都落地并逐字节回到原分段、跨两条覆盖带的批、同一列两次仍拦（现场宽度已不是记录时的）、
  用户改过宽度后撤销仍拦、用户只劈段没改宽度时撤销不再误拦（他劈出的段原样留着）；另有一条 `.xlsx` 落盘往返。
  本次门禁：`pnpm build` EXIT=0、`pnpm test` EXIT=0，逐包核对 sheet-io **161**（12 个文件）/ desktop **416**（37 个文件）/
  docx-io 296 / word-client 8 / word-tools 6 / server 2。

**提交 12 的实际结果**（`b87a88b` 引擎读面 + `dba6fb4` 新包 `packages/sheet-tools`）：

- 计划的落点只有 `packages/sheet-tools/src/index.ts`，实际**先补了引擎的三条读口**（`b87a88b`），否则工具层只能拿到
  `describeSheet`/`aggregateSheet`，模型连"这几行长什么样"都没法问：新增 `readRange`（区域 → 显示文本）、
  `findRows`（条件 → 行号与地址），并给 `describeSheet` 加 `maxColumns`。三条口都是"认不出来给理由串"而不是抛异常（B9）。
- **条件口径复用公式层**（§3.1）：`findRows` 直接用 `formula-eval.toCriteria`，所以工具筛出的行与 `COUNTIF` 数出的行
  同一套判定；带 `<f>` 的格按缓存值参与比较，没有缓存值就当空——Excel 重算前它本来就没有可比数字。
  `toCriteria` 因此从包内私有改成包内导出（**没有**进 `index.ts` 的公开面）。
- **上限写两处**：引擎侧是默认值（≤50 行样本、单格 ≤120 字、≤50 列画像、≤50 组/命中），schema 侧是硬边界（越界即拒收）。
  引擎的默认值给 CLI 用，schema 的边界给模型用，两边都有单测（B11）。
- 新包按 spec §三 的边界**只依赖 zod、不 import `sheet-io`**：它只判定"这套入参合不合法"，取值与统计全部由调用方
  拿校验过的入参去调引擎（B10）。列引用与条件的语义解释只在引擎里，schema 只约束长度，避免第二套判定。
- `parseToolArguments` 是包内唯一入口，永不抛异常（非法 JSON / 多余键 / 越界 / 不认识的工具名 → 错误串）。
  Word 侧同类逻辑目前是 `apps/server/src/agent.ts:98-123` 里逐工具内联的 try/catch；提交 13 **不许在桌面再抄一份**，
  要么复用这个入口，要么把 Word 侧也并进来（§3.2）。
- 工具定义与 schema 不许长歪：单测逐条比对两边的**工具名集合、参数键集合、必填项集合**（模型只能填它看得见的键）。
- **真实 `.xlsx` 上的落盘证据**（`sheet-file.test.ts` 新增 3 条，跑在阶段 1 夹具上）：表画像 `A1:J4` / 3 数据行 / 10 列，
  销量列 `sum=360` 与表里 `SUM(C2:C4)` 那条公式留下的缓存值**是同一个数**；`readRange` 出来的显示串仍按各格自己的
  数字格式（`120` 走内建 3、`0.12` 显示 `12.00%`、日期序列号显示成 `2026-01-31`、富文本 `粗蓝` 合成一句、空格拉成空串）；`findRows` 的 `>100` 命中
  `C2/C3`、`<>华东` 命中第 4 行、`3月` 命中 `A4`，表头行不进命中。
- 本次门禁：`pnpm build` EXIT=0、`pnpm test` EXIT=0，逐包核对 **sheet-io 176**（12 个文件，+15）/ **sheet-tools 9**（新包）/
  desktop 416 / docx-io 296 / word-client 8 / word-tools 6 / server 2。
- **已到位（提交 13 前半，`dba6fb4` `1f5ec47` `c52c1b7` `e69780e`）**：写侧只开 `sheet_submit_proposal` 一个出口、
  `ai:tool-turn` 的一轮转发、`src/ai/sheetAgent.ts` 的工具环（截断上送、非法入参不抛、提案→op、暂停/继续、8 轮上限）。
- **已到位（提交 13d，`f0f9853`）**：面板的表格分支——提案卡片（`Item` 多一支 `t:"sheet"`）、"应用到表格"的确认按钮、
  全自动那一路的直接落地与 toast 提示、暂停后卡上的"继续"。真实 app 实跑见 §八 与 B12/B14。
- **已到位（13e，按 §6.5 探针代码只留在 `tmp/`，不入主干）**：node 侧一次性探针（`tmp/sheet-agent-probe.mjs`）打真实端点跑"分析 / 计算 / 整理"各一条，
  工具序列与落地/撤销数字记在 B13。
- `sheet_create_chart` 按计划属阶段 5，不在本次工具面里（这条偏差在此显式记录）。

### 阶段 5 · 图表

**目标**：AI 选区出图；app 内可看（ECharts），`.xlsx` 内为原生可编辑图表（D5 spike 成功）或明确记录的贴图降级。
**完成判定**：`ChartSpec` 校验与渲染单测；写出的 chart part 读回结构断言；导出 PDF/docx 报告时图以 PNG 出现；WPS/Excel 可视验证标 BLOCKED 并给出结构级替代证据。
**进度（2026-10-13）**：写侧已落 `294ce88`——`ChartSpec` 校验单测（解析 3 条 + 落地 4 条）、chart part 的读回结构断言、
以及 P4 口径的**第三方交叉验证**（openpyxl 3.1.5 把 bar/line/pie 三份产物都读回，系列引用/缓存值/锚点/标题逐条对上）。
工具面与 app 内侧已落 `e19aece`：`sheet_create_chart` 进工具环，回合停在 `chart` 提案上，半自动挂"插入图表"卡片、
聊天模式只念引用与数值、全自动排进待写清单；待写清单不进工作簿载荷（图表是包部件不是格子），保存时由主进程逐张
`addChart`，插不进就整批取消（§2.4）。真端点实测（`tmp/chart-agent-probe.mjs`，见 §八 该行）模型确实自己调了这个工具，
折线图数值由引擎从现场表读成 `[100,120,150,150]`；**真实 app 也跑通**（§7.2，`tmp/chart-ui-probe2.mjs`：卡片→点"插入图表"→命令面板保存，数字见 §八）。
这条路上独立读取器还抓出一个 §2 红线缺陷并修在 `fa693d2`（§二 P8）。
未落：app 内 ECharts 查看、canvas→PNG 报告插图（都卡在 §6.4 那条"要不要引图表库"的表态）。Word/WPS 可视验证仍 `[!] BLOCKED`（本机无可视通道，`AGENTS.md` §8）。

### 阶段 6 · 收口：CLI、文档、规则落地

**目标**：`apps/cli` 增加 `ocs`（或 `oc sheet`）子命令复用 `sheet-io`+分析引擎；`README.md` 增补 Excel 小节；把 §一~§三 新增的不变量补进 `AGENTS.md`；
可选：`sheet:*` 的 docx 报告嵌入（把表格区域作为 Word 表写入，复用 `DocxFile.insertTable`）。
**完成判定**：三端（desktop/cli/docs）对同一份 `.xlsx` 行为一致；`AGENTS.md` §9 表格里新增项标清"已落地/待落地"。
**进度（2026-10-13）**：CLI 读侧已落 `a72196e`——`apps/cli/src/sheet.ts` 的 `oc sheet info|stats|read|find|aggregate|dedupe`，
**只做装配**：解析、单位换算、列画像、区域显示文本、聚合、条件筛选、重复行预览一律走 `sheet-io` 的公开入口
（`SheetFile.toWorkbook` + `analysis/**`），所以"三端对同一份 `.xlsx` 行为一致"是结构性的，不靠对拍；
每条都支持 `-s/--sheet`（表名或 `info` 里显示的 1 基序号）、`--no-header`、`--json`。
写侧**刻意不在 CLI 落地**：撤销、modlog、原子 `.tmp→rename` 都在桌面主进程那条路上，CLI 再写一份保存等于开第二个写入口，
`dedupe` 因此只预览并在末尾提示"删行请回桌面 app"。`apps/cli` 第一次有测试文件（3 条）。
同一批的还有文档与规则：`README.md` 的 Excel 小节、`oc sheet` 命令表、逐包通过数与 `已知边界`；
`AGENTS.md` 新加 §2.10（新建 OPC 部件三处一起登记）与 §2.11（数值一律由引擎算，模型只交引用），
§2.1/2.3/2.5/2.8/2.9 补上表格侧的单点与机检位置，§9 表格里把表格相关项标"已落地"并写出测试文件。
未落：可选的 docx 报告嵌入（把表格区域作为 Word 表写进 `.docx`）——它要的是写侧，与本阶段的"CLI 只读"不同，留给后续。

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
| 2 | 主进程按扩展名分派；新增 `sheet:*` handler；原子写 + `.bak` | `electron/fileHandlers.ts` `electron/sheetHandlers.ts` `electron/main.ts:37`（`DOCX_EXT` 改成 `OPENABLE_EXT`）、`package.json fileAssociations` | `bbf23c5`（计划"提交 6"） |
| 2 | preload 白名单加 `sheet:*`（同时改类型，禁万能 invoke） | `electron/preload.ts:98-103`（三条具名通道）+ `src/global.d.ts` | `bbf23c5`；`sheet:open` 后因死代码删除，见阶段 2 偏差 4 |
| 2 | store 与网格组件 | **实际**：`src/sheet/sheetGridModel.ts`（纯函数：几何、可见窗口、键盘步进、点→地址、改列宽）+ `src/sheet/SheetView.tsx`（单一组件，编辑框内联）；计划的 `sheetStore.ts`/`SheetGrid.tsx`/`CellEditor.tsx` 三件套未建，见阶段 2 偏差 1 | `ab27d0e` + `c8c3f3b`（计划"提交 7"） |
| 2 | 样式（唯一 CSS 体系，不新增入口） | `src/styles/sheet.css`（由组件自己引入，与 `footnotePane.css` 同款先例） | `c8c3f3b`（计划"提交 7"） |
| 2 | 列宽拖拽走 pointer capture；浮层用 fixed + `useAnchoredMenu` | 复用 `hooks/usePointerDrag.ts`；本阶段网格无浮层，`useAnchoredMenu` 无对应实现（阶段 4 有菜单时再落），见 D20 | `c8c3f3b`（计划"提交 8"） |
| 3 | `SheetCoreOp` 类型、`planEdit/runOps/revert` 的表格实现、行内容指纹 | **实际**：`packages/sheet-io/src/sheet-edit.ts`（+ 新增 `colFingerprints` 于 `xlsx-sheet.ts`）；计划的 `apps/desktop/src/ai/sheetEdit.ts` 未建，见阶段 3 偏差 1 | `3d1c357`（提交 9） |
| 3 | 撤销与修改历史接入（复用 modlog，不新建存储） | **实际**：`apps/desktop/src/sheet/sheetHistory.ts`（认形状、挑最近一条、倒回逆操作）+ `App.tsx` 的 `commitSheetOps`/`undoLastSheetEdit`；`electron/modlog.ts` 与 preload 通道一字未改 | `3d1c357`（提交 9） |
| 3 | 行列插入/删除的 op（C16 的字面要求） | `packages/sheet-io/src/sheet-edit.ts`（`insertTracksOp`/`deleteTracksOp` + `applyShiftOp` 位移与逆操作）；视图入口 `src/sheet/SheetView.tsx`（表头右键菜单，复用 `menuRect` 与 `.ctx-menu`） | `ec4bda4` + `e662914` |
| 3 | 保存路径不拍平富文本（§2.3 红线，实跑时才发现的既有缺陷） | `packages/sheet-io/src/xlsx-strings.ts`（run 指纹复用 + `reuseOrAppend` + `writeEntry` 回写颜色）、`xlsx-sheet.ts` 调用点 | `c099727` |
| 4 | 分析引擎（列类型推断、缺失统计、分组聚合、去重键、异常值） | `packages/sheet-io/src/analysis/**` | 提交 10 |
| 4 | 公式子集求值（D4 的函数白名单）+ 缓存值写入 | **实际**：`packages/sheet-io/src/formula-eval.ts`（单文件递归下降，计划的 `src/formula/**` 目录没建，理由见阶段 4 提交 11）+ `sheet-edit.ts` 的 `setCellFormulaOp` + `SheetView.tsx` 的 `=` 入口 | `0e7cb74`（提交 11） |
| 4 | 一批 op 里同一行的多条改动要都能落地（提交 12 的前置，见 §七） | `packages/sheet-io/src/sheet-edit.ts`（`applyOps` 的批内重新锚定 + `reanchorOp`/`rememberReanchor`） | `7de845f` |
| 4 | 一批里改多列列宽要都能落地（同一族的列轴问题，见 §七） | `packages/sheet-io/src/sheet-edit.ts`（`applyColumnWidthOp` 先精确整段、不一致时按现场重算覆盖带，`splitColumnBand`/`sameColumnWidth` 抽出复用） | 512fb0f |
| 4 | 工具 schema 与 JSON 工具定义 | **实际**：`packages/sheet-tools/src/index.ts`（四个读工具）+ 引擎侧先补三条读口 `packages/sheet-io/src/analysis/sheet-table.ts`（`readRange`/`findRows`/`describeSheet` 的 `maxColumns`），理由见阶段 4 提交 12 | `b87a88b` + `dba6fb4` + `cd65c5b`（提交 12） |
| 4 | 写侧工具 schema 与提案：`sheet_submit_proposal`（改格/改公式/插删行列/去重/列宽） | **实际**：`packages/sheet-tools/src/index.ts` 的 `SheetEdit` 四类 + `SubmitProposalArgs`。**已定的决策**：排序、拆合列、统一数字格式**不开**——现有四类 op 各自有安全的逆操作，第五类 move op 没建（要建得先解决"撤销整表前像"与 §2.5 的内容寻址），用一批 `setCellTextOp` 凑排序更会把用户后来的手工编辑牵连进去。边界写进工具描述与系统提示词，模型只能提这四项（理由见 §七同一行） | `dba6fb4` + `1f5ec47`（提交 13） |
| 4 | 桌面 agent：提示词、区域截断、提案装配、卡片渲染分支 | **实际**：`src/ai/sheetAgent.ts`（提示词、`sheetContextOf` 截断拼装、工具环、`sheetOpsOfProposal` 提案→op）+ `electron/aiHandlers.ts` 的 `ai:tool-turn` 一轮转发 + preload/`global.d.ts` 的跨界类型（§5.2 三处一起改）为 `c52c1b7` + `e69780e`（提交 13）。**卡片渲染分支已落**：`src/ai/AIPanel.tsx` 的表格分支（提案卡 / "应用到表格" / 全自动直接落地 / 暂停后"继续"），面板按模式各挂一份会话 key | `c52c1b7` + `e69780e` + `f0f9853` |
| 5 | `ChartSpec` + ECharts 渲染 + canvas→PNG | **已落**：`packages/sheet-io/src/chart-spec.ts`（`resolveChartSpec` 只收 A1 引用，数值/引用串/显示串一律引擎从现场表取，失败给一句理由；px→EMU 只在这一处）。**未落**：`src/sheet/ChartView.tsx`（ECharts 查看）、canvas→PNG——卡在 §6.4 的图表库表态（P6 同一条）；`e19aece` 里卡片只把引擎解析好的引用/数值/落点排成中文行给用户看，不预先造第二套画图代码 | `294ce88`（未落部分待 §6.4 有结论） |
| 5 | chart part 写入（`xl/charts` + drawing + rels + content-types 更新） | **实际**：`packages/sheet-io/src/xlsx-chart.ts`（纯 XML 拼装）+ `sheet-file.ts` 的 `SheetFile.addChart`（五件一起登记，`addedParts` 进 `toBytes`、`restoreSnapshot` 一起回退）+ `opc.ts:relativeTarget`（`resolveTarget` 的逆）。原计划的提交 14/15 合成一个 `294ce88`：解析与写入共用同一份测试文件，按文件粒度拆不开，硬拆只会造出编译不过的中间提交 | `294ce88` |
| 5 | `sheet_create_chart` 工具面 + app 内侧的图表卡片与保存落地 | **实际**：`src/ai/sheetAgent.ts`（工具 case、`{kind:"chart"}` 回合出口、系统提示词那段"只给区域，数值由引擎读"）+ `src/ai/AIPanel.tsx`（三种模式各一条路：卡片行/计数/按钮/已应用文案，引擎只给码与结构化引用，中文在渲染层拼，§5.7）+ `src/App.tsx` 的 `pendingCharts`（按 `partPath` 暂存、同表换掉前一张、换文档与新建都清空）+ `electron/sheetHandlers.ts` 的 `saveSheetDoc(..., charts)` 与 `global.d.ts` 的 `SheetSaveRequest.charts`（§5.2 三处一起改，未开新通道） | `e19aece` |
| 5 | 独立读取器复核保存路径时抓出的包结构缺陷（§二 P8） | `packages/sheet-io/src/sheet-file.ts`（`writeSharedStringsPart` 只在"原文件没有这张表"时补关系与 Override；追加关系那段收敛成 `relsWith`/`nextRelIdFor`，图表那条路复用同一处，§3.2）+ `xlsx-strings.ts` 三个常量 + `sheet-file.test.ts` 两条回归（inlineStr 来源、原有表逐字节不动） | `fa693d2` |
| 6 | CLI 子命令、README、AGENTS 更新 | **实际**：`apps/cli/src/sheet.ts`（六条只读子命令 + `registerSheetCommands`，纯装配）、`apps/cli/src/sheet.test.ts`（该包第一条测试）、`apps/cli/src/index.ts`（注册 + 描述里写明"只读分析 `.xlsx`"）、`apps/cli/package.json` 加 `@office-copilot/sheet-io`（唯一新增依赖，`AGENTS.md` §6.4：引擎已在仓内，不引第三方表格库）、`README.md`（`oc sheet` 命令表 + 示例、§三 下的 Excel 小节、逐包通过数、已知边界改写成"已落/未落"）、`AGENTS.md`（§0 范围、§2.10/§2.11 两条新不变量、§2.1/2.3/2.5/2.8/2.9 补表格侧、§3 禁止第二套 xlsx 实现、§5.1 依赖方向、§7.3 `tmp/` 没有 gitignore 条目、§8 两条环境事实、§9 机检状态） | `a72196e`（代码）+ 本提交（文档与规则） |

推进纪律：每阶段结束跑一次 `pnpm build > tmp/b.log 2>&1; echo EXIT=$?` 与 `pnpm test > tmp/t.log 2>&1; echo EXIT=$?`
（**重定向后再读退出码**，`AGENTS.md` §8 的假绿），并逐包核对通过数（`apps/cli`、`packages/shared` 本就无测试）。
每次改动完成立即提交并推送 `origin`，推送后 `git status -sb` + `git ls-remote origin main` 双重复核（`AGENTS.md` §1.6）。

---

## 六、验收标准列表（逐条打分，任一不过不得声称完成）

### A. 文档正确性（最高优先级，对应 `AGENTS.md` §2）

1. `[x]` **空改动往返**：读入 → 立刻写出，与原文件在所有部件上语义等价；未建模部件（条件格式、数据验证、透视表、外部链接、图表、主题）逐字节保留。
   **实测**：`sheet-file.test.ts`「打开后不改任何东西，写出的每个部件与原文件一致」把 13 个部件（含 `xl/externalLinks/externalLink1.xml` 与它的 rels）逐条 `toEqual` 原文。
   **逐特性对拍**（§二 P11，`bbc7eba`）：openpyxl 自写的 10 份"每份只压一类结构"的夹具走桌面真实保存路径后，再由 openpyxl 分别读原文件与产物逐字段比——10 份的结构审计写前写后都是 0 条，10/10 逐字段一致（首轮 9/10，`<c t="inlineStr"><is><r>` 的富文本被拍平，已修并有两条回归单测）。
2. `[x]` 打开真实世界工作簿不丢文本：含富文本 `<r>` run 的 sharedStrings 读出为纯文本时，写回后**内容不丢**（格式丢失需给出 warning，不许静默）。
   **实测**：夹具里那条两 run 的 `<si>`（粗体 + `FF0000FF` 蓝）读成纯文本 `粗蓝`，写回时整条 `rawSi` 原样搬（string 表只追加、不改已有条目）；第三方 exceljs 读回 `richText:[{bold:粗},{蓝}]`。
   **`bbc7eba` 补的另一条来源**：run 也可能不在字符串表里，而是整段待在格子里（`<c t="inlineStr"><is><r>`，openpyxl 的 `CellRichText` 就这么写）。这类格子原先只留拼接文本，保存一次就变成 `t="s"` 共享字符串并**新建一张 `xl/sharedStrings.xml`**，run 全丢；现在按同一条口径把整格原文挂 `rawXml` 原样回写，改值时 `setCellTextOp` 造的新格子天生不带它，照常覆盖得掉。openpyxl 读回产物：`CellRichText([TextBlock b=True '粗', '普通', TextBlock color=FF0000FF '蓝'])`，且 `xl/sharedStrings.xml` 不再凭空出现。
3. `[x]` 日期/数字读回正确：序列号 ↔ 日期只在 `sheet-io` 进出口换算一次；同一文件反复打开得到同一显示值。
   **实测**：`46053 → 2026-01-31`（`datetime.test.ts` 锚定真实样本序列号，含 1900 闰年 bug 的 `dayParts<=60` 分支）；自定义 `numFmtId=164` 与内建 `14` 都判成日期而值仍是序列号；`grep -rn "\bserial\b|serialTo|ToSerial|25569|epochDays" apps packages`（排除 `dist/`）在 `packages/sheet-io` **之外零命中**（本机实跑 exit 1），§零.2 的单点成立。
4. `[~]` 公式写回后 **Excel/WPS 打开能重算**（读回断言 `<f>` 存在且合法，`<v>` 为缓存值）；无法本地求值的函数在 UI 明示"需 Excel 重算"，不许显示过期数字当作结果。
   **阶段 1 到位的部分**：`<f>`+`<v>` 成对读写，跨簿引用 `'[1]Sheet1'!$A$2` 不被改写或多余转义；exceljs 读回公式 `SUM(C2:C6)` 与缓存值 `660`。
   **阶段 2 到位**：选中带公式的格时，编辑栏下方出现 `.sheet-recalc-hint`「公式显示的是文件里的缓存值，需在 Excel / WPS 里重算」（`SheetView.tsx`），显示的是缓存值而不是我们算的结果。
   **阶段 4 到位**（`0e7cb74`）：白名单函数由 `formula-eval.ts` 当场算出缓存值（D4 的写侧），算不出的**只写 `<f>` 不写 `<v>`**
   并回报 `recalcNeeded`；编辑栏提示分两句（"本地算不出来，格子里没有值" / "显示的是缓存值"）。落盘单测钉住
   `<c r="F2"><f>AVERAGE(C2:C4)</f><v>120</v></c>` 与 `<c r="F3"><f>TODAY()</f></c>` 两种形状；
   第三方 exceljs 交叉读回 `F2={formula:"AVERAGE(C2:C4)",result:120}`、`F3={formula:"TODAY()"}`（无 result），
   `xl/workbook.xml` 带 `<calcPr fullCalcOnLoad="1"/>`。
   **仍未到位**：Excel/WPS 真的重算这一半本机无法验证（§8：WPS 无可视通道），保持 BLOCKED。
5. `[~]` 保存走 `.tmp → rename`，原地写回保留 `.bak`；保存中断不会留下半个文件。
   **已证**：写回内容正确性有单测钉住（`sheet-file.test.ts`「跨界载荷」一条：改 `C2` → 重开底包 → `applySheetPayload` → `writeSheet` → 断言 `<c r="C2" s="1"><v>777</v></c>`，并在配不上 `partPath` 时整批不写）。
   **未证**：`electron/sheetHandlers.ts:saveSheetDoc` 里的 `.bak` 与 `.tmp → rename` 是主进程代码路径，本机没有 electron 进程级测试去模拟写到一半被中断，所以"中断不留半个文件"按代码推断而非实测。
   **`bdf3ba4` 加的一道闸**：出字节之后、写盘之前过一次结构审计，**只拦"这次改动新引入的"问题**（底包本来就不合规的第三方文件不该让用户存不了盘），拦下就整批不写、回一句中文，渲染层用已有的红 toast 显示（不加新 preload 通道）。审计自己读不懂陌生 XML 时按"判不了"放行——判不出绝不能挡住用户保存。判定逻辑有 3 条单测（`electron/sheetHandlers.test.ts`），"闸门不误伤"是实测：app 早前存过盘的 `chart-app-saved{,2}.xlsx` 连同 openpyxl 与我们 writer 的产物共 **7 份全部 0 条**。
6. `[x]` `.xlsm` / 含外部连接的包：按阶段 1 的结论明确拒绝或原样保留，二者之一并有测试，不许"看起来成功"。
   **实测结论**：含 VBA 宏（`xl/vbaProject.bin` 或宏 content-type）**明确拒绝**并给出可读消息；外部链接/透视表/图表/绘图**原样保留**并记进 `unsupported` 交给 UI 提示。各有单测（宏拒绝、`externalLinks` 部件逐字节 + `[1]` 引用活着）。
7. `[x]` 单位/颜色换算单点：全仓 grep 不出现第二处 serial/numFmt/颜色换算实现（§零.2）。
   **实测**：serial/numFmt 判定只在 `src/datetime.ts`，ARGB ↔ `#rrggbb` 只在 `src/color.ts`，列宽只存 `widthChars` 原样搬运（有单测钉住不做第二处换算）。
8. `[!] BLOCKED`（只剩"目视"这一半）用 Excel 或 WPS 实际打开一遍：本机装不了桌面 Excel，WPS 无可视验证通道（`AGENTS.md` §8、`docs/word-parity-spec.md` §8.12）。**现在的替代证据已经是五层**，比原来的三层强在"能机检、且真被办公软件读过"：
   ① 自读一致（我们的 parser 读回自己的产物）；② OOXML 结构断言；③ **第三方独立读取器**交叉读（exceljs 4.4.0、openpyxl 3.1.5）；
   ④ **结构审计**（`compat-audit.ts`，`6b770f5`）——把 Excel 会查的 ECMA-376 硬约束写成 8 个问题码，`oc sheet check <file>`（`72e95f6`）能单独当发布前闸门（有问题时退出码 1）；
   ⑤ **Numbers 实开**（§二 P9 ②，本机唯一的办公软件通道）：`tmp/chart-out-line.xlsx` → `tables=1 charts=1 A1=月份`，`tmp/chart-app-saved.xlsx`（app 保存那条路出的盘）→ `tables=1 charts=1`。
   审计的判据不是自证的：故意做一份坏文件（删掉画图部件的关系表 + 颠倒 `pageMargins`/`dimension`）时，**审计报 2 条，Numbers 打得开但 `charts=0`**（图静默丢失），好文件是 `charts=1`——两边对"这块内容没了"的结论一致。
   **仍标 BLOCKED 的部分**：Excel 打开后**画面上**对不对（图表配色、字体、列宽目视效果）、WPS 的行为，本机没有通道，须由装了 Excel/WPS 的机器补一次。

### B. AI 行为（对应 §零.4/§零.5）

9. `[x]`（读写工具六个，`dba6fb4` + `1f5ec47`）所有表格工具入参过 zod `.strict()`，非法入参返回给模型的是错误串而不是崩溃（对齐 `apps/server/src/agent.ts:117` 的做法）。
   **实测**：`parseToolArguments` 是唯一入口，11 条单测钉住非法 JSON / 空串 / 数组 / 多余键 / 越界 / 不认识的工具名都成串不成异常；
   写侧只开 `sheet_submit_proposal` 一个出口，条目只有四类自带逆操作的改动，`sort_rows` / `set_number_format` 这类提不出来（单测钉住被拒）。
   工具环这一侧同样不抛：`sheetAgent.test.ts` 里非法入参与幻觉方法名都变成回给模型的一句话，回路接着走。
10. `[x]`（引擎侧，`b87a88b`）**数字来自引擎**：单测固定住"聚合值由 `sheet-io` 计算"，模型给出的数字若与引擎不符，以引擎为准并记 warning。
    **实测**：`describeSheet`/`aggregateSheet`/`findRows`/`readRange` 的值全部由引擎算，并有 3 条用例跑在真实 `.xlsx` 夹具上（画像 `sum=360` 与表里 `SUM` 的缓存值同为一条断言）。
    工具环落地后（`e69780e`），模型能看到的数字只有引擎回包那一份，提示词里写死"不许凭样本心算整表结论"；
    **仍未做**：对模型最终答复里的数字与引擎值做**自动**交叉核对并记 warning——那需要在答复文本里再解析一遍数字，属下一阶段可选强化。
    13e 的人工交叉核对已过一遍：答复"华东 270 / 华北 90 / 合计 360"逐项与 `sheet_aggregate` 回包、与夹具落盘断言的 `sum=360` 同值。
11. `[x]`（上限已锁 `b87a88b` `dba6fb4`，拼装已落 `e69780e`）上送内容默认截断（表头 + ≤50 行样本 + 结构摘要），有单测锁死这个上限。
    **实测**：`readRange` 默认 ≤50 行、单格 ≤120 字，`describeSheet` 默认 ≤50 列并回报 `columnsTruncated`，聚合/命中默认 ≤50 条而计数不受影响；各有单测。
    `sheetContextOf` 的拼装由两条单测钉住：五行小表不出现截断说明；70 行的表只上送到第 50 行并写明"共 70 行、后面的行要接着读"。
12. `[x]` 半自动模式：提案必须用户点确认才落地；全自动模式的每步仍受 §2.5 撤销约束。
    **到位**：`runSheetAgent` 全程只读载荷，产出的是 `ops` 与 `notes`，写表发生在既有的 `commitSheetOps`（`App.tsx`）里——它才是唯一落地口，并写 modlog；
    代理自己没有任何写文件的通道（§2.8 的结构性保证）。面板分支在 `f0f9853`：卡片列出四类改动的可读行，半自动只有点"应用到表格"才交 `commitSheetOps`。
    **真实 app 实跑**（`tmp/sheet-ai-probe.mjs` 走 CDP 打构建产物，独立 `--user-data-dir`，不碰用户正在跑的实例）：
    半自动下卡片停在 `pending`，网格无变化；点"应用到表格"后卡片转 `applied`（"已写入表格 / 查看修改历史"），`C6` 出现引擎算出的 `360`，
    标题栏出脏点，Cmd+Z 把这条改动倒回去（网格尾部回到 `…3/31/2026, 销售, E17`）。全自动下不改网格也留有记录，`commitSheetOps` 那条 modlog 一字未改。
13. `[x]` 真实端点探针（一次性、`tmp/`）跑通"分析 / 整理 / 计算"各 ≥1 个任务，记录：模型实际调用的工具序列、失败样本、未转义引号这类解析坑（本仓已有 ark 回包先例）。
    **怎么跑的**：页面里包不了 `aiToolTurn`（contextBridge 属性不可配置，实测 `TypeError: Cannot redefine property`），
    所以探针在 node 里直接 `import` **同一个** `src/ai/sheetAgent.ts`（node 24 的 type stripping 吃得下，它唯一的相对 import 是纯类型），
    只把 `ask` 换成自己发的那一次 HTTP；工具定义、执行、提案→op 的翻译、落地与撤销全是产品代码本体。
    **实跑序列**（`tmp/sheet-agent-13e.json`，服务商 ark，逐条都是新开的 `.xlsx`）：
    - 分析：`sheet_aggregate{groupBy:[地区], sum/count 销量, hasHeader}` → 答复"华东 270（2 行）/ 华北 90（1 行）/ 合计 360"。2 轮，数字全部来自引擎回包（B10 的交叉核对：与阶段 4 夹具上那条 `sum=360` 断言同值）。
    - 计算：`sheet_submit_proposal{set_cell C6 =SUM(C2:C4)}` → 1 条 `setCellValue`，`runSheetOps` 落地 1/1，逆操作撤销 1/1。1 轮。
    - 整理：`sheet_find{销量, "<100"}` → `sheet_submit_proposal{delete_rows[4], set_column_width B 140px}` → 2 条 op，**落地 1/2**，
      被拦的是删行：夹具里 `SUM(C2:C4)` 的引用正好落在要删的带内，引擎按 `#REF!` 挡下并给出那句理由（面板同步报"部分写入（1/2）"）。撤销 1/1。
    **失败样本与坑**：三次都没复现"未转义引号"（`tmp/llm-probe` 那次的长文场景才出现），本轮三个 `arguments` 都能 `JSON.parse`；
    真实 app 里同一句整理指令重复跑过一次，第二次 `set_column_width` 因为现场宽度已经是 140 不再产生 op，卡片显示"1 处改动"却列了两行改动、
    末尾以"未写入"收尾——这是 `ops.length` 与 `edits.length` 两个口径的展示差异，不是漏改，落地侧行为（0 条 op 就不写）是对的。
14. `[x]` 暂停/继续行为与 Word 一致：中断后半截提案留在卡上，点"继续"能接着写。
    **已到位**（代理侧 `e69780e`）：表格的提案不是一段半截文本，所以这里的"暂停"停在**轮与轮之间**——`isPaused` 在下一轮请求发出前收住，`messages`（含已跑完的工具回包）原样带回，
    继续时把 transcript 传回去就不重发摘要（两条单测钉住：暂停时一次请求都没发；继续时请求体里的 messages 就是带回来的那份）。
    **真实 app 实跑**（`f0f9853`，探针 `tmp/sheet-ai-probe.mjs` 的 `PAUSE_MS` 分支；那次的 JSON 日志随探针产物一起清理，下面引的是当时打印的卡片原文）：
    发出整理指令 2.5 秒后点 composer 的"暂停"，卡片当场变成
    `0 处改动 | 已暂停，工具回合停在两轮之间。点「继续」接着跑。` 带"继续/忽略"两颗按钮；点"继续"后**同一张卡**收尾到提案结果，
    中间没有重复上送结构摘要（transcript 直接回灌），网格全程未被动过。

### C. 可撤销性与历史（对应 `AGENTS.md` §2.5）

15. `[x]` 每个 `SheetCoreOp` 都发出自己的逆操作；撤销按行内容指纹定位。
    **实测**（`3d1c357`）：`packages/sheet-io/src/sheet-edit.test.ts` 13 条 + 桌面 `sheetHistory.test.ts` 5 条。
    其中「落地后撤销，格子与已用区域都回到原样」「新写的格被撤销后整格消失，已用区域缩回原大小」钉住逆操作本身；
    「插入行把内容整体下移后，撤销仍命中原来那一行」钉内容寻址；`sheet-file.test.ts` 那条把整条链路做到真实 `.xlsx` 上（改→写盘→重开→撤销→再写盘→读回原样）。
16. `[x]` 插入/删除行列后再撤销早期改动，**不牵连**用户在这期间的手工编辑（用一条专门单测钉死）。
    **已到位**：`sheet-edit.test.ts` 现场被改过时报"认不出这一行"、`sheetHistory.test.ts` 撤销 A1 不把手工改过的 A2 退回去；
    插删的 op 本体在 `ec4bda4`（`insertTracksOp`/`deleteTracksOp` + `applyShiftOp` 位移与逆操作）落地，
    不再是"手工构造下移后的表来模拟"；入口（表头右键菜单）在 `e662914`。
    **实测**：真实 app 里插一行 → 保存（`A1:I5`）→ Cmd+Z → 再保存，`sheet1.xml` 与 `sharedStrings.xml` 与插入前逐字节一致。
17. `[x]` 无法安全回滚时明确报"部分撤销"；modlog 仍每文档 ≤60 条、路径不变。
    **实测**：`App.undoLastSheetEdit` 里 `outcome.failed.length` 时 toast「部分撤销（x/y）：理由…」，与 Word 那条（`ai/agent.ts` 的同款文案）一致；
    存储走 `electron/modlog.ts` 现有四条通道，`CAP = 60` 与 `userData/modlog/<slug(docKey)>.json` 一字未改（没有新建第二份历史，D18 因此仍是 `[x]`）。
    **未到位**：这条只能靠单测与代码路径证明——面板上看不到表格记录（阶段 3 偏差 3），真实 app 里按 Cmd+Z 的观感未验证（E26）。

### D. UI 与工程（对应 `AGENTS.md` §3/§4/§5）

18. `[x]` 没有复制出第二套撤销栈 / LLM 客户端 / 设置存储 / 样式体系 / SQLite 连接。
    **阶段 2 实测**：撤销栈一条都没有新建——行列插入/删除与撤销重做整体顺延到阶段 3（见"阶段 2 实际结果"的偏差说明），样式只新增 `src/styles/sheet.css` 一个文件并由组件自己引入（与 `footnotePane.css` 等同款先例一致）。
19. `[~]` 网格滚动与编辑在 5 万行 × 20 列的工作簿上不卡死。
    **实测数字**（一次性探针 `tmp/sheet-scale.mjs`，node 24 / Apple Silicon，纯 IO 侧不含渲染）：
    5 000 行 × 20 列 = 10 万格 → 打开 11ms + 解析 190ms，堆 185MB；20 000 行 = 40 万格 → 746ms，546MB；
    **50 000 行 × 20 列 = 100 万格 → 合计 2.04s（parse 1.89s），堆 1.2GB**，`zip` 29.2MB。
    结论：解析这一侧勉强踩在"≤2s"的线上，真正的瓶颈是**跨界载荷把 100 万个格子整个 structured-clone 过 IPC**，渲染层与主进程各存一份。
    **已到位**（`8ca38b1`，比原计划提前）：阈值判定只在 `sheet-io` 一处（`readOnlyReason` + `EDITABLE_CELL_LIMIT = 20 万格`），
    超阈值打开时把理由 toast 出来，`App` 的两个写入口（存 / 另存）先拦住，`SheetView` 里编辑框、删除键、列宽热区一律不接，名称框挂「只读」标记；
    `sheet-dto.test.ts` 4 条钉住"到上限仍可编辑 / 超一个就只读 / 数字不带千分位"。
    **仍未到位**：只读**不改变数据量**——整表 `cells` 仍然要经 IPC 进渲染层 state，100 万格那份的堆占用与滚动流畅度没在真实 app 里量过（E26）。
    虚拟滚动的渲染侧口径由两条钉住：`sheetGridModel.test.ts` 的 `visibleSlotRange`（窗口起止、夹在有效区间）与 Chromium 探针实测的槽位窗口；小表（几十行）本来就全量渲染，这是对的，不该出现空洞。
20. `[x]` 列宽拖拽用 `usePointerDrag` 的 pointer capture，热区带 `-webkit-app-region: no-drag`；浮层用 fixed + `useAnchoredMenu` 不出屏。
    **实测**：`SheetView.tsx` 的列宽拖拽与选区拖拽都走 `startPointerDrag`；探针里 `getComputedStyle('.sheet-col-resize').webkitAppRegion === "no-drag"`、`cursor === "col-resize"`。本阶段网格里没有浮层控件（右键菜单/公式面板属阶段 4），故 `useAnchoredMenu` 这条暂无对应实现，不是遗漏。
21. `[x]` 渲染层只经 `window.officeCopilot` 的白名单通道访问主进程；无 `nodeIntegration` 变化。
    **实测**：新增的是 `sheet:read` / `sheet:save` / `sheet:save-as` 三条具名通道，`preload.ts` 里逐条写返回类型（没有万能 `invoke`），`contextIsolation` 与 `nodeIntegration` 一字未改；`sourceXml` 不过 IPC（`sheet-dto.test.ts` 断言载荷里没有 `sourceXml` 键）。
22. `[x]` 新函数有中文 JSDoc（作用/入参含单位语义/失败语义）；命名有含义，无 `data/res/temp/flag`。
    **实测**：`geometry.ts`（px/chars/pt 单位都写进注释）、`sheetGridModel.ts`、`sheetHandlers.ts`、`SheetView.tsx` 逐函数自查过；失败语义写明"配不上 `partPath` 就整批不写""没有底包就回可读消息"。
23. `[x]` 无孤儿文件、无注释掉的代码块、无未被调用的导出。
    **实测**：`sheet:open` 通道与 `openSheetFile`（preload + `global.d.ts`）在"打开对话框同时接受两种文件、只回路径"改造后成了死代码，随即删掉；探针产物（`apps/desktop/tmp/sheet-probe.*`、`tmp/sheet-scale.mjs`）不入库。
    **阶段 3 补记**：`sheetGridModel.ts` 里被 op 层替换掉的 `withCellText`/`withColumnWidth` 连同它们的 5 条单测一起删（语义在 `sheet-io` 那边逐条重钉，不留第二份实现，§3.2/§3.5）；
    `index.ts` 的表格 op 公开面从 12 个导出收到 4 个函数 + `SheetCoreOp`/`SheetOpsResult` 两个类型，其余成员类型都能从这两个结果类型上取到，没人 import 就不摆着。

### E. 测试与门禁（对应 `AGENTS.md` §7）

24. `[x]` `pnpm build` 全绿（类型检查藏在 build 里，`pnpm test` 不能替代）。
   **实测**：`pnpm build > tmp/xlsx-probe/root-build.log 2>&1` → `BUILD_EXIT=0`，日志含 `packages/sheet-io build: Done`。
25. `[x]` `pnpm test` 全绿，且**逐包核对通过数**（新包 `sheet-io`/`sheet-tools` 必须有测试文件，不许靠 `passWithNoTests` 蒙过）。
   **实测**：`TEST_EXIT=0`；逐包核对 —— sheet-io **25**、docx-io 296、desktop 398、word-tools 6、word-client 8、server 2；`packages/shared`/`apps/cli` 仍 "No test files found"（本仓既有状态，不是本次引入的假绿）。
   **阶段 6 复测**：sheet-io **185** / desktop **439** / docx-io 296 / sheet-tools **12** / word-client 8 / word-tools 6 / **apps/cli 3**（该包第一条测试，上一条里的"cli 没有测试文件"从此不再成立）/ server 2；仍无测试的只剩 `packages/shared`。
   **区域取数收口后复测**（`3528019`）：sheet-io **188**（+3 条 `readRange` 收口用例）/ desktop **439** / docx-io 296 / sheet-tools 12 / word-client 8 / word-tools 6 / apps/cli 3 / server 2。
26. `[x]` `pnpm dev:desktop` 起真实 app 手工走通主流程，收尾报告看到什么；跑不起来就如实说明卡点，不许用单测冒充可视验收。
   **实测**：阶段 3/4/5 用临时 CDP harness 驱动**构建产物**里的真实 app（等价于 dev 版，且避开了用户自己在跑的实例），走完"开表→改格→保存→撤销→再保存→插图表→再保存"，看到的卡片文案、按钮、状态与落盘五件部件都记在 §八 对应行；表格模式下没有文档工具栏（`App.tsx` 的 `!isSheet`），保存走命令面板这条已写进 `AGENTS.md` §8 环境事实。
27. `[x]` 测试产物只落 `tmp/` 且清理；不新增任何二进制夹具入库（`AGENTS.md` §7.3）。
    **实测**：夹具由 `scripts/make-sample-sheet.ts` 在测试里现生成，仓库内无 `.xlsx`；探针与日志全在 `tmp/xlsx-probe/`（未跟踪）。
    **已清**（2026-10-13）：阶段 5 的 app 探针目录里那份 `config.json`（从用户 `userData/config.json` **复制来的、含真实 Key**）已删掉，`tmp/chart-ui-profile/` 与 `profile2/` 两份都在；全仓 grep 那段 Key 的明文（含 `tmp/`）零命中，两份探针脚本只在运行时从配置读 `provider.apiKey`，不把密钥写进源码。仍要记住的前提：`tmp/` **没有 `.gitignore` 条目**，所以它一直以待跟踪目录出现在 `git status` 里，任何 `git add -A` 都会把当时的探针产物（可能含复制来的配置）一起带上——这条已写进 `AGENTS.md` §7.3，禁用 `git add -A` 不是可选项。
28. `[x]` 提交信息中文、`<type>(<scope>)`、一事一提交，功能与格式化不夹带；已推送并复核远端。
    **实测**：阶段 6 拆成 `a72196e`（feat(cli)：只读子命令 + 该包第一条测试）与本提交（docs：README/AGENTS/spec 回填）两件事，各自推送后 `git status -sb` + `git ls-remote origin main` 双重复核；用户自己的未提交 WIP（`HistoryPanel.tsx`、`styles.css`）没有被夹带进任何一个提交。

---

## 七、风险与未决

| 风险 | 影响 | 缓解 |
| --- | --- | --- |
| 原生 chart part 手工生成失败率高 | 阶段 5 的核心卖点 | 阶段 0 就 spike，失败即降级为贴图并**显式记录**；不把它伪装成图表 |
| 未建模部件"逐字节保留"在 sharedStrings/styles 上做不到（它们互相引用，改一个值可能动 string 表） | 阶段 1 红线 | **已被实测消解**（§二 P2/P3/P5）：策略是定点替换 `<sheetData>` + string 表**只追加** + `cellXfs/numFmts/fonts` **只追加不重排**（位置索引重排会让用户格式错位）。单测必须钉住"改一个格不移动任何索引" |
| 大表全量载入内存 | 阶段 2 可用性 | 阶段 1 就定"稀疏 + 分 sheet 惰性解析"；超阈值（如 >20 万格）明确提示只读 |
| 本地公式求值与 Excel 结果不一致 | AI"数据计算"可信度 | 只支持白名单函数；UI 与 modlog 里标注缓存值来源；不一致的场景（循环引用、volatile 函数如 `TODAY()`）一律留空 `<v>` |
| **一批 op 里同一行改两格时第二条被"认不出这一行"跳过**（阶段 4 实测发现，`0e7cb74` 的交叉读探针钉出） | AI 的"加两列计算值""整表整理"都要每行发多条 op，按现状会**静默少改**，只剩 `skipped` 里一条理由 | **已解决**（`7de845f`）：批内记"旧指纹 → 改后的指纹"的锚点表，只在同一批内部生效；同一格两次、批外改动仍按 C15 拦住，四条单测钉住 |
| 同一族的**列轴**版本：一批里改多列列宽时，第一条把 `<col>` 分段劈开，第二条期望的分段形状对不上就被跳过 | "统一列宽 / 自动列宽"这类整理任务会静默少改若干列 | **已解决**（`512fb0f`）：现场分段与记录逐段一致时走精确路径整段换回（撤销能还原成原来那一段写法），不一致时只核对目标列自己的宽度没变、再按现场重算覆盖带；同一列两次、用户改过宽度仍按 C15 拦住，七条单测 + 一条 `.xlsx` 落盘往返钉住 |
| **整理写侧缺 op**：现有四类 `SheetCoreOp`（改值/改公式/列宽/插删行列）里没有"重排整行""拆一列成两列""统一数字格式"的构造 | 阶段 4 的"数据整理"只能做去重与列宽两项；排序若用一批改值凑出来，前像整表那么大且撤销语义要说清 | **已定案（提交 13，`1f5ec47`）**：只开四类能安全表达的改动，排序/拆合列/统一数字格式在 schema 里根本不存在（`sheet_submit_proposal` 的 discriminated union 拒掉未知 `type`），系统提示词里也写明"这些做不了，遇到就直说"。**没加**第五类 move op——那是为假想的未来做抽象，且撤销前像与 §2.5 的内容寻址要先重新设计。**残留风险**：用户真想要排序时会觉得产品缺功能，届时要按 spec 单独立一项再设计 op，不要临场用一批改值凑 |
| **新建 OPC 部件时只塞部件、不登记关系与内容类型**（阶段 5 实测发现：`xl/sharedStrings.xml` 从零新建时 openpyxl 读回 `IndexError`，Excel 判损坏，§二 P8） | 任何"来源表没用 string 表"的文件（openpyxl 与不少导出器全程 `t="inlineStr"`）存一次就坏一次，且单测只看自家读回是发现不了的 | **已解决**（`fa693d2`）：`writeSharedStringsPart` 把"部件 + workbook 关系 + `[Content_Types]` Override"绑成一个动作，只在原文件没有这张表时补；判据改成**第三方读取器审计我们的产物**（openpyxl 读回），而不是只信自家 parser——这条以后每次新增部件都适用 |
| **模型出图时不一定走 `sheet_create_chart`**（真端点实测：同一份表，"画成折线图"1 轮调图表工具，"画成柱状图"却改走 `sheet_aggregate`+`sheet_submit_proposal` 往 G1:H4 写辅助区） | "生成图表"这条卖点不可复现，验收里不能写"每次都会出图" | 现状可接受：两条路都是用户能看见并确认的（§2.8），辅助区那条照样可撤销；工具描述里已写明"要图就调这个、不要自己报数"。若后续要确定性，只能在**提示词或强制 tool_choice** 上想办法，不要为此往引擎加"猜意图"的分支 |
| **模型报一个表里没有的区域，取数就回一整块空格**（阶段 6 实测发现：CLI 探针 `oc sheet read -r Z99:ZZ2` 返回过 `Z2:ZZ51 98 行 × 677 列` 全是制表符） | 上送内容失控（§2.9），更糟的是模型看到"区域读得出来"就继续按幻觉的区域推理，验收 B10 的可信度整条塌掉 | **已解决**（`3528019`）：`readRange` 先把请求与已用区域**求交**再按上限截断，`ref` 回实际读出的范围、`rowCount`/`colCount` 回求交后的真实值；整块在数据区之外时给一句含实际数据区的理由（`「…」在这张表的数据区（A1:J4）之外`）而不是空格；列同样封顶 50 条。三条单测钉住，已用区域的写法抽成 `usedRangeOf` 供画像/理由/CLI 共用（§3.2） |
| **"我们的文件能被 Excel 打开"只能靠自家 parser 自证**（用户明确要求导出可在办公软件直接打开） | 自家 parser 与自家 writer 是同一套理解，写歪了会一起读歪（P8 那次 `IndexError` 就是外部读取器抓的）；单测全绿也可能是一份 Excel 拒收的文件 | **已解决**（`6b770f5`+`72e95f6`）：`auditXlsxPackage` 按 ECMA-376 的显式约束（§10.1.2/§9.2/§13.2.1/§18.3.1.100）机检，`oc sheet check` 有问题时退出码 1；同一套判定钉进单测（夹具/写回/加图表三种产物各 0 条），并用一份故意做坏的文件证明它不是橡皮图章（报 2 条，Numbers 打得开但 `charts=0`）。**以后每次新增 OPC 部件都跑一遍这条**（§2.10 的机检面） |
| **Numbers 打得开被当成"Excel 也没问题"** | Numbers 宽容：引用断了它照样开、把丢掉的内容静默省略，用它当判据会把坏文件放过去 | 已按实测划清口径（§二 P9）：Numbers 的读数只用作"第三方真的读过这份文件"的证据，**合规判定归结构审计**，两者结论不一致时（打得开但 `charts=0`）按审计走 |
| **审计全绿被当成"内容也没变瘦"**（§二 P11 实测：`<c t="inlineStr"><is><r>` 的 run 保存一次就被拍平成纯文本，而审计写前写后都是 0 条） | 审计只判规范明写的硬约束（顺序、引用、登记齐全），"文件照样能开、用户的格式没了"它一个字都不报；只看自家 parser 读回同样看不见，因为读的时候本来就没建模 | **已解决**（`bbc7eba`）：按"未建模内容挂 `rawXml` 原样回写"这条既有口径补上 `inlineStr` 那一支（§2.3），改值时 `setCellTextOp` 造的新格子天生不带它，覆盖照常能生效（两条回归单测）。判据也补了一层：合规判据之外再加**逐特性对拍**（openpyxl 自写 10 份夹具 → 桌面真实保存路径 → openpyxl 再读，现在 10/10 逐字段一致）。以后新增"读不懂就留着"的分支都按这条对拍法验，不要只跑审计 |
| 无法用 Word/WPS 目视验证 | 验收永远缺最后一环 | 沿用 Word 线的做法：结构级替代证据 + 明确 BLOCKED（`docs/word-parity-spec.md` §8.12） |
| 引入 Univer/SheetJS 的诱惑（"省几周"） | 破坏全局约束 3/6 | 本文 §二 已留否决理由；未来若要改选，必须先推翻 §零.6 再讨论 |

---

## 八、总验收表（随进度填写，键为提交哈希）

| 阶段 | 提交 | 判据编号 | 状态 | 实测/BLOCKED 说明 |
| --- | --- | --- | --- | --- |
| 0 | 本提交 | P1–P6 | `[x]`（P4 BLOCKED 已定替代证据，P6 顺延阶段 5） | 证据在 `tmp/xlsx-probe/run.log`（不入主干，`AGENTS.md` §7.3） |
| 1 | `e5d5ba4` | A1–A3、A6、A7、E24–E25、E27 | `[x]` | 25 条单测绿；第三方 exceljs 交叉读回值/公式缓存/日期 numFmt/合并/冻结/数据验证/富文本/外部链接部件全部一致 |
| 1 | `e5d5ba4` | A4 | `[~]`（阶段 2 之后仍是 `[~]`） | `<f>`+`<v>` 读写已测；UI 的"需 Excel 重算"提示已在阶段 2 落地（`.sheet-recalc-hint`）；"Excel 真的重算"那一半仍 BLOCKED |
| 1 | — | A8 | `[!] BLOCKED` | 本机 WPS 无可视通道（`AGENTS.md` §8），替代证据 = 自读一致 + OOXML 结构断言 + 第三方 exceljs 交叉读（上一条） |
| 2 | `bbf23c5` `ab27d0e` `c8c3f3b` | A5 | `[~]` | 写回内容正确性有单测钉住；`.bak` 与 `.tmp→rename` 在主进程里，本机没有 electron 进程级"写到一半中断"的测试，按代码推断 |
| 2 | 同上 | D18、D20–D23 | `[x]` | 没新建第二套撤销栈（行列操作与撤销整体顺延阶段 3）；pointer capture + `no-drag` 热区、3 条具名 IPC 通道、JSDoc、死代码（`sheet:open`/`openSheetFile`）清理逐条见 §六 C/D |
| 2 | `bbf23c5` `ab27d0e` `c8c3f3b` + `8ca38b1` | D19 | `[~]` | 10 万 → 100 万格的解析耗时/堆占用实测（190ms/185MB → 2.04s/1.2GB）；"超约 20 万格提示只读"的保护已在 `8ca38b1` 落地（判定单点在 `sheet-io`，写入口全拦）；但只读不减数据量，5 万行那份在真实 app 里的滚动/编辑手感仍未量（见 E26） |
| 2 | 同上 | E24–E25、E27 | `[x]` | 根 `pnpm build` EXIT=0；逐包核对 sheet-io **41** / desktop **416**；夹具现生成、探针产物只在 `tmp/` |
| 2 | 同上 | E26 | `[x]`（阶段 3 收尾补结） | 当时只有真实 Chromium 的样式/几何探针，没有 app 级通道；阶段 3 用临时 CDP harness 驱动构建产物里的桌面 app 走完了"开→改→存→撤销→再存"，见本表阶段 3 的 E26 行 |
| 3 | `3d1c357` | C15、C17 | `[x]` | 13 + 5 条单测钉住逆操作与内容寻址，另有「改→写盘→重开→撤销→再写盘→读回原样」落在真实 `.xlsx` 上；"部分撤销（x/y）"文案落地，modlog 存储与 60 条上限一字未改 |
| 3 | `ec4bda4` `e662914` | C16 | `[x]` | op 层：整行/整列插删各自带逆操作，`removed` 前像按 `target.index` 同框再平移；单测钉住"位移后撤销早期改动不牵连用户后来的手工编辑"。入口：表头右键菜单，真实 app 里插一行→保存（`A1:I5` + `.bak`）→Cmd+Z→再保存，**两个部件逐字节回到原样** |
| 3 | `c099727` | §2.3（写回不拍平局部格式） | `[x]` | 真实 app 保存时富文本格被重新 intern，颜色丢 + 追加重复 `<si>`；改成按 run 逐字段指纹复用，三条单测钉住（含"走桌面载荷那条路"）。实跑结论：`uniqueCount` 保持 11、`A2` 仍指原条目 |
| 3 | `ec4bda4` `e662914` `c099727` | E24–E25 | `[x]` | `pnpm build` EXIT=0、`pnpm test` EXIT=0；逐包核对 sheet-io **96** / desktop **416**（37 个文件）/ docx-io 296 / word-client 8 / word-tools 6 / server 2 |
| 3 | 同上 | E26 | `[x]` | 临时 CDP harness 驱动**构建产物**里的真实 app：开→拒绝（合并区被劈开，文件未动）→插入→保存→撤销→再保存回到逐字节一致；`sharedStringIndex`/`runs` 经 contextBridge 双向无损这一点也实测过（`PROBE_RECV` 那次）。踩到的坑（旧实例没退净 → 调试端口连上旧进程）写进阶段 3 收尾 |
| 4 | `470d370` | §零.4（数字来自引擎） | `[x]`（引擎侧） | 列画像 / 分组聚合 / 去重键落在 `packages/sheet-io/src/analysis/**`：模型只指定"按哪几列分组、对哪列做什么聚合"，值一律由引擎算；列引用解析不出时返回 `ok:false` + 理由，不抛异常 |
| 4 | `0e7cb74` | A4（写侧）、B9（求值器不崩） | `[~]`（A4 的"Excel 真重算"仍 BLOCKED） | 白名单函数当场算出 `<v>`，算不出的只写 `<f>`；两条形状有落盘单测，第三方 exceljs 交叉读到 `formula/result` 与"只有 formula"。逐包核对 sheet-io **152** / desktop **416**，`pnpm build`+`pnpm test` EXIT=0 |
| 4 | `7de845f` | C15 在批内的死角（同一行多条 op） | `[x]` | 交叉读探针实测：一批里改同一行的两格，第二条被"认不出这一行"跳过。`applyOps` 改成批内重新锚定（`reanchorOp`/`rememberReanchor`），四条单测钉住；探针重跑 `op ok: true`，exceljs 读到 `C2={formula:"C3+C4",result:240}`（见 §七） |
| 4 | `512fb0f` | C15 同族的列宽批内死角 | `[x]` | `applyColumnWidthOp` 原先要求现场 `<cols>` 分段与记录逐段一致，一批里改第二条就被跳过。现在精确路径优先、不一致时按现场重算覆盖带；七条单测 + 一条落盘往返（批内两列都落地、撤销后 `widthSegments` 与批前逐段相等） |
| 4 | `b87a88b` `dba6fb4` `cd65c5b` | B9、B10、B11（读侧） | `[x]`（读工具面；B10 的"与引擎不符记 warning"与 B11 的"拼装上送"仍待提交 13） | 引擎新增 `readRange`/`findRows`/`describeSheet` 的 `maxColumns`，条件判定复用 `formula-eval.toCriteria`（筛出的行与 `COUNTIF` 同一口径）；新包 `packages/sheet-tools` 只依赖 zod、不 import `sheet-io`，`parseToolArguments` 永不抛异常。三条读路径各有真实 `.xlsx` 夹具上的落盘用例。本次门禁：`pnpm build`/`pnpm test` EXIT=0，逐包核对 sheet-io **176**（12 个文件）/ **sheet-tools 9**（新包）/ desktop 416 / docx-io 296 / word-client 8 / word-tools 6 / server 2 |
| 4 | `c52c1b7` | §5.2（跨界一次改三处）、B12 的前置 | `[x]` | 新增 `ai:tool-turn`：主进程只转发一轮（收 `messages`+`tools`，回 `content`+`toolCalls`），工具在哪执行由渲染层决定，因为表格载荷只在那边。顺带修好一处潜伏缺陷——`runStream` 此前没把 `AbortController` 的 `signal` 接到 `fetch` 上，按"暂停"只是不再回抛事件，请求仍在后台跑到 180 秒；现在三条调用路径共用 `postChat`，单测钉住"取消真的中止 fetch"。本次门禁：`pnpm build`/`pnpm test` EXIT=0，逐包核对 desktop **423** / sheet-io 176 / sheet-tools 11 / docx-io 296 / word-client 8 / word-tools 6 / server 2 |
| 4 | `e69780e` | B9、B10、B11、B14（代理侧）、§零.4 | `[x]`（代理侧，UI 分支待 13d） | `src/ai/sheetAgent.ts` 跑工具环：上送只有"结构摘要 ≤50 列 + 样本 ≤50 行"，六件工具全部转给 `sheet-io` 执行，非法入参/幻觉方法名回一句话不抛；`sheet_submit_proposal` 的四类条目翻成 op，翻不动的逐条给原因，删行会让手写公式变 `#REF!` 时由引擎的现场核对挡住（单测实测到那条理由）。暂停/继续按"停在轮与轮之间"实现：`isPaused` 收住、`transcript` 带回、继续时不重发摘要，转满 8 轮没结论就停。本次门禁：`pnpm build`/`pnpm test` EXIT=0，逐包核对 desktop **437**（新增 `sheetAgent.test.ts` 14 条）/ sheet-io 176 / sheet-tools 11 / docx-io 296 / word-client 8 / word-tools 6 / server 2 |
| 4 | `f0f9853` | B12、B14（面板侧）、D 组（复用现有卡片样式） | `[x]` | `src/ai/AIPanel.tsx` 多一支 `t:"sheet"` 的条目与 `SheetPanelBridge`（只交载荷 + `onCommitOps`），`App.tsx` 把同一个 `commitSheetOps` 升级成返回 `SheetOpsResult \| null`，好让卡片能如实报"部分写入（x/y）"。零新增 CSS（复用 `.ai-card*`/`.op-list`），面板按模式各挂一份 `key` 让文档与表格的会话互不串台。**真实 app 实跑**（CDP 打构建产物 + 独立 `--user-data-dir`）：半自动提案卡 →"应用到表格"→ `C6` 出现引擎算出的 `360` → Cmd+Z 回到原样；全自动那条 toast"已直接写入表格，可在修改历史里撤销这次改动"；暂停后卡片写"已暂停，工具回合停在两轮之间"，"继续"在同一张卡上收尾。本次门禁：`pnpm build`/`pnpm test` EXIT=0，逐包核对 desktop **437** / sheet-io 176 / sheet-tools 11 / docx-io 296 / word-client 8 / word-tools 6 / server 2 |
| 4 | —（`tmp/` 探针，§6.5 不入主干） | B13、B10 的答复侧交叉核对 | `[x]` | node 侧 `tmp/sheet-agent-probe.mjs` import **同一个** `sheetAgent.ts`，只把 `ask` 换成自己发的 HTTP，三条任务各新开一份 `.xlsx`。序列与数字见 B13。顺带钉死一条环境事实：contextBridge 的属性不可配置，页面 JS 包不了 `aiToolTurn`，工具序列只能 node 侧抓。阶段 4 到此收口，阶段 5（图表）与阶段 6（CLI/文档）另起提交 |
| 5 | `294ce88` | D5 的写侧（阶段 0 就标了风险的 P7） | `[x]`（写侧；查看/工具面待下一个提交） | `resolveChartSpec` 只收 A1 引用，数值/引用串/显示串一律引擎从现场表取；`SheetFile.addChart` 五件一起登记（图表本体、画图部件、两层关系、content types 两条 Override），表部件只插一个 `<drawing r:id>`。**独立读取器复核**：`tmp/chart-out-{bar,line,pie}.xlsx` 用 openpyxl 3.1.5 读回，`charts=1`、`val='销售'!$C$2:$C$4` 缓存 `120/150/90`、类别缓存 `粗蓝/2月/3月`、锚点 `(col 4,row 1)`、含 `<&"` 的标题原样取回；一张表已有 `<drawing>` 时第二次请求返回 `drawing-exists`，不重排别人的关系。本次门禁：`pnpm build`/`pnpm test` EXIT=0，逐包核对 sheet-io **183**（新增 `xlsx-chart.test.ts` 7 条）/ sheet-tools 11 / desktop 437 / docx-io 296 / word-client 8 / word-tools 6 / server 2 |
| 5 | `e19aece` | D5 的工具面与 app 内侧（`sheet_create_chart` → 卡片 → 随保存落地） | `[x]`（工具面与保存侧；ECharts 查看与 canvas→PNG 见下一行） | 引擎侧 `sheetAgent.ts` 多一条 `sheet_create_chart` 与 `{kind:"chart"}` 回合出口；面板三种模式各一条路（半自动挂"插入图表"卡片、聊天只念引用与数值、全自动直接排进待写清单）；`App.tsx` 用 `pendingCharts` 按 `partPath` 暂存（图表是包部件不是格子，不进工作簿载荷；同表再来一张换掉前一张，因为 `CT_Worksheet` 只允许一个 `<drawing>`）；`saveSheetDoc(..., charts)` 逐张 `addChart`，插不进就整批取消并回一句中文（§2.4/§5.7）。**真端点实跑**（`tmp/chart-agent-probe.mjs`，node 侧 import 同一个 `sheetAgent.ts`）：说"画成折线图"时模型 1 轮直接调 `sheet_create_chart`，引擎从现场表读成 `cat='销售'!$A$2:$A$4`、值 `[100,120,150,150]`，写盘 7902 字节；同一份表上说"画成柱状图"，模型这次改走 `sheet_aggregate`+`sheet_submit_proposal`（往 G1:H4 写辅助区）——工具可达但不被确定性选中，这条如实记进 §七。**真实 app 实跑**（§7.2，CDP 打构建产物 + 独立 `--user-data-dir` + 本地 config，`tmp/chart-ui-probe2.mjs`）：`tmp/chart-app.xlsx` 打开 → 切半自动 → "按月份把销售额画成折线图" → 卡片标题「AI 表格提案 / 插入一张折线图」、计数「1 张图表」、按钮 `插入图表/放弃/忽略` → 点后状态变「已排进待写清单」并如实提示"撤销栈管不到图表，要回退请用同名 .bak" → 表格模式没有文档工具栏（`App.tsx:1210` 的 `!isSheet`），保存走命令面板（⌘K→"保存"）→ 落盘产物五件齐全（`xl/charts/chart1.xml`、`xl/drawings/drawing1.xml` + 它的 rels、`xl/worksheets/_rels/sheet1.xml.rels`、两条 Override）+ `.bak` 在。**独立读取器复核 app 产物**：openpyxl 3.1.5 读回 `charts=1`、`LineChart`、`val='销售'!$C$2:$C$5` 缓存 `100/120/150/150`、`cat='销售'!$A$2:$A$5`、锚点 `(col 6,row 1)`、`A1=月份 C2=100 C5=150`、`5×5` 未变形。本次门禁：`pnpm build`/`pnpm test` EXIT=0，逐包核对 sheet-io **185** / sheet-tools 12 / desktop **439**（新增 chart 那条与"区域取不到数"那条）/ docx-io 296 / word-client 8 / word-tools 6 / server 2 |
| 5 | `fa693d2` | §二 P8：只有 inlineStr 的来源表存盘后读方找不到字符串表（独立读取器抓出的 §2 红线） | `[x]` | 驱动 app 那条路时让 openpyxl 审计我们自己的产物，`tmp/chart-agent-out-折线图.xlsx` 直接 `IndexError: list index out of range`。比对 zip 清单定死原因：来源表全程 `<c t="inlineStr">`（openpyxl 就这么写）时包里没有 `xl/sharedStrings.xml`，我们新建部件却没登记 workbook 关系与 `[Content_Types]` Override。`writeSharedStringsPart` 只在"原文件没有这张表"时补这两处；原有表的结构逐字节不动（单测第二条钉住）。顺带把追加关系那段逻辑收敛成 `relsWith`/`nextRelIdFor` 给图表复用（§3.2）。重跑探针：`tmp/chart-app-saved.xlsx` 与二次保存的 `chart-app-saved2.xlsx`（幂等）都被 openpyxl 读回 `BarChart`、`cat='销售'!$A$2:$A$4`、缓存 `100/120/150`、`D2=已核对`；app 那次实跑的产物里 `xl/_rels/workbook.xml.rels` 也带上了这条关系 |
| 5 | 待起 | D5 剩下的一半：ECharts 查看（`src/sheet/ChartView.tsx`）与 canvas→PNG 报告插图（P6） | `[ ]` | 需要先答 §6.4："现有依赖能否覆盖"——仓内没有图表库，相邻先例只有 `apps/desktop/src/ai/diagramCanvas.ts:136 renderDiagram`（canvas 2D → PNG dataURL）。本次刻意不预先造第二套画图代码：卡片把引擎解析好的引用/数值/落点排成中文给人看，真正的图在 `.xlsx` 里是原生对象 |
| 6 | `a72196e` | 阶段 6 完成判定（三端一致）、D18、E24–E25、E27 | `[x]` | `apps/cli/src/sheet.ts` 六条只读子命令**只做装配**，import 的是桌面在用的同一份 `sheet-io`（不是各写一套再对拍），所以"三端对同一份 `.xlsx` 行为一致"结构性成立。同一份夹具上的实测数字：`oc sheet stats` → `C 销量 number 非空 3 互异 3  Σ 360 均值 120 中位 120 90~150`、`E 增长 … Σ 12.00% 缺失 2`、`F 备注 empty 缺失 3`；`oc sheet aggregate --by 地区 -m sum:销量 -m avg:销量` → `华东 2 行 sum(销量)=270 average(销量)=135` / `华北 1 行 sum=90 average=90` / `共 2 组`——与桌面网格显示的显示串、与 openpyxl 读回的缓存值同一口径（`3月` 那份 `sum=300` 也这样核过）。CLI 唯一新增依赖是仓内的 `@office-copilot/sheet-io`（§6.4：引擎已有，不引第三方表格库）。本次门禁：`pnpm build` EXIT=0（顺带抓出测试夹具里 `freeze: null` 的类型错——那个字段是"可选"不是"可空"）、`pnpm test` EXIT=0，逐包核对 **apps/cli 3**（该包第一条测试）/ sheet-io **185** / sheet-tools **12** / desktop **439** / docx-io 296 / word-client 8 / word-tools 6 / server 2 |
| 6 | `3528019` | B9、B11（区域取数的收口） | `[x]` | 阶段 6 的 CLI 冒烟里抓出的缺陷：`oc sheet read -r Z99:ZZ2` 回了 `Z2:ZZ51 98 行 × 677 列` 全空。修法是 `readRange` 先与已用区域求交、列封顶 50，整块在区外时给一句含实际数据区的理由（同一条夹具现在回 `「Z99:ZZ2」在这张表的数据区（A1:J4）之外…`，`A1:AZ60` 收口成 `A1:J4 共 4 行 × 10 列`）。`ref` 改成"实际读出的范围"、`rowCount`/`colCount` 改成"求交后的真实值"，`truncated` 行列任一方向收口都算——所以模型看得见"剩下的没上送"而不是"表就这么大"。已用区域的 A1 写法抽成 `usedRangeOf`（表画像 / 取数理由 / CLI `sheet info` 三处共用，§3.2），空表判定复用包内既有 `isSheetEmpty`（它此前是个没人调用的导出，§3.5）。本次门禁：`pnpm build` EXIT=0、`pnpm test` EXIT=0，逐包核对 sheet-io **188**（+3 条收口用例）/ desktop 439 / docx-io 296 / sheet-tools 12 / word-client 8 / word-tools 6 / apps/cli 3 / server 2 |
| 6 | 本提交（docs） | 阶段 6 完成判定（§9 标清已落地/待落地）、E28 | `[x]` | `README.md`：`oc sheet` 命令表与只读说明、§三 下的 Excel 小节（网格/公式/op 撤销/五个读工具/原生图表/验收口径）、示例脚本第 5 条用真实输出、逐包通过数改成 2026-10-13 的实测、`已知边界` 里 Excel 一行改写成"已落 / 未落"、四节里写明 `tmp/` **没有 `.gitignore` 条目**。`AGENTS.md`：§0 范围加表格、§2 新增 **2.10 新建 OPC 部件三处一起登记**与 **2.11 数值一律由引擎算、模型只交引用**，§2.1/2.3/2.5/2.8/2.9 补上表格侧的单点与机检位置，§3 禁止第二套 xlsx 实现，§5.1 依赖方向写清表格侧，§7.3/§8 补两条环境事实（独立读取器代替可视通道、表格模式的保存入口），§9 表格里表格相关项逐条标"已落地"并写出测试文件。**仍未落**（照实标 `[ ]`）：app 内图表预览与 canvas→PNG（上一条那一行）、CLI 写侧、`sheet:*` 的 docx 报告嵌入；A4/A8 的 Word/Excel 真重算与可视验证本机没有通道，保持 `[!] BLOCKED` |
| 6 | `6b770f5` | §二 P9/P10、A8 的机检那一半（用户要求：导出要能在办公软件里直接打开） | `[x]`（合规判定；目视仍 BLOCKED） | 新增 `packages/sheet-io/src/compat-audit.ts`：`auditXlsxPackage(bytes)` 把 Excel 打开时会查的显式硬约束翻成 8 个问题码——§10.1.2 每个部件都要被 `[Content_Types].xml` 覆盖（`part-not-declared` / `declared-part-missing`）、§9.2 关系 Target 必须存在且编号不重复、§13.2.1 部件里的 `r:id` 要在自己的关系表里查得到、§18.3.1.100 `<worksheet>` 子元素顺序与唯一必选的 `<sheetData>`、图表部件不许是空壳。依据用 MS-OFFMACRO2 公布的 schema 片段逐条核对过（不是"Excel 大概会喜欢"）。顺序判定**逐对相邻比**不用全程最大值，否则一处颠倒会把后面合规的元素全误报。引擎只出码不出文案（§5.7）。10 条单测：7 类问题各钉一条合成用例 + 我们自己的三种产物（原始夹具 / 改格写回 / 加图表写回）各 0 条。本次门禁：`pnpm build` EXIT=0、`pnpm test` EXIT=0，逐包核对 sheet-io **198**（+10）/ desktop 439 / docx-io 296 / sheet-tools 12 / word-client 8 / word-tools 6 / apps/cli 3 / server 2 |
| 6 | `72e95f6` | P9 ③ 的落地入口（体检要能被人和脚本用起来） | `[x]` | `oc sheet check <file>`：逐条报"部件 + 一句能照着改的中文"，有问题时**退出码 1**，能当发布前闸门。文案表是 `Record<CompatIssue["code"], string>`，新增问题码漏配文案直接编译不过（不是运行时才发现）。实跑：`tmp/chart-out-line.xlsx` → `没有发现会让办公软件判损坏的结构问题。`（EXIT=0）；故意做坏的 `tmp/broken.xlsx`（删掉 `xl/drawings/_rels/drawing1.xml.rels` + 把 `pageMargins` 插到 `dimension` 前）→ 报 2 条（EXIT=1）。**同一份坏文件交给 Numbers 打得开、但 `charts=0`，好文件 `charts=1`**——审计的判定与办公软件的实际后果对上了（§二 P9 的对照实验）。apps/cli 测试 3 → **4** |
| 6 | 本提交（docs） | P9/P10 取证入文、A8 改写、§七 两行、README 的验证口径 | `[x]` | `docs/excel-support-spec.md`：§二 新增 P9（本机三条可用通道：openpyxl 只读不判、Numbers 走 AppleScript 能问出 `charts=` 的实际读数、自写的结构审计；并划清"Numbers 打得开 ≠ Excel 认"）、P10（与 openpyxl 的 chart XML 逐元素对拍，我们不缺它写的任何东西，风险在引用与登记齐全度）；A8 从"三层替代证据"改成五层并说明**只剩目视那一半 BLOCKED**；§七 加两行（自家 parser 自证的假绿、把 Numbers 能开当成 Excel 没问题）。`AGENTS.md` §8 补 Numbers 可 AppleScript 驱动这条环境事实（含 `-2741` 的坑）。`README.md` 加"怎么确认导出的文件能在办公软件里打开"小节。**顺带修**上一条 `3528019` 行末重复的 `/ apps/cli 3` |
| 6 | `bdf3ba4` | A5 的写盘前闸门（P9 ③ 接到桌面那条路上） | `[x]` | `saveSheetDoc` 出字节后、`写 .tmp → rename` 之前过审计：**只拦本次改动新引入的问题**（底包原本就有的放行，否则第三方导出的文件会让人存不了盘），拦下整批不写并回一句中文，渲染层用已有的红 toast（§2.4 的事务语义，不新增 preload 通道）。审计自己读不懂陌生 XML 时按"判不了"放行——判不出不许挡住保存。3 条判定单测（`electron/sheetHandlers.test.ts`，`vi.mock("electron")` 沿用 `aiHandlers.test.ts` 的先例）钉住"旧的放行 / 新的必拦 / 同码不同部件不互抵"。**闸门不误伤是实测**：`chart-app-saved.xlsx`、`chart-app-saved2.xlsx`（app 从这条路存过盘的）+ `chart-app/chart-ref/cli-sample/chart-out-{bar,pie}.xlsx` 共 7 份 → 全部 0 条。本次门禁：`pnpm build` EXIT=0、`pnpm test` EXIT=0，逐包核对 desktop **442**（39 个文件，+3）/ sheet-io **198** / docx-io 296 / sheet-tools 12 / word-client 8 / word-tools 6 / apps/cli **4** / server 2 |
| 6 | `bbc7eba` | A1、A2、§2.3（未建模内容不许拍平）＋ 新增 §二 P11 | `[x]` | 逐特性对拍（`tmp/compat-sweep/`，§6.5 只在 tmp 不入主干）抓出的真实缺陷：openpyxl 自写 **10 份各只压一类结构**的夹具，走**桌面的真实保存路径**（`fromBytes` → `sheetPayloadOf` → `applySheetPayload` → `writeSheet` → `toBytes`）后再让 openpyxl 读回来逐字段比。首轮 **9/10 一致、1 份丢了内容**——`<c t="inlineStr"><is><r>` 的 run（粗体 + `FF0000FF`）只剩拼接文本 `粗普通蓝`，还凭空新建一张 `xl/sharedStrings.xml`；其余 9 份（合并+居中+边框、冻结+列宽+换行、日期与自定义 numFmt、跨行跨表公式、条件格式+数据验证、超链接+批注+定义名、表对象、原生图表、行高+隐藏列+分组+打印设置）一字未变。修法是既有口径补一支：解析 `inlineStr` 时也把整格原文挂上 `rawXml`，`cellXml` 原样搬回去，`countStringReferences` 从此不为 inline 文本凭空建部件；改值那条路不受影响（`setCellTextOp` 造的是新格子对象、天生不带 `rawXml`）。**修复后 10/10**：`rt-feat-richtext-empty.xlsx` 的 `A1` 与原文逐字节相同，openpyxl 读回 `CellRichText([b=True 粗, 普通, color=FF0000FF 蓝])`，`xl/sharedStrings.xml` 不再存在；10 份的审计写前写后都 0 条。两条回归单测（没改过的保留 / 改过的照常覆盖）。本次门禁：`pnpm build` EXIT=0、`pnpm test` EXIT=0，逐包核对 sheet-io **200**（14 个文件，+2）/ desktop 442 / docx-io 296 / sheet-tools 12 / word-client 8 / word-tools 6 / apps/cli 4 / server 2 |
| 6 | 本提交（docs） | P11 取证入文、A1/A2 回填、§七 加一行 | `[x]` | `docs/excel-support-spec.md`：§二 新增 P11（十份逐特性对拍的做法、首轮 9/10、抓到的那一条，以及"审计全绿不等于内容没变瘦"这条判据边界）；A1 补对拍结论、A2 补 inline 富文本这一支的来源与 openpyxl 的实测读数；§七 加一行"审计全绿被当成内容也没变瘦"并写明以后的验法。代码面的修复见上一行 |

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
