# Office 智能助手

三种形态，同一套 LLM 工具定义：

1. **桌面 app（功能最全）** —— Electron + Tiptap 编辑器，AI 直接改当前文档并留可撤销的修改历史，见「桌面应用」；它同时是一台 `.xlsx` 表格 app（网格编辑 + AI 分析/整理/计算 + 写原生图表），见「Excel 表格」。
2. **CLI（最省事，Word 和 WPS 都适用）** —— `oc` 命令直接修改 `.docx` 文件，无需 Office 宿主，输出文件用 Word / WPS / Pages / LibreOffice 都能打开；`oc sheet` 只读分析 `.xlsx`。
3. **Word 任务窗格加载项** —— 在 Word 桌面版/网页版里用自然语言改当前文档（需要 Microsoft Word）。

> 给 AI 代理看的改代码规范（提交格式、领域不变量、收尾自检）在仓库根目录的 **`AGENTS.md`**。

```
office-copilot/
├─ apps/
│  ├─ desktop/       # Electron 桌面 app：Tiptap 编辑器 + AI 直接改 .docx，另带 .xlsx 网格（功能最全）
│  ├─ cli/           # 命令行工具（Node，跨 Office/WPS）
│  ├─ task-pane/     # Office 加载项前端（React + Vite + Fluent UI）
│  └─ server/        # Node/Express：加载项用的 LLM 网关 + 版本历史
├─ packages/
│  ├─ docx-io/       # 直接读写 .docx（JSZip + OOXML），CLI 与桌面端共用
│  ├─ sheet-io/      # 直接读写 .xlsx：OPC 解析、保真回写、结构化 op、分析引擎、原生图表部件
│  ├─ sheet-tools/   # 表格侧暴露给 LLM 的工具定义 + zod 校验（共享）
│  ├─ word-client/   # Office.js Word API 原子操作封装，加载项用
│  ├─ word-tools/    # 暴露给 LLM 的工具定义 + zod 校验（共享）
│  └─ shared/        # 跨端类型与协议
├─ manifest/manifest.xml   # 仅加载项需要
├─ bin/              # oc / ocw 入口、AI改文档.app 等 macOS 集成
└─ AGENTS.md         # AI 代理改代码规范（提交格式、领域不变量、收尾自检）
```

依赖方向是单向的：`apps/*` 依赖 `packages/*`，`packages` 之间不互相依赖业务逻辑；
单位换算（px / 半磅 / twips）与颜色归一化只在 `packages/docx-io` 的进出口发生，
表格侧同理只在 `packages/sheet-io`（日期序列号 ↔ 显示文本、ARGB ↔ `#rrggbb`、px → EMU）。

---

## 零、最快上手（三选一）

```bash
pnpm install                       # 装依赖（原生构建默认跳过 keytar / msal，见「已知边界」）

pnpm --filter @office-copilot/cli build   # 形态 2：命令行
cd apps/cli && pnpm link --global && cd - # 之后可直接用 oc / ocw（可选）

pnpm dev:desktop                    # 形态 1：桌面 app（开发模式）
pnpm --filter @office-copilot/desktop pack:mac   # 打成 .dmg（产物在 apps/desktop/release/，不入库）

pnpm dev                            # 形态 3：任务窗格 + 后端（需要在 Word 里 sideload，见第二节）
```

三种形态都用同一份 `.env`（`cp .env.example .env`，填 `LLM_API_KEY`）。要用 `oc ai` / 桌面 AI 面板 /
任务窗格任一自然语言入口，都必须先配好这个 Key；纯命令行改文档（`oc replace` 等）不需要网络也不需要 Key。

---

## 一、CLI 使用（Word / WPS 都适用）

### 安装 & 构建

```bash
pnpm install
pnpm --filter @office-copilot/cli build
# 可选：把 oc 链到全局
cd apps/cli && pnpm link --global && cd -
```

之后直接用 `oc` 命令（或 `node apps/cli/dist/index.js`）。

### 命令一览

| 命令 | 作用 |
| --- | --- |
| `oc list <file.docx>` | 列出所有段落（下标、样式、文本预览） |
| `oc replace <file.docx> -f 甲方 -r 客户` | 全局替换文字；加 `--regex` 支持正则 |
| `oc insert-paragraph <file.docx> -t "..." -p end` | 在 `start/end/before/after` 插入段落，`before/after` 需配合 `-a <index>` |
| `oc insert-table <file.docx> -d '[["A","B"],["1","2"]]'` | 插入二维表格，默认首行加粗；`--no-header` 关闭 |
| `oc insert-image <file.docx> ./logo.png -w 200` | 在文末插入 PNG/JPEG 图片 |
| `oc ai <file.docx> "把所有'甲方'改成'客户'，并在末尾加签字栏"` | 调用 LLM 生成提案，确认后写入 |
| `oc sheet info <file.xlsx>` | 工作簿概览：每张表的已用区域、已用格数、警告与未建模结构 |
| `oc sheet stats <file.xlsx> [-s 销售]` | 列画像：类型推断、缺失数、众数、Σ/均值/中位/极值、离群值、"待 Excel 重算"的格数 |
| `oc sheet read <file.xlsx> -r A1:D20` | 按 A1 区域取数，输出的是显示文本（日期/百分比已换算） |
| `oc sheet find <file.xlsx> -c 销量 --criteria ">100"` | 按 COUNTIF 风格条件筛行（`>100`、`<=2`、`<>北京`、`北京`） |
| `oc sheet aggregate <file.xlsx> --by 地区 -m sum:销量 -m avg:销量` | 分组聚合，数值由引擎从现场表读（`avg` = `average`） |
| `oc sheet dedupe <file.xlsx> --keys 月份,地区` | 重复行预览：按这些键能安全删几行（删行请回桌面 app，才有撤销与 `.bak`） |
| `oc sheet check <file.xlsx>` | 兼容性体检：按 OOXML 硬约束查这份文件会不会被 Excel / WPS 判损坏（有问题时退出码 1，可当发布前闸门） |

`oc sheet` 全部是**只读**的：与桌面 app、AI 工具层跑同一份 `packages/sheet-io` 引擎，所以三端对同一个文件的
数字一致；每条都支持 `--json` 供脚本用。写侧刻意不在 CLI 落地（撤销、modlog、原子替换都在主进程那条路上）。

所有写命令默认：
- 原地写回时会先生成 `.bak` 备份（`--no-backup` 关闭）
- 用 `-o out.docx` 另存为新文件时不生成备份

### 自然语言修改（`oc ai`）

需要配置 LLM（OpenAI 兼容协议，默认 DeepSeek）：

```bash
cp .env.example .env
# 填入 LLM_API_KEY；改 LLM_BASE_URL / LLM_MODEL 可切到 OpenAI、Azure、Ollama 等
```

```bash
$ oc ai contract.docx
告诉助手怎么改这份文档：把所有"甲方"改成"客户"，并在末尾加一个 3 行 2 列的签字栏表格

修改提案：全局术语替换并追加签字栏
  - 甲方
  + 客户
  + 表格 3×2
    角色 | 签字
    甲方 | （签字）
    乙方 | （签字）
接受并写入文档？ (Y/n)
```

`-y` 可跳过确认（脚本场景用）。CLI 不依赖网络/Word 宿主，WPS 打开输出文件即可。

### 示例脚本

```bash
# 1. 看看文档结构
oc list contract.docx

# 2. 批量替换并另存
oc replace contract.docx -f 甲方 -r 客户 -o contract.revised.docx

# 3. 插入表格
oc insert-table contract.revised.docx -d '[["条款","责任方"],["交付","乙方"],["付款","甲方"]]'

# 4. 撤销：把 .bak 改回来即可
mv contract.docx.bak contract.docx

# 5. 表格只读分析（数字来自 sheet-io 引擎，与桌面 app、AI 工具层同一口径）
oc sheet stats sales.xlsx
oc sheet aggregate sales.xlsx --by 地区 -m sum:销量 -m avg:销量
  华东  2 行  sum(销量)=270  average(销量)=135
  华北  1 行  sum(销量)=90  average(销量)=90
共 2 组
```

### 跨软件兼容性

`docx-io` 输出的是标准 OOXML：用 **Microsoft Word、WPS Office、Pages、LibreOffice、Google Docs** 都能打开。反之，任何这些软件另存为的 `.docx` 也都能被 CLI 读取。

#### 在 WPS for Mac 里调用

Mac 版 WPS 不支持第三方加载项/VBA 调外部程序，因此无法在 WPS 里加一个真正的按钮。仓库里提供了两种"近似内嵌"的方式，都已配置好：

**A. `ocw` 命令（推荐日常使用）**

和 `oc` 参数完全一致，但任何修改命令执行完会自动用 WPS 打开结果：

```bash
ocw ai 合同.docx "把甲方改成客户并加签字栏"
# → 生成提案、你确认、写入、自动跳到 WPS 看结果
```

`bin/` 已加入 `~/.zshrc` 的 PATH。新开终端即可直接用 `oc` / `ocw`。

**B. 双击应用 `AI改文档.app`**

`bin/AI改文档.app` 是一个 AppleScript 打包的小程序：双击 → 选 .docx → 弹框输指令 → 自动跑 `oc ai -y` → 用 WPS 打开结果。

把它拖到 Dock 或 Finder 工具栏即可像普通 App 一样使用。第一次运行 macOS 可能提示"无法验证开发者"，右键 → 打开 一次即可。

**C. Finder 右键 Quick Action（可选）**

想要"右键 .docx → 服务 → AI 改文档"：

1. 打开 `/Applications/Utilities/Automator.app`
2. 新建 **快速操作（Quick Action）**
3. 顶部设置："工作流程当前" → **文档文件**，"位于" → **访客（Finder）**
4. 左侧搜索 **"运行 AppleScript"**，拖到右边，粘贴 `bin/ai-docx.applescript` 内容
5. 保存为 `AI 改文档`

之后在 Finder 右键任何 `.docx`，菜单"快速操作"下就有这一项。

### CLI 与任务窗格的差异

| 能力 | CLI | 任务窗格 |
| --- | --- | --- |
| 安装 | 一条命令构建 | 需要 sideload manifest |
| Word 宿主 | 不需要 | 需要 Microsoft Word |
| WPS 兼容 | ✅ | ❌ |
| 选中当前段落即改 | ❌（按下标/全文） | ✅ |
| 撤销 | `.bak` 文件 / Ctrl+Z | 应用内撤销栈 |
| 文档大纲 | 全文段落 | 前 200 段 |

---

## 二、Word 任务窗格加载项（可选）

### 工作流总览

```
用户在任务窗格输入指令
        │
        ▼
POST /api/word/command  (SSE)
        │
        ▼
后端 LLM tool-calling 循环
   ├─ 需要大纲 → word_list_paragraphs
   └─ 完成规划 → word_submit_proposal
        │
        ▼
任务窗格展示「修改提案 + diff」
        │  用户点「接受」
        ▼
word-client.applyEdits() 在 Word.run 内执行
        │  执行前 body.getOoxml() 拍快照
        ▼
写入完成，快照压入撤销栈 + 写入 SQLite
```

模型**只产生提案**，真正写入必须用户点确认，API Key 只在后端，不进加载项。

## 前置要求

- Node 20+（推荐 22+）、pnpm 9+；本仓库已用 pnpm 11 验证
- Microsoft 365 账号或 Office 2021+（桌面版 Word 或 [Word on the web](https://office.com)）
- 一个支持 tool-calling 的 LLM API Key（OpenAI 兼容协议；默认 DeepSeek）

## 快速开始

```bash
# 1. 安装依赖
pnpm install

# 2. 配置后端环境变量
cp .env.example .env
# 编辑 .env，填入 LLM_API_KEY；如用其它 OpenAI 兼容服务改 LLM_BASE_URL / LLM_MODEL

# 3. 启动前端 + 后端（并行）
pnpm dev
#    - 任务窗格： https://127.0.0.1:5173
#    - 后端：      https://127.0.0.1:3001 （health: /health）
```

首次启动 Vite 会用自签 HTTPS 证书。浏览器访问 `https://127.0.0.1:5173` 时点「高级 → 继续访问」。

### 4. 在 Word 里加载加载项

**桌面版 Word（推荐 macOS/Windows）**

```bash
pnpm sideload
```

`office-addin-debugging` 会注册 `manifest/manifest.xml` 并启动 Word。功能区「开始」选项卡里会出现「智能助手」按钮，点击打开任务窗格。

**Word 网页版**

1. 打开 [Word on the web](https://office.com) 并新建/打开一个文档；
2. 菜单：**插入 → 加载项 → 上传我的加载项**；
3. 选择本仓库的 `manifest/manifest.xml`；
4. 功能区会出现「打开智能助手」按钮。

> 要让桌面版信任自签证书，可运行一次 `pnpm exec office-addin-dev-certs install https://localhost:5173`。

## 常用脚本

| 脚本 | 作用 |
| --- | --- |
| `pnpm dev` | 并行启动前端 Vite 与后端 tsx watch |
| `pnpm build` | 构建所有包和应用 |
| `pnpm test` | 运行所有包/应用的 Vitest 单测 |
| `pnpm sideload` | 桌面 Word 加载 manifest |
| `pnpm manifest:validate` | 调用官方服务校验 manifest（需联网） |

## 能力清单

### 任务窗格 UI

- Chat 式对话，自然语言指令
- 助手回复展示**修改提案卡片**：
  - `replace_text`：旧/新对照
  - `insert_paragraph`：位置 + 内容
  - `insert_table`：表格预览（首行加粗）
- 每个提案有「拒绝 / 接受并写入」
- 「撤销」按钮恢复最近一次写入前的 OOXML 快照
- 快捷按钮：刷新大纲、末尾插段落、插入本地图片
- 显示当前选区字数、可折叠的文档大纲

### `word-client` 原子操作（packages/word-client）

| 函数 | 说明 |
| --- | --- |
| `listParagraphs()` | 列出段落（index、text、style） |
| `getSelection()` | 取当前选中文本 |
| `replaceText({find, replace, scope})` | 全文/选区内的字面量替换 |
| `insertParagraph({text, position, style})` | 在 before/after/start/end 插入段落 |
| `insertTable({rows, headerRow})` | 插入二维数组表格，首行加粗 |
| `insertImage({base64, width, height})` | 光标处或文末插入图片 |
| `getSnapshot()` / `restoreSnapshot()` | 通过 `range.getOoxml()` 实现快照/回滚 |
| `applyEdits(edits)` | 一个 `Word.run` 内批量执行并返回写入前快照 |

### LLM 工具（packages/word-tools）

- `word_list_paragraphs`：让模型查看大纲
- `word_submit_proposal`：打包修改提案

所有工具的入参都用 zod 校验；服务端只接受 OpenAI 兼容协议的 tool-call 输出，模型没有直接写入文档的通道。

### 后端 API

- `GET /health`：返回 `{ok, model, baseUrl}`
- `POST /api/word/command`：SSE 流，事件类型：
  - `status`：状态文本
  - `tool_call`：模型调用了哪个工具
  - `proposal`：一份完整的修改提案
  - `error` / `done`
- `GET /api/history/:sessionId`：当前会话的历史记录

## 三、桌面应用（apps/desktop）：AI 直接改文档

桌面端不再只让 AI 输出文字由你粘贴，而是由模型给出**结构化修改方案**，再按你选的模式决定要不要直接写进文档。

- 入口：AI 面板（⌥⌘L）；命令面板（⌘K）里 `AI：…` 一组命令；选中文字后的浮动气泡。
- 三种模式：`聊天` 只给建议和草稿、不碰文档；`半自动`（默认）先给 diff 预览，点「应用到文档」才写入；`全自动` 直接写入文档，不需要逐次确认。三者都会把改动记进修改历史，全自动同样可以一条撤销。
- 能力：改样式/格式（标题层级、对齐、字号、字体、颜色、行距、段前后距、缩进）、自动排版、纠正错别字、扩写/缩写/润色/改语气/翻译、新增文案、插入表格、插入 canvas 示意图（流程图/结构图，渲染成 PNG 后作为图片插入，导出 .docx 时一起带走）。
- 快捷芯片（扩写/润色/纠正错别字等）默认收起，点「快捷指令」展开；聊天模式下整块隐藏。
- `纠正错别字`、`插入表格`、`插入示意图` 走模型方案；`自动排版`（⌥⌘⇧F）是离线规则排版，同样记录为一次可撤销的修改。
- 样式单位：编辑器里字号一律存 CSS px（`"14px"`），`.docx` 里是半磅；行距/段前后距/缩进在编辑器里是 pt 与倍数，导出时换算成 twips。转换只在 `packages/docx-io` 的进出口做。颜色同理：标记里存 `#rrggbb`，粘贴 HTML 带进来的 `rgb(...)` 在导出时先归一化成 hex，读不懂的颜色直接不写（写进 `w:color` 会让 Word 认为文档损坏）。
- 工具栏第二排在「字号」右边有两个按钮：调色板图标 = 字体颜色，荧光笔图标 = 文字底色。点开会弹出配色格（第一格是「清除」）加一行「自定义颜色 / 自定义底色」，后者是系统原生取色器，不受预置色板限制。这两个菜单用 `position: fixed` 挂在按钮下方：工具栏那排是横向滚动容器，会把普通 `absolute` 子元素裁掉，之前点了像没反应。段落设置弹层同理。
- 整段改字符样式时，只补每段文字自己缺的属性，不会顺手抹掉红色字、荧光笔；如果一段里加粗/字号只有部分文字带（例如句首的加粗引导词），整段设置会被拒绝并提示「段内…只有部分文字带格式」，需要你先选中那段文字——宁可少改，也不把局部格式拍平。
- 每次写入都以「改动块」为粒度记入 `userData/modlog/<文档名>.json`（保留最近 60 条），在 **修改历史**（⌘K → 历史版本）里可以看每一次改动的 op 列表和 **git 风格 diff**：左右两个行号、`-`/`+` 分列、未改动的段落折叠成 `⋯ 未改动 N 行 ⋯`，改动行内部再按中文逐字/英文逐词高亮；只改样式（文本没变）时单独列出 `-样式：旧 / +样式：新`。
- **单条撤销**：diff 面板底部的「撤销这次修改」只回滚那一次 AI 改动。改动按内容定位（不是按位置），所以那之后你自己敲的其它修改不会被牵连；确实无法安全回滚的段落会明确报「部分撤销」而不是猜测。
- 写入是原子的：先按大纲模拟出预期结果，实际事务落地不一致就整批放弃，不会留下半套格式。
- 拖拽改尺寸：窗口四周 5px 是缩放热区（四个角 14px），鼠标移上去是 `↔`/`↕`/`⤡`/`⤢` 双向箭头，按住拖动改宽/高，边语义与最小尺寸夹取在主进程 `window:drag-resize` 里统一处理（最大化/全屏时忽略）。左侧大纲栏与右侧 AI 面板之间的分隔条同理改宽度，双击复位。两类拖拽都走 `hooks/usePointerDrag.ts` 的指针捕获：抓窗口边缘时指针立刻离开 web contents，`mouseup` 根本不会送达，早先用 `window` 上的 `mousemove/mouseup` 就会出现「松了手、下次划过还在改尺寸」；改成 `setPointerCapture` + 元素上的 `pointerup/pointercancel`，松手必然解绑。热区额外带 `-webkit-app-region: no-drag`，否则顶部那条会被标题栏当成拖动窗口。

```bash
pnpm --filter @office-copilot/desktop dev      # 开发
pnpm --filter @office-copilot/desktop test     # 引擎单测（op 寻址/逆操作/diff/排版/图形布局）
pnpm --filter @office-copilot/desktop build    # renderer + main + preload
```

### Excel 表格（同一 app 里的第二种文档：`.xlsx` 网格 + AI 分析/整理/计算 + 原生图表）

完整规格（选型研究、架构决策 D1–D6、阶段 0–6、验收清单 A–E、逐提交实测数字）在 **`docs/excel-support-spec.md`**，这里只说"能做什么、怎么做到的"。

- 打开：文件 → 打开（⌘O）选 `.xlsx`；主进程按扩展名分派到 `sheet:*` 一组具名 IPC（`electron/sheetHandlers.ts`），渲染层仍只经 preload 白名单访问 Node。
- 网格：`src/sheet/SheetView.tsx` + `sheetGridModel.ts`（几何、可见窗口、键盘步进、点→地址都是纯函数）。双击格内编辑，表头右键插/删行列，列宽拖拽走指针捕获。
- 公式：以 `=` 开头当公式，`sheet-io` 里单文件递归下降求值。按 D4 只写公式与缓存值、**不实现重算引擎**，白名单外的公式留给 Excel 重算（列画像会显示"待 Excel 重算 N 格"）。
- 保真：只建模认得的部分，未建模部件原样回写；富文本 run 的颜色在保存时按 run 指纹复用，不会把局部格式拍平（`AGENTS.md` §2.3 同款红线）。
  整段 inline 写法的富文本（`<c t="inlineStr"><is><r>`，openpyxl 与不少导出器就这么写）我们没有建模，**整格原文原样搬**，改过值才重写。
  这条是逐特性对拍抓出来的：让 openpyxl 自写 10 份"每份只压一类结构"的 `.xlsx`，走桌面那条真实保存路径存一遍，再让 openpyxl 读回来逐字段比（spec §二 P11）。
- 撤销：每次改动是**自带逆操作**的结构化 op，撤销按内容寻址而不是按位置，所以不会牵连你后来的手工编辑；存储与文档侧共用同一份 `userData/modlog/<文档名>.json`。
- AI（⌥⌘L 面板在表格模式下走 `src/ai/sheetAgent.ts`）：五个只读工具 `sheet_describe`（列画像）、`sheet_read_range`（区域显示文本）、`sheet_aggregate`（分组聚合）、`sheet_find`（COUNTIF 风格条件）、`sheet_duplicates`（重复行），加上 `sheet_create_chart` 与 `sheet_submit_proposal`。**所有数值由引擎从现场表算，模型只引用不许心算**——这条写在工具描述里。
- 提案能落地四类：改格（`set_cell`，文本以 `=` 开头当公式）、列宽、插行、删行。排序 / 拆合列 / 统一数字格式刻意不开（理由见 spec §七）。
- 图表是**原生 chart part**而不是贴图（D5）：`xl/charts/chartN.xml` + 画图部件 + 两级关系 + `[Content_Types].xml` 两条 Override，五件一起登记；保存时逐张落地，插不进就整批取消。Excel / WPS 里能当图表点开改。
- 三种模式沿用文档侧：半自动先出卡片（把引擎读到的引用与数值排成中文行）再点「插入图表」；全自动排进待写清单；聊天模式只念数字不落地。
- 怎么确认导出的 `.xlsx` 能在办公软件里打开（本机没装 Microsoft Excel / LibreOffice，所以是三层证据，spec §二 P9/P10）：
  ① `oc sheet check <file.xlsx>` —— 按 ECMA-376 的显式硬约束机检（内容类型是否覆盖每个部件、关系是否指向真实存在的部件、
  `r:id` 有没有落点、`<worksheet>` 子元素顺序），引擎里同一份判定是 `auditXlsxPackage`，我们自己的三种产物（夹具 / 改格写回 /
  加图表写回）都是 0 条；② **openpyxl 交叉读**（`PYTHONPATH=tmp/pylibs python3`）——独立读取器对系列引用 / 缓存值 / 锚点 / 图表类型；
  ③ **Numbers 实开**（本机唯一能驱动的办公软件）：`osascript -e 'tell application "Numbers" to open POSIX file "/abs/path.xlsx"'`，
  再问 `count of charts of sheet 1`、`value of cell "A1" of table 1`。
  桌面 app 的保存本身已经带这道闸：`saveSheetDoc` 在 `写 .tmp → rename` **之前**审一遍，只拦"这次改动新引入的"问题
  （底包原本就不合规的第三方文件照样能存），拦下就整批不写、弹一句中文，磁盘上的文件一个字都不动。
  这三层能证明"文件合规、且第三方办公软件真把它读成了表格和图"，**证明不了配色与字体的画面效果**——那一条照旧 `[!] BLOCKED`，
  要装了 Excel / WPS 的机器补一次目视。注意 Numbers 很宽容：一份引用断链的文件它照样打开、只是 `charts=0`（图静默丢了），
  所以"Numbers 打得开"不等于合规，合规判定以 ① 为准。

## 项目结构要点

- **不在前端持有 LLM Key**：前端只通过后端 SSE 调用模型；CORS 只允许任务窗格来源。
- **写入默认要用户确认**：模型产出的是 `Proposal`，点「接受」才执行；执行前自动抓 OOXML 快照。（桌面 AI 面板的「全自动」模式跳过这一步，但每次改动仍记入修改历史，可单独撤销。）
- **会话历史持久化**：`apps/server/data/history.sqlite`（better-sqlite3，WAL 模式）。
- **选区三态处理**：有选区→选区内改；只有光标→光标处插；都没有→作用于 body。
- **regex 替换首版不支持**：Word 原生 `search` 不支持 JS 正则；UI 会给友好提示。后续可做"先 search 到字面量再批量替换"。

## 测试

```bash
pnpm test      # 逐包跑 vitest
pnpm build     # 类型检查藏在 build 里（desktop / task-pane 走 tsc --noEmit）
```

当前各包通过数（2026-10-13 实测）：`sheet-io` 200、`desktop` 442、`docx-io` 296、`sheet-tools` 12、`word-client` 8、`word-tools` 6、`apps/cli` 4、`server` 2。

- `word-tools`：6 个 zod schema 校验用例
- `word-client`：8 个 Office.js mock 用例（列出段落、替换、段落/表格插入、快照回滚、批量 `applyEdits` 等）
- `docx-io`：296 个用例，覆盖 OOXML 读入 / 写出往返、单位与颜色换算
- `sheet-io`：185 个用例（A1 地址与列号、日期序列号与显示文本、ARGB↔hex、列宽/行高几何、公式子集求值与引用解析、`sharedStrings` 富文本 run 往返、结构化 op 的逆操作与批内重新锚定、分析引擎（列画像 / 聚合 / 条件定位 / 重复行 / 区域截断）、`ChartSpec` 校验、chart part 的读回结构断言、新建 `sharedStrings` 时三处一起登记）
- `sheet-tools`：12 个工具入参校验用例（JSON 语法错、缺字段、多字段、越界都回错误串而不是抛异常）
- `apps/cli`：3 个用例（挑表按名字与 1 基序号、`--metric` 的 `avg`→`average`、分组键逗号切分）
- `server`：mock `fetch`，验证 tool-calling 循环产出提案 / 纯聊天状态
- `desktop`：439 个用例。文档侧原有：`ai/docEdit` 内容寻址 + 逆操作 + 单条撤销、`ai/diff` 中文分词 diff + git 风格行号/折叠上下文、`ai/DiffView` 渲染出的 diff 结构、`ai/diagramSpec` 图形布局、`editor/formatDocument` 离线排版幂等且只改格式、`editor/docBlocks.live` 用 App 真实 schema（`editorExtensions.ts`）跑格式写入（局部加粗/红字/高亮不被整段抹平、列表项改对齐不破坏 `bulletList`、排版后撤销能还原原文档）、`editor/templates` 公文/论文模板套版、`hooks/useAnchoredMenu` 弹层贴边不出屏、`hooks/usePointerDrag` 捕获指针后松手必然解绑（面板宽度的方向与夹取也在这里）。表格侧新增：`sheet/sheetGridModel` 几何/可见窗口/键盘步进/点→地址/改列宽、`sheet/sheetHistory` 认形状并倒回逆操作、`ai/sheetAgent` 的工具环与提案→op（含图表回合）

> 注意：根 `pnpm test` 是 `pnpm -r` + `passWithNoTests: true`，**某个包一个用例都没有也是全绿**
> （`packages/shared` 目前没有测试文件）。验收时按上面的逐包通过数核对，别只看总退出码；
> 也不要把 `pnpm test` 用管道接 `tail`（退出码会变成 `tail` 的）。

## 联调建议（首次验收脚本）

打开一份测试合同，在任务窗格里依次试：

1. **全局术语替换**：`把所有的"甲方"改成"客户"`
2. **结构化插入**：`在文末加一个 3 行 2 列的表格，表头是"条款/责任方"，下面两行随便填示例`
3. **插入段落**：`在开头加一段正式的前言`
4. **改写语气**：先选一段，再输入 `把这段改成更正式的商务语气`
5. **撤销**：点「撤销」验证恢复

每条都先看 diff 再「接受」。

## 已知边界 & 下一阶段

- **原生模块不构建**：`pnpm-workspace.yaml` 把 `keytar`、`@azure/msal-node-extensions`、
  `@microsoft/m365agentstoolkit-cli` 的构建脚本设为跳过（`allowBuilds: false`），这样 `pnpm install`
  不需要 Xcode 工具链、`pnpm dev` 能直接跑起来。代价是加载项里依赖系统钥匙串鉴权的分支可能不可用；
  真遇到再把这些项改回 `true` 重新 `pnpm install`。

- **regex 替换**：未支持，需要在客户端做字面量探测。
- **重度 PDF 编辑**：下一阶段会走 LibreOffice headless 把 PDF 转 .docx，在 Word 里编辑后再导回；复杂排版评估商业 SDK（Aspose / PDFTron）。
- **Excel**：读入 + 保真回写 + 手动编辑 + 结构化 op 与撤销 + AI 分析/整理/计算 + 原生图表部件已落，见「Excel 表格」与 `docs/excel-support-spec.md`。**还没做的**：app 内的图表预览视图（`.xlsx` 里已是原生图表，只是本 app 不画它，卡在"要不要引图表库"这条表态上）、CLI 写侧（撤销、modlog、原子替换都在桌面主进程那条路上，CLI 只读）、公式重算引擎（按 D4 只写公式与缓存值）、透视表 / 条件格式 / 数据验证的编辑（读入时原样保留，不建模）、Excel / WPS 打开后的**画面效果**（本机没装 Microsoft Excel，WPS 无可视通道，`AGENTS.md` §8，这条标 BLOCKED）。**"文件能不能被办公软件打开"不再靠猜**：`oc sheet check` 机检 OOXML 硬约束 + openpyxl 交叉读 + Numbers 实开，三层证据见上面「怎么确认导出的 `.xlsx` 能在办公软件里打开」。任务窗格加载项不接表格。
- **PPT**：未开工，原计划的 `packages/ppt-client` 还没建。
- **多人/企业部署**：目前是单机自托管；后续加 SSO、审计日志、按租户的 Key 管理。
- **图标**：`apps/task-pane/public/assets/icon-*.png` 目前是占位 1×1 PNG，上架前替换为正式品牌图标。

## 安全提示

- 加载项只申请 `ReadWriteDocument` 权限。
- 发给 LLM 的文档内容默认截断（前 200 段，每段最多 120 字；选区最多 1000 字）。UI 可加显式开关控制是否发送全文。
- 图片走前端 `FileReader → Base64`，不经后端。
- `.env`、`certs/`、`apps/server/data/*.sqlite` 均已在 `.gitignore`。

---

## 四、仓库与协作约定

- 远端：`git@github.com:dcc123456/office-copilot.git`（`main`）。改动完成即提交并推送，不在本地堆积。
- **不入库的东西**：`node_modules`、`dist`、`apps/desktop/release/`（`.dmg` / `.app` 构建产物，单个可达
  数百 MB）、`certs/**`（除 `.gitkeep`）、`.env`、`apps/server/data/*.sqlite`。
  `tmp/`（测试与探针产物、夹具、验证用的 `.xlsx`）**没有 `.gitignore` 条目**，所以它每次都以未跟踪目录出现在
  `git status` 里——按 `AGENTS.md` §7.3 靠人工不暂存，谁也不要 `git add -A`。
- 提交信息：描述部分用中文，格式 `<type>(<scope>): <做了什么 + 为什么>`；一个提交只做一件事。
- AI 代理与人类同样受 **`AGENTS.md`** 约束，尤其是第 2 节「领域不变量」（单位/颜色换算只在
  `packages/docx-io` 与 `packages/sheet-io` 各自的进出口、局部格式不许被拍平、新建 OPC 部件要三处一起登记、
  写入原子且可撤销）。收尾自检要求逐包核对测试结果，
  并注意根 `pnpm test` 对"某包没有测试用例"也是全绿。

## 五、与 Microsoft Word 的功能差距（联网核对 + 本机实测）

研究口径：先按 Word 桌面版（Microsoft 365 / Word 2019 中文界面）的功能面列清单，再逐条在本仓
按**能力**搜实现（不是按名字搜），有实现的给出仓库路径。结论分四档：
**必须**（不做就称不上"Word 替代品"，且能落到现有引擎）/ **建议**（真实文档工作常用）/
**观察**（Word 有但本产品定位不需要）/ **不加**（与"本地改写 + 出可用 .docx"的目标冲突）。

一手来源：Microsoft 支持文档
[在粘贴文本时控制其格式](https://support.microsoft.com/zh-CN/Word/control-the-formatting-when-you-paste-text)
（粘贴三态：保留源格式 / 合并格式 / 只保留文本）、
[格式化您的 Word 文件](https://support.microsoft.com/zh-tw/word/format-your-word-document)、
[Word 中用于"智能剪切和粘贴"设置的选项](https://support.microsoft.com/kb/297680)；
右键菜单随对象变化的行为取自
[Word右键菜单，你用好了么？](https://www.sohu.com/a/158622654_417040)（文字/表格/图片三种落点给
不同命令，剪贴板命令的可用态随选区变化）。**转述不可信**：菜单里的每一条命令都以本仓能真实落地的
编辑器命令为准来对齐，落不了地的（如"新建批注"）就不进菜单，只在下面这张表里记账。

| Word 功能 | 本项目现状 | 判断 | 依据 |
| --- | --- | --- | --- |
| 右键编辑菜单（剪切/复制/粘贴三态、格式刷、字体/段落对话框、超链接、图片环绕） | **本次已补**：`apps/desktop/src/editor/contextMenu.ts` + `components/EditorContextMenu.tsx` | 必须 | Word 最高频的入口；缺它时表格右键菜单也压根没挂载（`TableContextMenu` 曾被 import 未渲染） |
| 段落对话框的完整字段（对齐/行距/缩进之外：大纲级别、段前分页、与下段同页、段中不分页、孤行控制、对齐到网格） | 只有对齐/行距/段前后/左缩进/首行缩进（`editor/extensions/ParagraphFormatting.ts`、`components/ParagraphPopover.tsx`） | 必须 | 公文/论文套版全靠这些开关；OOXML 里就是 `w:keepNext`/`w:pageBreakBefore`/`w:widowControl`，docx-io 已能读写同类 `w:pPr`，成本主要在 schema |
| 悬挂缩进 / 右缩进 | 无（只有 `indent`、`indentFirst`） | 必须 | 参考文献、公文回执行都靠悬挂缩进 |
| 上标 / 下标、字符缩放、文字方向（纵排）、项目符号自定义字符、多级列表编号 | 全无上标下标扩展（`package.json` 未装 `@tiptap/extension-subscript/superscript`）；字距 `letterSpacing` 有属性无入口（本次已在「字体」面板给出） | 建议 | 科技/学术文档常用；StarterKit 已含 `Code`，补两个 mark 即可，但要同步 `EDITOR_EXTENSIONS` + docx 往返用例（AGENTS §5.3） |
| 具名样式与样式库（样式窗格、修改样式、清除格式） | 无具名样式；只有"正文/标题 1-6"下拉 + 整篇模板 `editor/templates.ts` | 必须 | 这是与 Word 最大的一处结构性差距：Word 文档的样式名是跨文档一致排版与生成目录/交叉引用的前提。清除所有格式本次已进右键菜单 |
| 查找替换：按格式替换、通配符、转到 | `components/FindBar.tsx` 只有文本 + 区分大小写 + 全部替换（AI 侧 `replace_text` 支持 regex，见 `packages/word-tools/src/index.ts`） | 建议 | 中文文档日常够用；通配符属于进阶，先不做 |
| 拼写检查 / 语法检查 / 同义词 / 自动更正 | 明确关掉（`App.tsx` 的 `spellcheck:"false"`），只有 AI"纠正全文错别字"命令 | 建议（用 AI 路线，不做红波浪线） | 本项目已有 LLM 通道，中文分词拼写库（如 Hunspell）不适配中文，装原生依赖也违背 AGENTS §6.4 |
| 批注 / 修订（跟踪修订、接受/拒绝、显示标记） | 无批注、无修订；有 diff 视图与修改历史（`ai/diff.ts`、`components/HistoryPanel.tsx`） | 必须 | 协同改稿的实际前提。我们的 modlog + 内容寻址撤销（AGENTS §2.5）已经是"半个修订"，缺的是 OOXML `w:ins/w:del` 的读写与 UI |
| 比较文档 / 并排查看 | 无 | 建议 | AI 改写前后对比已由 diff 视图覆盖，Word 那份"比较两版文档"属于人肉核对场景，优先级低于修订 |
| 超链接：打开 / 编辑 / 删除、书签、交叉引用、题注、目录、索引、引文 | 链接插入/编辑/删除有；右键新增"打开超链接"（经协议白名单，`components/linkPrompt.ts`）；书签/交叉引用/目录域全无（`editor/extensions/Field.ts` 只支持页码/日期等内联域） | 必须（目录 + 交叉引用） | 长文档刚需，且 `w:fldSimple` 通道已存在，补 TOC 域是同一套机制 |
| 表格：排序、文字转表格、表格属性（对齐/缩进/边框底纹/平均分布） | 行列增删、合并拆分、表头、垂直对齐、单元格底色有（`components/TableContextMenu.tsx`）；排序与文字↔表格互转无 | 建议 | 中文材料里"文字转表格"很常用，实现只在 `docx-io` 加一层解析 |
| 图片：裁剪、文字环绕可视化编辑、另存为图片、替代文字 | 环绕属性有（`Drawing.ts` 的 `wrap`）但此前无 UI，右键菜单已补"文字环绕"；裁剪/替代文字无 | 建议 | 环绕属性能写进 OOXML 但看不见效果，先在右键菜单给入口，裁剪要引入画布交互，成本另计 |
| 页眉页脚 / 页码 / 分节符 / 多栏 / 水印 / 页面边框 / 首字下沉 / 行号 | 页眉页脚与页码有（`HeaderFooterEditor.tsx` + `renderHeaderFooter.ts`）；分节符（同一文档多套页眉）、分栏、水印、页面边框、首字下沉、行号无 | 必须（分节、分栏） | 公文与论文的硬需求；分节还牵动 docx-io 的 `w:sectPr` 多节往返 |
| 封面页、主题（字体/配色主题）、文档保护、宏 | 无封面页部件；`uiStore` 的"主题"只是界面皮肤，不是 Word 主题；无文档保护、无宏 | 观察 | 宏与本产品"AI 改写"的定位冲突；文档保护是协同议题 |
| 阅读视图 / Web 版式 / 大纲视图 | 只有分页排版 + 大纲侧栏 | 不加 | 桌面形态的目标是改写与导出，多视图会稀释交付物（可用的 .docx） |
| 自动图文集 / 快速部件 / 自动图文集库 | 无；有模板级美化排版与页眉页脚留底 | 观察 | 与"模板 + AI"两条路线重叠，先用模板覆盖 |
| AI 反向差距：Word 有 Copilot（改写/起草/总结），本项目已具备扩写/缩写/润色/排版 | 已覆盖，本次并入右键菜单 | — | 这是我们相对 Word 的差异化，不属于差距 |

本轮落地范围（一次只推进一个能力，AGENTS §0）：**只做右键编辑菜单**。字体/段落对话框复用了
既有面板（`ParagraphFields` 从 `ParagraphPopover` 抽出，字号候选与颜色表抽到
`editor/fontOptions.ts`），粘贴三态经主进程 `clipboard:read` 读剪贴板，"合并格式"的实现是
"插入后清掉该段的 `textStyle` 标记"——保留结构（加粗/列表/标题）而去掉直接字符格式，语义对齐
Microsoft 对 Merge Formatting 的描述。批注、修订、目录、样式库、分节分栏按上表的"必须"档排队。

### 5.1 二轮补充：哪些是"用户常用"，以及三条实测出来的保真缺陷

研究口径与上表一致，但这一轮多问两句：**这个功能谁在用、什么文档在用**（用来排优先级），
以及**本项目是"读得进来但没入口"还是"两边都没有"**（前者只缺 UI，成本一个面板；后者要动
`docx-io` 的 schema，成本一档文档往返）。缩进四项（左/右/首行/悬挂）是 Word 段落对话框的
标准字段，Microsoft 官方对"右缩进"的描述就是"在段落右侧移动所选间距量"
（[调整缩进和间距](https://support.microsoft.com/zh-cn/office/%E8%B0%83%E6%95%B4%E7%BC%A9%E8%BF%9B%E5%92%8C%E9%97%B4%E8%B7%9D-36239d98-14c5-411e-a880-1ddf25d65cd6)），
悬挂缩进则是参考文献表的标准做法
（[Word 参考文献悬挂缩进](https://www.php.cn/faq/1439924.html)）；
[使用 Word 撰写优秀论文](https://support.microsoft.com/zh-hk/topic/%E4%BD%BF%E7%94%A8-microsoft-word-%E6%92%B0%E5%AF%AB%E5%84%AA%E7%A7%80%E7%9A%84%E8%AB%96%E6%96%87-52cc22e6-7a7a-4fd9-834d-fb6521451af1)
把样式化标题、页码/字数、插入引文与自动生成书目列为论文流程，但**通篇未用"必备"这类强制表述**——
所以"常用"只能按文档场景排序，不能当成官方结论。

**常用度（按场景排，越靠前越该先做）**

| 排名 | Word 功能 | 谁在用 | 本项目属于哪一类 | 判断 |
| --- | --- | --- | --- | --- |
| 1 | 段落对话框全字段：右缩进、悬挂缩进、大纲级别、与下段同页、孤行控制 | 所有排版场景，公文/论文尤重 | 右缩进**两边都没有**；悬挂**读得进来无入口**；大纲级别等同页类**两边都没有** | 必须（且要先补保真，见 5.2） |
| 2 | 具名样式库 + 样式窗格 | 长文档、模板复用、目录的前提 | 两边都没有 | 必须（结构性差距，见上表） |
| 3 | 目录 / 题注 / 交叉引用 / 脚注尾注 | 论文、报告、标书 | 两边都没有（`editor/extensions/Field.ts` 只有内联页码/日期域） | 必须 |
| 4 | 批注 / 修订跟踪 | 协同改稿（甲方、导师、编辑） | 两边都没有；我们的 modlog + 内容寻址撤销是"半个修订" | 必须 |
| 5 | 制表位与前导符（左/右/居中/小数 + 点线） | 目录行、签名行、简历 | 两边都没有；`w:tab` 字符读得进来但没有制表位定义 | 建议 |
| 6 | 上标 / 下标 | 科技、数学、法律条款 | 两边都没有（未装 `@tiptap/extension-subscript/superscript`） | 建议 |
| 7 | 分栏 / 水印 / 页面边框 / 首字下沉 / 行号 | 试卷、海报、红头文件 | 分栏两边都没有；页面边框可由 `rawPpr` 透传但无入口 | 建议（分栏 → 必须） |
| 8 | 自动图文集 / 快速部件 | 公文、合同的重复段落 | 有"模板 + AI"两条路线覆盖同一诉求 | 观察 |
| 9 | 多级列表自定义编号/项目符号字符 | 论文图表编号、法规条款 | 列表有，自定义符号无 | 观察 |

### 5.2 实测保真缺陷（本轮唯一证据来源：docx-io 往返单测探针，跑完即删）

这三条不是"功能缺失"，是**已经在丢用户的格式**，优先级高于上表任何新功能。

| 现象 | 实测输入 → 输出 | 定位 | 结论 |
| --- | --- | --- | --- |
| **段落右缩进在保存后永久消失** | `<w:ind w:left="300" w:right="200"/>` → PM attrs 只有 `{ indent: 15, rawPpr }` → 再写出 `<w:ind w:left="300"/>` | 读侧 `ooxml-pm.ts:437-455` 只解析 `w:left/w:firstLine/w:hanging`；写侧 `paragraphXml` 的 `cleanRaw`（`:980-990`）把整个 `<w:ind>` 从 rawPpr 里抹掉，再由 `synthesizedFormatXml`（`:1027-1040`）按 attrs 重建 | **缺陷**。右缩进没有对应的编辑器属性，被"抹掉重建"这一手直接丢弃，违反 AGENTS §2.3「不许把局部格式拍平」 |
| **悬挂缩进往返正确，但面板会把它清零** | `<w:ind w:left="300" w:hanging="420"/>` → `{ indent: 15, indentFirst: -21 }` → `w:hanging="420"`（数值一致） | 负 `indentFirst` 表示悬挂（`ooxml-pm.ts:445-454`、`:1034-1038`）；但 `components/ParagraphPopover.tsx` 的数值框写死 `Math.max(0, …)`，且「段落」面板只有"首行缩进"一格 | **缺陷 + 缺入口**：用户在段落面板里碰一下首行缩进，悬挂就变 0。要按 Word 那样做成"特殊格式：无/首行/悬挂"下拉 |
| **分隔线（horizontalRule）导出成未定义样式** | `{ type: "horizontalRule" }` → `<w:p><w:pPr><w:pStyle w:val="HorizontalRule"/></w:pPr></w:p>` | `ooxml-pm.ts:933`；本仓不生成这个样式定义，读侧也没有任何分支产出该节点（`w:pBdr` 全程未被解析，仅靠 rawPpr 透传） | **推断为缺陷**：Word 对未定义的 `w:pStyle` 会退回 Normal，线就不见了。**本机无 Word/WPS，标 BLOCKED 未实测**；正确写法是段落底部边框 `<w:pBdr><w:bottom …/></w:pBdr>` |

顺带记两处不一致（不是缺陷，动这块代码时要一并收拾）：`pPrXml`（`ooxml-pm.ts:1255-1260`）和
`synthesizedFormatXml` 是**两份** pt→twips 的 `w:ind` 写法，而前者没有 hanging 分支，负的
`indentFirst` 走那条路会写成 `w:firstLine="-420"`；`w:ind` 的 `w:start`/`w:end` 同义属性完全不读，
仓库 fixture 里没出现过这种写法，属风险项而非已观测缺陷。

