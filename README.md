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
| `oc ppt info <file.pptx>` | 演示文稿概览：幻灯片数、页面尺寸（EMU 原样）、每张幻灯片的形状号/名字/文字 |
| `oc ppt check <file.pptx>` | 兼容性体检：同样的 OPC 硬约束 + `.pptx` 专属三条（`cSld > spTree`、指向版式的关系、几何 off/ext 属性成对且 `<a:xfrm>` 两块都在），有问题时退出码 1 |
| `oc pdf check <file.pdf>` | PDF 结构体检：按 ISO 32000-1 的显式硬约束查这份文件会不会被阅读器判损坏（头尾标记、`/Root` 与页树 `/Count` 对得上、内容流解得出来），并数出有多少显示文字读不出因而改不了（有问题时退出码 1） |

`oc sheet` 与 `oc ppt` 全部是**只读**的：与桌面 app、AI 工具层跑同一份 `packages/sheet-io` / `packages/pptx-io` 引擎，所以三端对同一个文件的
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
- 公式：以 `=` 开头当公式，`sheet-io` 里单文件递归下降求值。按 D4 只写公式与缓存值、**不实现重算引擎**，白名单外的公式留给 Excel 重算（列画像会显示"待 Excel 重算 N 格"）。改动落在我们手里时，被牵连的公式缓存由引擎当场重算并一起写进 `<v>`（决策 D8：刷新本身就是一条带逆操作的 op，撤销与插删/排序这类几何批不刷），网格显示与产物从此同一个数；算不出的（白名单外、`TODAY()` 这类 volatile、跨表引用）**原样保留**，不抹掉用户没碰过的格子。
- 保真：只建模认得的部分，未建模部件原样回写；富文本 run 的颜色在保存时按 run 指纹复用，不会把局部格式拍平（`AGENTS.md` §2.3 同款红线）。
  整段 inline 写法的富文本（`<c t="inlineStr"><is><r>`，openpyxl 与不少导出器就这么写）我们没有建模，**整格原文原样搬**，改过值才重写。
  这条是逐特性对拍抓出来的：让 openpyxl 自写 10 份"每份只压一类结构"的 `.xlsx`，走桌面那条真实保存路径存一遍，再让 openpyxl 读回来逐字段比（spec §二 P11）。
- 撤销：每次改动是**自带逆操作**的结构化 op，撤销按内容寻址而不是按位置，所以不会牵连你后来的手工编辑；存储与文档侧共用同一份 `userData/modlog/<文档名>.json`。
- AI（⌥⌘L 面板在表格模式下走 `src/ai/sheetAgent.ts`）：五个只读工具 `sheet_describe`（列画像）、`sheet_read_range`（区域显示文本）、`sheet_aggregate`（分组聚合）、`sheet_find`（COUNTIF 风格条件）、`sheet_duplicates`（重复行），加上 `sheet_create_chart` 与 `sheet_submit_proposal`。**所有数值由引擎从现场表算，模型只引用不许心算**——这条写在工具描述里。
- 提案能落地五类：改格（`set_cell`，文本以 `=` 开头当公式）、列宽、插行、删行、按一列重排整行（`sort_rows`，只交区域与键列，行序由引擎算；带内有公式或这些行在区域外还有内容时整条拒绝）。拆合列 / 统一数字格式刻意不开（理由见 spec §七）。
- 图表是**原生 chart part**而不是贴图（D5）：`xl/charts/chartN.xml` + 画图部件 + 两级关系 + `[Content_Types].xml` 两条 Override，五件一起登记；保存时逐张落地，插不进就整批取消。Excel / WPS 里能当图表点开改。
- 三种模式沿用文档侧：半自动先出卡片（把引擎读到的引用与数值排成中文行，并在卡片上**用 Chart.js 画出预览图**，点「插入图表」之前就能看见结果）；全自动排进待写清单；聊天模式只念数字不落地。预览只是给人看的，`.xlsx` 里存的一直是原生 chart part。
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

当前各包通过数（2026-10-15 实测，合计 1 310）：`docx-io` 296、`sheet-io` 245、`desktop` 534、`pptx-io` 78、`pdf-io` 70、`pptx-tools` 20、`opc` 17、`pdf-tools` 13、`sheet-tools` 12、`word-client` 8、`word-tools` 6、`apps/cli` 5、`shared` 4、`server` 2。

- `word-tools`：6 个 zod schema 校验用例
- `word-client`：8 个 Office.js mock 用例（列出段落、替换、段落/表格插入、快照回滚、批量 `applyEdits` 等）
- `docx-io`：296 个用例，覆盖 OOXML 读入 / 写出往返、单位与颜色换算
- `sheet-io`：245 个用例（A1 地址与列号、日期序列号与显示文本、ARGB↔hex、列宽/行高几何、公式子集求值与引用解析、`sharedStrings` 富文本 run 往返、整格 `inlineStr` 富文本原样回写、结构化 op 的逆操作与批内重新锚定、**第五类 op：按一列重排整块行**（比较器四条规则、样式/行高/单行合并区随记录走、带公式与探出区域一律拒绝、撤销就是把两份镜像对调）、分析引擎（列画像 / 聚合 / 条件定位 / 重复行 / 区域截断）、`ChartSpec` 校验、chart part 的读回结构断言、新建 `sharedStrings` 时三处一起登记、`cellXfs` 只追加不重排与「改数字格式不许拍平字体/底色/边框/对齐」、键入的百分比显示回原话、**改动落地后把受牵连的公式缓存重算进 `<v>`**（D8：刷新是一条自带逆操作的 op，撤销批一律不刷；D9 把几何批那条边界收窄成「插删里**引用串被实际改写**的那些格要刷」——区间被删带穿过时收窄成 `SUM(A2:A3)`，旧的 `<v>` 就自相矛盾了）、**外部工作簿引用不再否决本地插删**（D9：`'[1]Sheet1'!$A$2` 的坐标系在别的文件里，整段原样带过；跨表 `!`、数组常量 `{}`、`#` 仍然整条拒绝）、**整列统一数字格式**（D10：新 kind 只改 `numberFormat`，值/公式/样式索引一概不动，格式码只认准入表——内建码、内建表盖不住的小数位、自定义日期写法，其余回 null 并给中文理由；按列展开由引擎逐格判定，文本格与格式已对的格不出 op；写盘只换格子副本，不许把上层载荷的 op 前像一起改掉））
- `sheet-tools`：12 个工具入参校验用例（JSON 语法错、缺字段、多字段、越界都回错误串而不是抛异常）
- `apps/cli`：5 个用例（挑表按名字与 1 基序号、`--metric` 的 `avg`→`average`、分组键逗号切分、`oc ppt` 那条装配面：引擎的问题码念成一句能照着改的中文，`detail` 拼在冒号后）
- `shared`：4 个用例（`parseToolArguments` 那一条判定全仓只有一份：JSON 语法错 / 缺字段 / 多字段 / 越界各回一句中文理由而不是抛异常；PDF 体检的七个问题码念成中文——页级必须带「第 N 页 · 」、整份文件级别的不带、`detail` 能关掉好塞进 toast、计数那句按红黄两档各数一次；幻灯片那一页读不懂的三种说法也只存一份，面板 / `oc ppt info` / 模型收到的 reason 三处念同一句）
- `opc`：17 个用例（格式无关那一层：包与关系读写、通用 XML 序列化往返不许丢读不懂的东西、§2.10 新建部件的关系与内容类型登记原语、`newlyIntroducedIssues` 只对本次事务引入的问题负责，外加 `serializePackage` 那两条——`.docx`/`.xlsx`/`.pptx` 三处 `toBytes` 现在共用这一份序列化，产物里**只有文件部件、没有目录项**，新增部件时文件部件清单与原文件逐条一致）
- `pptx-io`：78 个用例（读模型把占位符记成"自己没写几何"而不烘焙继承值，**继承的那一档按 D3 逐轴从版式算到母版**（`effectiveBoxPx`，缺哪一轴就沿哪条链往上找，找不到就不给数）、EMU↔px 与半磅↔pt 只在 IO 边换算、保真回写逐字回原、**一页读不懂时回三种失败码**（部件缺失 / XML 解析失败 / 没有形状树，不再一律回空数组把"读不动"说成"这页是空的"）、七个带逆操作的 op（文字 / 字体 / 颜色 / 几何 / 叠放次序 / 加文本框 / 删形状）、写进读回单位与颜色形态一致、结构审计与 `oc ppt check`；审计那三条含几何那一类——`<a:off>`/`<a:ext>`（组合形状是 `<a:chOff>`/`<a:chExt>`）的 x、y、cx、cy 不许只写一半，`<a:xfrm>` 也不许少 `off|chOff` / `ext|chExt` 整块，与写侧那条拒绝配成一对。注入一份半个 `<a:off y>` 的真实 deck 时 `oc ppt check` 退出码 1、python-pptx 同一份文件抛 `InvalidXmlError`，而 WPS 自带模板与 python-pptx 默认模板等 6 份真文件各 0 条（不误报））
- `pptx-tools`：20 个工具入参校验用例（七类改动的判别式联合、strict 拒多字段、坏 JSON 与越界回中文理由而不是抛异常、zod 与给模型的 JSON `oneOf` 逐项相等）
- `pdf-io`：70 个用例（读模型与 W1 改写 21 + AI 读面 16 + 内容流的文字定位 15 + 逻辑文本行 10 + 结构体检 8。`.pdf` 读模型 + D8 的 W1 档改写与撤销。钉的四条：① 文字认得**两种串形态**（`(literal) Tj` 与 `<hex> Tj`），且要**递归进 Form XObject**（表格与模板类产物的正文常在那一层）；② `/FlateDecode` 是 **zlib 框不是 raw deflate**——第一版按 raw 解，两份真实文件读出 0 条 run 而全部未压缩夹具照样绿，所以压缩流那条夹具是回归网；CID/Identity-H 按两字节一码解，缺 `/ToUnicode` 的字形**如实给 `null` 而不是猜一个字**；③ 改写只允许**同长度整条替换**，装不回去就整批不写（`length-mismatch`）、子集里没有那个字给 `char-outside-subset`（不许拿别的字顶）、定位键对不上现场模型给 `run-not-found`（不许改错地方），三条失败都不产出字节（§2.4）；落地走 **`7.5.6` 增量更新**：原文件前缀逐字不变、只追加 delta。④ 撤销走**同一份写入实现**（`revertPdfTextEdits`，不平行长第二条路）：成功那次会把 `{target, fromText, toText}` 一起交出去，撤销就是把 `fromText` 换回去，但**先按内容比对**——那条 run 现在的文字若已经不是这批写出去的样子（后来的编辑覆盖过）就判 `text-drifted` 并整批不写，所以撤销永不牵连用户后来的手工编辑（§2.5）；同长度替换让 `start` 与 `streamRef` 在增量存回后仍然稳定，定位键才能直接复用。第五类是 W2 第一刀新加的**文字定位**用例（`content-stream.test.ts`）：扫描器跟住 `cm`/`q`/`Q`/`BT`/`Tm`/`Td`/`TD`/`TL`/`T*`/`'`/`"`，把每条 run 的字形原点与**生效字号**带进读模型，钉的都是容易想当然的那几条——`T*` 是 `0 -Tl Td`（往下走要**减** leading）、`Td` 在**行矩阵**上累加而不是替换当前笔位、`"` 前两个数是词/字间距**不是 leading**、`cm` 要复合进 CTM、`q`/`Q` 把文字状态与图形状态一起存回、`BT` 只清两个矩阵不清字体与 leading、操作数不够时保持上一个值而不拿 0 顶。第六类是 W2 第二刀的**逻辑文本行**（`pdf-model.test.ts`，`pdfTextLinesOf`）：真实中文产物一条 run 只有一个字，所以"改一句话"读到的必须先是一行。断行的判据是从两份真文件的分布里量出来的，不是拍的——同一条**坐标层**（页 + 内容流 + Form）内、`|Δy| ≤ 0.6pt` 算同一行（WPS 每字形一次 `TD`，实测同一条视觉行内相邻字形的 y 最多差 0.119pt，而真是两条行时最近的两行差 15.12pt），x 方向的空隙超过 **4 倍生效字号**就断句（行内实测 p50=1.00em、p90=1.01em，而两栏之间的栏间距实测 9–17em，中间是一条空的带；**固定 pt 数不行**：这份文件里行内最大 21pt、栏间最小 94pt，25pt 恰好蒙得住，可标题「个人简历」每字进 22pt 而正文只有 10.5pt，同一份文件两种尺度混在一起，字号再大一号的标题就会被固定阈值从中间劈开——按 em 归一才与字号无关）。两条不许动的规矩：① **分组只改怎么断行，不改文字本身**——把各行首尾相接必须等于把各 run 首尾相接，一条不许多也不许少（所以段间不插空格，中文会多出看不见的一个字）；② `describePdf` 的概要按行拼，`readPdfPageText` **仍按 run 逐条列**，模型提交改动的把手还是页号 + 页内 run 序号，行号是第三刀之后**并列**的第二把（两套坐标各自都能定位，不是替换掉 run 序号）。实测：WPS 那份 6 页两栏简历 5196 条 run → 206 行（平均 25.2 run/行），第 1 页 611 条 → 29 行，左栏 x 90…142.4 与右栏 289.2…388.5 在同一个 y=615.0 上**没有粘成一行**，最长那几行（65/54/53 段）读出来是完整的句子；Word/`printToPDF` 那份一条 run 一行，14 条 → 14 行，不多合并。**第七类是 W2 第三刀的整行那一档**（`analysis.test.ts` +6、`pdf-file.test.ts` +1）：`pdfLineEditsOf(model, 页号, 行号, 新文字)` 把一行换新拆成该行各 run 的等长替换——按每条 run **原有的字数**依次切片，所以中文那行六个字回六条各一字、Word 那行回一条整行，`streamRef`/`start` 全由引擎填；行内只要有一条改不了（内联内容流 / 认不出原文）就**整行动**并点名卡在哪一类，字数不等、与原文一样、行号与页号越界各回一句中文理由（§2.4 整批语义在代理层之前就守住）。新加的 `cidPerGlyphFixturePdf()` 是这份能力唯一的真形状夹具：四个 CID 写成**四条 `Tj`**（每字后跟一条 `8 0 Td`），因为四个 CID 挤一条 `Tj` 时"一行一条 run"根本没有可切的对象——那条用例读回的是**逐字形的真形状落盘**：前缀字节逐字不变、只追加 delta、行外一条 run 不动。**第八类是 `oc pdf check` 的结构体检**（`check.ts`，8 条）：判定住在引擎里、只出问题码 + 数字，中文文案在装配层（`packages/shared` 那一份，`oc pdf check` 与桌面打开时念的是同一套，§3.2），与 `auditPptxPackage` 同一套纪律。查的都是 ISO 32000-1 的**显式**硬约束而不是风格偏好——§7.5.2 的第一行 `%PDF-1.x` 与每段结尾的 `%%EOF`、trailer 的 `/Root` 与 `/Pages` 页树、§7.7.3.4 的 `/Count` 必须等于 `/Kids` 长度、§7.3.5 声明了 `/Filter` 就必须解得出来。**加密与数字签名刻意没查**（本机没有能产出这两类文件的工具链，写一条测不到的分支不如不写，归 spec D8 的 W4）。坏文件夹具一律**等长改写**真夹具字节：PDF 的交叉引用按字节偏移寻址，改一个字节数就会连带一堆假问题。**`doc.catalog` 那条判据是实测撞出来的**：截断到只剩一节的文件 `PDFDocument.load` 照样成功，`catalog` 却是 `undefined`——不判这一条，整份体检就直接抛出去。内容流解压那几件公用件（`decodeStream`/`streamOf`/`contentKeys`）从 `pdf-file.ts` 抽进 `streams.ts`，读侧与体检侧共用同一份判定（§3.2 同一逻辑出现第二次就必须抽）。**包的依赖方向有一条硬规矩（§5.2）**：带 `node:zlib` 与 `@cantoo/pdf-lib` 的 `pdf-file.ts` 只给主进程，渲染层只准走子入口 `@office-copilot/pdf-io/pure`（`pdf-model.ts` 的类型与判定 + `analysis.ts` 的 AI 读面）——走过主入口的那一次，Vite 把 `require("zlib")` 打进浏览器包，真 app 里 PDF 面板整页白屏，而 `pnpm build` 与 `pnpm test` 全绿。真文件实测见 `docs/office-parity-spec.md` D8 与 §八）
- `pdf-tools`：13 个工具入参校验用例（`pdf_describe` / `pdf_read_page_text` 两条读工具与 `pdf_submit_proposal` 那条写工具的判别式联合、strict 拒多字段、页号 1 基与 runIndex/lineIndex 越界的中文理由、zod 与给模型的 JSON `oneOf` 逐项相等；写面开着**两类**条目 `set_run_text`（逐条）与 `set_line_text`（整行，模型只报页号 + 行号，拆成哪几条 run 归引擎），**模型坐标只有页号 + 页内 run 序号或行号**，引擎的 `streamRef`/`withinForm`/`start` 与"这一行由哪几条 run 拼成"都不出现在工具面里）
- `server`：mock `fetch`，验证 tool-calling 循环产出提案 / 纯聊天状态
- `desktop`：534 个用例。pdf 那条新增：`electron/pdfHandlers` 出入口（钉的是**落地只追加 delta**——原文件前缀逐字节不变、撤销写回旧文字、那条 run 后来又被改过时整条撤不动、任一条改不动整批不写盘）、`pdf/pdfRuns` 草稿守卫（内联流与解不出原文的 run 不给编辑框，字节预算**不在渲染层判**）、`pdf/pdfHistory` 记录形状（同一 docKey 下四类记录互不认错）、`ai/pdfAgent` 的工具环与提案→定位键（**引擎坐标不上送**：模型只交页号与页内 run 序号；页码越界、runIndex 越界、判不可改、与原文相同、字数不等这五种各给一句中文 `notes` 而不是把坏条目送去写）、`pdf/rendererIoEntry` 那条 §5.2 机检（`src/` 下任何文件 import `@office-copilot/pdf-io` 主入口即失败）文档侧原有：`ai/docEdit` 内容寻址 + 逆操作 + 单条撤销、`ai/diff` 中文分词 diff + git 风格行号/折叠上下文、`ai/DiffView` 渲染出的 diff 结构、`ai/diagramSpec` 图形布局、`editor/formatDocument` 离线排版幂等且只改格式、`editor/docBlocks.live` 用 App 真实 schema（`editorExtensions.ts`）跑格式写入（局部加粗/红字/高亮不被整段抹平、列表项改对齐不破坏 `bulletList`、排版后撤销能还原原文档）、`editor/templates` 公文/论文模板套版、`hooks/useAnchoredMenu` 弹层贴边不出屏、`hooks/usePointerDrag` 捕获指针后松手必然解绑（面板宽度的方向与夹取也在这里）、`electron/fsutil` 的原子写（走真文件系统临时目录：`rename` 那一刻目标仍是整份旧内容、失败不留 `.tmp`）。表格侧新增：`sheet/sheetGridModel` 几何/可见窗口/键盘步进/点→地址/改列宽、`sheet/sheetHistory` 认形状并倒回逆操作、落地条数按**提交的那批**算（引擎会额外追加公式缓存刷新 op，D8；拿它的条数当分子会说出"部分写入 2/2"）、`ai/sheetAgent` 的工具环与提案→op（含图表回合、按一列重排整行、引擎拒绝时理由原样进 `notes` 而不是静默丢掉半份提案）、`sheet/chartPreviewModel` 的引用→数据→图型映射与图例/标题显隐规则。幻灯片侧新增：`slides/slideDeckModel` 把草稿摊成待保存改动（多 run 段落一条都不交出、形状没写自己的几何就不交几何改动、删形状时把它没保存的草稿一起丢掉）、`ai/slideAgent` 的工具环与提案→op（页号/形状号/段/run 号越界逐条给中文理由、已在叠放某一端的不交出、同批已删的形状后面那条不交出、几何继承自版式的形状只给单轴时不交出）

> 注意：根 `pnpm test` 是 `pnpm -r` + `passWithNoTests: true`，**某个包一个用例都没有也是全绿**
> （当前唯一没用例的包是 `apps/task-pane`：实测日志里它打印「No test files found, exiting with code 0」）。验收时按上面的逐包通过数核对，别只看总退出码；
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
- **重度 PDF 编辑**：规格已立（`docs/office-parity-spec.md` 决策 D8，分甲/乙两档）。**本日实测到的硬事实**：Keynote 导出的 PDF 里正文是 CID 十六进制串（`<0ebd1dcc12a8> Tj`）+ 子集字体（`AAAAAB+PingFangSC-Regular`），2 个嵌入字体只有 1 个 `/ToUnicode`——所以"改正文"不是字符串替换，要先解决 CID→Unicode 反查与子集字体合并；本机 `qpdf`/`pdftotext`/`mutool`/`gs`/`soffice` **全部没有**，`MuPDF` 是 AGPL-3.0（红线，不搬），可用的是 MIT 的 `pdf-lib` 与 Apache-2.0 的 `pdfjs-dist`。旧计划里"LibreOffice headless 转 .docx 再导回"这条本机走不通（没装）。
- **Excel**：读入 + 保真回写 + 手动编辑 + 结构化 op 与撤销 + AI 分析/整理/计算 + 原生图表部件已落，见「Excel 表格」与 `docs/excel-support-spec.md`。**还没做的**：把图表导出成 PNG 插进 docx/PDF 报告（app 内的预览已落，`23db5af`，用 Chart.js 只画给人看，`.xlsx` 里仍是原生图表）、CLI 写侧（撤销、modlog、原子替换都在桌面主进程那条路上，CLI 只读）、公式重算引擎（按 D4 只写公式与缓存值）、透视表 / 条件格式 / 数据验证的编辑（读入时原样保留，不建模）、Excel / WPS 打开后的**画面效果**（本机没装 Microsoft Excel，WPS 无可视通道，`AGENTS.md` §8，这条标 BLOCKED）。**"文件能不能被办公软件打开"不再靠猜**：`oc sheet check` 机检 OOXML 硬约束 + openpyxl 交叉读 + Numbers 实开，三层证据见上面「怎么确认导出的 `.xlsx` 能在办公软件里打开」。**大表的实测边界**（`294fc1f`，真实 app + CDP 取数）：行/列表头与表体一样只画可见窗口，所以 60 万格那份滚动 20 步 502ms、0 个长任务、方向键 13.6 ms/次；但**只读模式不减数据量**——整表仍要经 IPC 进渲染层，那份 60 万格是 250MB 堆 / 1.4GB RSS、打开 6.2s，再大一档要分片载入才算解决；**改一格是 358ms**（10 万格可编辑那份），因为每次编辑整表载荷往返一次主进程，不是局部写。任务窗格加载项不接表格。
- **PPT**：规格与取证见 `docs/office-parity-spec.md`；已按 D2 红线推进——**阶段 0 抽出 `packages/opc`**（通用 XML + OPC 读 + §2.10 关系/内容类型登记写侧，docx-io 的 `xml.ts` 变薄壳、sheet-io 不再反向依赖 docx-io，全仓回到一套 OPC 实现），**阶段 2 建 `packages/pptx-io`**：读模型（`presentation.xml` 的 `sldIdLst` 顺序 + `sldSz` EMU 尺寸 + 形状树，占位符按 D3 只记 `hasOwnGeometry` 绝不烘焙继承几何）与保真读写器（没改的部件一个不碰，含二进制逐字带回）已落，4 条单测覆盖。**回写保真不再靠自证**：本机 python-pptx 1.0.2 交叉读一份真实 38 部件的 `.pptx` 往返产物，形状文字读回一致、**38 个部件全部字节级一致（0 个部件变化）**。同一判据后来又拿**最富的真文件**跑过一遍：WPS 自带模板 `normal_mac.pptx`（108 个文件部件，其中 **65 个 `ppt/tags/*` 我们一件都没建模**）只打开不编辑存回，**部件清单逐条一致、内容逐字一致（对不上的部件 0 个）**，产物里不再有 OPC 目录项（zip 序列化全仓只有 `@office-copilot/opc` 的 `serializePackage` 一份，见 `AGENTS.md` §3）。**四个结构化编辑 op 已落引擎**：`pptx-io/edit.ts` 有 run 级文字（只动目标 `<a:t>`）、字体（`b`/`sz`，pt×100 只在引擎边换算）、颜色（替换该 run 的 `<a:solidFill>`，`#rrggbb`）与形状几何（移动/缩放，写 `<p:spPr><a:xfrm>` off/ext，px×9525→EMU；组合形状拒、占位符只在明示覆盖时才写不烘焙）四类，都不拍平邻居 run（§2.3），共用定位（`locateShape`→`locateRun`），op 存整张幻灯片前后像做内容寻址撤销（§2.5）。15 条引擎单测覆盖 apply/revert 与越界拒绝，本机 python-pptx 交叉读真实 deck 逐项对上（文字变、粗体邻居保住、字号/颜色/位置读回一致、只 `slide1.xml` 一个部件变、`apply→revert` 后 38 部件全回原样）。**AI 工具面已落**：照 `word-tools`/`sheet-tools` 先例新建 `packages/pptx-tools`（纯 zod，不 import 引擎），`ppt_describe`/`ppt_read_slide_text` 读 + `ppt_submit_proposal` 写（提案条数与引擎那七个带逆操作的 op 一一对应），20 条校验单测覆盖坏 JSON/缺字段/越界/未知 type/不给属性都回中文理由且永不抛异常（§2.8）。**改完的文件能不能被办公软件打开有机检面了**：与格式无关的五条 OPC 判定从表格侧抽进 `@office-copilot/opc`（`auditOpcPackage`，`AGENTS.md` §3.2 不许长第二套；`auditXlsxPackage` 只留 `.xlsx` 专属两条，sheet-io 245 条单测一条未改，证明抽取行为等价），`pptx-io` 的 `auditPptxPackage` 在其上补 `.pptx` 专属三条（`<p:sld>` 必须有 `cSld > spTree`、幻灯片必须有指向 slideLayout 的关系、几何元素 `<a:off>`/`<a:ext>`（组合形状是 `chOff`/`chExt`）的 x、y 与 cx、cy 不许只写一半——§20.1.1.15/§20.1.1.14 那四条 required attribute，`CT_Transform2D` 又要求 `off|chOff` 与 `ext|chExt` 两块都在，少整块同样判损坏；写侧的拒绝由 `edit.ts` 守着，这一条是落盘前的机检面，两边都按本地名比对），审计单测从 7 条长到 12 条；python-pptx 生成的真实 40 部件 deck 与我们改过文字后的写回都审计 0 条。**CLI 端已有 `oc ppt info` / `oc ppt check`**（与桌面同一份引擎，只读；`check` 有问题退出码 1，能当脚本闸门）。**视觉通道是本机第一次拿到的**：Keynote 14.2 能打开我们生成的 `.pptx`、`export … as slide images` 出的 PNG 里图表/表格底色都真渲染；
2026-10-14 这条通道第一次用在**我们自己改过文字后的写回产物**上——Keynote 读出 2 页、`object text` 读回"改过的标题"（没改的那页原样），
导出的 PNG 里第一页居中黑字确实是"改过的标题"、副标题仍是继承版式的灰字，所以占位符继承样式在真办公软件里没有丢。**幻灯片视图与 AI 那一环也已经接上**：桌面 `apps/desktop/src/slides/**` 是列表式编辑（文字/字号/颜色/位置尺寸四类，落地范围由 `slideDeckModel.pendingEdits` 说一不二，多条 run 的段落不许整段覆盖；每个形状行还排着那四个层级按钮 `上移一层 / 下移一层 / 置于顶层 / 置于底层`，这一条**不进草稿、点一次当场落盘**，因为 `up`/`down` 相对的是当前顺序、攒进草稿点第二次会退化成只挪一层，而到头的那两头由 `layerEdgeNote` 直接置灰并在悬停提示里说理由），手工「保存」与 AI 提案卡共用 `App.tsx` 里同一个 `commitSlideEdits` 口和同一份撤销凭据——每次落地把整页前/后像经 `slides/slideHistory.ts` 写进与文档、表格**同一份** `userData/modlog/<文档名>.json`（没有第二套历史栈），撤销从历史里取最新一条没撤销的记录逆序送回，所以杀掉进程重开同一份 `.pptx` 也还能撤（真 app 实测：保存后历史里 1 条 `幻灯片改动`，重启后「撤销上一次保存」照旧可点，撤完磁盘回原字、记录转 `undone`、按钮回禁用）；点「撤销上一次保存」只回已落盘的那批，**不吞用户还没保存的草稿**（草稿存的是绝对值，回滚后 `pendingEdits` 对着新现场重算，恰好等于现场值的那条自己就不交了）；写回前像前逐页核对现场（`revertSlideEdits` 只在现场与该 op 后像逐字一致时才动手），那一页后来又被改过就整条不落并如实报「部分撤销（x/y）」；`apps/desktop/src/ai/slideAgent.ts` 跑 `ppt_describe → ppt_read_slide_text → ppt_submit_proposal` 的工具回合，把模型视角的"第几页第几个形状第几段第几截"换成引擎坐标，对不上现场模型的条目**逐条不交出并给中文理由**而不是整批静默不落（13 条单测，其中一条钉住"该形状已在最顶/最底时别交出"）。第一类结构 op 也已经进引擎与提案面：`move_shape_layer`（上移/下移/置于顶层/置于底层）只重排 `<p:spTree>` 里的形状子元素——那个文档顺序就是叠放次序，依据是 ECMA-376 §19.3.5 的 `CT_GroupShape` 子元素序列加 python-pptx 源码里那句 "first shape … backmost in z-order and the last shape is topmost"；它照样走同一条 op 通道，所以撤销与前像核对一句没改（一页两形状的夹具：改完自己的读侧与 python-pptx 都读到 `3 -> 2`，撤完都回到 `2 -> 3`，三份文件结构审计各 0 条）。**真 app 里那四个按钮实测**（一页三形状，`tmp/probe/slide-layer-ui.mjs`）：最底那行的「下移一层 / 置于底层」是禁用态、最顶那行的「上移一层 / 置于顶层」是禁用态；点最底那行的「上移一层」磁盘当场变 `3 -> 2 -> 4`、列表跟着重排、那个形状的四个按钮同时解禁，中途留的那条未保存文字草稿一直是 `1 处待保存`；再点「置于底层」回到 `2 -> 3 -> 4`；点「撤销上一次保存」只回最后一次点击（磁盘回到 `3 -> 2 -> 4`）而草稿没被吞；最后点「保存」，磁盘上的 `<a:t>` 依次是 `正文乙改过 | 标题甲 | 正文丙`，与新叠放顺序一致。事后 `auditPptxPackage` 0 条，python-pptx 独立读到叠放 `[3, 2, 4]`。提案卡与表格那张收成同一张 `ToolPlanCard`，真 app 里实测过：卡片只算能落的两条、越界那条跟在理由后面，点「应用到幻灯片」当场写盘（脏点不出现），幻灯片面板的「撤销上一次保存」把标题与那一截颜色都退回原样，邻居 run 一字不动。**增删形状这批也已进引擎与视图**：第 1 页标题行有「加文本框」，每个形状行末尾有「删形状」，两条都与那四个层级按钮一样**不进草稿、点一次当场落盘**（新建的骨架逐字照 python-pptx 的 `CT_Shape._textbox_sp_tmpl()`，也就是 PowerPoint 插入文本框自己写的那一段；形状号取**这一页部件内**已有 `@id` 的最大值 +1——本仓文档先前写过"整份文档唯一"，按一手源码纠正是**以每页部件为界**：`oxml/shapes/groupshape.py` 的 `max_shape_id` 扫的就是这一页的 `//@id`，实测真实文件两页的 id 各是 `[1,2,3]` 与 `[1,2]`，本来就跨页重复），删形状只从 `<p:spTree>` 摘那一个元素、不动 `.rels` 与 `[Content_Types].xml`（OPC 只要求被引用的关系存在），所以撤销、前像核对、modlog 三条通道一句没改。**真 app 实测**（`tmp/probe/slide-add-remove-ui.mjs`）：加完磁盘形状序列由 `[2,3]` 变 `[2,3,4]`，新行标签是"文本框"、四个几何框里是 `96/96/480/144` px、段输入框可直接编辑；把它改成 `Q4 目标` 点保存磁盘跟着变；点「删形状」磁盘回到 `[2,3]`，且整页**去掉 XML 声明后与原文件逐字一致**；点「撤销上一次保存」形状连同那句 `Q4 目标` 一起回来；事后 `auditPptxPackage` 0 条、只有 `ppt/slides/slide1.xml` 被改过，python-pptx 独立读到 `id=4 name='TextBox 3' type=TEXT_BOX text='Q4 目标'`、尺寸 1.0/1.0/5.0/1.5 英寸，并且它排在 `shapes` 末尾（最顶）。删掉的形状会连带丢掉它自己那几行未保存的草稿——形状号会被下一个新建的形状回收，留着旧草稿等于把字静默写进新框。**增删形状这两条也已经开进 AI 的提案面**（先前只给模型五类、引擎其实已经有七个 op，那是一处没人说的收窄）：`ppt_submit_proposal` 现在收 `add_text_box`（页号 + 四个几何值 + 文字，全部必填，strict 拒 `shapeId`——新形状号是引擎落地时按该页现有最大值 +1 分配的，模型在这一份提案里引用不到自己刚要建的框）与 `remove_shape`（页号 + 形状号），给模型看的 JSON `oneOf` 与 zod 联合由一条单测钉住必须一致（只改校验忘了改声明，模型就提不了那条）。装配面多一条同批守卫：这一批里已经删掉的形状后面再被指到，那条只进 `notes` 不落地——按顺序执行时它必然定位失败，而失败是**整批不写**的（§2.4），不能让用户白点一次确认。**真 app 实测**（`tmp/probe/slide-ai-structure-ui.mjs`，端点是本机 8899 桩 + 假 Key，验装配面不是模型行为）：四条提案里卡片只算「2 处改动」，两条拦下的各带一句理由（「在同一批里已经被删掉」「没有第 9 页」），新文本框那行写的是「第 1 页 加文本框「谢谢观看」（96,420px 起 480×120px）」；点「应用到幻灯片」磁盘从 `[2:Title 1, 3:TextBox 2]` 变 `[2:Title 1, 4:TextBox 3]`、`<a:t>` 只剩 `标题甲 | 谢谢观看`，脏点不出现；点「撤销上一次保存」后 `slide1.xml` 与原文件**连 XML 声明都逐字一致**（`byte:true`）。**其余尚未开工**：幻灯片部件解析失败现在被折成"这张幻灯片没有可读出的形状"（本批做视图按钮夹具时撞到：少一个 `</a:p>` 的那一页面板说的是"没有形状"，该如实分成两句）、历史面板按格式渲染幻灯片与表格那两类记录（现在它一律按文档记录画，幻灯片记录的逐条改动列表是空的）、图表部件、真端点模型行为探针。
- **PDF**：规格与取证见 `docs/office-parity-spec.md` D8（dcc 裁决"要真·正文改写，不许收窄"）。**第一档 W1 已落 `packages/pdf-io`**：读模型认得内容流里**两种串形态**（`(literal) Tj` 与 `<hex> Tj`）、**递归进 Form XObject**、`/FlateDecode` 按 **zlib 框**解（第一版当 raw deflate 解，两份真实文件读出 0 条 run 而未压缩夹具全绿——压缩夹具因此固化为回归网）、CID/Identity-H 两字节一码，缺 `/ToUnicode` 的字形**如实给 `null` 而不是猜一个字**。改写只允许**同长度整条替换**（PDF 没有回流），装不回去整批不写：`length-mismatch` / `char-outside-subset`（子集里没这个字，不许拿别的字顶上）/ `run-not-found`（定位键对不上现场模型，不许改错地方），失败都不产出字节（§2.4）。落地走 **ISO 32000-1 `7.5.6` 增量更新**（`@cantoo/pdf-lib`，MIT，无原生依赖；坑：只 `context.assign` 会静默丢掉，必须 `snapshot.markRefForSave(ref)`），**原文件前缀逐字不变、只追加 delta**。真文件实测：WPS 简历 6 页 5196 条 run 全解出；`printToPDF` 产物把"段：这是"改成"是这：段"，**前缀 33 706 字节逐字不变 + delta 1 657 字节**，Apple **PDFKit** 读回新文字、**QuickLook 出的 PNG 肉眼核对**只有那四个 CID 字形换位且无豆腐块（pdf 线是本机**有**视觉通道的格式，与 WPS 不同）。**一条实测出来的排队理由**：真实中文产物是一条 run 一个字（每字形一次 `Tj` 做字距），所以"改一句话"必须先拼逻辑文本行 → 归 W2 前置；**撤销已在引擎侧闭环**（`revertPdfTextEdits` 与写入共用同一份实现，按内容比对、对不上就整批不写，所以撤那一条不会把后来的编辑一起冲掉，§2.5；真文件实测：改完再撤，**撤销那一次同样是前缀逐字不变 + 只追加 delta**，读回是原句）。**桌面装配与 modlog 落盘已落（2026-10-15）**：`electron/pdfHandlers.ts` 开 `pdf:read/edit/revert/save-as` 四条通道（preload 白名单同步，渲染层只送定位键与新文字），`src/pdf/{pdfRuns,PdfView,pdfHistory}` 是判定层 + 视图 + 历史形状；打开对话框、拖入、Finder「打开方式」三条入口都认 `.pdf`。真 app 实跑（`dist` + CDP，一份第三方 `printToPDF` 简历）：面板列出 14 段可改文字 → 改成同长度换位 → 保存 toast「追加 402 字节增量，原来的字节没重排」→ **磁盘前缀逐字节不变、新文字在、旧文字一条不剩** → `userData/modlog` 里一条「PDF 改动」→ 点撤销后**旧文字回到文件里**、那条记录标成已撤销。**AI 工具面已落（2026-10-15）**：`packages/pdf-tools` 三条工具（`pdf_describe` 给页与可改条数的画像、`pdf_read_page_text` 逐条列出该页文字与"原位只有 N 字节"、`pdf_submit_proposal` 只收 `set_run_text`），**模型看到的坐标是页号（1 基）+ 页内 run 序号（0 基），引擎的 `streamRef`/`withinForm`/`start` 绝不上送**；`apps/desktop/src/ai/pdfAgent.ts` 跑工具回合（≤8 轮）并把模型坐标换成引擎定位键，翻不动的五种情形（页码越界 / runIndex 越界 / 这条改不了 / 与原文相同 / 字数不等）各留一句中文理由而不是把坏条目送去写。落地仍是 `App.commitPdfEdits` → `pdf:edit` 那一个口（§2.8：AI 只给提案，你点「应用到 PDF」才写文件），所以 modlog 那条「PDF 改动」、撤销凭据、失败文案都只有一份；卡片不新开第三张，PDF 是加进文档/表格/幻灯片共用的那张工具回合卡。pdf 模式下 docx 才有的快捷指令与选区改写入口一并关掉（它们改的是编辑器，对着 PDF 点了会误导），引导语换成 PDF 那条。**真 app 用真服务商实跑（2026-10-15，Key 只在 `tmp/` 的探针 profile 里，绝不入库）**：「把 Jane Q. Doe 的名和姓调个顺序」→ 卡片自己写出"保持 11 个字符长度不变"，落地**前缀逐字不变 + 追加 402 字节**、旧署名一条不剩、modlog 一条「PDF 改动」；「把职位行末尾的 Corp 改成 Comp」→ 目标是页内**第 6 条**（不是第一条，靠位置猜不出来），坐标对、同样只追加 delta；「SUMMARY 改成 32 磅红色」**没有提案卡**，模型直说做不了并点名 `set_run_text` 只能等长整条替换。**观测不到的一条**：想在渲染层数工具回合数（包 `window.officeCopilot.aiToolTurn`）包不上——contextBridge 交出的对象在页面主世界只读、赋值静默无效，这正好是 §5.2 那层保护在起作用；要看回合数与工具名得在主进程侧接计数。**W2 第一刀已落（2026-10-15）**：内容流扫描器跟住文字定位状态（`cm`/`q`/`Q`/`BT`/`Tm`/`Td`/`TD`/`TL`/`T*`/`'`/`"`），每条 run 带出**字形原点**（PDF 用户空间 pt、y 朝上）与**生效字号**，读模型与真文件都核对过——Word/Google 那份简历 14 条 run 的 x 恒为 40、y 从 740 递减到 420，字号 28/14/12/11/10/9 与肉眼一致；**WPS 那份是逐字形发一条 `TD`**（`0.05 0 0 -0.05 253.559 114.89 Tm /FT8 439 Tf` + `<1D2C>Tj 441.594 -0 TD…`），所以原点每个字都不同、同一行的 **y 相同**，而字号必须**乘过 Tm 的 0.05** 才是看到的 21.95/15.95/10.45pt（照抄字面 439 会让下一个读字段的人排错版）。实测还钉住了第二刀的一条判据：那份 6 页两栏简历里有 **19 行**在同一个 y 上 x 断开超过 25pt（最大 171pt，左栏 90…142 / 右栏 289…341），所以**只按 y 分组必然把两栏串成一行**，拼行必须再按 x 断句。**W2 第二刀已落（2026-10-15）**：`pdfTextLinesOf` 把逐字形的 run 拼成逻辑文本行（同一坐标层 + `|Δy| ≤ 0.6pt` + x 空隙 ≤ 4 倍生效字号），`describePdf` 的 AI 概要改成**按行读**——那份 WPS 简历第 1 页 611 条 run 现在上送的是「姓名：丁长长」「出生日期：1994.01.16」这样的整行，不再是 611 个单字；分组不许改动文字本身（行拼接 = run 拼接，一条不许多），`readPdfPageText` 与提交改动的坐标仍是页号 + 页内 run 序号。真文件核对：5196 run → 206 行、两栏在同一个 y 上没粘、0 条"本该同行却被拆开"的单字行。**W2 第三刀已落（2026-10-15）**：逻辑行进 AI 的**编辑**面——`pdf_submit_proposal` 的写面开成两类条目，新增 `set_line_text`（页号 + 页内行号 + 新文字），`pdfLineEditsOf` 按该行每条 run 原有的字数依次切片，模型提交**一条**就把中文那一行的六个字各换一遍，而 `streamRef`/`start` 仍旧只在引擎里出现（§2.1 换算单点：模型侧数不了一次字节，也不许有第二份字数账）。确认卡片头部的条数按**模型提交的那一批**算（`units`），一条整行显示成"1 处改动 · 第 N 页 · 第 K 行改为「…」"，与表格侧 D8"按提交的那批数 op"同一条判据；行内有一条改不了就整行不动、字数不等与行号越界各回一句理由，都在代理层拦下，不把注定失败的条目送进 `pdf:edit`。**WPS 那份简历的实测**（`tmp/pdf-w3/probe.mjs`，跑 `dist`）：概要里读到「姓名：丁长长」是第 1 页第 2 行、由 6 条 run 拼成 → 一次提交 `set_line_text` → 引擎回 6 条等长替换 → 落盘**前缀 334 523 字节逐字不变、只追加 3 344 字节 delta**、行外 605 条 run 一字未动、读回该行是新文字 → 再按内容寻址撤销这一次整行改动，**撤销那一次同样只追加 delta**，读回是原句「姓名：丁长长」。`pdfAgent` 的画像首行现在同时报"可改条数"与"可改行数"，系统提示词直接告诉模型"真实中文产物一个字就是一条 run，改一句话一律报行号"。**这一档的真端点实跑也过了（2026-10-15，`tmp/pdf-w3/app-probe.mjs`，Key 只在探针 profile 里）**：对着**同一份 WPS 简历**（第 1 页 611 条 run）说"把「姓名：丁长长」那行的名字倒过来写"，模型交回来的是**一条 `set_line_text`（`lineIndex=2`）**而不是六条 `set_run_text`——卡片头部「1 处改动 · 第 1 页 · 第 2 行改为「姓名：长长丁」」，落地后**前缀逐字不变、只追加 3 342 字节**、该行仍是 6 条 run、总 run 数 5 196 不变，点「撤销上一次保存」再追加 3 343 字节把原句换回来（截图 `tmp/pdf-w3/app.png`）。**如实记**：模型正文里把行号念成了"第 3 行"（按 1 基数数），但送进工具的 `lineIndex` 是 2、与引擎行序一致，落点以入参为准。**`oc pdf check` 已落（2026-10-15）**：结构体检住在 `packages/pdf-io/src/check.ts`（`auditPdfStructure`，与 `auditPptxPackage` 同一套纪律：引擎只出码与数字，文案在 CLI），CLI 只做装配（`apps/cli/src/pdf.ts`），有问题退出码 1，能当发布前闸门。真文件实测：那份 WPS 简历 **0 条**（不误报），`printToPDF` 产物如实报「1113 条显示文字里 181 条认不出内容（字体没带 /ToUnicode）」，砍掉尾巴的文件报 `%%EOF` 缺失 + 目录查不到。**`print:pdf` 旁路已具备、A10 按现通路关闭（2026-10-15 现跑，`tmp/pdf-w5/`）**：经与桌面 `printToPDF` 相同的选项产一份中英 A4 PDF，用**我们自己的 `readPdfModel`**（那条旁路）读回 1 页 59 run 能定位——即"读页/文本块/字体"对自家导出产物是通的；生效字体名是合规的 `/AAAAAA+PingFangSC-Regular`，当年 `/SourceHanSansCN-Regular-6087` 那种违反 §9.6.4 的命名在当前 Chromium 通路复现不出（违规名才是 MuPDF/PDFKit 拒字体程序、中文渲染成豆腐的根因，故旧记的 A10 判为不再适用）。**同一次实跑翻出的残留**属上游字体数据、非本仓解析器：PingFang SC 的 ToUnicode 反查表把 长/高 等少数字映进了康熙部首码位（CMap 字面写着 `<1CE9> <2ED3>`、`<1E5B> <2FBC>`），字形画面对（走 CID→字形与 ToUnicode 无关），只有文本抽取/复制/无障碍那层拿到部首字，同一份里拉丁整行也读回残缺。**尚未开工**：若要修旁路那两处抽取残缺，得改产 PDF 侧的字体来源或 ToUnicode，属新立一档。**打开时的提示条已落**：面板顶部一条（红 = 可能打不开或丢内容、黄 = 打得开但有一部分改不了），健康的文件一句不出；读模型都建不起来的那份把红色那两句念进 toast。
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
  `tmp/`（测试与探针产物、夹具、验证用的 `.xlsx`）**已于 2026-10-14 经批准进 `.gitignore`**（`591ed89`）。
  口径没变：探针产物只准落在 `tmp/`、自己负责清理，里面可能躺着复制来的探针配置（含真实 Key），
  所以照 `AGENTS.md` §7.3 逐文件点名 `git add`，谁也不要 `git add -A` / `git add .`。
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

