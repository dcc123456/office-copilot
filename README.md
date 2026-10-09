# Office 智能助手

两种形态，同一套 LLM 工具定义：

1. **CLI（推荐，Word 和 WPS 都适用）** —— `oc` 命令直接修改 `.docx` 文件，无需 Office 宿主，输出文件用 Word / WPS / Pages / LibreOffice 都能打开。
2. **Word 任务窗格加载项** —— 在 Word 桌面版/网页版里用自然语言改当前文档（需要 Microsoft Word）。

```
office-copilot/
├─ apps/
│  ├─ cli/           # 命令行工具（Node，跨 Office/WPS）
│  ├─ task-pane/     # Office 加载项前端（React + Vite + Fluent UI）
│  └─ server/        # Node/Express：加载项用的 LLM 网关 + 版本历史
├─ packages/
│  ├─ docx-io/       # 直接读写 .docx（JSZip + OOXML），CLI 用
│  ├─ word-client/   # Office.js Word API 原子操作封装，加载项用
│  ├─ word-tools/    # 暴露给 LLM 的工具定义 + zod 校验（共享）
│  └─ shared/        # 跨端类型与协议
└─ manifest/manifest.xml   # 仅加载项需要
```

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

## 桌面应用（apps/desktop）：AI 直接改文档

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

## 项目结构要点

- **不在前端持有 LLM Key**：前端只通过后端 SSE 调用模型；CORS 只允许任务窗格来源。
- **写入默认要用户确认**：模型产出的是 `Proposal`，点「接受」才执行；执行前自动抓 OOXML 快照。（桌面 AI 面板的「全自动」模式跳过这一步，但每次改动仍记入修改历史，可单独撤销。）
- **会话历史持久化**：`apps/server/data/history.sqlite`（better-sqlite3，WAL 模式）。
- **选区三态处理**：有选区→选区内改；只有光标→光标处插；都没有→作用于 body。
- **regex 替换首版不支持**：Word 原生 `search` 不支持 JS 正则；UI 会给友好提示。后续可做"先 search 到字面量再批量替换"。

## 测试

```bash
pnpm test
```

包含：

- `word-tools`：5 个 zod schema 校验用例
- `word-client`：8 个 Office.js mock 用例（列出段落、替换、段落/表格插入、快照回滚、批量 `applyEdits` 等）
- `server`：mock `fetch`，验证 tool-calling 循环产出提案 / 纯聊天状态
- `desktop`：88 个用例（`ai/docEdit` 内容寻址 + 逆操作 + 单条撤销、`ai/diff` 中文分词 diff + git 风格行号/折叠上下文、`ai/DiffView` 渲染出的 diff 结构、`ai/diagramSpec` 图形布局、`editor/formatDocument` 离线排版幂等且只改格式、`editor/docBlocks.live` 用 App 真实 schema（`editorExtensions.ts`）跑格式写入：局部加粗/红字/高亮不被整段抹平、列表项改对齐不破坏 `bulletList`、排版后撤销能还原原文档、`editor/templates` 公文/论文模板套版、`hooks/useAnchoredMenu` 弹层贴边不出屏、`hooks/usePointerDrag` 捕获指针后松手必然解绑（面板宽度的方向与夹取也在这里））

## 联调建议（首次验收脚本）

打开一份测试合同，在任务窗格里依次试：

1. **全局术语替换**：`把所有的"甲方"改成"客户"`
2. **结构化插入**：`在文末加一个 3 行 2 列的表格，表头是"条款/责任方"，下面两行随便填示例`
3. **插入段落**：`在开头加一段正式的前言`
4. **改写语气**：先选一段，再输入 `把这段改成更正式的商务语气`
5. **撤销**：点「撤销」验证恢复

每条都先看 diff 再「接受」。

## 已知边界 & 下一阶段

- **regex 替换**：未支持，需要在客户端做字面量探测。
- **重度 PDF 编辑**：下一阶段会走 LibreOffice headless 把 PDF 转 .docx，在 Word 里编辑后再导回；复杂排版评估商业 SDK（Aspose / PDFTron）。
- **Excel / PPT**：新增 `packages/excel-client`、`packages/ppt-client`，复用 server 的 tool-calling 循环。
- **多人/企业部署**：目前是单机自托管；后续加 SSO、审计日志、按租户的 Key 管理。
- **图标**：`apps/task-pane/public/assets/icon-*.png` 目前是占位 1×1 PNG，上架前替换为正式品牌图标。

## 安全提示

- 加载项只申请 `ReadWriteDocument` 权限。
- 发给 LLM 的文档内容默认截断（前 200 段，每段最多 120 字；选区最多 1000 字）。UI 可加显式开关控制是否发送全文。
- 图片走前端 `FileReader → Base64`，不经后端。
- `.env`、`certs/`、`apps/server/data/*.sqlite` 均已在 `.gitignore`。
