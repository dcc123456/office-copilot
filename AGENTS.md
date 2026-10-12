# AGENTS.md — Office Copilot 项目规范（对 AI 代理与人类同等生效）

本文件是**约束清单**，不是建议清单。每条规则后面标注它的**强制手段**：
`[机检]` = 有构建 / 测试 / 钩子自动拦截；`[验收]` = 写进本次改动的验收条目，逐项打分；
`[纪律]` = 只能靠执行者自觉，因此每次都必须在收尾自检里显式回答。

冲突时的优先级：**本文件 > 当前任务的 plan / 用户明示指令 > 个人习惯**。
若某条规则与当前计划的既定决策不可调和，**先停下来提出冲突，不要自行打破规则继续写**。

---

## 0. 项目是什么（一句话，防止跑偏）

面向 Word / WPS 文档的**本地文档改写工具**：主形态是 **Electron 桌面 app**（`apps/desktop`，
Tiptap 编辑器 + 直接读写 `.docx`），另有 CLI（`apps/cli`，`oc` / `ocw`）与 Word 任务窗格加载项
（`apps/task-pane`），三者共用 `packages/*` 的工具与 IO 层。它**不是** web 站点、**不是** SaaS 后台、
**不是**脚本集合。任何"先做成网页再说"的实现都是错的。

交付物永远是**一个能被 Word / WPS / Pages / LibreOffice 正常打开的 `.docx`**。凡是"页面看起来对但导出
的文档坏了"的改动，一律视为缺陷，不是可接受的妥协。

执行纪律：**一次只推进一个能力**，先有结论（研究 / 复现 / 判据），再写代码，最后按 §7.4 逐项验收。

---

## 1. 提交规范（Commit）

| #   | 规则 | 强制 |
| --- | --- | --- |
| 1.1 | 提交信息的**描述部分必须用中文**。type / scope 是机器解析的枚举，保持小写英文。 | `[纪律]` |
| 1.2 | 格式：`<type>(<scope>): <中文subject>`，subject 必须说清"做了什么 + 为什么"——是功能就写实现了什么功能，是缺陷就写修复了什么 bug / 什么场景下出问题。 | `[纪律]` + `[验收]` |
| 1.3 | 禁止 `update`、`fix bug`、`临时提交`、`wip`、`顺手改一下` 这类无信息量的 subject。 | `[纪律]` |
| 1.4 | 一个提交只做一件事。功能与格式化改动**必须分开提交**，不许夹带无关重构。 | `[纪律]` |
| 1.5 | subject ≤ 60 个汉字；正文可选，用 `-` 列点说明动机与影响面。 | `[纪律]` |
| 1.6 | 每次修改完成后必须立即提交并推送远端（`origin` = `git@github.com:dcc123456/office-copilot.git`），不允许代码在本地堆积。推送后必须**复核远端 ref** 再声称已推送。 | `[纪律]`（本仓尚无钩子，见 §9） |

**type 取值**：`feat` `fix` `chore` `docs` `refactor` `test` `perf` `build` `ci` `style`
**scope**：包名或功能区 —— `desktop` `cli` `server` `task-pane` `docx-io` `word-client` `word-tools` `shared` `manifest` `docs`，kebab-case。

示例（正确）：

```
feat(docx-io): 导出时把 rgb() 颜色归一化为 hex，避免 Word 判定文档损坏

- 粘贴 HTML 带进来的 rgb(...) 直接写进 w:color 会让 Word 报"文件损坏"
- 转换只在 docx-io 进出口做，编辑器内一律存 #rrggbb
```

```
fix(desktop): 修复拖拽窗口边缘松手后仍在改尺寸

- 指针离开 web contents 时 mouseup 根本不送达，改用 setPointerCapture
```

反例（禁止）：`chore: update`、`fix: 修 bug`、`feat: wip`。

---

## 2. 领域不变量（改动前必须先读，这类错误不可挽回）

违反本节任一条，产出的 `.docx` 会被 Word 判为损坏或静默丢格式，属于**最高优先级红线**。

| #   | 不变量 | 强制 |
| --- | --- | --- |
| 2.1 | **单位换算只允许发生在 `packages/docx-io` 的进出口**。编辑器内字号一律 CSS px（`"14px"`），`.docx` 里是半磅；行距/段前后距/缩进在编辑器是 pt 与倍数，导出换算成 twips（`TWIP_TO_PX = 96/1440`，A4 = 11906×16838 twips @96DPI）。渲染层、CLI、AI 工具层**禁止**出现第二处换算。 | `[纪律]` + `[验收]` |
| 2.2 | **颜色一律 `#rrggbb`**。标记里存 hex；读不懂的颜色值宁可不写，也不要原样塞进 `w:color`。 | `[验收]` |
| 2.3 | **不许把局部格式拍平**。整段改字符样式时只补每段自己缺的属性；一段内加粗/字号只有部分文字带时，必须拒绝并提示用户，禁止静默覆盖红色字、荧光笔、加粗。 | `[机检]` `apps/desktop/src/editor/docBlocks.live.test.ts` |
| 2.4 | **写入必须原子**：先按大纲模拟出预期结果，实际事务落地不一致就整批放弃；文件写入走 `写 path + ".tmp" → fs.rename`（`apps/desktop/electron/fileHandlers.ts`）。禁止直接 `writeFile` 覆盖用户正在打开的文件。 | `[机检]` 桌面引擎单测 + `[纪律]` |
| 2.5 | **每次改动可撤销**：以"改动块"为粒度记入 `userData/modlog/<文档名>.json`（保留最近 60 条），每个 CoreOp 必须发出自己的逆操作；撤销按**内容寻址**而不是按位置，禁止牵连用户后来的手工编辑。无法安全回滚时明确报"部分撤销"，不许猜。 | `[机检]` `apps/desktop/src/ai/docEdit.test.ts` |
| 2.6 | **离线排版必须幂等且只改格式**：`editor/formatDocument` 跑两遍结果一致，且不得改动任何文字内容。 | `[机检]` `formatDocument` / `templates` 单测 |
| 2.7 | **没有模板角色的块 = 不产生任何 op**：套版/美化时，判定不了角色的段落一律原样保留（例如封面居中带里放着标题、分节标题保持原样式）。禁止为了"排得好看"改动用户已有的版式意图。 | `[验收]` |
| 2.8 | **模型只产生提案，不直接写文档**：LLM 输出必须过 zod 校验的工具 schema；任务窗格与桌面"半自动"必须用户点确认才落地（桌面"全自动"是用户显式选择的例外，但仍受 2.5 撤销约束）。 | `[机检]` `packages/word-tools` schema 校验 |
| 2.9 | **发给 LLM 的文档内容默认截断**（前 200 段、每段 ≤120 字、选区 ≤1000 字）。禁止为了方便改成整篇上送。 | `[纪律]` |

---

## 3. 代码复用与简洁（本项目最重要的工程美德）

**以逻辑复用为荣，以堆积代码为耻。**

| #   | 规则 | 强制 |
| --- | --- | --- |
| 3.1 | 写任何新函数之前，先在**同包**、`packages/shared`、`packages/docx-io`、`packages/word-tools` 里按**能力**搜索是否已有实现（`retry`/`parse`/`locate`/`convert`/`diff`/`undo`），不只按名字搜。已有就复用或扩展，不要复制改写一份。 | `[纪律]` + `[验收]` |
| 3.2 | 同一逻辑**出现第二次**就必须抽公共层；第三次还在复制视为缺陷，必须修掉才算完成。 | `[验收]` |
| 3.3 | 新能力优先落到**引擎层**（`packages/*` 或 `apps/desktop/src/{ai,editor}`），界面只做装配。桌面 / CLI / 任务窗格三端共用的逻辑禁止各长一套。 | `[验收]` |
| 3.4 | 优先**扩展现有模块的接口**，而不是新建平行模块。新增包 / 新增顶层目录需要在动手前说明理由。 | `[验收]` |
| 3.5 | 发现无用代码（未被调用的导出、死分支、被替换的旧实现、注释掉的代码块）**及时删除**，不要"先留着"。 | `[纪律]`（无 eslint，见 §9） |
| 3.6 | 代码尽可能简洁：不做超出需求的功能、不为假想的未来做抽象、不加不会发生的异常处理。边界校验只在系统边界做（用户输入、`.docx` 文件内容、LLM 输出）。 | `[纪律]` + `[验收]` |
| 3.7 | 抽象门槛：三行相似代码好过一个过早的抽象；但同一逻辑出现第二次（3.2）就必须抽。冲突时按"是否已经在做同一件事"判断，不按行数判断。 | `[纪律]` |

**禁止引入第二套同类基础设施**（违反即退回）：第二套 docx 读写、第二套单位/颜色换算、第二套 diff 引擎、
第二套撤销栈、第二套 LLM 客户端、第二套设置存储、第二个 SQLite 连接、第二套样式体系。
跨包能力一律经包名（`@office-copilot/*`）的公开入口引用，**禁止深 import 到别人的 `src/`**
（现存唯一例外：`apps/desktop/electron/fileHandlers.ts` → `@office-copilot/docx-io/src/pm-schema.js`，
不要再新增这类口子）。

---

## 4. 注释与命名

| #   | 规则 | 强制 |
| --- | --- | --- |
| 4.1 | **每个新增/修改的函数都要有中文注释**，说明：作用、入参（含单位与边界含义，如"px，非 pt"）、返回值（含失败时的语义）。用 JSDoc 形式。 | `[验收]` + `[纪律]` |
| 4.2 | 注释写**为什么**和非直观约束，不复述代码字面。`// 返回 x` 这类废话注释视为违规。 | `[纪律]` |
| 4.3 | 复杂流程在函数头部写 1–3 行流程说明；一行注释只允许一行，禁止段落式大文档注释块。 | `[纪律]` |
| 4.4 | **变量命名必须有含义**：读名字就知道它装什么、属于哪个阶段。禁止 `data`、`res`、`temp`、`item`、`flag`、`obj2`、`a/b/c` 与无语义缩写；循环内单字母仅允许在 ≤3 行的极短作用域。 | `[纪律]` + `[验收]` |
| 4.5 | 命名风格：`camelCase` 变量、`PascalCase` 类型与类、`kebab-case` 文件名、常量 `UPPER_SNAKE_CASE`；布尔量用 `is/has/should/can` 前缀。测试文件与被测文件同名加 `.test.ts`，同目录。 | `[纪律]`（无 eslint，见 §9） |
| 4.6 | 术语与 `README.md` 保持一致：提案（Proposal）/ 改动块（op）/ 快照（snapshot）/ 撤销栈 / 修改历史（modlog）/ 大纲（paragraphs）。代码里可用英文，但概念要能对上。 | `[纪律]` |

---

## 5. 项目结构与前端

| #   | 规则 | 强制 |
| --- | --- | --- |
| 5.1 | monorepo 分包见 `apps/*` 与 `packages/*`，一个包只做一件事。依赖方向只能是 `apps → packages`，`packages/shared` 谁都能依赖但不许反向；**禁止 app 依赖 app**、禁止 `word-client` 被桌面/CLI 依赖、禁止 `docx-io` 被任务窗格依赖。 | `[纪律]` |
| 5.2 | 渲染层**不碰 Node 能力**：只经 `window.officeCopilot`（`electron/preload.ts` 的 contextBridge 白名单）访问主进程。`contextIsolation: true` + `nodeIntegration: false` 是不可动摇的底线；新增跨界调用必须同时改 preload 的类型与主进程 handler，不许开万能 `invoke` 透传。 | `[机检]` TS 类型 + `[纪律]` |
| 5.3 | 新的编辑器节点 / 标记必须注册进**唯一**的 `EDITOR_EXTENSIONS`（`apps/desktop/src/editor/extensions/editorExtensions.ts`），并同步补一条 `docBlocks.live.test.ts` 真实 schema 用例 + `packages/docx-io` 的导入导出往返用例。只在临时数组里测过的用例不算数。 | `[验收]` |
| 5.4 | 样式：桌面端只用 `src/styles/*.css`（`tokens.css` 存设计变量）这一套；任务窗格只用 Fluent UI（`@fluentui/react-components`）。**禁止**新增 CSS-in-JS、CSS Modules、第四套样式入口或内联 `style={}` 做视觉布局。 | `[纪律]` |
| 5.5 | 需要固定/浮层的交互控件时，注意横向滚动容器会裁掉 `position: absolute` 子元素——这类菜单/弹层用 `position: fixed` 挂到视口坐标（工具栏颜色弹层、段落设置弹层既有先例），并保证不出屏（`hooks/useAnchoredMenu`）。 | `[机检]` `useAnchoredMenu` 单测 |
| 5.6 | 指针拖拽（改宽度、改窗口尺寸）一律走 `hooks/usePointerDrag.ts` 的 `setPointerCapture` + 元素上的 `pointerup/pointercancel`，禁止再用 `window` 上的 `mousemove/mouseup`（松手不送达会留下"下次划过还在改"的缺陷）。热区必须带 `-webkit-app-region: no-drag`。 | `[机检]` `usePointerDrag` 单测 |
| 5.7 | 页面文案：当前产品 UI 为**中文单语**，允许中文字面量；但不得把用户可见文案散进新的工具函数/引擎层（引擎只输出结构化 op 与错误码，文案在渲染层组装）。若未来要加第二语言，必须先立 i18n 方案再动。 | `[纪律]` |
| 5.8 | 目录不允许长期存在孤儿文件与"临时"目录；未完成的实现不进主干，或在验收条目里明确标 `[ ]`。 | `[纪律]` |

---

## 6. 方案先行：新能力必须先研究

**实现项目里没有的能力之前，必须先做外部取证（官方文档 / GitHub / 检索），确认方案可行且是当前主流做法，再写代码。**

| #   | 规则 | 强制 |
| --- | --- | --- |
| 6.1 | 研究结论写进当次改动说明或 `README.md` 对应小节：候选方案、**否决理由**、被验证的一手来源（ECMA-376 章节 / Office.js 文档 / 仓库路径）、以及**本机实测的结论**。 | `[验收]` |
| 6.2 | **文档转述不可信，必须以实测为准**。涉及 OOXML / Office.js / Tiptap / Electron 的 API 形态时，先读它的 `.d.ts` 或产物，或用最小样例（`make-sample.mjs`）验证，再写调用代码。 | `[纪律]` + `[验收]` |
| 6.3 | `.docx` 相关改动必须用真实产物验证：生成 / 打开 → **用 Word 或 WPS 实际打开一遍**，或至少跑一次读回解析看往返一致。只看编辑器 UI 或只看 JSON 断言不足以证明文档没坏。 | `[验收]` |
| 6.4 | 引入新依赖前先回答"现有依赖能否覆盖"；能覆盖就不装。原生编译依赖（keytar、msal-node-extensions 等）在本仓被**故意跳过构建**（见 §8），不许顺手改成允许。 | `[纪律]` |
| 6.5 | spike / 一次性调查代码不进主干仓库；结论进文档，代码删掉或放到工作区之外。 | `[纪律]` |

---

## 7. 测试与收尾自检

**禁止"测试绿了就宣称做完"。** 本仓当前以单元测试 + 真实 schema 的 live 测试为主，没有浏览器/CDP harness；
因此"能看到"的证据由**真实 `.docx` 产物**与桌面 app 实际运行来提供。

| #   | 规则 | 强制 |
| --- | --- | --- |
| 7.1 | 每个涉及写入路径（op、格式、模板、导入导出）的改动，必须补/改对应单测：桌面引擎（`src/ai`、`src/editor`）、`docx-io`、`word-tools`、`word-client`、`server`。测试用本地 fixture，禁止访问真实 LLM 端点（mock `fetch`）。 | `[机检]` `pnpm test` |
| 7.2 | UI 类改动必须 `pnpm dev:desktop` 起真实 app 走一遍主流程（选中→改→撤销→导出 docx→用 Word/WPS 打开），并在收尾时报告看到了什么；跑不起来就如实说明卡在哪，不许用单测冒充可视验收。 | `[纪律]` + `[验收]` |
| 7.3 | **测试产生的文件一律不入库**：临时 `.docx`、截图、探针输出只能写到 `tmp/`（自己负责清理）。仓库里图片只允许两种路径：`apps/task-pane/public/assets/**`（应用素材）与文档明确要求的品牌/说明图。 | `[纪律]`（本仓尚无 pre-commit，见 §9） |
| 7.4 | **收尾自检清单**（每次改完逐条回答，任一不过不得提交、不得声称完成）：<br>① `pnpm build` + `pnpm test` 全绿，贴实际命令与输出（注意 ⑧ 的假绿）<br>② §2 领域不变量逐条自查：单位、颜色、局部格式、原子写、可撤销、幂等<br>③ 复用检查：新逻辑搜过同能力实现吗？出现第二处复制了吗？（§3）<br>④ 死代码检查：有没有未被调用的导出、被替换的旧实现、注释掉的代码块（§3.5）<br>⑤ 新增编辑器能力：`EDITOR_EXTENSIONS` 注册了吗？docx 往返测了吗？（§5.3）<br>⑥ `.docx` 真的用 Word / WPS 打开验证过吗？结论写清楚是"实测"还是"推断"<br>⑦ 提交与推送是否完成（§1.6），commit message 是否符合 §1<br>⑧ 暂存区里有没有测试临时产物、证书、`.env`（§8） | `[验收]` |

`[!]` 是合法状态——环境不具备（例如本机没装 Word、无法验证 Windows 行为）时**如实标 BLOCKED**，
禁止用推测写成"已验证"。诚实的空项比虚假的全绿有用。

---

## 8. 安全与环境事实（已实测，勿重复试错）

**安全红线（不可为赶进度让步）**

1. **API Key 只留在后端 / 本地配置**：任务窗格前端不得持有 LLM Key（只经后端 SSE 调用）；桌面设置落
   `userData/config.json`，`.env` 已在 `.gitignore`。
2. 加载项只申请 `ReadWriteDocument` 权限，不要新增更大权限范围。
3. `certs/**` 除 `.gitkeep` 外一律不入库；自签证书由 `office-addin-dev-certs` 生成到 `~/.office-addin-dev-certs/`，
   不要把它复制进仓库提交。
4. 备份语义不许省：原地写回默认先出 `.bak`，去掉这层保护需要用户明示要求。
5. 图片走前端 `FileReader → Base64`，不经后端上传。

**环境事实**

- 本机是 **macOS（Apple Silicon）+ zsh**。两条由此产生的硬约束：
  ① **Windows / WPS-Windows 运行期验证在本机无法完成**，相关条目一律标 BLOCKED；
  ② **zsh 里未加引号的 glob 会让整条链式命令当场中止**——`grep --include=*.tsx` 报 `no matches found` 时
  `&&` 后面的命令根本不执行。用 Glob/Grep 工具，或把 glob 加引号。
- **本机装有 WPS（`/Applications/wpsoffice.app`），但看不到它的渲染结果**：驱动它要辅助功能权限
  （`osascript` 实测报 `-1728`），整屏截图只拿到壁纸与菜单栏，文档窗口不可捕获；没装 LibreOffice。
  所以"用 Word / WPS 打开验证"这一类条目写 `[!] BLOCKED` 时，原因是"看不到画面"而**不是**"没装"——
  别把"没装"当成结论重复写进文档或提交说明（详见 `docs/word-parity-spec.md` §8.12）。
- **桌面探针跑的是 `dist`，不是 `src`**：`npx electron apps/desktop` 不带 `VITE_DEV_SERVER_URL` 时
  `electron/main.ts` 走 `loadFile(dist/index.html)`，所以改完 src **必须先 `pnpm build`**（或
  `pnpm --filter @office-copilot/desktop exec vite build`）再跑探针，否则"数字一模一样"是旧包造成的假象。
  `curl http://127.0.0.1:5174/src/...` 只能证明 dev server 是新的，证明不了 app 跑的是它（§7.2 的实跑验收常踩这条）。
- **门禁命令必须重定向后再读退出码**：`pnpm test > tmp/t.log 2>&1; echo "EXIT=$?"`。把 `pnpm test` 管道接
  `tail` 时退出码属于 `tail`，会出现套件红过仍拿到 0 的假绿。
- **`pnpm test` 根命令是 `passWithNoTests: true` + `pnpm -r`**：某个包没测试也会"全绿"。**必须逐包核对通过数**
  （`apps/cli`、`packages/shared` 当前没有测试文件），不要只看总退出码就声称测过了。
- 类型检查没有独立脚本，**它藏在 build 里**（`desktop` 的 `build:renderer` 是 `tsc --noEmit && vite build`，
  `task-pane` 的 build 是 `tsc --noEmit`）。改完至少跑 `pnpm build`，只跑 `pnpm test` 不等于类型过。
- `pnpm-workspace.yaml` 的 `allowBuilds` 里 `keytar`、`@azure/msal-node-extensions`、
  `@microsoft/m365agentstoolkit-cli` 是**故意设为 false** 的（跳过原生构建以解阻塞 `pnpm dev`）。
  副作用：Word 加载项的某些鉴权路径可能失效。**这不是待修项，不要在没被要求时"顺手改回 true"**；
  若 sideload 真出问题，先报告这个已知跳过，再决定是否恢复构建。
- 本机 `~/postcss.config.js` 会拉进 Tailwind 让 Vite 构建失败（本仓**不用 Tailwind**）。所有
  Vite/vitest 配置里的 `css.postcss.plugins: []` 就是为了挡这个，新增配置必须照做，不要"清理掉这段奇怪的代码"。
- 端口与形态：任务窗格 `https://127.0.0.1:5173`、后端 `https://127.0.0.1:3001`（`/health`）、桌面 renderer dev `5174`。
  桌面 app 单实例锁按 `userData` 目录算——**不要用 kill 去清用户正在跑的 app**；要并存第二个实例就换 userData 目录与端口。
- 会话历史：`apps/server/data/history.sqlite`（better-sqlite3，WAL）；已在 `.gitignore`，不要提交。
- 远端已配置：`origin` → `git@github.com:dcc123456/office-copilot.git`（普通 SSH 22 端口实测可推）。
  每次推送后 `git status -sb` + `git ls-remote origin main` 双重复核再声称完成。

---

## 9. 机检落地状态（避免误以为规则已被工具强制执行）

本文件里的 `[机检]` 分两类。**已落地**的可以立刻依赖；**待落地**的在实现之前仍然只是 `[纪律]`，
执行者必须靠 §7.4 的自检清单人工核对，不得声称"工具会拦"。

| 规则 | 手段 | 状态 | 落地位置 |
| --- | --- | --- | --- |
| §2.1/2.2 单位与颜色换算单点 | `docx-io` 往返单测 | **已落地（部分）** | `packages/docx-io` 测试，覆盖常见属性；新增属性时必须补对应用例 |
| §2.3 局部格式不被拍平 | live 测试（真实 `EDITOR_EXTENSIONS` schema） | **已落地** | `apps/desktop/src/editor/docBlocks.live.test.ts` |
| §2.4/2.5 op 逆操作与内容寻址撤销 | 单测 | **已落地** | `apps/desktop/src/ai/docEdit.test.ts`、`diff.test.ts` |
| §2.6 离线排版幂等、只改格式 | 单测 | **已落地** | `apps/desktop/src/editor/formatDocument.test.ts`、`templates.test.ts` |
| §2.8 工具入参校验 | zod | **已落地** | `packages/word-tools` |
| §5.2 渲染层不碰 Node | contextBridge 白名单 + TS 类型 | **已落地** | `apps/desktop/electron/preload.ts` |
| §5.5/5.6 弹层贴边、指针捕获 | 单测 | **已落地** | `hooks/useAnchoredMenu.test.ts`、`hooks/usePointerDrag.test.ts` |
| §1 提交格式（中文描述 / type(scope)） | `commit-msg` 钩子 | **待落地** | 尚未创建 `.githooks/commit-msg`，当前靠人工 |
| §7.3 测试产物不入库 | `pre-commit` 路径白名单 | **待落地** | 尚未创建 `.githooks/pre-commit`，当前靠人工 |
| §3.5 死代码 / §4.5 命名规范 | eslint（`no-unused-vars`、naming-convention） | **待落地** | 本仓**没有任何 lint/format 工具**（无 eslint / prettier / biome），当前靠人工 |

> 待落地项在它们变成 `[机检]` 之前，AGENTS.md 是唯一防线，所以 §7.4 的自检清单不许跳过。
