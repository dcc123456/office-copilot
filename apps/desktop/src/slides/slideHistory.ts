/**
 * 幻灯片改动与修改历史（modlog）之间的那一层（`docs/office-parity-spec.md` 阶段 4 的 modlog 落盘）。
 *
 * 存储复用文档那条已有的四个通道（`electron/modlog.ts`），不新建第二份历史（`AGENTS.md` §3 禁止第二套
 * 同类基础设施）。这一层只管幻灯片专属的**记录形状**：什么算一条幻灯片改动、怎么挑出最近一条没撤销的。
 * 回滚的真身不在这里——记录里的 op 送回主进程，由 `pptx-io` 的 `revertSlideEdits` 逐页核对现场再写前像（§2.5）。
 */
import type { SlideEditOp } from "@office-copilot/pptx-io";

/**
 * 一次幻灯片改动在历史里的样子。
 *
 * 只存 op（每张幻灯片的整页前/后像），不存整包快照：历史上限 60 条，而一份 `.pptx` 里的图片与母版
 * 能让快照到几十 MB，`userData/modlog/<文档名>.json` 会直接失控。
 */
export interface SlideEditRecord {
  id: string;
  ts: number;
  label: string;
  summary: string;
  /** 这批改动落地的 op，顺序与当时交给主进程的一致（撤销要逆序送回）。 */
  ops: SlideEditOp[];
  /** 已经撤销过：一条记录不许撤销两次，否则会把用户后来的编辑再倒回去一遍。 */
  undone?: boolean;
}

/** 记进历史的标签：手工改幻灯片与 AI 改幻灯片都留痕，与文档那条"AI 修改"平级。 */
const SLIDE_EDIT_LABEL = "幻灯片改动";

/**
 * 判断一条历史记录是不是幻灯片的。
 *
 * 历史文件里文档、表格、幻灯片三类记录混在同一个 docKey 下，而 `listModifications` 返回的是 `unknown[]`，
 * 不校验就用会拿到 undefined 字段。幻灯片 op 的唯一签名是每条都带整页的 `beforeXml`/`afterXml`
 * （文档的 `CoreOp` 与表格的 `SheetCoreOp` 都没有这两个字段）。
 *
 * @param value 从主进程读回来的一条记录。
 * @returns `ops` 是非空数组且每条都像幻灯片 op 才算。
 */
export function isSlideEditRecord(value: unknown): value is SlideEditRecord {
  if (typeof value !== "object" || value === null) return false;
  const record = value as Partial<SlideEditRecord>;
  if (!Array.isArray(record.ops) || record.ops.length === 0) return false;
  return record.ops.every((op) => {
    if (typeof op !== "object" || op === null) return false;
    const slideOp = op as Partial<SlideEditOp>;
    return typeof slideOp.partPath === "string" && typeof slideOp.beforeXml === "string" && typeof slideOp.afterXml === "string";
  });
}

/**
 * 从一堆记录里挑出最近一次还没撤销的幻灯片改动。
 *
 * @param records `listModifications` 的原始返回（顺序不保证，这里自己按 ts 比）。
 * @returns 最新的幻灯片记录；没有可撤销的改动时返回 null。
 */
export function pickLatestSlideEdit(records: unknown[]): SlideEditRecord | null {
  const candidates = records.filter(isSlideEditRecord).filter((record) => !record.undone);
  if (!candidates.length) return null;
  return candidates.reduce((latest, record) => (record.ts > latest.ts ? record : latest), candidates[0]!);
}

/**
 * 把一次幻灯片落地结果收成一条历史记录。
 *
 * @param ops 主进程回包里那批已落地的 op（含前/后像）。
 * @param summary 一句话说明这次改了什么，历史面板直接显示。
 * @returns 可直接交给 `appendModification` 的记录。
 */
export function slideEditRecordOf(ops: SlideEditOp[], summary: string): SlideEditRecord {
  return {
    id: crypto.randomUUID(),
    ts: Date.now(),
    label: SLIDE_EDIT_LABEL,
    summary,
    ops,
  };
}
