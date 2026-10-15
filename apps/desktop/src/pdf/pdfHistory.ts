/**
 * PDF 改动与修改历史（modlog）之间的那一层（`docs/office-parity-spec.md` 阶段 7 的装配那一半）。
 *
 * 存储复用文档那条已有的四个通道（`electron/modlog.ts`），不新建第二份历史（`AGENTS.md` §3 禁止第二套
 * 同类基础设施）；幻灯片那条的先例是 `src/slides/slideHistory.ts`，这一层只管 pdf 专属的**记录形状**：
 * 什么算一条 PDF 改动、怎么挑出最近一条没撤销的。回滚的真身不在这里——记录里的 `applied` 送回主进程，
 * 由 `pdf-io` 的 `revertPdfTextEdits` 对着现场核对旧文字再写回（§2.5）。
 */
import type { PdfAppliedEdit } from "@office-copilot/pdf-io/pure";

/**
 * 一次 PDF 改动在历史里的样子。
 *
 * 只存"改了哪条 run、从什么字改成什么字"，不存整包字节：增量更新本来就只追加几十到几千字节，
 * 而历史上限 60 条，存整份文件会让 `userData/modlog/<文档名>.json` 直接失控。
 */
export interface PdfEditRecord {
  id: string;
  ts: number;
  label: string;
  summary: string;
  /** 这批改动落地的凭据，顺序与当时交给主进程的一致（撤销按同样的顺序换回旧文字）。 */
  applied: PdfAppliedEdit[];
  /** 已经撤销过：一条记录不许撤销两次，否则会把用户后来的编辑再倒回去一遍。 */
  undone?: boolean;
}

/** 记进历史的标签：手工改 PDF 与以后 AI 改 PDF 都留痕，与文档、幻灯片那两条平级。 */
const PDF_EDIT_LABEL = "PDF 改动";

/**
 * 判断一条历史记录是不是 pdf 的。
 *
 * 历史文件里四类记录混在同一个 docKey 下，而 `listModifications` 返回的是 `unknown[]`，
 * 不校验就用会拿到 undefined 字段。pdf 记录的唯一签名是它带 `applied`，且每条都有 `fromText`/`toText`
 * 与四段定位键（文档的 `inverses`、表格与幻灯片的 `ops` 都没有这个形状）。
 *
 * @param value 从主进程读回来的一条记录。
 * @returns `applied` 是非空数组且每条都像 pdf 改动才算。
 */
export function isPdfEditRecord(value: unknown): value is PdfEditRecord {
  if (typeof value !== "object" || value === null) return false;
  const record = value as Partial<PdfEditRecord>;
  if (!Array.isArray(record.applied) || record.applied.length === 0) return false;
  return record.applied.every((entry) => {
    if (typeof entry !== "object" || entry === null) return false;
    const edit = entry as Partial<PdfAppliedEdit>;
    return (
      typeof edit.fromText === "string" &&
      typeof edit.toText === "string" &&
      typeof edit.target === "object" &&
      edit.target !== null &&
      typeof edit.target.pageIndex === "number" &&
      typeof edit.target.streamRef === "string" &&
      typeof edit.target.start === "number"
    );
  });
}

/**
 * 从一堆记录里挑出最近一次还没撤销的 pdf 改动。
 *
 * @param records `listModifications` 的原始返回（顺序不保证，这里自己按 ts 比）。
 * @returns 最新的 pdf 记录；没有可撤销的改动时返回 null。
 */
export function pickLatestPdfEdit(records: unknown[]): PdfEditRecord | null {
  const candidates = records.filter(isPdfEditRecord).filter((record) => !record.undone);
  if (!candidates.length) return null;
  return candidates.reduce((latest, record) => (record.ts > latest.ts ? record : latest), candidates[0]!);
}

/**
 * 把一次 pdf 落地结果收成一条历史记录。
 *
 * @param applied 主进程回包里那批已落地的改动（含改前的文字）。
 * @param summary 一句话说明这次改了什么，历史面板直接显示。
 * @returns 可直接交给 `appendModification` 的记录。
 */
export function pdfEditRecordOf(applied: PdfAppliedEdit[], summary: string): PdfEditRecord {
  return {
    id: crypto.randomUUID(),
    ts: Date.now(),
    label: PDF_EDIT_LABEL,
    summary,
    applied,
  };
}
