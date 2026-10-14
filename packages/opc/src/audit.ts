/**
 * OPC 包的**通用结构审计**：`.docx`/`.xlsx`/`.pptx` 共用的那几条"办公软件一打开就判损坏"的硬约束，只在这一处实现（`AGENTS.md` §3，别长第二套）。
 *
 * 为什么抽在这一层：`sheet-io` 的 `compat-audit.ts` 里有五条判定其实与具体格式无关——
 * 内容类型表覆盖、Override 指向存在、关系 Target 可解析、关系 Id 不重复、部件里引用的 `r:id` 查得到，
 * 它们只碰 `OpcPackage` 的通用字段。再给 `.pptx`（以及未来的 `.docx`）用就是第二次、第三次复制，所以提到 `opc`。
 * 各格式**自己的**元素顺序判定（如 `<worksheet>` 的 `xsd:sequence`）仍留在各自包里，把这里的通用结果拼上去。
 * 同层还有 `newlyIntroducedIssues`：保存前"只拦本次新引入的问题"那道判定，两种格式一模一样，也抽在这里。
 *
 * 判定依据取 ECMA-376 / ISO-IEC 29500-1 的显式约束（OPC 见 Part 2 §10.1.2、§9.2、§13.2.1），不取"Office 大概会喜欢"：
 * - `[Content_Types].xml` 必须覆盖每个部件（Part 2 §10.1.2）；
 * - 每条非 External 关系的 Target 必须解析到真实存在的部件（Part 2 §9.2）；
 * - 同一张关系表里 `@Id` 不许重复（Part 2 §13.2.1）；
 * - 部件里引用的每个 `r:id` 必须在它自己的关系表里存在（Part 1 §13.2.1 同理）。
 */
import { readText, resolveTarget, type OpcPackage } from "./opc.js";

/** 通用（格式无关）的结构问题码。 */
export type OpcAuditCode =
  | "part-not-declared"
  | "declared-part-missing"
  | "relationship-target-missing"
  | "relationship-id-duplicate"
  | "unresolved-relationship-id";

/** 一条通用结构问题；`part` 是包内部件路径，`detail` 供渲染层拼中文句子（`AGENTS.md` §5.7）。 */
export interface OpcAuditIssue {
  code: OpcAuditCode;
  part: string;
  detail: string;
}

/**
 * 扫出部件文本里引用到的关系 Id（`r:id="..."` 这一族属性）。
 *
 * @param text 部件 XML 原文。
 * @returns 去重后的引用 Id 列表；前缀不是 `r` 的 `id=` 属性（如 `c:idx val`）不会被误收。
 */
export function referencedRelIds(text: string): string[] {
  const ids = new Set<string>();
  for (const match of text.matchAll(/\br:id="([^"]+)"/g)) ids.add(match[1] ?? "");
  return [...ids].filter(Boolean);
}

/**
 * 审计一份 OPC 包与格式无关的几类完整性：内容类型覆盖、关系可解析、`r:id` 落点。
 *
 * @param pkg 已 `loadPackage` 的包（调用方各自持有，避免重复解 zip）。
 * @returns 问题数组，空数组代表这几条通用约束都满足。不推测渲染结果，只判确定会被判损坏的项。
 */
export async function auditOpcPackage(pkg: OpcPackage): Promise<OpcAuditIssue[]> {
  const issues: OpcAuditIssue[] = [];
  const entrySet = new Set(pkg.entryNames);
  const partNames = pkg.entryNames.filter((name) => !name.endsWith("/"));

  // §10.1.2：每个部件都要被 Default（按扩展名）或 Override（按部件名）覆盖，漏了办公软件直接不认这个部件。
  for (const part of partNames) {
    if (part === "[Content_Types].xml") continue;
    const extension = part.slice(part.lastIndexOf(".") + 1).toLowerCase();
    if (pkg.overrideParts.has(part) || pkg.contentTypes.has(extension)) continue;
    issues.push({ code: "part-not-declared", part, detail: extension });
  }
  // Override 指向不存在的部件同样是"包内容类型表说谎"，打开时会报错。
  for (const part of pkg.overrideParts.keys()) {
    if (!entrySet.has(part)) issues.push({ code: "declared-part-missing", part, detail: "" });
  }

  // §9.2：每条非 External 关系的 Target 必须解析到真实存在的部件，同一张关系表里 Id 不许重复。
  const idsBySourcePart = new Map<string, Set<string>>();
  for (const [sourcePart, rels] of pkg.relsBySourcePart) {
    const seen = new Set<string>();
    for (const rel of rels) {
      if (seen.has(rel.id)) issues.push({ code: "relationship-id-duplicate", part: sourcePart, detail: rel.id });
      seen.add(rel.id);
      if (rel.targetMode === "External") continue;
      const target = resolveTarget(sourcePart, rel.target);
      if (!entrySet.has(target)) {
        issues.push({ code: "relationship-target-missing", part: sourcePart, detail: `${rel.id} -> ${target}` });
      }
    }
    idsBySourcePart.set(sourcePart, seen);
  }

  for (const part of partNames) {
    if (part === "[Content_Types].xml" || part.endsWith(".rels")) continue;
    const text = await readText(pkg.zip, part);
    if (!text) continue;
    // 部件里引用的每个 `r:id` 都要能在**它自己**的关系表里查到（没有关系表就是全都查不到）。
    const ids = idsBySourcePart.get(part) ?? new Set<string>();
    for (const id of referencedRelIds(text)) {
      if (!ids.has(id)) issues.push({ code: "unresolved-relationship-id", part, detail: id });
    }
  }
  return issues;
}

/**
 * 只留"本次事务新引入的"结构问题——保存前那道闸门的判定，`.xlsx` 与 `.pptx` 同一条规则。
 *
 * @param baseIssues 底包（磁盘上原来那份）审计出来的问题。
 * @param outIssues 准备写出的字节审计出来的问题。
 * @returns 底包里没有的那些。底包本来就不合规（第三方导出器留下的）不该让用户没法保存，
 *   写盘这一侧只对**本次事务**负责（`AGENTS.md` §2.4）。
 */
export function newlyIntroducedIssues<Issue extends OpcAuditIssueFields>(baseIssues: Issue[], outIssues: Issue[]): Issue[] {
  const seen = new Set(baseIssues.map(issueKey));
  return outIssues.filter((issue) => !seen.has(issueKey(issue)));
}

/** 一条问题的身份：码 + 部件 + 细节。同一个包审两遍，同一条问题会得到同一个键。 */
interface OpcAuditIssueFields {
  code: string;
  part: string;
  detail: string;
}

/** 一条问题的身份键。 */
function issueKey(issue: OpcAuditIssueFields): string {
  return `${issue.code}@${issue.part}@${issue.detail}`;
}
