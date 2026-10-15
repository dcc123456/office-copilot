/**
 * `.pptx` 的保真读写器：开包、拿只读模型、按部件改写、回写整包。
 *
 * 保真策略与 `sheet-io` 的 `SheetFile` 一致（spec D2/§2.3）：**没改过的部件一个都不碰**——JSZip 里原样留着，
 * 回写时连二进制部件（`printerSettings/*.bin`、`docProps/thumbnail.jpeg`、嵌入的 `.xlsx`）都逐字带回。
 * 只有显式 `writePartText` 改写过的文本部件才重新序列化，改一张幻灯片只重写那一张的部件。
 */
import { loadPackage, readText, serializePackage, type OpcPackage, type OpcRelationship } from "@office-copilot/opc";
import { readModelFromPackage, type PresentationModel } from "./read.js";

export class PptxFile {
  private constructor(
    private readonly pkg: OpcPackage,
    private readonly presentation: PresentationModel,
    private readonly rewrittenParts: Set<string>,
  ) {}

  /**
   * 打开一份 `.pptx`。
   *
   * @param bytes 原始字节。
   * @returns 读写器；不是合法 zip 时由 `loadPackage` 抛错。
   */
  static async open(bytes: Uint8Array): Promise<PptxFile> {
    const pkg = await loadPackage(bytes);
    const presentation = await readModelFromPackage(pkg);
    return new PptxFile(pkg, presentation, new Set());
  }

  /** 只读的演示文稿模型（幻灯片清单 + 形状树）。 */
  get model(): PresentationModel {
    return this.presentation;
  }

  /** 本次改写过的部件路径（已排序），验收与审计用它核对"到底动了哪几件"。 */
  get changedParts(): string[] {
    return [...this.rewrittenParts].sort();
  }

  /**
   * 读某个文本部件的当前正文。
   *
   * @param partPath 包内路径。
   * @returns 文本；部件不存在返回 undefined。
   */
  async partText(partPath: string): Promise<string | undefined> {
    return readText(this.pkg.zip, partPath);
  }

  /**
   * 改写一个文本部件（幻灯片、母版、内容类型等）。
   *
   * @param partPath 包内路径。
   * @param content 新的 XML 文本。
   */
  writePartText(partPath: string, content: string): void {
    this.pkg.zip.file(partPath, content);
    this.rewrittenParts.add(partPath);
  }

  /**
   * 从包里删掉一个部件（撤销"新建部件"那一档用的——前像为 null 的 op 就是这个语义）。
   *
   * 为什么走 `delete zip.files[name]` 而不是 `zip.remove(name)`：`remove` 是**按前缀递归**删的，
   * 而 `ppt/notesSlides/notesSlide1.xml` 恰好是 `ppt/notesSlides/_rels/…` 之外又一个前缀陷阱——
   * `notesSlide1.xml.rels` 不以 `.xml` 结尾但 `startsWith("…notesSlide1.xml")` 为真，一次 `remove` 会把
   * 它的关系表一起删没。与 `opc` 的 `serializePackage` 清目录项时是同一条理由。
   *
   * @param partPath 包内路径；不存在时静默（撤销的现场核对在调用方，那里已经比过现场与后像）。
   */
  removePart(partPath: string): void {
    if (this.pkg.zip.files[partPath]) delete this.pkg.zip.files[partPath];
    this.rewrittenParts.add(partPath);
  }

  /**
   * 包里的条目清单（含目录项，`opc` 的 `nextPartIndex` 按前缀与后缀数，目录项不会误收）。
   *
   * @returns 已排序的条目名，新建部件避开重名用。
   */
  get partNames(): string[] {
    return this.pkg.entryNames;
  }

  /**
   * 某个源部件已登记的关系。
   *
   * @param sourcePartPath 承载关系表的部件路径，如 `ppt/presentation.xml`。
   * @returns 关系清单（`id`/`type`/`target` 原样），没有该关系表时为空数组。注意这是**打开时**解析的那一份，
   *   本次新追加的关系不在里面——要判"这一页是不是已经有备注页"得读关系表当前文本（`partText`）。
   */
  relationshipsOf(sourcePartPath: string): OpcRelationship[] {
    return this.pkg.relsBySourcePart.get(sourcePartPath) ?? [];
  }

  /** 供 op 生产者拼登记文本用的那份包（`opc` 的 `relsWith` / `contentTypesWith` 都吃它）。 */
  get package(): OpcPackage {
    return this.pkg;
  }

  /**
   * 回写整包字节。
   *
   * @returns `.pptx` 的 Uint8Array；没改过任何部件时，解包后每个文件部件内容与原文件逐字一致，
   *   且经 `serializePackage` 不写目录项（新增幻灯片/图片时不会多出 `ppt/`、`ppt/slides/`）。
   */
  async toBytes(): Promise<Uint8Array> {
    return serializePackage(this.pkg.zip);
  }
}
