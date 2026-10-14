/**
 * `.pptx` 的保真读写器：开包、拿只读模型、按部件改写、回写整包。
 *
 * 保真策略与 `sheet-io` 的 `SheetFile` 一致（spec D2/§2.3）：**没改过的部件一个都不碰**——JSZip 里原样留着，
 * 回写时连二进制部件（`printerSettings/*.bin`、`docProps/thumbnail.jpeg`、嵌入的 `.xlsx`）都逐字带回。
 * 只有显式 `writePartText` 改写过的文本部件才重新序列化，改一张幻灯片只重写那一张的部件。
 */
import { loadPackage, readText, type OpcPackage } from "@office-copilot/opc";
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
   * 回写整包字节。
   *
   * @returns `.pptx` 的 Uint8Array；没改过任何部件时，解包后每个部件内容与原文件逐字一致。
   */
  async toBytes(): Promise<Uint8Array> {
    return this.pkg.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
  }
}
