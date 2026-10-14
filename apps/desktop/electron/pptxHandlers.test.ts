/**
 * 幻灯片那条出入口（`docs/office-parity-spec.md` 阶段 3）。
 *
 * 钉两件事：① 落地的事务语义——改对了只动该动的那个部件、撤销写回前像、定位失败时磁盘一个字节都不动（§2.4/§2.5）；
 * ② 跨界载荷里模型是刷新过的，渲染层拿到的就是磁盘上的样子。
 * "只拦本次新引入的结构问题"那道判定本身已上抽到 `@office-copilot/opc`，用例在 `opc.test.ts`（§3.2）。
 * 夹具用 `@office-copilot/pptx-io/testing` 那份最小 deck，与引擎单测共用。
 */
import { describe, expect, it, vi } from "vitest";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { type PresentationModel } from "@office-copilot/pptx-io";
import { buildDeckBytes } from "@office-copilot/pptx-io/testing";
import { PPTX_EXT, editPptxDoc, openPptxDoc, revertPptxDoc } from "./pptxHandlers.js";

vi.mock("electron", () => ({ dialog: {} }));

/** 从模型里取第一页标题形状（`id=2`）的文字。 */
function titleOf(model: PresentationModel): string {
  const shape = model.slides[0]!.shapes.find((item) => item.shapeId === 2)!;
  return shape.paragraphTexts.join("");
}

/** 重新打开磁盘上那份，问标题现在是什么——用来证明落地的确实是送回来的那个模型。 */
async function titleOnDisk(file: string): Promise<string> {
  const reopened = await openPptxDoc(file);
  if (!reopened.ok) throw new Error("重开失败：" + reopened.message);
  return titleOf(reopened.presentation);
}

/** 把最小 deck 落到一个临时文件里，返回路径（测试自己负责清）。 */
async function writeDeck(name: string): Promise<string> {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "oc-pptx-"));
  const file = path.join(dir, name);
  await fs.writeFile(file, await buildDeckBytes());
  return file;
}

/** 标题那条 run 的定位坐标。 */
const TITLE_RUN = { partPath: "ppt/slides/slide1.xml", shapeId: 2, paraIndex: 0, runIndex: 0 };

describe("幻灯片出入口的扩展名分派", () => {
  it("只认 .pptx，老的二进制 .ppt 进不来", () => {
    expect(PPTX_EXT.test("演示文稿.pptx")).toBe(true);
    expect(PPTX_EXT.test("/tmp/a.PPTX")).toBe(true);
    expect(PPTX_EXT.test("/tmp/a.ppt")).toBe(false);
  });
});

describe("幻灯片出入口的落地与撤销", () => {
  it("打开→改文字：只动那张幻灯片，磁盘上就是新文字，并留一份 .bak", async () => {
    const file = await writeDeck("改字.pptx");
    const opened = await openPptxDoc(file);
    expect(opened.ok && titleOf(opened.presentation)).toBe("标题甲");

    const edited = await editPptxDoc(file, file, [{ kind: "text", target: TITLE_RUN, text: "改过的标题" }]);
    expect(edited.ok).toBe(true);
    if (!edited.ok) return;
    expect(edited.changedParts).toEqual(["ppt/slides/slide1.xml"]);
    // 返回的模型是重读磁盘那份算出来的，渲染层据此刷新，不用再发一次 read。
    expect(titleOf(edited.presentation)).toBe("改过的标题");
    expect(await titleOnDisk(file)).toBe("改过的标题");
    await expect(fs.stat(`${file}.bak`)).resolves.toBeTruthy();
  });

  it("撤销写回前像：op 由主进程现做并随成功结果送回，渲染层原样送回即可回到改前", async () => {
    const file = await writeDeck("撤销.pptx");
    const edited = await editPptxDoc(file, file, [{ kind: "text", target: TITLE_RUN, text: "改过的标题" }]);
    expect(edited.ok && edited.ops).toHaveLength(1);
    const reverted = await revertPptxDoc(file, file, edited.ok ? edited.ops : []);
    expect(reverted.ok && titleOf(reverted.presentation)).toBe("标题甲");
    expect(await titleOnDisk(file)).toBe("标题甲");
  });

  it("同批里后面的改动看得到前面已落地的结果；任一条定位失败就整批不写盘", async () => {
    const file = await writeDeck("两批.pptx");
    const twice = await editPptxDoc(file, file, [
      { kind: "text", target: TITLE_RUN, text: "第一次" },
      { kind: "text", target: TITLE_RUN, text: "第二次" },
    ]);
    expect(twice.ok && titleOf(twice.presentation)).toBe("第二次");
    expect(twice.ok && twice.ops).toHaveLength(2);

    const before = await fs.readFile(file);
    const bad = await editPptxDoc(file, file, [
      { kind: "text", target: TITLE_RUN, text: "不该出现" },
      { kind: "text", target: { ...TITLE_RUN, shapeId: 999 }, text: "定位不到" },
    ]);
    expect(!bad.ok && bad.message).toContain("找不到形状 id=999");
    expect((await fs.readFile(file)).equals(before)).toBe(true);
  });
});
