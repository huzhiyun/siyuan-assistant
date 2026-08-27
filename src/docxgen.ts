/*
 * 中间块 → docx Document 生成器（纯模块，docx 库 Node/浏览器双端可用）
 * 对齐 Hermes siyuan-to-docx 导出器的核心规范：
 *  - 标题纯净文本 + Word 原生多级编号（方法 A）
 *  - 首行缩进 0.74cm、宋体正文、黑体标题、10.5pt 小四
 *  - NodeThematicBreak → 分页符
 *  - 表格：全边框、表头灰底加粗居中
 *  - 图片：宽度上限 560px（≈5.8in）、高度上限 720px（7.5in），按比例缩放
 */
import {
    Document,
    Paragraph,
    TextRun,
    HeadingLevel,
    Table,
    TableRow,
    TableCell,
    TableLayoutType,
    WidthType,
    ShadingType,
    BorderStyle,
    PageBreak,
    AlignmentType,
    ImageRun,
    Numbering,
    LevelFormat,
    Packer,
    type IImageOptions,
} from "docx";
import type { SyBlock, InlineRun } from "./blocks";
import { inlineToMd } from "./blocks";

export interface ImageData {
    buffer: ArrayBuffer;
    widthPx: number;
    heightPx: number;
    type: "png" | "jpg" | "gif" | "bmp";
}

export interface DocxGenOptions {
    /** Word 原生多级编号（方法 A），默认 true */
    autoNumberHeadings: boolean;
    /** 正文首行缩进 0.74cm，默认 true */
    firstLineIndent: boolean;
    /** 标题是否加粗黑体，默认 true */
    boldHeadings: boolean;
    /** 图片尺寸上限（px） */
    maxImageWidth: number;
    maxImageHeight: number;
}

export interface ExportContent {
    blocks: SyBlock[];
    /** 块类型为 image 时，src → 已下载的图片数据（缺省则输出占位文本） */
    images: Map<string, ImageData>;
}

const MAX_W = 560;
const MAX_H = 720;
const FIRST_LINE_INDENT = 420; // twips ≈ 0.74cm
const BODY_FONT = { ascii: "Times New Roman", eastAsia: "宋体", hAnsi: "Times New Roman" };
const HEADING_FONT = { ascii: "Arial", eastAsia: "黑体", hAnsi: "Arial" };

/** 解析 markdown 内联标记（**bold** / *italic* / `code` / ^sup^）→ runs */
export function runsFromMd(md: string): InlineRun[] {
    const out: InlineRun[] = [];
    let i = 0;
    const n = md.length;
    while (i < n) {
        if (md.startsWith("**", i)) {
            const end = md.indexOf("**", i + 2);
            if (end > i) {
                out.push({ text: md.slice(i + 2, end), bold: true });
                i = end + 2;
                continue;
            }
        }
        if (md.startsWith("`", i)) {
            const end = md.indexOf("`", i + 1);
            if (end > i) {
                out.push({ text: md.slice(i + 1, end), code: true });
                i = end + 1;
                continue;
            }
        }
        if (md.startsWith("^", i)) {
            const end = md.indexOf("^", i + 1);
            if (end > i) {
                out.push({ text: md.slice(i + 1, end), sup: true });
                i = end + 1;
                continue;
            }
        }
        if (md.startsWith("*", i)) {
            const end = md.indexOf("*", i + 1);
            if (end > i) {
                out.push({ text: md.slice(i + 1, end), italic: true });
                i = end + 1;
                continue;
            }
        }
        out.push({ text: md[i] });
        i++;
    }
    return out;
}

function runToTextRun(run: InlineRun): TextRun {
    return new TextRun({
        text: run.text,
        bold: run.bold,
        italics: run.italic,
        font: run.code ? { ascii: "Consolas", eastAsia: "宋体", hAnsi: "Consolas" } : BODY_FONT,
        superScript: run.sup,
    });
}

function runsToParagraphs(text: string, opts: DocxGenOptions): Paragraph[] {
    const inlineRuns = runsFromMd(text);
    const paras: Paragraph[] = [];
    let current: TextRun[] = [];
    const flush = () => {
        if (current.length > 0) {
            paras.push(
                new Paragraph({
                    children: current,
                    indent: opts.firstLineIndent ? { firstLine: FIRST_LINE_INDENT } : undefined,
                    spacing: { line: 360 }, // 1.5 倍行距
                })
            );
            current = [];
        }
    };
    for (const r of inlineRuns) {
        const parts = r.text.split("\n");
        for (let k = 0; k < parts.length; k++) {
            if (k > 0) {
                flush();
            }
            if (parts[k]) {
                current.push(runToTextRun({ ...r, text: parts[k] }));
            }
        }
    }
    flush();
    return paras;
}

function headingPara(level: number, text: string, opts: DocxGenOptions): Paragraph {
    const headingMap: Record<number, (typeof HeadingLevel)[keyof typeof HeadingLevel]> = {
        1: HeadingLevel.HEADING_1,
        2: HeadingLevel.HEADING_2,
        3: HeadingLevel.HEADING_3,
        4: HeadingLevel.HEADING_4,
        5: HeadingLevel.HEADING_5,
        6: HeadingLevel.HEADING_6,
    };
    const runs: TextRun[] = [];
    for (const run of runsFromMd(text)) {
        runs.push(
            new TextRun({
                text: run.text,
                bold: opts.boldHeadings ? true : run.bold,
                color: "000000",
                font: HEADING_FONT,
            })
        );
    }
    return new Paragraph({
        heading: headingMap[Math.min(Math.max(level, 1), 6)],
        numbering: opts.autoNumberHeadings
            ? { reference: "syass-heading", level: Math.min(Math.max(level, 1), 6) - 1 }
            : undefined,
        children: runs,
        spacing: { before: 240, after: 120, line: 360 },
    });
}

function tablePara(rows: string[][], opts: DocxGenOptions): Table {
    const nCols = Math.max(...rows.map((r) => r.length));
    // A4 text area: 12240 twips page width − 2×1440 twips margins.
    // Fixed DXA widths make macOS Quick Look honor the same layout as Word.
    const tableWidth = 9360;
    const firstColIsNumber = rows.every((r) => /^\s*(?:\d+|[一二三四五六七八九十]+)[.、]?\s*$/.test(r[0] || ""));
    const firstWidth = firstColIsNumber && nCols > 1 ? 720 : 0;
    const otherWidth = Math.floor((tableWidth - firstWidth) / (nCols - (firstWidth ? 1 : 0)));
    const colWidths = Array.from({ length: nCols }, (_, i) => i === 0 && firstWidth ? firstWidth : otherWidth);
    const borders = {
        top: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
        bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
        left: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
        right: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
        insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
        insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "000000" },
    };
    return new Table({
        width: { size: tableWidth, type: WidthType.DXA },
        columnWidths: colWidths,
        layout: TableLayoutType.FIXED,
        borders,
        rows: rows.map((r, ri) => {
            const isHeader = ri === 0;
            return new TableRow({
                children: r.map((cellText, ci) => {
                    const cellRuns = runsFromMd(cellText).map((run) =>
                        new TextRun({
                            text: run.text,
                            bold: isHeader || run.bold,
                            size: 18, // 9pt 表体（投标规范）
                            font: BODY_FONT,
                        })
                    );
                    return new TableCell({
                        width: { size: colWidths[ci] || otherWidth, type: WidthType.DXA },
                        shading: isHeader
                            ? { type: ShadingType.CLEAR, fill: "D9D9D9" }
                            : undefined,
                        margins: { top: 60, bottom: 60, left: 100, right: 100 },
                        children: [
                            new Paragraph({
                                children: cellRuns.length > 0 ? cellRuns : [new TextRun({ text: "", size: 18, font: BODY_FONT })],
                                alignment: isHeader ? AlignmentType.CENTER : AlignmentType.LEFT,
                                spacing: { line: 300 },
                            }),
                        ],
                    });
                }),
            });
        }),
    });
}

function imagePara(block: Extract<SyBlock, { type: "image" }>, data: ImageData | undefined): Paragraph {
    if (!data) {
        return new Paragraph({
            children: [new TextRun({ text: `[图片缺失: ${block.src}]`, italics: true, font: BODY_FONT })],
        });
    }
    // 按比例缩放：宽上限 MAX_W，高上限 MAX_H
    let w = data.widthPx;
    let h = data.heightPx;
    if (w <= 0 || h <= 0) {
        w = MAX_W;
        h = Math.round(MAX_W * 0.75);
    }
    const scale = Math.min(1, MAX_W / w, MAX_H / h);
    w = Math.round(w * scale);
    h = Math.round(h * scale);
    const options: IImageOptions = {
        type: data.type,
        data: data.buffer,
        transformation: { width: w, height: h },
    };
    return new Paragraph({
        children: [new ImageRun(options)],
        alignment: AlignmentType.CENTER,
        spacing: { before: 120, after: 120 },
    });
}

/** 中间块数组 → docx Document */
export function blocksToDocument(content: ExportContent, opts: Partial<DocxGenOptions> = {}): Document {
    const o: DocxGenOptions = {
        autoNumberHeadings: true,
        firstLineIndent: true,
        boldHeadings: true,
        maxImageWidth: MAX_W,
        maxImageHeight: MAX_H,
        ...opts,
    };
    const children: (Paragraph | Table)[] = [];
    for (const b of content.blocks) {
        switch (b.type) {
            case "heading":
                children.push(headingPara(b.level, b.text, o));
                break;
            case "paragraph": {
                const paras = runsToParagraphs(b.text, o);
                children.push(...paras);
                break;
            }
            case "table":
                children.push(tablePara(b.rows, o));
                break;
            case "list":
                for (const item of b.items) {
                    children.push(
                        new Paragraph({
                            children: runsFromMd(item).map(runToTextRun),
                            bullet: { level: 0 },
                            indent: { left: 360 },
                            spacing: { line: 360 },
                        })
                    );
                }
                break;
            case "code":
                for (const line of b.code.split("\n")) {
                    children.push(
                        new Paragraph({
                            children: [
                                new TextRun({ text: line, font: { ascii: "Consolas", eastAsia: "宋体", hAnsi: "Consolas" }, size: 18 }),
                            ],
                            indent: { left: 360 },
                            spacing: { line: 300 },
                        })
                    );
                }
                break;
            case "math":
                children.push(
                    new Paragraph({
                        children: [new TextRun({ text: b.code, font: { ascii: "Cambria Math", eastAsia: "宋体" } })],
                        alignment: AlignmentType.CENTER,
                    })
                );
                break;
            case "quote":
                children.push(
                    new Paragraph({
                        children: [new TextRun({ text: b.text, italics: true, font: BODY_FONT })],
                        indent: { left: 360 },
                        spacing: { line: 360 },
                    })
                );
                break;
            case "thematic":
                children.push(new Paragraph({ children: [new PageBreak()] }));
                break;
            case "image":
                children.push(imagePara(b, content.images.get(b.src)));
                break;
            case "html":
                children.push(
                    new Paragraph({
                        children: [new TextRun({ text: b.text, font: BODY_FONT })],
                        spacing: { line: 360 },
                    })
                );
                break;
        }
    }
    return new Document({
        numbering: o.autoNumberHeadings
            ? {
                  config: [
                      {
                          reference: "syass-heading",
                          levels: Array.from({ length: 6 }, (_, i) => {
                              const text =
                                  i === 0
                                      ? "%1."
                                      : Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join(".") + ".";
                              return {
                                  level: i,
                                  format: LevelFormat.DECIMAL,
                                  text,
                                  alignment: AlignmentType.START,
                                  style: {
                                      run: { color: "000000", font: HEADING_FONT },
                                      paragraph: {
                                          indent: { left: 240 * (i + 1), hanging: 240 },
                                      },
                                  },
                              };
                          }),
                      },
                  ],
              }
            : undefined,
        styles: {
            default: {
                document: {
                    run: { font: BODY_FONT, size: 21 }, // 10.5pt 小四
                    paragraph: {
                        spacing: { line: 360 },
                        indent: o.firstLineIndent ? { firstLine: FIRST_LINE_INDENT } : undefined,
                    },
                },
            },
        },
        sections: [
            {
                properties: {
                    page: {
                        margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }, // 2.54cm
                    },
                },
                children,
            },
        ],
    });
}

/** 浏览器端：Document → Blob 下载 */
export async function documentToBlob(doc: Document): Promise<Blob> {
    return Packer.toBlob(doc);
}

export { inlineToMd };
