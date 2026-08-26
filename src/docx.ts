/*
 * DOCX → Markdown 解析器（纯模块，无浏览器依赖，可在 Node 中单测）
 * 从 Hermes skill docx-table-import-siyuan / docx-section-to-siyuan-import 的实战经验提炼：
 *  - 图片按 rId 提取（imageN ≠ 第 N 张）
 *  - gridSpan 补空 cell 保证 | 数量一致
 *  - vMerge continue 行跳过（思源普通表格不支持纵向合并）
 *  - 全空行跳过
 */
import JSZip from "jszip";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const V = "urn:schemas-microsoft-com:vml";

export type UploadFn = (blob: Blob, name: string, dir: string) => Promise<string>;

interface ImgRef {
    rId: string;
    alt: string;
    target: string; // docx 内部路径，如 word/media/image1.png
    ext: string;
}

interface DocxResult {
    md: string;
    title: string;
    uploadFailures: number; // 上传失败被跳过的图片数（0 表示全部成功）
}

function escPipe(s: string): string {
    return s.replace(/\|/g, "\\|");
}

function firstChildNS(el: Element, ns: string, local: string): Element | null {
    for (const child of Array.from(el.childNodes)) {
        if (child.nodeType === 1) {
            const c = child as Element;
            if (c.namespaceURI === ns && c.localName === local) {
                return c;
            }
        }
    }
    return null;
}

function findBlipRId(el: Element): string {
    const blips = el.getElementsByTagNameNS(A, "blip");
    for (const blip of Array.from(blips)) {
        const rid = blip.getAttributeNS(R, "embed");
        if (rid) {
            return rid;
        }
    }
    const imgs = el.getElementsByTagNameNS(V, "imagedata");
    for (const img of Array.from(imgs)) {
        const rid = img.getAttributeNS(R, "id");
        if (rid) {
            return rid;
        }
    }
    return "";
}

/** 递归提取 run 级内容：文本 / 换行 / 图片引用 */
function docxRunsToText(
    container: Element,
    addRef: (rId: string, alt: string) => void
): Array<{ type: "text" | "img" | "pageBreak"; markdown: string; rId: string }> {
    const out: Array<{ type: "text" | "img" | "pageBreak"; markdown: string; rId: string }> = [];
    const wrap = (t: string, bold: boolean, italic: boolean, sup: boolean, sub: boolean): string => {
        let s = t.replace(/\u00a0/g, " ");
        if (bold) {
            s = `**${s}**`;
        }
        if (italic) {
            s = `*${s}*`;
        }
        if (sup) {
            s = `^${s}^`;
        }
        if (sub) {
            s = `~${s}~`;
        }
        return s;
    };
    const walk = (el: Element, bold: boolean, italic: boolean, sup: boolean, sub: boolean) => {
        for (const node of Array.from(el.childNodes)) {
            if (node.nodeType === 3) {
                const t = node.textContent || "";
                if (t) {
                    out.push({ type: "text", markdown: wrap(t, bold, italic, sup, sub), rId: "" });
                }
                continue;
            }
            if (node.nodeType !== 1) {
                continue;
            }
            const child = node as Element;
            const local = child.localName || "";
            if (local === "r") {
                let cb = bold;
                let ci = italic;
                let cs = sup;
                let csub = sub;
                const rPr = firstChildNS(child, W, "rPr");
                if (rPr) {
                    cb = bold || !!firstChildNS(rPr, W, "b");
                    ci = italic || !!firstChildNS(rPr, W, "i");
                    const va = firstChildNS(rPr, W, "vertAlign");
                    if (va) {
                        const val = va.getAttributeNS(W, "val");
                        cs = val === "superscript";
                        csub = val === "subscript";
                    }
                }
                walk(child, cb, ci, cs, csub);
            } else if (local === "t") {
                const t = node.textContent || "";
                if (t) {
                    out.push({ type: "text", markdown: wrap(t, bold, italic, sup, sub), rId: "" });
                }
            } else if (local === "br") {
                const breakType = child.getAttributeNS(W, "type") || child.getAttribute("w:type") || "";
                out.push({ type: breakType === "page" ? "pageBreak" : "text", markdown: breakType === "page" ? "" : "\n", rId: "" });
            } else if (local === "tab") {
                out.push({ type: "text", markdown: "  ", rId: "" });
            } else if (local === "drawing" || local === "pict") {
                const blip = findBlipRId(child);
                if (blip) {
                    let alt = "";
                    // docPr 在 wp 命名空间（wordprocessingDrawing），通配查找兼容各种实现
                    const docPrs = child.getElementsByTagNameNS("*", "docPr");
                    if (docPrs.length > 0) {
                        alt = docPrs[0].getAttribute("descr") || docPrs[0].getAttribute("name") || "";
                    }
                    addRef(blip, alt);
                    out.push({ type: "img", markdown: "", rId: blip });
                }
            } else if (local === "hyperlink") {
                walk(child, bold, italic, sup, sub);
            } else if (local === "bookmarkStart" || local === "bookmarkEnd" || local === "proofErr" || local === "instrText" || local === "fldChar") {
                // 忽略书签、校对标记和 Word 域指令；域结果中的可见 w:t 仍会正常保留
            } else {
                walk(child, bold, italic, sup, sub);
            }
        }
    };
    walk(container, false, false, false, false);
    return out;
}

/** 段落 → markdown 块（含标题、列表、硬分页、内联格式与图片占位符） */
function docxParaToMd(
    p: Element,
    addRef: (rId: string, alt: string) => void,
    styleLevels: Map<string, number>
): { md: string } {
    let level = 0;
    let listLevel: number | null = null;
    const pPr = firstChildNS(p, W, "pPr");
    if (pPr) {
        const pStyle = firstChildNS(pPr, W, "pStyle");
        if (pStyle) {
            const style = pStyle.getAttributeNS(W, "val") || "";
            level = styleLevels.get(style) || 0;
            if (!level) {
                const m = /^(?:Heading|heading|标题)\s*(\d)$/.exec(style);
                if (m) {
                    level = parseInt(m[1], 10);
                } else if (/^Title$/i.test(style)) {
                    level = 1;
                }
            }
        }
        const numPr = firstChildNS(pPr, W, "numPr");
        const ilvl = numPr && firstChildNS(numPr, W, "ilvl");
        if (ilvl) {
            listLevel = parseInt(ilvl.getAttributeNS(W, "val") || "0", 10) || 0;
        }
    }
    const runs = docxRunsToText(p, addRef);
    const lines: string[] = [];
    const imgs: string[] = [];
    let text = "";
    const flushText = () => {
        const content = text.trim();
        text = "";
        if (!content) {
            return;
        }
        if (level > 0) {
            const headText = content.replace(/\*\*/g, "");
            if (headText) {
                lines.push(`${"#".repeat(level)} ${headText}`);
            }
        } else if (listLevel !== null) {
            lines.push(`${"  ".repeat(listLevel)}1. ${content}`);
        } else {
            lines.push(content);
        }
    };
    for (const run of runs) {
        if (run.type === "img") {
            imgs.push(`%%SYASS_IMG:${run.rId}%%`);
        } else if (run.type === "pageBreak") {
            flushText();
            lines.push("---");
        } else {
            text += run.markdown;
        }
    }
    flushText();
    lines.push(...imgs);
    return { md: lines.join("\n\n") };
}

/** 表格 → markdown 表格（gridSpan 补空、vMerge continue 跳过、全空行跳过） */
function docxTableToMd(tbl: Element, addRef: (rId: string, alt: string) => void): string {
    const rows: string[][] = [];
    const imgs: string[] = [];
    for (const tr of Array.from(tbl.getElementsByTagNameNS(W, "tr"))) {
        const cells: string[] = [];
        for (const tc of Array.from(tr.childNodes)) {
            if (tc.nodeType !== 1 || (tc as Element).localName !== "tc") {
                continue;
            }
            const tcEl = tc as Element;
            let gridSpan = 1;
            let vMergeContinue = false;
            const tcPr = firstChildNS(tcEl, W, "tcPr");
            if (tcPr) {
                const gs = firstChildNS(tcPr, W, "gridSpan");
                if (gs) {
                    const v = parseInt(gs.getAttributeNS(W, "val") || "1", 10);
                    if (v > 0) {
                        gridSpan = v;
                    }
                }
                const vm = firstChildNS(tcPr, W, "vMerge");
                if (vm) {
                    const val = vm.getAttributeNS(W, "val");
                    if (val !== "restart") {
                        vMergeContinue = true;
                    }
                }
            }
            if (vMergeContinue) {
                cells.push("");
                for (let i = 1; i < gridSpan; i++) {
                    cells.push("");
                }
                continue;
            }
            const runs = docxRunsToText(tcEl, addRef);
            let cellText = "";
            for (const run of runs) {
                if (run.type === "img") {
                    imgs.push(`%%SYASS_IMG:${run.rId}%%`);
                } else {
                    cellText += run.markdown;
                }
            }
            cellText = escPipe(cellText.replace(/\s*\n+\s*/g, " ").trim());
            cells.push(cellText);
            for (let i = 1; i < gridSpan; i++) {
                cells.push("");
            }
        }
        if (cells.some((c) => c.trim() !== "")) {
            rows.push(cells);
        }
    }
    if (rows.length === 0) {
        return "";
    }
    const nCols = Math.max(...rows.map((r) => r.length));
    const lines = rows.map((r) => {
        while (r.length < nCols) {
            r.push("");
        }
        return `|${r.join("|")}|`;
    });
    lines.splice(1, 0, `|${Array(nCols).fill("---").join("|")}|`);
    if (imgs.length) {
        lines.push("", ...imgs);
    }
    return lines.join("\n");
}

/** styles.xml 中 paragraph styleId → Word outlineLvl（0-based）映射到 Markdown H1-H6。 */
function parseStyleLevels(stylesXml: string, parser: DOMParser): Map<string, number> {
    const levels = new Map<string, number>();
    const stylesDoc = parser.parseFromString(stylesXml, "application/xml");
    for (const style of Array.from(stylesDoc.getElementsByTagNameNS(W, "style"))) {
        if (style.getAttributeNS(W, "type") !== "paragraph") {
            continue;
        }
        const styleId = style.getAttributeNS(W, "styleId");
        const pPr = firstChildNS(style, W, "pPr");
        const outline = pPr && firstChildNS(pPr, W, "outlineLvl");
        const rawLevel = outline && outline.getAttributeNS(W, "val");
        if (styleId && rawLevel !== null && rawLevel !== undefined) {
            const level = parseInt(rawLevel, 10) + 1;
            if (level >= 1 && level <= 6) {
                levels.set(styleId, level);
            }
        }
    }
    return levels;
}

/**
 * 解析 docx 文件内容
 * @param ab        docx 文件的 ArrayBuffer
 * @param baseName  用于 assets 子目录命名（如文档名去扩展名）
 * @param upload    图片上传回调，返回思源 assets 相对路径（返回 "" 表示该图失败，跳过不阻塞整体）
 * @param opts      可选：{ concurrency: 并发上传数（默认 4）; onProgress: (done, total) => void }
 */
export async function parseDocx(
    ab: ArrayBuffer,
    baseName: string,
    upload: UploadFn,
    opts?: { concurrency?: number; onProgress?: (done: number, total: number) => void }
): Promise<DocxResult> {
    const zip = await JSZip.loadAsync(ab);
    const docXmlFile = zip.file("word/document.xml");
    if (!docXmlFile) {
        throw new Error("不是有效的 docx（缺少 word/document.xml）");
    }
    const docXml = await docXmlFile.async("string");
    const parser = new DOMParser();
    const doc = parser.parseFromString(docXml, "application/xml");
    const stylesFile = zip.file("word/styles.xml");
    const styleLevels = stylesFile
        ? parseStyleLevels(await stylesFile.async("string"), parser)
        : new Map<string, number>();

    const relsMap = new Map<string, string>();
    const relsFile = zip.file("word/_rels/document.xml.rels");
    if (relsFile) {
        const relsXml = await relsFile.async("string");
        const relsDoc = parser.parseFromString(relsXml, "application/xml");
        for (const rel of Array.from(relsDoc.getElementsByTagName("Relationship"))) {
            const id = rel.getAttribute("Id");
            const target = rel.getAttribute("Target");
            if (id && target) {
                relsMap.set(id, target);
            }
        }
    }

    const mediaRefs: ImgRef[] = [];
    const seenRefs = new Set<string>();
    const addRef = (rId: string, alt: string) => {
        if (!rId || seenRefs.has(rId)) {
            return;
        }
        const target = relsMap.get(rId) || "";
        if (!/^media\//.test(target)) {
            return;
        }
        seenRefs.add(rId);
        const extMatch = /\.(\w+)$/.exec(target);
        const ext = extMatch ? extMatch[1] : "png";
        mediaRefs.push({ rId, alt: alt || `image_${mediaRefs.length + 1}`, target: `word/${target}`, ext });
    };

    const body = doc.getElementsByTagNameNS(W, "body")[0];
    if (!body) {
        throw new Error("docx 正文为空");
    }
    const blocks: string[] = [];
    let firstH1 = "";
    for (const child of Array.from(body.childNodes)) {
        if (child.nodeType !== 1) {
            continue;
        }
        const el = child as Element;
        const local = el.localName || "";
        if (local === "p") {
            const res = docxParaToMd(el, addRef, styleLevels);
            if (res.md.trim()) {
                blocks.push(res.md);
                if (!firstH1 && /^# /.test(res.md)) {
                    firstH1 = res.md.replace(/^#\s*/, "").trim();
                }
            }
        } else if (local === "tbl") {
            const res = docxTableToMd(el, addRef);
            if (res.trim()) {
                blocks.push(res);
            }
        }
    }

    const dir = `assets/${baseName}`;
    const urlMap = new Map<string, string>();
    // ⚠️ 并发受限上传：全量 Promise.all 会瞬间打出几百个请求（内核/浏览器都扛不住，且单个请求
    // 挂起会导致 Promise.all 永不 resolve → 按钮永久置灰）。改为固定并发 worker + 单图失败跳过。
    const concurrency = opts?.concurrency && opts.concurrency > 0 ? opts.concurrency : 4;
    const total = mediaRefs.length;
    const onProgress = opts?.onProgress || (() => {});
    let done = 0;
    let cursor = 0;
    let uploadFailures = 0;
    const workers = Array.from({length: Math.min(concurrency, mediaRefs.length)}, async () => {
        while (cursor < mediaRefs.length) {
            const ref = mediaRefs[cursor++];
            try {
                const zf = zip.file(ref.target);
                if (!zf) {
                    urlMap.set(ref.rId, "");
                } else {
                    const blob = await zf.async("blob");
                    const url = await upload(blob, `image_${ref.rId.replace(/^rId/i, "")}.${ref.ext}`, dir);
                    if (url) {
                        urlMap.set(ref.rId, url);
                    } else {
                        uploadFailures++;
                        urlMap.set(ref.rId, "");
                    }
                }
            } catch (e) {
                uploadFailures++;
                urlMap.set(ref.rId, ""); // 单图失败不阻塞整体，md 里该图占位符会被替换为空
            }
            done++;
            onProgress(done, total);
        }
    });
    await Promise.all(workers);

    let md = blocks.join("\n\n");
    md = md.replace(/%%SYASS_IMG:([^%]+)%%/g, (_: string, rId: string) => {
        const url = urlMap.get(rId);
        if (!url) {
            return "";
        }
        const ref = mediaRefs.find((m) => m.rId === rId);
        return `![${ref ? ref.alt : ""}](${url})`;
    });

    return { md, title: firstH1, uploadFailures };
}
