/*
 * 思源 getDoc HTML → 中间块表示（纯模块，无浏览器依赖，Node 可单测）
 * 解析逻辑从 siyuan-assistant v0.1.0 的节级导出提炼，
 * markdown 渲染（节级导出）与 docx 渲染（导出 Word）共用。
 *
 * 依赖全局 DOMParser（浏览器内置；Node 测试注入 @xmldom/xmldom）。
 * 注意：思源 getDoc 输出是良构 XML（lxml 可解析），统一用 application/xml 解析。
 */

export interface InlineRun {
    text: string;
    bold?: boolean;
    italic?: boolean;
    code?: boolean;
    sup?: boolean;
}

export type SyBlock =
    | { type: "heading"; level: number; text: string }
    | { type: "paragraph"; text: string }
    | { type: "table"; rows: string[][] }
    | { type: "list"; items: string[] }
    | { type: "code"; code: string }
    | { type: "math"; code: string }
    | { type: "quote"; text: string }
    | { type: "thematic" }
    | { type: "image"; src: string; alt: string }
    | { type: "html"; text: string };

/** 解析 getDoc content 为 DOM 根元素（多根 div 片段包一层根；思源块图标用 xlink 前缀需声明） */
export function parseHtml(content: string): Element {
    const parser = new DOMParser();
    return parser.parseFromString(
        `<div id="syass-root" xmlns:xlink="http://www.w3.org/1999/xlink">${content}</div>`,
        "application/xml"
    ).documentElement;
}

/** 取所有带 data-node-index 的 div（标准 API，兼容 xmldom） */
function nodeIndexedDivs(root: Element): HTMLElement[] {
    return Array.from(root.getElementsByTagName("div")).filter((el) =>
        el.hasAttribute("data-node-index")
    ) as HTMLElement[];
}

/** 按 data-node-index 范围原地裁剪 DOM（从最深到最浅删除） */
export function sliceElement(root: Element, start: number, end: number): void {
    const all = nodeIndexedDivs(root);
    const keepIds = new Set<string>();
    for (const el of all) {
        const idx = parseInt(el.getAttribute("data-node-index") || "0", 10);
        if (idx >= start && idx <= end) {
            keepIds.add(el.getAttribute("data-node-id") || "");
        }
    }
    const toRemove = all
        .filter((el) => !keepIds.has(el.getAttribute("data-node-id") || ""))
        .sort(
            (a, b) =>
                nodeIndexedDivs(b).length -
                nodeIndexedDivs(a).length
        );
    for (const el of toRemove) {
        if (el.parentElement) {
            el.parentElement.removeChild(el);
        }
    }
}

/** 块内联内容 → 结构化 runs（docx 渲染用） */
export function inlineToRuns(el: Element): InlineRun[] {
    const out: InlineRun[] = [];
    const push = (text: string, bold?: boolean, italic?: boolean, code?: boolean, sup?: boolean) => {
        if (!text) {
            return;
        }
        out.push({ text, bold, italic, code, sup });
    };
    const walk = (node: Element, bold: boolean, italic: boolean, code: boolean, sup: boolean) => {
        for (const child of Array.from(node.childNodes)) {
            if (child.nodeType === 3) {
                push(child.textContent || "", bold, italic, code, sup);
                continue;
            }
            if (child.nodeType !== 1) {
                continue;
            }
            const c = child as Element;
            const tag = (c.localName || "").toLowerCase();
            const dt = c.getAttribute("data-type") || "";
            if (tag === "span" && dt.includes("strong")) {
                walk(c, true, italic, code, sup);
            } else if (tag === "span" && dt.includes("em")) {
                walk(c, bold, true, code, sup);
            } else if (tag === "span" && dt.includes("code")) {
                walk(c, bold, italic, true, sup);
            } else if (tag === "span" && (dt.includes("u") || dt.includes("mark") || dt.includes("kbd") || dt.includes("tag"))) {
                walk(c, bold, italic, code, sup);
            } else if (tag === "sup") {
                walk(c, bold, italic, code, true);
            } else if (tag === "sub") {
                walk(c, bold, italic, code, false);
            } else if (tag === "a") {
                walk(c, bold, italic, code, sup);
            } else if (tag === "strong" || tag === "b") {
                walk(c, true, italic, code, sup);
            } else if (tag === "em" || tag === "i") {
                walk(c, bold, true, code, sup);
            } else if (tag === "img") {
                const src = c.getAttribute("data-src") || c.getAttribute("src") || "";
                if (src) {
                    push(`![${c.getAttribute("alt") || ""}](${src})`);
                }
            } else if (tag === "br") {
                push("\n");
            } else {
                walk(c, bold, italic, code, sup);
            }
        }
    };
    walk(el, false, false, false, false);
    return mergeRuns(out);
}

/** 合并相邻同格式 run */
export function mergeRuns(runs: InlineRun[]): InlineRun[] {
    const out: InlineRun[] = [];
    for (const r of runs) {
        const last = out[out.length - 1];
        if (
            last &&
            last.bold === r.bold &&
            last.italic === r.italic &&
            last.code === r.code &&
            last.sup === r.sup
        ) {
            last.text += r.text;
        } else {
            out.push({ ...r });
        }
    }
    return out;
}

/** 块内联内容 → markdown 文本（节级导出用；strong/em/code/u/a/img） */
export function inlineToMd(el: Element): string {
    let out = "";
    for (const child of Array.from(el.childNodes)) {
        if (child.nodeType === 3) {
            out += child.textContent || "";
            continue;
        }
        if (child.nodeType !== 1) {
            continue;
        }
        const node = child as Element;
        const tag = (node.localName || "").toLowerCase();
        const dataType = node.getAttribute("data-type") || "";
        if (tag === "span" && dataType.includes("strong")) {
            out += `**${inlineToMd(node)}**`;
        } else if (tag === "span" && dataType.includes("em")) {
            out += `*${inlineToMd(node)}*`;
        } else if (tag === "span" && dataType.includes("code")) {
            out += `\`${inlineToMd(node)}\``;
        } else if (tag === "span" && dataType.includes("u")) {
            out += inlineToMd(node);
        } else if (tag === "a") {
            const href = node.getAttribute("data-href") || node.getAttribute("href") || "";
            const text = inlineToMd(node);
            out += href && !href.startsWith("siyuan://") ? `[${text}](${href})` : text;
        } else if (tag === "img") {
            const src = node.getAttribute("data-src") || node.getAttribute("src") || "";
            if (src) {
                out += `![${node.getAttribute("alt") || ""}](${src})`;
            }
        } else if (tag === "br") {
            out += "\n";
        } else {
            out += inlineToMd(node);
        }
    }
    return out.replace(/\s+/g, " ").trim();
}

function escPipe(s: string): string {
    return s.replace(/\|/g, "\\|");
}

/** 取元素文本：innerText 优先（浏览器渲染文本），xmldom 等无 innerText 时退 textContent */
function elText(el: Element): string {
    const ie = (el as HTMLElement).innerText;
    return ie !== undefined && ie !== null ? ie : el.textContent || "";
}

/** 表格元素 → rows（跳过全空行，半角 | 转义） */
export function tableToRows(table: Element): string[][] {
    const rows: string[][] = [];
    for (const tr of Array.from(table.getElementsByTagName("tr"))) {
        const cells: string[] = [];
        for (const child of Array.from(tr.childNodes)) {
            if (child.nodeType !== 1) {
                continue;
            }
            const td = child as Element;
            const local = (td.localName || "").toLowerCase();
            if (local !== "th" && local !== "td") {
                continue;
            }
            let text = elText(td).replace(/\s*\n+\s*/g, " ").trim();
            text = escPipe(text);
            cells.push(text);
        }
        if (cells.some((c) => c !== "")) {
            rows.push(cells);
        }
    }
    return rows;
}

/** DOM → 中间块数组（递归遍历，对齐 v0.1.0 的 nodeToMd 行为） */
export function htmlToBlocks(root: Element, skipCallout: boolean): SyBlock[] {
    const blocks: SyBlock[] = [];
    const walk = (el: Element, depth: number) => {
        for (const child of Array.from(el.childNodes)) {
            if (child.nodeType === 3) {
                continue;
            }
            if (child.nodeType !== 1) {
                continue;
            }
            const node = child as Element;
            const tag = (node.localName || "").toLowerCase();
            const dataType = node.getAttribute("data-type") || "";
            if (dataType === "NodeHeading") {
                const subtype = node.getAttribute("data-subtype") || "";
                const m = /^h([1-6])$/.exec(subtype);
                const level = m ? parseInt(m[1], 10) : 2;
                const text = inlineToMd(node).trim().replace(/\*\*/g, "");
                if (text) {
                    blocks.push({ type: "heading", level, text });
                }
            } else if (dataType === "NodeParagraph") {
                const text = inlineToMd(node).trim();
                if (text) {
                    // 段落内嵌图片（![alt](src)）拆成独立 image 块，避免 docx 里显示字面文本
                    const imgRe = /!\[([^\]]*)\]\(([^)]+)\)/g;
                    let m: RegExpExecArray | null;
                    let last = 0;
                    let hasImg = false;
                    while ((m = imgRe.exec(text)) !== null) {
                        hasImg = true;
                        const before = text.slice(last, m.index).trim();
                        if (before) {
                            blocks.push({ type: "paragraph", text: before });
                        }
                        blocks.push({ type: "image", src: m[2].trim(), alt: m[1] });
                        last = m.index + m[0].length;
                    }
                    if (!hasImg) {
                        blocks.push({ type: "paragraph", text });
                    } else {
                        const after = text.slice(last).trim();
                        if (after) {
                            blocks.push({ type: "paragraph", text: after });
                        }
                    }
                }
            } else if (dataType === "NodeTable") {
                const rows = tableToRows(node);
                if (rows.length > 0) {
                    blocks.push({ type: "table", rows });
                }
            } else if (dataType === "NodeList" || dataType === "NodeListItem") {
                const items = Array.from(node.childNodes)
                    .filter((c) => c.nodeType === 1)
                    .map((c) => c as Element)
                    .filter((c) => {
                        const dt = c.getAttribute("data-type") || "";
                        const local = (c.localName || "").toLowerCase();
                        return dt === "NodeListItem" || local === "li";
                    });
                if (items.length === 0) {
                    walk(node, depth);
                    continue;
                }
                const listItems: string[] = [];
                for (const item of items) {
                    const md = inlineToMd(item as Element).trim();
                    if (md) {
                        listItems.push(md);
                    }
                }
                if (listItems.length > 0) {
                    blocks.push({ type: "list", items: listItems });
                }
                // List item paragraphs are already represented by this list block.
                // Do not recurse into them, otherwise every item is exported again
                // as a normal paragraph after the list.
            } else if (dataType === "NodeCodeBlock") {
                const code = elText(node).replace(/\n+$/, "");
                blocks.push({ type: "code", code });
            } else if (dataType === "NodeMathBlock") {
                blocks.push({ type: "math", code: elText(node).trim() });
            } else if (dataType === "NodeBlockquote") {
                const text = inlineToMd(node).trim();
                if (text) {
                    blocks.push({ type: "quote", text });
                }
            } else if (dataType === "NodeCallout") {
                if (!skipCallout) {
                    const text = inlineToMd(node).trim();
                    if (text) {
                        blocks.push({ type: "quote", text });
                    }
                }
            } else if (dataType === "NodeThematicBreak") {
                blocks.push({ type: "thematic" });
            } else if (dataType === "NodeImage") {
                const img = node.getElementsByTagName("img")[0];
                if (img) {
                    const src = img.getAttribute("data-src") || img.getAttribute("src") || "";
                    const alt = img.getAttribute("alt") || "";
                    if (src) {
                        blocks.push({ type: "image", src, alt });
                    }
                }
            } else if (dataType === "NodeHTMLBlock") {
                const text = inlineToMd(node).trim();
                if (text) {
                    blocks.push({ type: "html", text });
                }
            } else if (tag === "img") {
                const src = node.getAttribute("data-src") || node.getAttribute("src") || "";
                if (src) {
                    blocks.push({ type: "image", src, alt: node.getAttribute("alt") || "" });
                }
            } else {
                walk(node, depth);
            }
        }
    };
    walk(root, 0);
    return blocks;
}

/** 中间块 → markdown */
export function mdFromBlocks(blocks: SyBlock[]): string {
    const lines: string[] = [];
    for (const b of blocks) {
        switch (b.type) {
            case "heading":
                lines.push(`${"#".repeat(b.level)} ${b.text}`);
                break;
            case "paragraph":
                lines.push(b.text);
                break;
            case "table": {
                if (b.rows.length === 0) {
                    break;
                }
                const nCols = Math.max(...b.rows.map((r) => r.length));
                const out = b.rows.map((r) => {
                    while (r.length < nCols) {
                        r.push("");
                    }
                    return `|${r.join("|")}|`;
                });
                out.splice(1, 0, `|${Array(nCols).fill("---").join("|")}|`);
                lines.push(out.join("\n"));
                break;
            }
            case "list":
                for (const item of b.items) {
                    lines.push(`- ${item}`);
                }
                break;
            case "code":
                lines.push("```", b.code, "```");
                break;
            case "math":
                lines.push("$$", b.code, "$$");
                break;
            case "quote":
                lines.push(b.text.split("\n").map((l) => `> ${l}`).join("\n"));
                break;
            case "thematic":
                lines.push("---");
                break;
            case "image":
                lines.push(`![${b.alt}](${b.src})`);
                break;
            case "html":
                lines.push(b.text);
                break;
        }
    }
    return lines.join("\n\n").replace(/\n{3,}/g, "\n\n").trim();
}
