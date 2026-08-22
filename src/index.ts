/*
 * SiYuan Assistant (思源助手)
 * v0.1.0 — 前端插件，minAppVersion 3.6.4
 *
 * 功能：
 *  1. DOCX 导入（JSZip 解析 → 图片上传 → createDocWithMd → renameDoc 设标题）
 *  2. 标题扁平化（当前文档 H2~H6 全部降级为 H5，H1 不动）
 *  3. 图片宽度设置（setBlockAttrs custom-data-width-percent，自动横85%/竖50%）
 *  4. 节级导出（getDoc HTML 按 data-node-index 切片 → markdown 复制）
 */
import {
    Plugin,
    showMessage,
    confirm,
    Dialog,
    Menu,
    fetchPost,
    getFrontend,
} from "siyuan";
import { parseDocx } from "./docx";
import "./index.scss";

// ---------- WordprocessingML 命名空间 ----------

interface ApiResp {
    code: number;
    msg: string;
    data: any;
}

function api(path: string, payload: any): Promise<ApiResp> {
    return new Promise((resolve, reject) => {
        fetchPost(path, payload, ((resp: any) => {
            if (resp && resp.code === 0) {
                resolve(resp as ApiResp);
            } else {
                reject(new Error((resp && resp.msg) || `API ${path} 失败`));
            }
        }) as any);
    });
}

/** 思源资产访问 URL（渲染进程内可用） */
function assetUrl(src: string): string {
    const token = window.localStorage.getItem("token") || "";
    const clean = src.replace(/^assets\//, "");
    return `${location.origin}/assets/${clean}?token=${encodeURIComponent(token)}`;
}

async function copyText(text: string): Promise<boolean> {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch (e) {
        try {
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.opacity = "0";
            document.body.appendChild(ta);
            ta.select();
            const ok = document.execCommand("copy");
            document.body.removeChild(ta);
            return ok;
        } catch (e2) {
            return false;
        }
    }
}


/** 图片块信息（用于宽度设置） */
interface ImgBlockInfo {
    id: string;
    src: string;
    width: number;
    height: number;
    pct: number;
}

interface HeadingInfo {
    idx: number;
    level: number;
    text: string;
}

export default class SiYuanAssistant extends Plugin {
    private isMobile = false;

    async onload() {
        this.isMobile = ["mobile", "browser-mobile"].includes(getFrontend());
        this.addIcons(`<symbol id="iconSyassImport" viewBox="0 0 32 32"><path d="M3 6h9l2 2h9a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V6z"/><path d="M16 12v8M12 16l4 4 4-4" fill="none" stroke="currentColor" stroke-width="2"/></symbol>
<symbol id="iconSyassFlatten" viewBox="0 0 32 32"><path d="M16 4v15M10 13l6 6 6-6" fill="none" stroke="currentColor" stroke-width="2"/><path d="M6 26h20" stroke="currentColor" stroke-width="2"/></symbol>
<symbol id="iconSyassImage" viewBox="0 0 32 32"><path d="M4 6h24v20H4z" fill="none" stroke="currentColor" stroke-width="2"/><circle cx="11" cy="11" r="2.5" fill="currentColor"/><path d="M4 22l7-7 6 6 4-4 7 7" fill="none" stroke="currentColor" stroke-width="2"/></symbol>
<symbol id="iconSyassSection" viewBox="0 0 32 32"><path d="M4 4h16v6h-2V6H6v20h12v-4h2v6H4z" fill="currentColor"/><path d="M20 8l6 6-6 6-1.4-1.4 3.6-3.6H11v-2h11.2l-3.6-3.6z" fill="currentColor"/></symbol>`);

        this.addCommand({
            langKey: "cmdDocxImport",
            hotkey: "⇧⌘I",
            callback: () => {
                this.openDocxImportDialog();
            },
        });
        this.addCommand({
            langKey: "cmdFlatten",
            hotkey: "⇧⌘J",
            callback: () => {
                this.openFlattenDialog();
            },
        });
        this.addCommand({
            langKey: "cmdImgWidth",
            hotkey: "⇧⌘K",
            callback: () => {
                this.openImageWidthDialog();
            },
        });
        this.addCommand({
            langKey: "cmdSection",
            hotkey: "⇧⌘L",
            callback: () => {
                this.openSectionExportDialog();
            },
        });
    }

    onLayoutReady() {
        const topBarElement = this.addTopBar({
            icon: "iconSyassSection",
            title: "思源助手",
            position: "right",
            callback: () => {
                const menu = new Menu("syassTopBar", () => {});
                menu.addItem({
                    icon: "iconSyassImport",
                    label: "DOCX 导入…",
                    click: () => {
                        this.openDocxImportDialog();
                    },
                });
                menu.addItem({
                    icon: "iconSyassFlatten",
                    label: "标题降级为 H5…",
                    click: () => {
                        this.openFlattenDialog();
                    },
                });
                menu.addItem({
                    icon: "iconSyassImage",
                    label: "图片宽度…",
                    click: () => {
                        this.openImageWidthDialog();
                    },
                });
                menu.addItem({
                    icon: "iconSyassSection",
                    label: "节级导出（复制 Markdown）…",
                    click: () => {
                        this.openSectionExportDialog();
                    },
                });
                menu.open({
                    x: topBarElement.getBoundingClientRect().x,
                    y: topBarElement.getBoundingClientRect().y + 30,
                });
            },
        });
        void topBarElement;
    }

    onunload() {}

    // ================================================================
    // 通用工具
    // ================================================================

    private currentEditor(): IPluginEditor | null {
        const anyThis = this as any;
        return anyThis.getEditor ? (anyThis.getEditor() as IPluginEditor) : null;
    }

    private currentDocId(): string {
        const editor = this.currentEditor();
        if (!editor || !editor.protyle || !editor.protyle.block) {
            throw new Error("请先打开一个文档");
        }
        return editor.protyle.block.rootID;
    }

    private newDialog(title: string, content: string, width?: string): Dialog {
        return new Dialog({
            title,
            content,
            width: this.isMobile ? "92vw" : width || "560px",
        });
    }

    private static escPipe(s: string): string {
        return s.replace(/\|/g, "\\|");
    }

    /** 取块 markdown 的标题级别（# 数量），非标题返回 0 */
    private static headingLevelOfMd(md: string): number {
        const m = /^(#{1,6})\s/.exec(md);
        return m ? m[1].length : 0;
    }

    // ================================================================
    // 功能 1：DOCX 导入
    // ================================================================

    private openDocxImportDialog() {
        const dlg = this.newDialog(
            "DOCX 导入",
            `<div class="b3-dialog__content">
  <div class="fn__block">
    <label class="fn__block b3-label">目标笔记本</label>
    <select id="syassNotebook" class="b3-select fn__block"></select>
  </div>
  <div class="fn__block">
    <label class="fn__block b3-label">文档标题（留空自动取文件名）</label>
    <input id="syassDocTitle" class="b3-text-field fn__block" placeholder="文档标题">
  </div>
  <div class="fn__block">
    <label class="fn__block b3-label">选择 .docx 文件</label>
    <input id="syassDocxFile" type="file" accept=".docx" class="fn__block">
  </div>
  <div class="fn__block">
    <label class="fn__block"><input type="checkbox" id="syassKeepH1"> 正文保留一级标题（不勾选则第一个 H1 作为文档标题，正文不含 H1）</label>
  </div>
</div>
<div class="b3-dialog__action">
  <button class="b3-button b3-button--cancel" id="syassCancel">取消</button>
  <span class="fn__space"></span>
  <button class="b3-button b3-button--text" id="syassGo">开始导入</button>
</div>`
        );
        const sel = dlg.element.querySelector("#syassNotebook") as HTMLSelectElement;
        const fileInput = dlg.element.querySelector("#syassDocxFile") as HTMLInputElement;
        const titleInput = dlg.element.querySelector("#syassDocTitle") as HTMLInputElement;
        const keepH1 = dlg.element.querySelector("#syassKeepH1") as HTMLInputElement;
        const goBtn = dlg.element.querySelector("#syassGo") as HTMLButtonElement;

        api("/api/notebook/lsNotebooks", {})
            .then((resp) => {
                const notebooks: Array<{ id: string; name: string }> = resp.data.notebooks || [];
                for (const nb of notebooks) {
                    const opt = document.createElement("option");
                    opt.value = nb.id;
                    opt.textContent = nb.name;
                    sel.appendChild(opt);
                }
            })
            .catch((e) => {
                showMessage(`加载笔记本列表失败: ${e.message}`);
            });

        dlg.element.querySelector("#syassCancel").addEventListener("click", () => dlg.destroy());
        goBtn.addEventListener("click", async () => {
            const file = fileInput.files && fileInput.files[0];
            if (!file) {
                showMessage("请先选择 .docx 文件");
                return;
            }
            const notebook = sel.value;
            if (!notebook) {
                showMessage("请选择目标笔记本");
                return;
            }
            goBtn.disabled = true;
            goBtn.textContent = "导入中…";
            try {
                const title = (titleInput.value || file.name.replace(/\.docx$/i, "")).trim();
                const parsed = await this.parseDocxFile(file);
                const docId = await this.importDocx(parsed, notebook, title, keepH1.checked);
                showMessage(`导入成功：${title}（${docId}）`);
                dlg.destroy();
            } catch (e) {
                showMessage(`导入失败: ${(e as Error).message}`);
                goBtn.disabled = false;
                goBtn.textContent = "开始导入";
            }
        });
    }

    private async parseDocxFile(file: File): Promise<{ md: string; title: string }> {
        const ab = await file.arrayBuffer();
        const baseName = file.name.replace(/\.docx$/i, "");
        return parseDocx(ab, baseName, (blob, name, dir) => this.uploadImage(blob, name, dir));
    }
    private async importDocx(
        parsed: { md: string; title: string },
        notebook: string,
        title: string,
        keepH1: boolean
    ): Promise<string> {
        let body = parsed.md;
        if (!keepH1) {
            // 去掉与标题相同的开头 H1，避免正文重复
            const re = new RegExp(`^#\\s+${SiYuanAssistant.escapeRegex(title)}\\s*\\n+`);
            body = body.replace(re, "");
            body = body.trim();
        }
        const resp = await api("/api/filetree/createDocWithMd", {
            notebook,
            path: "",
            markdown: body,
        });
        const docId = resp.data as string;
        if (!docId) {
            throw new Error("createDocWithMd 未返回文档 ID");
        }
        // 设置文档标题（标题与正文 H1 独立）
        await api("/api/filetree/renameDoc", { id: docId, title });
        return docId;
    }

    private static escapeRegex(s: string): string {
        return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    /** 段落 → markdown 行（含内联格式与图片占位符） */
    private async uploadImage(blob: Blob, name: string, dir: string): Promise<string> {
        const token = window.localStorage.getItem("token") || "";
        const fd = new FormData();
        fd.append("assetsDirPath", dir);
        fd.append("file[]", blob, name);
        const paths = ["/api/upload", "/api/asset/upload", "/api/upload2"];
        for (const path of paths) {
            try {
                const resp = await fetch(path, {
                    method: "POST",
                    headers: { Authorization: "Token " + token },
                    body: fd,
                });
                const json: any = await resp.json().catch((): null => null);
                if (json && json.code === 0 && json.data && json.data[0] && json.data[0].url) {
                    return json.data[0].url as string;
                }
            } catch (e) {
                // 继续尝试下一个端点
            }
        }
        throw new Error(`上传图片失败: ${name}`);
    }

    // ================================================================
    // 功能 2：标题扁平化（H2~H6 → H5）
    // ================================================================

    private async openFlattenDialog() {
        let docId = "";
        try {
            docId = this.currentDocId();
        } catch (e) {
            showMessage((e as Error).message);
            return;
        }
        let headings: Array<{ id: string; level: number; text: string }> = [];
        try {
            const resp = await api("/api/query/sql", {
                stmt: `SELECT id, markdown FROM blocks WHERE root_id = '${docId}' AND type = 'h'`,
            });
            headings = (resp.data as Array<{ id: string; markdown: string }>)
                .map((b) => {
                    const level = SiYuanAssistant.headingLevelOfMd(b.markdown || "");
                    return {
                        id: b.id,
                        level,
                        text: (b.markdown || "").replace(/^#{1,6}\s*/, "").replace(/\*\*/g, "").trim(),
                    };
                })
                .filter((h) => h.level > 0);
        } catch (e) {
            showMessage(`查询标题失败: ${(e as Error).message}`);
            return;
        }
        const targets = headings.filter((h) => h.level >= 2);
        if (targets.length === 0) {
            showMessage("当前文档没有 H2 及以上的标题，无需降级");
            return;
        }
        const countByLevel: Record<number, number> = {};
        for (const h of targets) {
            countByLevel[h.level] = (countByLevel[h.level] || 0) + 1;
        }
        const summary = Object.keys(countByLevel)
            .sort()
            .map((l) => `H${l}×${countByLevel[parseInt(l, 10)]}`)
            .join(" / ");
        confirm(
            "标题降级",
            `当前文档标题统计：${summary}<br>将把 H2~H6 共 <b>${targets.length}</b> 个标题全部降级为 <b>H5</b>（H1 文档标题不动）。确定执行？`,
            async () => {
                let ok = 0;
                for (const h of targets) {
                    try {
                        await api("/api/block/updateBlock", {
                            id: h.id,
                            dataType: "markdown",
                            data: `##### ${h.text}`,
                        });
                        ok++;
                    } catch (e) {
                        console.warn("降级失败:", h.id, e);
                    }
                }
                showMessage(`标题降级完成：${ok}/${targets.length} 个已变为 H5`);
            }
        );
    }

    // ================================================================
    // 功能 3：图片宽度
    // ================================================================

    private async openImageWidthDialog() {
        let docId = "";
        try {
            docId = this.currentDocId();
        } catch (e) {
            showMessage((e as Error).message);
            return;
        }
        let rows: Array<{ id: string; markdown: string }> = [];
        try {
            const resp = await api("/api/query/sql", {
                stmt: `SELECT id, markdown FROM blocks WHERE root_id = '${docId}' AND (type = 'p' OR type = 'h') AND markdown LIKE '%![%'`,
            });
            rows = resp.data || [];
        } catch (e) {
            showMessage(`查询图片失败: ${(e as Error).message}`);
            return;
        }
        if (rows.length === 0) {
            showMessage("当前文档没有图片块");
            return;
        }
        const imgBlocks: ImgBlockInfo[] = [];
        for (const row of rows) {
            const clean = (row.markdown || "").replace(/\{:.*?\}\s*$/g, "");
            const m = /!\[([^\]]*)\]\(([^)]+)\)/.exec(clean);
            if (m) {
                imgBlocks.push({ id: row.id, src: m[2].trim(), width: 0, height: 0, pct: 0 });
            }
        }
        if (imgBlocks.length === 0) {
            showMessage("当前文档没有找到可识别的图片");
            return;
        }

        const dlg = this.newDialog(
            "图片宽度",
            `<div class="b3-dialog__content">
  <div class="fn__block">共找到 <b>${imgBlocks.length}</b> 张图片</div>
  <div class="fn__block">
    <label class="fn__block b3-label">模式</label>
    <label class="fn__block"><input type="radio" name="syassMode" value="auto" checked> 自动（横图 85% / 竖图 50%）</label>
    <label class="fn__block"><input type="radio" name="syassMode" value="custom"> 自定义百分比</label>
    <input id="syassPct" type="number" min="1" max="100" value="85" class="b3-text-field fn__block" style="display:none" placeholder="百分比，如 85">
  </div>
  <div class="fn__block">
    <div class="syass__imglist" id="syassImgList"></div>
  </div>
</div>
<div class="b3-dialog__action">
  <button class="b3-button b3-button--cancel" id="syassCancel">取消</button>
  <span class="fn__space"></span>
  <button class="b3-button b3-button--text" id="syassGo">应用</button>
</div>`,
            "620px"
        );

        const listEl = dlg.element.querySelector("#syassImgList") as HTMLDivElement;
        const pctInput = dlg.element.querySelector("#syassPct") as HTMLInputElement;
        const radios = dlg.element.querySelectorAll('input[name="syassMode"]');
        for (const radio of Array.from(radios)) {
            radio.addEventListener("change", () => {
                pctInput.style.display = (radio as HTMLInputElement).value === "custom" ? "" : "none";
            });
        }

        // 渲染列表 + 测量宽高（自动模式需要）
        for (const info of imgBlocks) {
            const item = document.createElement("div");
            item.className = "syass__imgitem";
            item.innerHTML = `<span class="syass__imgname">${SiYuanAssistant.escapeHtml(info.src)}</span><span class="syass__imgdim">测量中…</span>`;
            listEl.appendChild(item);
            const dimEl = item.querySelector(".syass__imgdim") as HTMLSpanElement;
            const img = new Image();
            img.onload = () => {
                info.width = img.naturalWidth;
                info.height = img.naturalHeight;
                info.pct = info.width >= info.height ? 85 : 50;
                dimEl.textContent = `${info.width}×${info.height} → ${info.pct}%`;
            };
            img.onerror = () => {
                info.pct = 85;
                dimEl.textContent = "尺寸未知 → 85%";
            };
            img.src = assetUrl(info.src);
        }

        dlg.element.querySelector("#syassCancel").addEventListener("click", () => dlg.destroy());
        dlg.element.querySelector("#syassGo").addEventListener("click", async () => {
            const mode = (dlg.element.querySelector('input[name="syassMode"]:checked') as HTMLInputElement).value;
            const custom = parseInt(pctInput.value, 10);
            const goBtn = dlg.element.querySelector("#syassGo") as HTMLButtonElement;
            goBtn.disabled = true;
            goBtn.textContent = "应用中…";
            let ok = 0;
            for (const info of imgBlocks) {
                const pct = mode === "custom" ? Math.min(100, Math.max(1, custom || 85)) : info.pct || 85;
                try {
                    await api("/api/attr/setBlockAttrs", {
                        id: info.id,
                        attrs: { "custom-data-width-percent": `${pct}%` },
                    });
                    ok++;
                } catch (e) {
                    console.warn("设置宽度失败:", info.id, e);
                }
            }
            showMessage(`图片宽度设置完成：${ok}/${imgBlocks.length}`);
            dlg.destroy();
        });
    }

    private static escapeHtml(s: string): string {
        return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
    }

    // ================================================================
    // 功能 4：节级导出（切片 → markdown 复制）
    // ================================================================

    private async openSectionExportDialog() {
        let docId = "";
        try {
            docId = this.currentDocId();
        } catch (e) {
            showMessage((e as Error).message);
            return;
        }
        let content = "";
        try {
            const resp = await api("/api/filetree/getDoc", { id: docId });
            content = resp.data.content || "";
        } catch (e) {
            showMessage(`读取文档失败: ${(e as Error).message}`);
            return;
        }
        const wrap = document.createElement("div");
        wrap.innerHTML = content;
        const headings: HeadingInfo[] = [];
        for (const el of Array.from(wrap.querySelectorAll<HTMLElement>("div[data-node-index]"))) {
            if (el.dataset.type !== "NodeHeading") {
                continue;
            }
            const idx = parseInt(el.dataset.nodeIndex || "0", 10);
            const subtype = el.dataset.subtype || "";
            const level = /^h([1-6])$/.exec(subtype) ? parseInt(RegExp.$1, 10) : 0;
            const text = ((el as HTMLElement).innerText || "").replace(/\s+/g, " ").trim().slice(0, 40);
            if (level > 0) {
                headings.push({ idx, level, text });
            }
        }
        if (headings.length === 0) {
            showMessage("当前文档没有标题，无法按节导出");
            return;
        }

        const opts = headings
            .map((h) => `<option value="${h.idx}">H${h.level} · ${SiYuanAssistant.escapeHtml(h.text)}</option>`)
            .join("");
        const dlg = this.newDialog(
            "节级导出",
            `<div class="b3-dialog__content">
  <div class="fn__block">
    <label class="fn__block b3-label">起始节</label>
    <select id="syassFrom" class="b3-select fn__block">${opts}</select>
  </div>
  <div class="fn__block">
    <label class="fn__block b3-label">结束节（含）</label>
    <select id="syassTo" class="b3-select fn__block">${opts}</select>
  </div>
  <div class="fn__block">
    <label class="fn__block"><input type="checkbox" id="syassSkipCallout" checked> 跳过提示块（callout / 引用标注）</label>
  </div>
</div>
<div class="b3-dialog__action">
  <button class="b3-button b3-button--cancel" id="syassCancel">取消</button>
  <span class="fn__space"></span>
  <button class="b3-button b3-button--text" id="syassGo">复制 Markdown</button>
</div>`,
            "620px"
        );
        const fromSel = dlg.element.querySelector("#syassFrom") as HTMLSelectElement;
        const toSel = dlg.element.querySelector("#syassTo") as HTMLSelectElement;
        toSel.selectedIndex = Math.min(headings.length - 1, 4);
        const skipCallout = dlg.element.querySelector("#syassSkipCallout") as HTMLInputElement;

        dlg.element.querySelector("#syassCancel").addEventListener("click", () => dlg.destroy());
        dlg.element.querySelector("#syassGo").addEventListener("click", async () => {
            const start = parseInt(fromSel.value, 10);
            const end = parseInt(toSel.value, 10);
            const md = this.sliceAndConvert(content, start, end, skipCallout.checked);
            if (!md.trim()) {
                showMessage("切片结果为空");
                return;
            }
            const ok = await copyText(md);
            showMessage(ok ? `已复制 ${md.length} 字符到剪贴板` : "复制失败，请手动复制");
            dlg.destroy();
        });
    }

    /** 按 node-index 范围切片 HTML，并转 markdown */
    private sliceAndConvert(content: string, start: number, end: number, skipCallout: boolean): string {
        const wrap = document.createElement("div");
        wrap.innerHTML = content;
        const all = Array.from(wrap.querySelectorAll<HTMLElement>("div[data-node-index]"));
        const keepIds = new Set<string>();
        for (const el of all) {
            const idx = parseInt(el.dataset.nodeIndex || "0", 10);
            if (idx >= start && idx <= end) {
                keepIds.add(el.dataset.nodeId || "");
            }
        }
        // 从最深到最浅删除，避免父先删子无谓遍历
        const toRemove = all
            .filter((el) => !keepIds.has(el.dataset.nodeId || ""))
            .sort((a, b) => b.querySelectorAll("div[data-node-index]").length - a.querySelectorAll("div[data-node-index]").length);
        for (const el of toRemove) {
            if (el.parentElement) {
                el.parentElement.removeChild(el);
            }
        }
        const lines: string[] = [];
        this.nodeToMd(wrap, lines, 0, skipCallout);
        return lines.join("\n\n").replace(/\n{3,}/g, "\n\n").trim();
    }

    private nodeToMd(el: Element, lines: string[], depth: number, skipCallout: boolean): void {
        for (const child of Array.from(el.childNodes)) {
            if (child.nodeType === 3) {
                continue; // 结构文本节点忽略
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
                const text = this.inlineToMd(node).trim();
                if (text) {
                    lines.push(`${"#".repeat(level)} ${text.replace(/\*\*/g, "")}`);
                }
            } else if (dataType === "NodeParagraph") {
                const md = this.inlineToMd(node).trim();
                if (md) {
                    lines.push(md);
                }
            } else if (dataType === "NodeTable") {
                const md = this.tableElToMd(node);
                if (md) {
                    lines.push(md);
                }
            } else if (dataType === "NodeList" || dataType === "NodeListItem") {
                this.listElToMd(node, lines, depth, skipCallout);
            } else if (dataType === "NodeCodeBlock") {
                const code = ((node as HTMLElement).innerText || "").replace(/\n+$/, "");
                lines.push("```", code, "```");
            } else if (dataType === "NodeMathBlock") {
                lines.push("$$", ((node as HTMLElement).innerText || "").trim(), "$$");
            } else if (dataType === "NodeBlockquote") {
                const md = this.inlineToMd(node).trim();
                if (md) {
                    lines.push(md.split("\n").map((l) => `> ${l}`).join("\n"));
                }
            } else if (dataType === "NodeCallout") {
                if (!skipCallout) {
                    const md = this.inlineToMd(node).trim();
                    if (md) {
                        lines.push(md.split("\n").map((l) => `> ${l}`).join("\n"));
                    }
                }
            } else if (dataType === "NodeThematicBreak") {
                lines.push("---");
            } else if (dataType === "NodeImage") {
                const img = node.querySelector("img");
                if (img) {
                    const src = img.getAttribute("data-src") || img.getAttribute("src") || "";
                    const alt = img.getAttribute("alt") || "";
                    if (src) {
                        lines.push(`![${alt}](${src})`);
                    }
                }
            } else if (dataType === "NodeHTMLBlock") {
                const md = this.inlineToMd(node).trim();
                if (md) {
                    lines.push(md);
                }
            } else if (tag === "img") {
                const src = node.getAttribute("data-src") || node.getAttribute("src") || "";
                if (src) {
                    lines.push(`![${node.getAttribute("alt") || ""}](${src})`);
                }
            } else {
                this.nodeToMd(node, lines, depth, skipCallout);
            }
        }
    }

    /** 块内联内容 → markdown（strong/em/code/u/a/img） */
    private inlineToMd(el: Element): string {
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
                out += `**${this.inlineToMd(node)}**`;
            } else if (tag === "span" && dataType.includes("em")) {
                out += `*${this.inlineToMd(node)}*`;
            } else if (tag === "span" && dataType.includes("code")) {
                out += `\`${this.inlineToMd(node)}\``;
            } else if (tag === "span" && dataType.includes("u")) {
                out += this.inlineToMd(node);
            } else if (tag === "a") {
                const href = node.getAttribute("data-href") || node.getAttribute("href") || "";
                const text = this.inlineToMd(node);
                out += href && !href.startsWith("siyuan://") ? `[${text}](${href})` : text;
            } else if (tag === "img") {
                const src = node.getAttribute("data-src") || node.getAttribute("src") || "";
                if (src) {
                    out += `![${node.getAttribute("alt") || ""}](${src})`;
                }
            } else if (tag === "br") {
                out += "\n";
            } else {
                out += this.inlineToMd(node);
            }
        }
        return out.replace(/\s+/g, " ").trim();
    }

    /** 思源表格 HTML → markdown 表格 */
    private tableElToMd(table: Element): string {
        const rows: string[][] = [];
        for (const tr of Array.from(table.querySelectorAll("tr"))) {
            const cells: string[] = [];
            for (const td of Array.from(tr.querySelectorAll("th, td"))) {
                let text = ((td as HTMLElement).innerText || "").replace(/\s*\n+\s*/g, " ").trim();
                text = SiYuanAssistant.escPipe(text);
                cells.push(text);
            }
            if (cells.some((c) => c !== "")) {
                rows.push(cells);
            }
        }
        if (rows.length === 0) {
            return "";
        }
        const nCols = Math.max(...rows.map((r) => r.length));
        const out = rows.map((r) => {
            while (r.length < nCols) {
                r.push("");
            }
            return `|${r.join("|")}|`;
        });
        out.splice(1, 0, `|${Array(nCols).fill("---").join("|")}|`);
        return out.join("\n");
    }

    /** 列表 → markdown 列表（简单单层） */
    private listElToMd(el: Element, lines: string[], depth: number, skipCallout: boolean): void {
        const items = Array.from(el.querySelectorAll(":scope > div[data-type='NodeListItem'], :scope > li"));
        if (items.length === 0) {
            this.nodeToMd(el, lines, depth, skipCallout);
            return;
        }
        for (const item of items) {
            const itemEl = item as HTMLElement;
            const md = this.inlineToMd(itemEl).trim();
            const indent = "  ".repeat(Math.min(depth, 6));
            if (md) {
                lines.push(`${indent}- ${md}`);
            }
            this.nodeToMd(itemEl, lines, depth + 1, skipCallout);
        }
    }
}
