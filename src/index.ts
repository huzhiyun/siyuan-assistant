/*
 * SiYuan Assistant (思源助手)
 * v0.2.0 — 前端插件，minAppVersion 3.6.4
 *
 * 功能：
 *  1. DOCX 导入（JSZip 解析 → 图片上传 → createDocWithMd，标题走 path 参数）
 *  2. 导出 Word（getDoc HTML → 中间块 → docx 库生成 → 浏览器下载）
 *  3. 标题扁平化（当前文档 H2~H6 全部降级为 H5，H1 不动）
 *  4. 图片宽度设置（setBlockAttrs custom-data-width-percent，自动横85%/竖50%）
 *  5. 节级导出（getDoc HTML 按 data-node-index 切片 → markdown 复制）
 */
import {
    Plugin,
    showMessage,
    confirm,
    Dialog,
    Menu,
    fetchPost,
    getFrontend,
    getActiveEditor,
} from "siyuan";
import { parseDocx, type ImageExtent } from "./docx";
import { rotatedRasterDimensions } from "./image-layout";
import { parseHtml, sliceElement, htmlToBlocks, mdFromBlocks } from "./blocks";
import { blocksToDocument, documentToBlob, type ImageData } from "./docxgen";
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

/** 带超时的 fetch：超时 abort，避免请求挂起导致界面永久等待 */
function fetchWithTimeout(url: string, init: RequestInit, ms: number): Promise<Response> {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), ms);
    return fetch(url, { ...init, signal: ctrl.signal }).finally(() => clearTimeout(timer));
}

/** 带超时的 api 调用（超时报错，恢复按钮） */
function apiWithTimeout(path: string, payload: any, ms: number): Promise<ApiResp> {
    return new Promise((resolve, reject) => {
        const timer = setTimeout(() => reject(new Error(`API ${path} 超时（${ms / 1000}s）`)), ms);
        api(path, payload).then(
            (r) => { clearTimeout(timer); resolve(r); },
            (e) => { clearTimeout(timer); reject(e); }
        );
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
    private reuseImages = true;

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
            langKey: "cmdDocxExport",
            hotkey: "⇧⌘D",
            callback: () => {
                this.openDocxExportDialog();
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
                    icon: "iconSyassSection",
                    label: "导出 Word…",
                    click: () => {
                        this.openDocxExportDialog();
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
        // Official SiYuan SDK route: works for top-bar commands where the
        // plugin instance itself has no editor binding.
        const activeEditor = getActiveEditor(true) as any;
        const fromActiveEditor = activeEditor?.block?.rootID || activeEditor?.protyle?.block?.rootID;
        if (fromActiveEditor) return fromActiveEditor;
        const editor = this.currentEditor();
        const fromEditor = editor?.protyle?.block?.rootID;
        if (fromEditor) return fromEditor;
        const active = document.querySelector(".layout__wnd--active .protyle-wysiwyg[data-node-id]")
            || document.querySelector(".protyle--focus .protyle-wysiwyg[data-node-id]");
        const fromDom = active?.getAttribute("data-node-id") || "";
        if (fromDom) return fromDom;
        throw new Error("未找到当前活动文档；请点击文档编辑区后重试");
    }

    private newDialog(title: string, content: string, width?: string): Dialog {
        return new Dialog({
            title,
            content,
            width: this.isMobile ? "92vw" : width || "560px",
        });
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
    <label class="fn__block"><input type="checkbox" id="syassReuseImages" checked> 复用相同图片（SHA-256，避免重复上传）</label>
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
        const reuseImages = dlg.element.querySelector("#syassReuseImages") as HTMLInputElement;
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
            goBtn.textContent = "开始导入";
            this.reuseImages = reuseImages.checked;
            try {
                const title = (titleInput.value || file.name.replace(/\.docx$/i, "")).trim();
                goBtn.textContent = "正在解析 docx…";
                const parsed = await this.parseDocxFile(file, (done, total) => {
                    goBtn.textContent = `正在上传图片 ${done}/${total}…`;
                });
                goBtn.textContent = "正在创建文档…";
                const docId = await this.importDocx(parsed, notebook, title, keepH1.checked);
                const warn = parsed.uploadFailures > 0 ? `（${parsed.uploadFailures} 张图片上传失败已跳过）` : "";
                showMessage(`导入成功：${title}${warn}`);
                dlg.destroy();
            } catch (e) {
                showMessage(`导入失败: ${(e as Error).message}`);
                goBtn.disabled = false;
                goBtn.textContent = "开始导入";
            }
        });
    }

    private async parseDocxFile(
        file: File,
        onProgress?: (done: number, total: number) => void
    ): Promise<{ md: string; title: string; uploadFailures: number; centeredImageIndexes: number[]; centeredTableIndexes: number[] }> {
        const ab = await file.arrayBuffer();
        const baseName = file.name.replace(/\.docx$/i, "");
        return parseDocx(ab, baseName, (blob, name, dir, rotationDegrees, extent) => this.uploadImage(blob, name, dir, rotationDegrees, extent), {
            concurrency: 4,
            onProgress,
        });
    }
    private async importDocx(
        parsed: { md: string; title: string; centeredImageIndexes: number[]; centeredTableIndexes: number[] },
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
        // ⚠️ 实测（v3.6.4）：renameDoc 返回 code=0 但静默不生效；
        // 正确姿势 = 标题放 createDocWithMd 的 path 末段（skill create-document.js 同款）
        const safeTitle = title.replace(/\//g, "／").trim();
        // 大文档 markdown 可能几百 KB，内核建块耗时较长，给 5 分钟超时兜底（避免永久等待）
        const resp = await apiWithTimeout("/api/filetree/createDocWithMd", {
            notebook,
            path: `/${safeTitle}`,
            markdown: body,
        }, 300000);
        const docId = resp.data as string;
        if (!docId) {
            throw new Error("createDocWithMd 未返回文档 ID");
        }
        await this.applyCenteredDocxLayout(docId, parsed);
        return docId;
    }

    /** Persist Word center alignment on the corresponding imported image/table blocks. */
    private async applyCenteredDocxLayout(
        docId: string,
        parsed: { centeredImageIndexes: number[]; centeredTableIndexes: number[] }
    ): Promise<void> {
        if (parsed.centeredImageIndexes.length === 0 && parsed.centeredTableIndexes.length === 0) return;
        const safeDocId = docId.replace(/'/g, "''");
        let imageBlocks: any[] = [];
        let tableBlocks: any[] = [];
        const expected = parsed.centeredImageIndexes.length + parsed.centeredTableIndexes.length;
        // createDocWithMd returns before large documents necessarily finish materializing
        // every block. Wait briefly and retry rather than reporting the created document
        // as a failed import.
        for (let attempt = 0; attempt < 10; attempt++) {
            const resp = await api("/api/query/sql", {
                stmt: `SELECT id, type, markdown FROM blocks WHERE root_id = '${safeDocId}' AND type IN ('p', 't') ORDER BY sortkey ASC`,
            });
            const blocks = Array.isArray(resp.data) ? resp.data : [];
            imageBlocks = blocks.filter((block: any) => block.type === "p" && /!\[[^\]]*\]\(/.test(block.markdown || ""));
            tableBlocks = blocks.filter((block: any) => block.type === "t");
            const found = parsed.centeredImageIndexes.filter((index) => imageBlocks[index]?.id).length + parsed.centeredTableIndexes.filter((index) => tableBlocks[index]?.id).length;
            if (found === expected) break;
            await new Promise((resolve) => setTimeout(resolve, 500));
        }
        const ids = [
            ...parsed.centeredImageIndexes.map((index) => imageBlocks[index]?.id),
            ...parsed.centeredTableIndexes.map((index) => tableBlocks[index]?.id),
        ].filter((id): id is string => typeof id === "string" && id.length > 0);
        if (ids.length !== expected) {
            console.warn("部分居中布局未定位，文档已正常创建", { expected, found: ids.length, docId });
            return;
        }
        await Promise.all(ids.map(async (id) => {
            await api("/api/attr/setBlockAttrs", { id, attrs: { "custom-syass-align": "center" } });
            const attrs = await api("/api/attr/getBlockAttrs", { id });
            if (attrs.data?.["custom-syass-align"] !== "center") {
                throw new Error(`居中属性写入未确认: ${id}`);
            }
        }));
    }

    private static escapeRegex(s: string): string {
        return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    }

    /** 把 OOXML 图形旋转烧录到像素；思源图片块不保留 Word 的 a:xfrm 元数据。 */
    private async rotateImageBlob(blob: Blob, rotationDegrees: number, extent?: ImageExtent): Promise<Blob> {
        if (!rotationDegrees || blob.type === "image/gif" || typeof createImageBitmap !== "function") {
            return blob;
        }
        try {
            const bitmap = await createImageBitmap(blob);
            const dimensions = rotatedRasterDimensions(bitmap.width, bitmap.height, rotationDegrees, extent);
            const canvas = document.createElement("canvas");
            canvas.width = dimensions.width;
            canvas.height = dimensions.height;
            const ctx = canvas.getContext("2d");
            if (!ctx) {
                return blob;
            }
            ctx.translate(canvas.width / 2, canvas.height / 2);
            ctx.rotate((rotationDegrees * Math.PI) / 180);
            ctx.drawImage(bitmap, -dimensions.sourceWidth / 2, -dimensions.sourceHeight / 2, dimensions.sourceWidth, dimensions.sourceHeight);
            bitmap.close();
            const type = blob.type === "image/png" ? "image/png" : "image/jpeg";
            return await new Promise<Blob>((resolve) => canvas.toBlob((out) => resolve(out || blob), type, 0.92));
        } catch (e) {
            console.warn(`图片旋转处理失败，保留原图: ${rotationDegrees}°`, e);
            return blob;
        }
    }

    private async sha256(blob: Blob): Promise<string> {
        const bytes = await crypto.subtle.digest("SHA-256", await blob.arrayBuffer());
        return Array.from(new Uint8Array(bytes), b => b.toString(16).padStart(2, "0")).join("");
    }

    /** 上传图片到思源 assets，返回 assets 相对路径；失败重试后仍失败返回 ""（不阻塞整体导入） */
    private async uploadImage(blob: Blob, name: string, dir: string, rotationDegrees = 0, extent?: ImageExtent): Promise<string> {
        const normalizedBlob = await this.rotateImageBlob(blob, rotationDegrees, extent);
        const hash = this.reuseImages ? await this.sha256(normalizedBlob) : "";
        const cacheKey = hash ? `syass:image:${hash}` : "";
        const cached = cacheKey ? localStorage.getItem(cacheKey) : "";
        if (cached) return cached;
        const token = window.localStorage.getItem("token") || "";
        const fd = new FormData();
        fd.append("assetsDirPath", dir);
        fd.append("file[]", normalizedBlob, name);
        // 实测（v3.6.4）：/api/upload 404 不存在；正确端点是 /api/asset/upload，
        // 响应为 data.succMap = { 原文件名: "assets/.../改名.png" }
        // 旧版本兼容保留 /api/upload + data[0].url 分支
        const paths = ["/api/asset/upload", "/api/upload"];
        let lastErr: string = "";
        for (const path of paths) {
            for (let attempt = 0; attempt < 2; attempt++) {
                try {
                    const resp = await fetchWithTimeout(path, {
                        method: "POST",
                        headers: { Authorization: "Token " + token },
                        body: fd,
                    }, 30000); // 30s 超时，避免请求挂起
                    const json: any = await resp.json().catch((): null => null);
                    if (json && json.code === 0) {
                        if (json.data && json.data.succMap && json.data.succMap[name]) {
                            const uploaded = json.data.succMap[name] as string;
                            if (cacheKey) localStorage.setItem(cacheKey, uploaded);
                            return uploaded;
                        }
                        if (json.data && json.data[0] && json.data[0].url) {
                            return json.data[0].url as string;
                        }
                    }
                    lastErr = `HTTP ${resp.status}`;
                } catch (e) {
                    lastErr = (e as Error).message || String(e);
                    // 继续尝试（下一次或下一个端点）
                }
            }
        }
        console.warn(`上传图片失败（已跳过）: ${name} — ${lastErr}`);
        return "";
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
        const wrap = parseHtml(content);
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
            const root = parseHtml(content);
            sliceElement(root, start, end);
            const blocks = htmlToBlocks(root, skipCallout.checked);
            const md = mdFromBlocks(blocks);
            if (!md.trim()) {
                showMessage("切片结果为空");
                return;
            }
            const ok = await copyText(md);
            showMessage(ok ? `已复制 ${md.length} 字符到剪贴板` : "复制失败，请手动复制");
            dlg.destroy();
        });
    }

    // ================================================================
    // 功能 2：导出 Word（docx 库，浏览器生成下载）
    // ================================================================

    private async openDocxExportDialog() {
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
        const root = parseHtml(content);
        const headings: HeadingInfo[] = [];
        for (const el of Array.from(root.querySelectorAll<HTMLElement>("div[data-node-index]"))) {
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
        const hasHeadings = headings.length > 0;
        const opts = headings
            .map((h) => `<option value="${h.idx}">H${h.level} · ${SiYuanAssistant.escapeHtml(h.text)}</option>`)
            .join("");

        const dlg = this.newDialog(
            "导出 Word",
            `<div class="b3-dialog__content">
  <div class="fn__block">
    <label class="fn__block b3-label">导出范围</label>
    <label class="fn__block"><input type="radio" name="syassRange" value="all" checked> 全文</label>
    <label class="fn__block"><input type="radio" name="syassRange" value="section" ${hasHeadings ? "" : "disabled"}> 当前节${hasHeadings ? "" : "（文档无标题）"}</label>
  </div>
  <div class="fn__block" id="syassSectionBlock" style="display:none">
    <label class="fn__block b3-label">起始节</label>
    <select id="syassFrom" class="b3-select fn__block">${opts}</select>
    <label class="fn__block b3-label">结束节（含）</label>
    <select id="syassTo" class="b3-select fn__block">${opts}</select>
  </div>
  <div class="fn__block">
    <label class="fn__block"><input type="checkbox" id="syassSkipCallout" checked> 跳过提示块（callout / 引用标注，投标终稿用）</label>
    <label class="fn__block"><input type="checkbox" id="syassAutoNum" checked> Word 原生多级编号（1. / 1.1 / 1.1.1）</label>
  </div>
</div>
<div class="b3-dialog__action">
  <button class="b3-button b3-button--cancel" id="syassCancel">取消</button>
  <span class="fn__space"></span>
  <button class="b3-button b3-button--text" id="syassGo">导出并下载</button>
</div>`,
            "620px"
        );
        const rangeRadios = dlg.element.querySelectorAll('input[name="syassRange"]');
        const sectionBlock = dlg.element.querySelector("#syassSectionBlock") as HTMLDivElement;
        for (const radio of Array.from(rangeRadios)) {
            radio.addEventListener("change", () => {
                sectionBlock.style.display = (radio as HTMLInputElement).value === "section" ? "" : "none";
            });
        }
        const fromSel = dlg.element.querySelector("#syassFrom") as HTMLSelectElement;
        const toSel = dlg.element.querySelector("#syassTo") as HTMLSelectElement;
        if (toSel) {
            toSel.selectedIndex = Math.min(headings.length - 1, 4);
        }
        const skipCallout = dlg.element.querySelector("#syassSkipCallout") as HTMLInputElement;
        const autoNum = dlg.element.querySelector("#syassAutoNum") as HTMLInputElement;
        const goBtn = dlg.element.querySelector("#syassGo") as HTMLButtonElement;

        dlg.element.querySelector("#syassCancel").addEventListener("click", () => dlg.destroy());
        goBtn.addEventListener("click", async () => {
            const range = (dlg.element.querySelector('input[name="syassRange"]:checked') as HTMLInputElement).value;
            goBtn.disabled = true;
            goBtn.textContent = "生成中…";
            try {
                const root2 = parseHtml(content);
                if (range === "section") {
                    sliceElement(root2, parseInt(fromSel.value, 10), parseInt(toSel.value, 10));
                }
                const blocks = htmlToBlocks(root2, skipCallout.checked);
                // 下载图片（并发）
                const images = new Map<string, ImageData>();
                const imgBlocks = blocks.filter((b) => b.type === "image") as Array<{
                    type: "image";
                    src: string;
                    alt: string;
                }>;
                await Promise.all(
                    imgBlocks.map(async (b) => {
                        if (images.has(b.src)) {
                            return;
                        }
                        const data = await this.loadImageData(b.src);
                        if (data) {
                            images.set(b.src, data);
                        }
                    })
                );
                // 文件名：hpath 末段
                let fileName = "导出文档";
                try {
                    const hp = await api("/api/filetree/getHPathByID", { id: docId });
                    const seg = ((hp.data as string) || "").split("/").filter(Boolean);
                    if (seg.length > 0) {
                        fileName = seg[seg.length - 1];
                    }
                } catch (e) {
                    // 忽略，用默认名
                }
                const doc = blocksToDocument(
                    { blocks, images },
                    { autoNumberHeadings: autoNum.checked }
                );
                const blob = await documentToBlob(doc);
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `${fileName}.docx`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                setTimeout(() => URL.revokeObjectURL(url), 3000);
                showMessage(`已导出 ${fileName}.docx（${blocks.length} 块，${images.size} 张图）`);
                dlg.destroy();
            } catch (e) {
                showMessage(`导出失败: ${(e as Error).message}`);
                goBtn.disabled = false;
                goBtn.textContent = "导出并下载";
            }
        });
    }

    /** 下载思源图片并探测尺寸 */
    private async loadImageData(src: string): Promise<ImageData | null> {
        try {
            const resp = await fetch(assetUrl(src));
            if (!resp.ok) {
                return null;
            }
            const buffer = await resp.arrayBuffer();
            const blob = new Blob([buffer]);
            const url = URL.createObjectURL(blob);
            const dims = await new Promise<{ w: number; h: number }>((resolve) => {
                const img = new Image();
                img.onload = () => resolve({ w: img.naturalWidth, h: img.naturalHeight });
                img.onerror = () => resolve({ w: 0, h: 0 });
                img.src = url;
            });
            URL.revokeObjectURL(url);
            const ext = (/\.(\w+)$/.exec(src) || [])[1]?.toLowerCase() || "png";
            const type = ext === "jpg" || ext === "jpeg" ? "jpg" : ext === "gif" ? "gif" : ext === "bmp" ? "bmp" : "png";
            return { buffer, widthPx: dims.w, heightPx: dims.h, type };
        } catch (e) {
            return null;
        }
    }
}
