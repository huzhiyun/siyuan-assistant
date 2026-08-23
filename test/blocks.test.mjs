/*
 * blocks.ts 单测（Node + xmldom）
 * 覆盖：标题/段落格式/表格/列表/图片/callout/分隔线 → 中间块 → markdown
 */
import { createRequire } from "module";
const require = createRequire(import.meta.url);

const { DOMParser } = require("@xmldom/xmldom");
globalThis.DOMParser = DOMParser;
const { parseHtml, sliceElement, htmlToBlocks, mdFromBlocks, inlineToRuns, inlineToMd } = require("/tmp/syass_blocks.cjs");

const HTML = `<div data-type="NodeDocument">
  <div data-node-id="h1" data-node-index="1" data-type="NodeHeading" data-subtype="h2"><div contenteditable="true">1.1 服务范围</div></div>
  <div data-node-id="p1" data-node-index="2" data-type="NodeParagraph"><div contenteditable="true">这是<span data-type="strong">加粗</span>和<span data-type="em">斜体</span>和<span data-type="code">代码</span>。</div></div>
  <div data-node-id="t1" data-node-index="3" data-type="NodeTable"><table><thead><tr><th><div contenteditable="true">列一</div></th><th><div contenteditable="true">列二</div></th></tr></thead><tbody><tr><td><div contenteditable="true">a|b</div></td><td><div contenteditable="true">b</div></td></tr></tbody></table></div>
  <div data-node-id="l1" data-node-index="4" data-type="NodeList"><div data-type="NodeListItem"><div contenteditable="true">项目一</div></div><div data-type="NodeListItem"><div contenteditable="true">项目二</div></div></div>
  <div data-node-id="img1" data-node-index="5" data-type="NodeImage"><img data-src="assets/test/a.png" alt="测试图"/></div>
  <div data-node-id="c1" data-node-index="6" data-type="NodeCallout"><div class="callout-title">提示</div><div contenteditable="true">这是提示</div></div>
  <div data-node-id="hr1" data-node-index="7" data-type="NodeThematicBreak"></div>
  <div data-node-id="h2" data-node-index="8" data-type="NodeHeading" data-subtype="h3"><div contenteditable="true">1.1.1 细节</div></div>
</div>`;

let pass = 0;
let fail = 0;
function assert(name, cond, extra = "") {
    if (cond) {
        pass++;
        console.log(`  ✓ ${name}`);
    } else {
        fail++;
        console.log(`  ✗ ${name} ${extra}`);
    }
}

// 1. htmlToBlocks
const root = parseHtml(HTML);
const blocks = htmlToBlocks(root, true);
const types = blocks.map((b) => b.type).join(",");
assert("块类型序列正确", types === "heading,paragraph,table,list,image,thematic,heading", `got: ${types}`);

const h1 = blocks[0];
assert("heading 解析", h1.type === "heading" && h1.level === 2 && h1.text === "1.1 服务范围");
const p = blocks[1];
assert("paragraph 内联 md", p.type === "paragraph" && p.text === "这是**加粗**和*斜体*和`代码`。", JSON.stringify(p));
const t = blocks[2];
assert("表格 rows", t.type === "table" && t.rows.length === 2 && t.rows[1][0] === "a\\|b", JSON.stringify(t.rows));
const l = blocks[3];
assert("列表 items", l.type === "list" && l.items.length === 2 && l.items[0] === "项目一");
const img = blocks[4];
assert("图片 src/alt", img.type === "image" && img.src === "assets/test/a.png" && img.alt === "测试图");
assert("callout 已跳过(skipCallout=true)", !blocks.some((b) => b.type === "quote"));
const th = blocks[5];
assert("分隔线", th.type === "thematic");

// 2. callout 保留模式
const blocks2 = htmlToBlocks(parseHtml(HTML), false);
assert("callout 保留为 quote", blocks2.some((b) => b.type === "quote" && (b.text || "").includes("这是提示")));

// 3. inlineToRuns
const paraEl = Array.from(parseHtml(HTML).getElementsByTagName("div")).find(
    (el) => el.getAttribute("data-node-id") === "p1"
);
const runs = inlineToRuns(paraEl);
const joined = runs.map((r) => r.text).join("");
assert("runs 文本完整", joined.includes("这是加粗和斜体和代码。"), joined);
const boldRun = runs.find((r) => r.bold && r.text === "加粗");
assert("加粗 run", !!boldRun);
const codeRun = runs.find((r) => r.code && r.text === "代码");
assert("代码 run", !!codeRun);
const emRun = runs.find((r) => r.italic && r.text === "斜体");
assert("斜体 run", !!emRun);
// 相邻同格式合并
const merged = require("/tmp/syass_blocks.cjs").mergeRuns([
    { text: "a", bold: true },
    { text: "b", bold: true },
    { text: "c" },
]);
assert("mergeRuns 合并相邻", merged.length === 2 && merged[0].text === "ab");

// 4. mdFromBlocks
const md = mdFromBlocks(blocks);
assert("md 含标题", md.includes("## 1.1 服务范围"));
assert("md 含表格", md.includes("|列一|列二|") && md.includes("|a\\|b|b|"));
assert("md 含列表", md.includes("- 项目一"));
assert("md 含图片", md.includes("![测试图](assets/test/a.png)"));
assert("md 含分隔线", md.includes("---"));

// 5. sliceElement
const root3 = parseHtml(HTML);
sliceElement(root3, 3, 5);
const blocks3 = htmlToBlocks(root3, true);
const types3 = blocks3.map((b) => b.type).join(",");
assert("切片 3-5（表格+列表+图片）", types3 === "table,list,image", `got: ${types3}`);

// 6. inlineToMd
const md2 = inlineToMd(paraEl);
assert("inlineToMd 输出", md2 === "这是**加粗**和*斜体*和`代码`。", md2);

console.log(`\n结果: ${pass} 通过 / ${fail} 失败`);
process.exit(fail > 0 ? 1 : 0);
