/*
 * docxgen.ts 单测（docx 库 Node 可跑）
 * 覆盖：标题多级编号/表格/图片占位/分页符/中文样式 → Packer.toBuffer → 检查 XML
 */
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const JSZip = require("jszip");
const { blocksToDocument, runsFromMd } = require("/tmp/syass_docxgen.cjs");
const { Packer } = require("docx");

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

// 1. runsFromMd 解析
const runs = runsFromMd("前**加粗**中`代码`后*斜体*末");
assert("runsFromMd 数量", runs.length === 7, JSON.stringify(runs));
assert("加粗解析", runs[1].bold && runs[1].text === "加粗");
assert("代码解析", runs[3].code && runs[3].text === "代码");
assert("斜体解析", runs[5].italic && runs[5].text === "斜体");

// 2. 生成文档
const blocks = [
    { type: "heading", level: 1, text: "第一部分 概述" },
    { type: "heading", level: 2, text: "1.1 服务范围" },
    { type: "paragraph", text: "这是**加粗**正文段落。" },
    { type: "table", rows: [["列一", "列二"], ["a", "b"]] },
    { type: "list", items: ["项目一", "项目二"] },
    { type: "image", src: "assets/x.png", alt: "图" },
    { type: "thematic" },
];
const doc = blocksToDocument({ blocks, images: new Map() }, {});
const buffer = await Packer.toBuffer(doc);
const zip = await JSZip.loadAsync(buffer);
const xml = await zip.file("word/document.xml").async("string");

assert("docx 生成非空", buffer.byteLength > 1000, `${buffer.byteLength}B`);
assert("标题 Heading1 样式", xml.includes('w:val="Heading1"'), xml.slice(0, 300));
assert("标题 Heading2 样式", xml.includes('w:val="Heading2"'));
assert("多级编号引用 numPr", xml.includes("<w:numPr>") || xml.includes("<w:numPr/>"));
assert("正文加粗 run", xml.includes('<w:b/>') || xml.includes('<w:b '));
assert("表格存在", xml.includes("<w:tbl>"));
assert("表格边框", xml.includes("<w:tblBorders>"));
assert("表头灰底", xml.includes("D9D9D9"));
assert("图片缺失占位", xml.includes("[图片缺失: assets/x.png]"));
assert("分页符", xml.includes('w:type="page"'));
assert("中文字体 eastAsia", xml.includes("w:eastAsia=\"宋体\"") || xml.includes('w:eastAsia="宋体"'));

// 3. numbering.xml 存在
const numberingXml = await zip.file("word/numbering.xml").async("string").catch(() => "");
assert("numbering.xml 多级配置", numberingXml.includes("syass-heading") || numberingXml.includes("abstractNum"));

// 4. 关闭自动编号：标题不带编号，但列表 bullet 仍有 numPr（属正常）
const countNumPr = (s) => (s.match(/<w:numPr>/g) || []).length;
const doc2 = blocksToDocument({ blocks, images: new Map() }, { autoNumberHeadings: false });
const buf2 = await Packer.toBuffer(doc2);
const zip2 = await JSZip.loadAsync(buf2);
const xml2 = await zip2.file("word/document.xml").async("string");
assert(
    "关闭编号后 numPr 减少（标题不编号，仅列表 bullet）",
    countNumPr(xml2) < countNumPr(xml),
    `${countNumPr(xml2)} vs ${countNumPr(xml)}`
);

console.log(`\n结果: ${pass} 通过 / ${fail} 失败`);
process.exit(fail > 0 ? 1 : 0);
