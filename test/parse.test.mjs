/*
 * DOCX 解析器单元测试（Node 环境）
 * 构造真实 OOXML 结构的 docx zip → parseDocx → 断言 markdown 输出
 * 覆盖：Title/Heading 映射、加粗斜体、表格 gridSpan 补空、全空行跳过、图片 rId 提取与上传
 */
import { createRequire } from "module";
const require = createRequire(import.meta.url);

const { DOMParser } = require("@xmldom/xmldom");
globalThis.DOMParser = DOMParser;

const JSZip = require("jszip");
const { parseDocx } = require("/tmp/syass_docx.cjs");

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml">
<w:body>
  <w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>封面标题</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>第一部分 总述</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>1.1 服务范围</w:t></w:r></w:p>
  <w:p><w:pPr><w:pStyle w:val="42"/></w:pPr><w:r><w:t>数字样式标题</w:t></w:r></w:p>
  <w:p><w:r><w:t>分页前</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>分页后</w:t></w:r></w:p>
  <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>第一项</w:t></w:r></w:p>
  <w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> INCLUDEPICTURE &quot;https://example.invalid/image.png&quot; \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>域后正文</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
  <w:p><w:r><w:rPr><w:b w:val="0"/><w:i w:val="0"/></w:rPr><w:t>明确非粗斜</w:t></w:r></w:p>
  <w:p><w:r><w:t>这是</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>加粗</w:t></w:r><w:r><w:t>和</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>斜体</w:t></w:r><w:r><w:t>文本。</w:t></w:r></w:p>
  <w:tbl>
    <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>合并单元格A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t>C</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>D</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>E</w:t></w:r></w:p></w:tc></w:tr>
    <w:tr><w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t></w:t></w:r></w:p></w:tc></w:tr>
  </w:tbl>
  <w:p>
    <w:r><w:drawing>
      <wp:inline><wp:docPr id="0" name="截图1.png" descr="架构截图"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline>
    </w:drawing></w:r>
  </w:p>
  <w:sectPr/>
</w:body>
</w:document>`;

const RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="42">
    <w:name w:val="Custom Heading 3"/>
    <w:pPr><w:outlineLvl w:val="2"/></w:pPr>
  </w:style>
</w:styles>`;

// 1x1 红色 PNG
const PNG_B64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const PNG = Buffer.from(PNG_B64, "base64");

async function buildTestZip() {
    const zip = new JSZip();
    zip.file("word/document.xml", DOCUMENT_XML);
    zip.file("word/styles.xml", STYLES_XML);
    zip.file("word/_rels/document.xml.rels", RELS_XML);
    zip.file("word/media/image1.png", PNG);
    return zip.generateAsync({ type: "arraybuffer" });
}

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

const ab = await buildTestZip();
const uploadCalls = [];
const { md, title } = await parseDocx(ab, "testdoc", (blob, name, dir) => {
    uploadCalls.push({ name, dir, size: blob.size });
    return `${dir}/${name}`;
});

console.log("=== 标题 ===");
assert("title = 封面标题 (Title 样式优先)", title === "封面标题", `got: ${title}`);
assert("H1 块 # 第一部分 总述", md.includes("# 第一部分 总述"));
assert("H2 块 ## 1.1 服务范围", md.includes("## 1.1 服务范围"));
assert("数值样式通过 outlineLvl 还原 H3", md.includes("### 数字样式标题"), md);

console.log("=== 结构语义 ===");
assert("硬分页符转换为独立 thematic break", /(?:^|\n)---(?:\n|$)/.test(md), md);
assert("numPr 转换为有序列表", /(?:^|\n)1\. 第一项(?:\n|$)/.test(md), md);
assert("INCLUDEPICTURE 域指令不进入正文", !md.includes("INCLUDEPICTURE") && !md.includes("MERGEFORMAT"), md);
assert("域后的可见正文保留", md.includes("域后正文"), md);

console.log("=== 内联格式 ===");
assert("显式 w:val=0 的粗斜体开关必须保持普通文本", md.split("\n").includes("明确非粗斜"), md.split("\n").find(l => l.includes("明确非粗斜")));
assert("加粗 **加粗**", md.includes("这是**加粗**和*斜体*文本。"), md.split("\n").find(l => l.includes("加粗")));

console.log("=== 表格 ===");
const tblLines = md.split("\n\n").find(b => b.includes("合并单元格A"));
assert("表格存在", !!tblLines);
assert("gridSpan 补空: |合并单元格A||B|", tblLines && tblLines.includes("|合并单元格A||B|"));
assert("分隔行 |---|---|---|", tblLines && tblLines.includes("|---|---|---|"));
assert("第二行 |C|D|E|", tblLines && tblLines.includes("|C|D|E|"));
assert("全空行已跳过（无 ||||||）", tblLines && !tblLines.includes("||||||"));

console.log("=== 图片 ===");
assert("图片行 ![](assets/testdoc/image_5.png)", md.includes("![架构截图](assets/testdoc/image_5.png)"));
assert("上传回调收到正确参数", uploadCalls.length === 1 && uploadCalls[0].name === "image_5.png" && uploadCalls[0].dir === "assets/testdoc", JSON.stringify(uploadCalls));
assert("上传 blob 是 PNG", uploadCalls.length === 1 && uploadCalls[0].size === PNG.length);

console.log(`\n结果: ${pass} 通过 / ${fail} 失败`);
process.exit(fail > 0 ? 1 : 0);
