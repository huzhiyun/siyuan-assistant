# 思源助手 (SiYuan Assistant)

思源笔记一体化工具插件，把 Hermes 思源 skill 的高频操作搬进思源内一键完成，无需再开 Hermes、连内网、处理 token/代理。

## 功能

1. **DOCX 导入**（命令 `⇧⌘I`）
   - 解析 Word 文档：标题（Heading1-9 / Title 样式）、段落、内联格式（加粗/斜体/上标/下标）、表格（gridSpan 自动补空、vMerge 合并行跳过、全空行跳过）、图片（按 rId 提取并上传到 assets）
   - 图片先上传、再 createDocWithMd 建文档（标题通过 path 参数设置，v3.6.4 的 renameDoc 静默失效不可用）
   - 默认把第一个 H1 当文档标题、正文不含 H1（可勾选保留）

2. **导出 Word**（命令 `⇧⌘D`）
   - 思源文档 → .docx，纯前端生成（docx 库），浏览器直接下载，无需 python/pandoc
   - 导出范围：全文 / 当前节（按标题选起止）
   - Word 原生多级编号（1. / 1.1 / 1.1.1，可关）、标题纯净文本、宋体正文 10.5pt + 首行缩进 0.74cm、黑体标题、1.5 倍行距
   - 表格全边框 + 表头灰底加粗、图片按比例缩放（宽≤560px 高≤720px）、分隔线转分页符、可跳过提示块（callout）

3. **标题降级为 H5**（命令 `⇧⌘J`）
   - 当前文档 H2~H6 全部降级为 H5（H1 文档标题不动），方便合入父文档

4. **图片宽度设置**（命令 `⇧⌘K`）
   - 批量设置当前文档所有图片宽度：自动模式按横竖比例（横图 85% / 竖图 50%），或自定义百分比
   - 通过 `setBlockAttrs` 的 `custom-data-width-percent` 实现

5. **节级导出**（命令 `⇧⌘L`）
   - 按标题选择起始/结束节，切片后转 Markdown 复制到剪贴板
   - 可跳过提示块（callout / 引用标注），适合合入投标文档

顶栏右侧也有「思源助手」菜单，四个功能都有入口。

## 安装

思源本地插件 = 把插件目录放进工作空间 `data/plugins/`（集市 UI **没有**手动导入 zip 的选项）：

1. 获取 `siyuan-assistant-v0.2.0.zip`（NAS: `/hermes/输出/`）
2. 在思源所在机器（工作空间如 `/siyuan/workspace/`）执行：
   ```bash
   mkdir -p /siyuan/workspace/data/plugins/siyuan-assistant
   unzip siyuan-assistant-v0.2.0.zip -d /siyuan/workspace/data/plugins/siyuan-assistant/
   ```
   （zip 内是 plugin.json/index.js 等直接文件，解压后目录即插件目录，目录名必须等于 plugin.json 的 `name`）
3. 重启思源（或刷新），打开 设置 → 集市 → 已下载 → 启用「思源助手」

Docker 环境：`docker cp siyuan-assistant-v0.2.0.zip siyuan:/siyuan/workspace/` 后容器内 unzip，或 `docker exec siyuan mkdir -p /siyuan/workspace/data/plugins/siyuan-assistant && docker exec siyuan unzip ...`

## 构建

```bash
npm install
npm run build
# 产物: package.zip（含 plugin.json / index.js / index.css / i18n / icon.png / preview.png / README）
```

## 已知限制

- 目标版本：思源 v3.6.4+（`minAppVersion: 3.6.4`）
- 文档标题通过 `createDocWithMd` 的 `path` 参数设置（v3.6.4 实测 `renameDoc` 返回 code=0 但静默不生效，已弃用）
- 图片上传使用 `/api/asset/upload`（v3.6.4 实测 `/api/upload` 返回 404），响应按 `data.succMap` 解析，兼容旧版 `data[0].url` 格式
- 导出 Word 为纯前端 docx 库生成，与 Hermes 的 python-docx 精修版（封面/暗标格式/mermaid 渲染等）有差距，投标终稿建议仍走 Hermes 导出
- DOCX 标题层级按 Word 样式名映射（Heading1-9 / 标题1-9 / Title）；Word 自动编号不解析为文本
- 段内换行（`<w:br/>`）转为段落内换行，导入后可能呈现为多行段落
- 下划线、删除线不映射为 Markdown（Lute 不支持标准下划线语法）
- 表格纵向合并（vMerge）在思源普通表格中不支持，导入时合并行以空单元格补齐
- 节级导出的 Markdown 转换是轻量实现，复杂嵌套（如列表套列表）可能简化
- 图片上传端点做了多路径兼容（/api/asset/upload、/api/upload），如遇失败请反馈实际版本

## 致谢

- 模板：[siyuan-note/plugin-sample](https://github.com/siyuan-note/plugin-sample)
- 思路来自 Hermes `siyuan-unified` / `docx-table-import-siyuan` / `siyuan-heading-flatten-workflow` / `siyuan-section-export` 等 skill 的实战沉淀
