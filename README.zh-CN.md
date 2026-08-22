# 思源助手 (SiYuan Assistant)

思源笔记一体化工具插件，把 Hermes 思源 skill 的高频操作搬进思源内一键完成，无需再开 Hermes、连内网、处理 token/代理。

## 功能

1. **DOCX 导入**（命令 `⇧⌘I`）
   - 解析 Word 文档：标题（Heading1-9 / Title 样式）、段落、内联格式（加粗/斜体/上标/下标）、表格（gridSpan 自动补空、vMerge 合并行跳过、全空行跳过）、图片（按 rId 提取并上传到 assets）
   - 图片先上传、再 createDocWithMd 建文档、renameDoc 设置文档标题（标题与正文 H1 独立）
   - 默认把第一个 H1 当文档标题、正文不含 H1（可勾选保留）

2. **标题降级为 H5**（命令 `⇧⌘J`）
   - 当前文档 H2~H6 全部降级为 H5（H1 文档标题不动），方便合入父文档

3. **图片宽度设置**（命令 `⇧⌘K`）
   - 批量设置当前文档所有图片宽度：自动模式按横竖比例（横图 85% / 竖图 50%），或自定义百分比
   - 通过 `setBlockAttrs` 的 `custom-data-width-percent` 实现

4. **节级导出**（命令 `⇧⌘L`）
   - 按标题选择起始/结束节，切片后转 Markdown 复制到剪贴板
   - 可跳过提示块（callout / 引用标注），适合合入投标文档

顶栏右侧也有「思源助手」菜单，四个功能都有入口。

## 安装

1. 构建或获取 `package.zip`
2. 思源 → 设置 → 集市 → 手动安装插件（或解压到 `工作空间/data/plugins/siyuan-assistant/`）
3. 在集市「已下载」中启用

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
- DOCX 标题层级按 Word 样式名映射（Heading1-9 / 标题1-9 / Title）；Word 自动编号不解析为文本
- 段内换行（`<w:br/>`）转为段落内换行，导入后可能呈现为多行段落
- 下划线、删除线不映射为 Markdown（Lute 不支持标准下划线语法）
- 表格纵向合并（vMerge）在思源普通表格中不支持，导入时合并行以空单元格补齐
- 节级导出的 Markdown 转换是轻量实现，复杂嵌套（如列表套列表）可能简化
- 图片上传端点做了多路径兼容（/api/upload、/api/asset/upload），如遇失败请反馈实际版本

## 致谢

- 模板：[siyuan-note/plugin-sample](https://github.com/siyuan-note/plugin-sample)
- 思路来自 Hermes `siyuan-unified` / `docx-table-import-siyuan` / `siyuan-heading-flatten-workflow` / `siyuan-section-export` 等 skill 的实战沉淀
