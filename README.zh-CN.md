# 思源助手

[思源笔记](https://github.com/siyuan-note/siyuan) 本地插件：把日常最常用的
DOCX 导入/导出、文档整理动作，集中到顶栏的「思源助手」菜单里。

> English summary: see `README.md`.

## 功能

| 命令 | 快捷键 | 作用 |
| --- | --- | --- |
| 导入 DOCX | `⇧⌘I` | 解析标题（Heading 1–9 / Title）、段落、内联粗体/斜体/下划线/上标/下标、表格（gridSpan 补空、跳过 vMerge 合并行、跳过全空行）和图片（按 rId 提取后上传到 `assets`）。通过 `createDocWithMd` 建文档，`renameDoc` 设标题；默认以第一个 H1 作为文档标题。 |
| 标题降级 | `⇧⌘J` | 把当前文档的 H2–H6 统一降为 H5（H1 不动），方便子文档合入父文档时不影响父文档大纲。 |
| 设置图片宽度 | `⇧⌘K` | 批量为图片块写 `custom-data-width-percent` —— 自动（横图 85% / 竖图 50%）或自定义百分比。 |
| 导出节 | `⇧⌘L` | 按标题范围切片当前文档，转为 Markdown 复制到剪贴板，可选跳过 callout。 |
| 导出 Word | （菜单） | 在浏览器内把当前文档（或手动输入的 docid）导出为 `.docx`，无需服务端、无需 `python-docx`、无需 pandoc。Word 原生多级编号、表格固定宽度（兼容 Word/WPS/mac 预览）、图片/表格居中、中文排版默认。 |

五个功能在顶栏 **思源助手** 菜单里同样可用。

## 环境要求

- 思源 `>= 3.6.4`（`plugin.json` 中的 `minAppVersion`）。
- 桌面端、移动端、浏览器端均可，见 `plugin.json` 的 `frontends` 字段。

## 安装

思源本地插件 = 把插件目录放进工作空间 `data/plugins/`，**集市界面没有**「手动
安装 zip」的选项。

1. 从 [Releases](../../releases) 页下载 `siyuan-assistant-v0.3.13.zip`（或
   用 `npm run build` 自己构建）。
2. 在思源所在机器上（工作空间如 `/siyuan/workspace/`）：

   ```bash
   mkdir -p /siyuan/workspace/data/plugins/siyuan-assistant
   unzip siyuan-assistant-v0.3.13.zip -d /siyuan/workspace/data/plugins/siyuan-assistant/
   ```

   zip 内是 `plugin.json` / `index.js` 等直接文件；**目录名必须等于**
   `plugin.json` 里的 `name` 字段。

3. 重启思源（或刷新），然后 **设置 → 集市 → 已下载 → 启用「思源助手」**。

**Docker**：

```bash
docker cp siyuan-assistant-v0.3.13.zip siyuan:/siyuan/workspace/
docker exec siyuan mkdir -p /siyuan/workspace/data/plugins/siyuan-assistant
docker exec siyuan unzip -o /siyuan/workspace/siyuan-assistant-v0.3.13.zip \
  -d /siyuan/workspace/data/plugins/siyuan-assistant/
```

## 从源码构建

```bash
npm install
npm run typecheck   # tsc --noEmit
npm test            # esbuild 打包 + node 测试运行器（69 用例）
npm run build       # webpack 生产构建 → dist/ + package.zip
```

`npm run build` 产出 `package.zip`（约 850 KB），把它部署到
`data/plugins/siyuan-assistant/` 即可。

## 目录结构

```
src/
  index.ts          # 插件入口：顶栏菜单与命令注册
  docx.ts           # DOCX 解析（OOXML → 类 Markdown 块）
  docxgen.ts        # 纯前端 .docx 生成器
  blocks.ts         # 共享块模型 + 思源 getDoc HTML → 块
  image-layout.ts   # DOCX 导入的图片旋转 / SHA-256 复用
test/
  parse.test.mjs    # DOCX 解析回归
  blocks.test.mjs   # 块模型回归
  image-layout.test.mjs
  docxgen.test.mjs
webpack.config.js   # 生产构建
plugin.json         # 思源插件清单
```

## 许可

MIT —— 见 `LICENSE`。
