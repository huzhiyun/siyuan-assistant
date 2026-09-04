# SiYuan Assistant

An all-in-one plugin for [SiYuan Note](https://github.com/siyuan-note/siyuan)
that brings the most-used DOCX import / export operations and a few quality-of-
life document tools into a single top-bar menu.

> 思源笔记的本地插件：在思源里完成 DOCX 导入、标题降级、图片宽度、节级导出
> 与 Word 导出。详细中文说明见 `README.zh-CN.md`。

## Features

| Command | Shortcut | What it does |
| --- | --- | --- |
| Import DOCX | `⇧⌘I` | Parse headings (Heading 1–9 / Title), paragraphs, inline bold / italic / underline / sup / sub, tables (gridSpan padding, vMerge skip, empty-row drop), and images (rId-based extraction → upload to `assets`). Create the document via `createDocWithMd` and set the title via `renameDoc`. The first H1 becomes the document title by default. |
| Flatten headings | `⇧⌘J` | Demote the current document's H2–H6 to H5 (H1 is left alone) so a sub-document can be merged into a parent without disturbing the parent's outline. |
| Set image width | `⇧⌘K` | Batch-set `custom-data-width-percent` on every image — auto (landscape 85% / portrait 50%) or a custom percent. |
| Export section | `⇧⌘L` | Slice the current document by heading range, convert to Markdown, copy to clipboard. Callout blocks can be skipped. |
| Export Word | (menu) | Convert the current document (or a manually-selected docid) to `.docx` in the browser — no server, no `python-docx`, no pandoc. Native multi-level numbering, fixed-width tables for Word/WPS/mac Quick Look, centered image/table blocks, Chinese typography defaults. |

All five features are also available from the top-right **SiYuan Assistant**
menu.

## Requirements

- SiYuan `>= 3.6.4` (the `minAppVersion` declared in `plugin.json`).
- Desktop, mobile, or browser frontends — see `frontends` in `plugin.json`.

## Install

A local SiYuan plugin is a folder inside the workspace `data/plugins/`. The
marketplace UI does **not** have a "install local zip" option.

1. Download `siyuan-assistant-v0.3.13.zip` from the [Releases](../../releases)
   page (or build your own with `npm run build`).
2. On the machine running SiYuan (workspace e.g. `/siyuan/workspace/`):

   ```bash
   mkdir -p /siyuan/workspace/data/plugins/siyuan-assistant
   unzip siyuan-assistant-v0.3.13.zip -d /siyuan/workspace/data/plugins/siyuan-assistant/
   ```

   The zip contains `plugin.json`, `index.js`, etc. directly — the folder
   name **must** equal the `name` field in `plugin.json`.

3. Restart SiYuan (or refresh), then **Settings → Marketplace → Downloaded →
   enable "SiYuan Assistant"**.

**Docker**:

```bash
docker cp siyuan-assistant-v0.3.13.zip siyuan:/siyuan/workspace/
docker exec siyuan mkdir -p /siyuan/workspace/data/plugins/siyuan-assistant
docker exec siyuan unzip -o /siyuan/workspace/siyuan-assistant-v0.3.13.zip \
  -d /siyuan/workspace/data/plugins/siyuan-assistant/
```

## Build from source

```bash
npm install
npm run typecheck   # tsc --noEmit
npm test            # esbuild bundles + node test runners (69 cases)
npm run build       # webpack production build → dist/ + package.zip
```

`npm run build` produces `package.zip` (≈850 KB) which is the file you deploy
into `data/plugins/siyuan-assistant/`.

## Project layout

```
src/
  index.ts          # plugin entry: top-bar menu, command registration
  docx.ts           # DOCX parser (OOXML → Markdown-ish blocks)
  docxgen.ts        # pure-frontend .docx generator
  blocks.ts         # shared block model + SiYuan getDoc HTML → blocks
  image-layout.ts   # image rotation / hash reuse for DOCX import
test/
  parse.test.mjs    # DOCX parser regressions
  blocks.test.mjs   # block model regressions
  image-layout.test.mjs
  docxgen.test.mjs
webpack.config.js   # production build
plugin.json         # SiYuan plugin manifest
```

## License

MIT — see `LICENSE`.
