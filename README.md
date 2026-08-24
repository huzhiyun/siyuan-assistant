# SiYuan Assistant

All-in-one plugin for SiYuan, porting frequent Hermes skill operations into the app itself.

## Features

1. **Import DOCX** (command `⇧⌘I`): headings (Heading1-9 / Title), paragraphs, inline bold/italic/sup/sub, tables (gridSpan padding, vMerge skip, empty rows skipped), images (extracted by rId and uploaded to assets). Creates doc via createDocWithMd, sets title via renameDoc. First H1 becomes the doc title by default.
2. **Flatten headings to H5** (command `⇧⌘J`): demote H2~H6 of current doc to H5 (H1 untouched) for easy merge into parent docs.
3. **Set image width** (command `⇧⌘K`): batch set image widths via setBlockAttrs `custom-data-width-percent` — auto (landscape 85% / portrait 50%) or custom percent.
4. **Export section** (command `⇧⌘L`): slice current doc by heading range, convert to Markdown, copy to clipboard. Callouts can be skipped.

All four features are also available from the top-right "SiYuan Assistant" menu.

## Install

A local SiYuan plugin = a folder inside the workspace `data/plugins/` (the marketplace UI has **no** "install local zip" option):

1. Get `siyuan-assistant-v0.2.0.zip` (NAS: `/hermes/输出/`)
2. On the machine running SiYuan (workspace e.g. `/siyuan/workspace/`):
   ```bash
   mkdir -p /siyuan/workspace/data/plugins/siyuan-assistant
   unzip siyuan-assistant-v0.2.0.zip -d /siyuan/workspace/data/plugins/siyuan-assistant/
   ```
   (the zip contains plugin.json/index.js directly; the folder name must equal the `name` field in plugin.json)
3. Restart SiYuan (or refresh), then Settings → Marketplace → Downloaded → enable "SiYuan Assistant"

Docker: `docker cp siyuan-assistant-v0.2.0.zip siyuan:/siyuan/workspace/` then unzip inside the container.

## Build

```bash
npm install
npm run build
# Output: package.zip
```

## Known limitations

- Target: SiYuan v3.6.4+ (`minAppVersion: 3.6.4`)
- Doc title is set via the `path` parameter of `createDocWithMd` (verified on v3.6.4: `renameDoc` returns code=0 but silently does nothing, so it is not used)
- Image upload uses `/api/asset/upload` (verified on v3.6.4: `/api/upload` returns 404), parsing `data.succMap`, with fallback to legacy `data[0].url` format
- DOCX heading level mapped from Word style names (Heading1-9 / 标题1-9 / Title); Word auto-numbering is not extracted as text
- Line breaks (`<w:br/>`) become paragraph line breaks after import
- Underline / strikethrough not mapped to Markdown
- Table vMerge rows are padded with empty cells (SiYuan plain tables don't support rowspan)
- Section export Markdown conversion is lightweight; deep nesting may be simplified
- Image upload tries multiple endpoints (/api/upload, /api/asset/upload) for compatibility

## Credits

- Template: [siyuan-note/plugin-sample](https://github.com/siyuan-note/plugin-sample)
- Logic distilled from Hermes skills: siyuan-unified, docx-table-import-siyuan, siyuan-heading-flatten-workflow, siyuan-section-export
