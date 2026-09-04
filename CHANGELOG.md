# Changelog

## v0.3.13 (2026-09-04)

First public release of the SiYuan Assistant plugin after the v0.1.0–v0.3.13
internal cycle. Highlights of the cycle:

- **DOCX import** with correct rId-based image extraction, gridSpan/vMerge
  table handling, outline / page-break / list recovery, and robust handling of
  the `w:val="0"` bold/italic disabler.
- **DOCX import hardening** — bounded 4-way concurrent uploads, 30s per-image
  timeout with retry, single-image failure no longer blocks the import, and a
  progress indicator.
- **Image fidelity** — Word `a:xfrm/@rot` is rasterized at upload time and
  width/height are swapped for 90°/270° images; SHA-256 reuse avoids duplicate
  uploads on re-import; `·` markers become Markdown list items.
- **Inline formatting** — adjacent `****` runs are merged so SiYuan does not
  display raw Markdown, and `<u>` is not allowed to bleed bold into the
  following block.
- **Heading flatten (⇧⌘J)** — demote H2–H6 to H5 for clean merge into a parent
  document.
- **Image width (⇧⌘K)** — batch set `custom-data-width-percent` (auto
  landscape 85% / portrait 50%, or custom percent).
- **Section export (⇧⌘L)** — slice the current document by heading range and
  copy Markdown to the clipboard, with optional callout skipping.
- **Word export** — pure-frontend docx generation with native multi-level
  numbering, fixed-width tables for Word/WPS/mac Quick Look compatibility,
  centered image/table attributes, and Chinese typography defaults.

All four primary commands are also available from the top-bar "SiYuan
Assistant" menu. See `README.md` / `README.zh-CN.md` for usage and the test
suite (`npm test`) for the regression coverage that ships with each fix.

## Notes for the first public tag

- This is the first tag — the `siyuan-note/plugin-sample` history that the
  repository was forked from has been squashed out of the changelog.
- The v0.3.13 plugin package (`package.zip`) is attached to the GitHub
  release of the same name.
- License: MIT. The original template's LICENSE copyright has been updated
  to the plugin author.
