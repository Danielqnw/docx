# Changelog

## v0.12.0

### Enhancements (template filling)

- Cross-run text substitution that matches placeholders split across multiple runs while preserving the first matched run's formatting:
  - `Paragraph#substitute(match, replacement, multiline: false)` — `match` accepts `String` or `Regexp` (capture groups supported)
  - `Document#substitute_across_runs(match, replacement, multiline: false)` — applies to every paragraph including table cells
- Multiline replacement: with `multiline: true`, `"\n"` in the replacement becomes a soft line break (`<w:br/>`)
- Insert images at text-only placeholders (no pre-existing image required), for both paragraphs and table cells:
  - `Document#insert_image_at_placeholder(placeholder, source, options = {})`
  - `Document#insert_images_at_placeholder(placeholder, sources = [], options = {})`
  - Auto-registers `word/media`, the relationship in `document.xml.rels`, and the `[Content_Types].xml` default; `[Content_Types].xml` is now loaded and written back on save/stream
- Replace images by placeholder in ordinary paragraphs (not just table cells):
  - `Document#replace_image_by_placeholder(placeholder, source, options = {})` (the existing `*_in_table` method is preserved and now delegates to a shared implementation)
- Checkbox / selection state:
  - `Document#set_checkbox(locator, checked:, checked_glyph:, unchecked_glyph:)` — character-glyph checkboxes (e.g. `☐`→`☑`, or `[ ]`→`[x]`)
  - `Document#check_content_control(tag_or_alias, checked:)` — Word content-control checkboxes (`w:sdt` + `w14:checkbox`), syncs the displayed glyph
- High-level data-driven rendering:
  - `Document#render(text:, images:, checkboxes:, content_controls:, tables:, multiline:, strip_unfilled:, strict:, image_options:)`
- Placeholder cleanup:
  - `Document#strip_unfilled_placeholders(pattern: /\{\{.*?\}\}/)` — cross-run aware, clears leftover placeholders without touching filled content

All new capabilities are additive; existing public APIs (`substitute`, `replace_image*`, table/merge) are unchanged.

## v0.11.0

### Enhancements

- Add logical-grid table APIs for reading and writing merged cells:
  - `Table#cell_at(row, col)` — logical coordinate lookup (any coordinate within a merge returns the top-left anchor cell; out of bounds returns `nil`)
  - `Table#each_cell` — iterate each anchor cell once with logical coordinates
  - `Table#merged?(row, col)` — whether a coordinate belongs to a merged region
  - `Table#merge_cells(row0, col0, row1, col1)` — merge a rectangular range (inclusive); single-cell ranges are a no-op
  - `Table#unmerge_cells(row, col)` — split from the anchor cell; no-op on an unmerged anchor
  - `TableCell#colspan`, `#rowspan`, `#merged?`, `#merge_anchor?`, `#merge_continuation?` — merge read attributes
  - `TableCell#unmerge!` — split via the parent table
- Add `Docx::Errors::InvalidMergeRange`, `MergeConflict`, and `InvalidMergeTarget` for invalid merge/unmerge operations
- Merge and unmerge are round-trip safe on save/reopen; `row_count` and `column_count` are unchanged

### Bug fixes

- Fix `Table#columns` returning misaligned cells in tables with merged cells
- Fix `Table#columns` incorrectly including nested table `w:tc` elements; columns are now built from the logical grid with child-axis positioning
- `columns[col].cells[row]` now equals `cell_at(row, col)`; this may break code that relied on the previous (incorrect) behavior

## v0.7.0

### Enhancements

- Adds to_xml to Document [#116](https://github.com/ruby-docx/docx/pull/116)
- fix runs text not changed after update [#120](https://github.com/ruby-docx/docx/pull/120)

### Bug fixes

- Passing a Nokogiri::XML::Node as the second parameter to Node.new is deprecated [#121](https://github.com/ruby-docx/docx/pull/121)

### Chores

- Add Ruby 3.1 to the CI matrix by petergoldstein [#122](https://github.com/ruby-docx/docx/pull/122)

## v0.6.2

### Bug fixes

- Fix `Docx::Document#to_s` fails when given file has `document22.xml.rels` [#112](https://github.com/ruby-docx/docx/pull/112), [#106](https://github.com/ruby-docx/docx/pull/106)

## v0.6.1

### Bug fixes

- Use `Zip::File#glob` to match any `document.xml` [#104](https://github.com/ruby-docx/docx/pull/104)

### Chores

- Enable Coverall's coverage report [#102](https://github.com/ruby-docx/docx/pull/102)
- Add table write example to README.md [#99](https://github.com/ruby-docx/docx/pull/99)
- Replace Travis CI build with GitHub Action [#98](https://github.com/ruby-docx/docx/pull/98)
- Add ruby 3.0 to versions for testing on Travis CI [#97](https://github.com/ruby-docx/docx/pull/97)

## v0.6.0

### Enhancements

- Added support for hyperlinks (implemented [#70](https://github.com/ruby-docx/docx/pull/70) again) by ollieh-m and gopeter [#92](https://github.com/ruby-docx/docx/pull/92)

### Chores

- Drop ruby 2.4 from supporeted versions by satoryu [#93](https://github.com/ruby-docx/docx/pull/93)
- Refactoring `spec_helper` by satoryu [#90](https://github.com/ruby-docx/docx/pull/90)
- Starts measuring code coverage with coveralls by satoryu [#88](https://github.com/ruby-docx/docx/pull/88)

## v0.5.0

### Enhancements

- Added opening streams and outputting to a stream [#66](https://github.com/ruby-docx/docx/pull/66)
- Added supports for Office 365 files [#85](https://github.com/ruby-docx/docx/pull/85)

### Bug fixes

- `Docx::Document` handles a docx file without styles.xml [#81](https://github.com/ruby-docx/docx/pull/81)
- Fixes insert text before after were switched [#84](https://github.com/ruby-docx/docx/pull/84)

## v0.4.0

### Enhancements

- Implement substitute method on TextRun class. [#75](https://github.com/ruby-docx/docx/pull/75)

### Improvements

- Updates dependencies. [#72](https://github.com/ruby-docx/docx/pull/72), [#77](https://github.com/ruby-docx/docx/pull/77)
- Fix: #paragraphs grabs paragraphs in tables. [#76](https://github.com/ruby-docx/docx/pull/76)
- Updates supported ruby versions. [#78](https://github.com/ruby-docx/docx/pull/78)
