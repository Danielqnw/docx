# docx

[![Gem Version](https://badge.fury.io/rb/docx.svg)](https://badge.fury.io/rb/docx)
[![Ruby](https://github.com/ruby-docx/docx/workflows/Ruby/badge.svg)](https://github.com/ruby-docx/docx/actions?query=workflow%3ARuby)
[![Coverage Status](https://coveralls.io/repos/github/ruby-docx/docx/badge.svg?branch=master)](https://coveralls.io/github/ruby-docx/docx?branch=master)
[![Gitter](https://badges.gitter.im/ruby-docx/community.svg)](https://gitter.im/ruby-docx/community?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)

**English** | [简体中文](README.zh-CN.md)

> A Ruby library/gem for reading and writing Microsoft Word `.docx` files.

It lets you work with a document's content (paragraphs, bookmarks, tables, images, styles) through a friendly object model instead of **hand-editing the underlying Office Open XML**.

For the full API reference, option tables, and more examples, see the **[docs site](docs/index.html)** (Chinese / English toggle).

## Features

| Capability | What you can do |
| --- | --- |
| 📖 Read content | Iterate paragraphs and bookmarks, render paragraphs to HTML |
| 📂 Open from anywhere | Open from a file path or from an in-memory buffer / IO object |
| 📊 Tables | Read rows / columns / cells, copy rows, substitute placeholder text |
| 🔗 Cell merging | Merge / unmerge rectangular regions on a logical grid, with safe `gridSpan` / `vMerge` handling |
| 🖼️ Image replacement | Replace by relationship id, archive path, or placeholder text, including batch replacement in table cells |
| ✏️ Text substitution | Replace text while preserving formatting, with optional regex captures |
| 🧩 Template filling | Cross-run substitution, insert images at text-only placeholders, checkboxes, multiline breaks, and a high-level `render` entry |
| 📎 Cross-document import | Import body nodes from another doc while isolating style / numbering / media / bookmark ids |
| 🎨 Styles | Add, modify, and remove paragraph / character styles |
| 🔧 Low-level access | Reach the underlying `Nokogiri` nodes when you need finer control |

## Prerequisites

- Ruby 2.6 or later

## Installation

Add the following line to your application's Gemfile:

```ruby
gem 'docx'
```

Then run `bundle install`. Or install it yourself directly:

```shell
gem install docx
```

## Quick Start

```ruby
require 'docx'

doc = Docx::Document.open('example.docx')

doc.paragraphs.each { |p| puts p.to_s }

doc.paragraphs.each do |p|
  p.each_text_run { |run| run.substitute('{{name}}', 'Alice') }
end
doc.save('example-edited.docx')
```

> [!NOTE]
> Examples below assume `require 'docx'`. More reading patterns (buffers, HTML rendering, etc.) are in [Docs · Reading](docs/index.html#reading).

## Common recipes

### Tables: copy a row and fill values

```ruby
doc = Docx::Document.open('tables.docx')
table = doc.tables[0]
last_row = table.rows.last

new_row = last_row.copy
new_row.insert_before(last_row)
new_row.cells.each do |cell|
  cell.paragraphs.each do |paragraph|
    paragraph.each_text_run { |text| text.substitute('_placeholder_', 'replacement') }
  end
end

doc.save('tables-edited.docx')
```

Logical-grid merge / unmerge (`cell_at`, `merge_cells`, `unmerge_cells`) is covered in [Docs · Tables](docs/index.html#tables).

### Images: replace by relationship or placeholder

```ruby
doc = Docx::Document.open('with-images.docx')

doc.replace_image('rId5', 'replacement.png')
doc.replace_image_by_placeholder_in_table('{{photo_a}}', 'replacement.png', fit: :cover)

doc.save('with-images-edited.docx')
```

Options such as `fit` / `width` / `height` and batch insertion are in [Docs · Images](docs/index.html#images).

### Template filling: one `render` call

Fill a `.docx` template with runtime data. Keys are placeholder names; no host/business semantics are assumed.

```ruby
doc = Docx::Document.open('template.docx')

doc.render(
  text:             { '{{name}}' => 'Alice', '{{bio}}' => "line 1\nline 2" },
  images:           { '{{photo}}' => 'avatar.png' },
  checkboxes:       { 'Option A' => true },
  content_controls: { 'opt_a' => true },
  tables:           [{ placeholder_row: '{{row}}',
                       rows: [ { '{{city}}' => 'Paris' }, { '{{city}}' => 'Tokyo' } ] }],
  multiline:      true,
  strip_unfilled: true
)

doc.save('report.docx')
```

Split APIs (cross-run substitution, insert-at-placeholder, checkboxes, etc.) are in [Docs · Template filling](docs/index.html#template-filling).

### Cross-document import: move body nodes with isolation

A plain `dup` only copies `document.xml`, so style / numbering / image / bookmark ids break. The import APIs remap those ids and pull over the definitions they depend on:

```ruby
target = Docx::Document.open('target.docx')
source = Docx::Document.open('source.docx')

importer = Docx::Merge::Importer.new(target, source)
anchor = target.paragraphs.last.node

source.tables.each do |table|
  imported = importer.import(table.node)
  anchor.add_previous_sibling(imported)
end

target.save('merged.docx')
```

Convenience helpers: `import_before` / `import_after` / `import_node`. Isolation rules and `fallback:` are in [Docs · Cross-document import](docs/index.html#cross-document-import).

### Styles: define once, apply to paragraphs

```ruby
doc = Docx::Document.open('example.docx')

style = doc.styles_configuration.add_style('Red', name: 'Red', font_color: 'FF0000', font_size: 20)
style.bold = true
doc.paragraphs.each { |p| p.style = 'Red' }

doc.save('styled.docx')
```

The full style attribute list is in [Docs · Styles](docs/index.html#styles).

## Advanced

When the wrapped API is not enough, drop down to `#node` / `#xpath` on the underlying `Nokogiri` nodes — see [Docs · Advanced](docs/index.html#advanced).

Error types under `Docx::Errors::*` are listed in [Docs · Error reference](docs/index.html#error-reference).

## Development

```shell
bundle install
bundle exec rspec
```

### TODO

- Calculate element formatting based on values present in element properties as well as properties inherited from parents
- Inherit default formatting when new elements are inserted
- Implement formattable elements
- More convenient methods for inserting multi-line text at a bookmark (inserting paragraph nodes after the bookmark's paragraph)

## Contributing

Bug reports and pull requests are welcome on GitHub at <https://github.com/ruby-docx/docx>. Please make sure the test suite passes (`bundle exec rspec`) first.

## License

This gem is available as open source under the terms in [LICENSE.md](LICENSE.md).
