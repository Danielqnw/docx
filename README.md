# docx

[![Gem Version](https://badge.fury.io/rb/docx.svg)](https://badge.fury.io/rb/docx)
[![Ruby](https://github.com/ruby-docx/docx/workflows/Ruby/badge.svg)](https://github.com/ruby-docx/docx/actions?query=workflow%3ARuby)
[![Coverage Status](https://coveralls.io/repos/github/ruby-docx/docx/badge.svg?branch=master)](https://coveralls.io/github/ruby-docx/docx?branch=master)
[![Gitter](https://badges.gitter.im/ruby-docx/community.svg)](https://gitter.im/ruby-docx/community?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)

**English** | [简体中文](README.zh-CN.md)

> A Ruby library/gem for reading and writing Microsoft Word `.docx` files.

It lets you work with a document's content (paragraphs, bookmarks, tables, images, styles) through a friendly object model instead of **hand-editing the underlying Office Open XML**.

## Features

| Capability | What you can do |
| --- | --- |
| 📖 Read content | Iterate paragraphs and bookmarks, render paragraphs to HTML |
| 📂 Open from anywhere | Open from a file path or from an in-memory buffer / IO object |
| 📊 Tables | Read rows / columns / cells, copy rows, substitute placeholder text |
| 🔗 Cell merging | Merge / unmerge rectangular regions on a logical grid, with safe `gridSpan` / `vMerge` handling |
| 🖼️ Image replacement | Replace by relationship id, archive path, or placeholder text, including batch replacement in table cells |
| ✏️ Text substitution | Replace text while preserving formatting, with optional regex captures |
| 🎨 Styles | Add, modify, and remove paragraph / character styles |
| 🔧 Low-level access | Reach the underlying `Nokogiri` nodes when you need finer control |

## Table of Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Reading](#reading)
  - [Paragraphs and bookmarks](#paragraphs-and-bookmarks)
  - [Opening from a buffer](#opening-from-a-buffer)
  - [Rendering HTML](#rendering-html)
- [Tables](#tables)
  - [Reading tables](#reading-tables)
  - [Writing to tables](#writing-to-tables)
  - [Merging and unmerging cells](#merging-and-unmerging-cells)
- [Images](#images)
- [Writing and substituting text](#writing-and-substituting-text)
- [Styles](#styles)
  - [Style attributes](#style-attributes)
- [Advanced: raw node access](#advanced-raw-node-access)
- [Error reference](#error-reference)
- [Development](#development)
- [Contributing](#contributing)
- [License](#license)

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

The snippet below shows the common "open → read → edit → save" round trip:

```ruby
require 'docx'

# Open an existing document
doc = Docx::Document.open('example.docx')

# Read every paragraph
doc.paragraphs.each { |p| puts p.to_s }

# Edit and save under a new name
doc.paragraphs.each do |p|
  p.each_text_run { |run| run.substitute('{{name}}', 'Alice') }
end
doc.save('example-edited.docx')
```

> [!NOTE]
> Every example below assumes you have already required the gem with `require 'docx'`.

## Reading

### Paragraphs and bookmarks

Once a document is open, you can iterate its paragraphs and bookmarks separately:

```ruby
# Create a Docx::Document object for an existing docx file
doc = Docx::Document.open('example.docx')

# Retrieve and display paragraphs
doc.paragraphs.each do |p|
  puts p
end

# Retrieve and display bookmarks, returned as a hash keyed by bookmark name
doc.bookmarks.each_pair do |bookmark_name, bookmark_object|
  puts bookmark_name
end
```

A paragraph responds to `to_s` (plain text) and `to_html`, and exposes its text runs via `each_text_run`.

### Opening from a buffer

You don't need a file on disk — an in-memory buffer or any IO-like object works too. This is handy for documents fetched over HTTP or received as web uploads:

```ruby
# Create a Docx::Document object from a remote file / StringIO / uploaded file
doc = Docx::Document.open(buffer)

# Everything about reading is the same as shown above
```

### Rendering HTML

Convert paragraphs to HTML strings for display on the web:

```ruby
doc = Docx::Document.open('example.docx')
doc.paragraphs.each do |p|
  puts p.to_html
end
```

## Tables

### Reading tables

Use `doc.tables` to get every table, then address cells by row or by column:

```ruby
doc = Docx::Document.open('tables.docx')

first_table = doc.tables[0]
puts first_table.row_count
puts first_table.column_count
puts first_table.rows[0].cells[0].text
puts first_table.columns[0].cells[0].text

# Iterate through every table
doc.tables.each do |table|
  table.rows.each do |row| # Row-based iteration
    row.cells.each do |cell|
      puts cell.text
    end
  end

  table.columns.each do |column| # Column-based iteration
    column.cells.each do |cell|
      puts cell.text
    end
  end
end
```

There are two ways to address cells; understanding the difference matters when working with merged cells:

| Approach | Syntax | Behaviour |
| --- | --- | --- |
| **Physical access** | `table.rows[i].cells[j]` | Follows the actual `w:tc` elements in each row; a row with merged cells has fewer physical cells |
| **Logical access** | `table.cell_at(row, col)` | Follows the *logical grid*; any coordinate inside a merged region returns that region's anchor cell, so the grid stays rectangular |

### Writing to tables

The example below copies a table's last row, inserts it as a new row, and substitutes its placeholder text:

```ruby
doc = Docx::Document.open('tables.docx')

# Iterate over each table
doc.tables.each do |table|
  last_row = table.rows.last

  # Copy the last row and insert a new one before it
  new_row = last_row.copy
  new_row.insert_before(last_row)

  # Substitute text in each cell of the new row
  new_row.cells.each do |cell|
    cell.paragraphs.each do |paragraph|
      paragraph.each_text_run do |text|
        text.substitute('_placeholder_', 'replacement value')
      end
    end
  end
end

doc.save('tables-edited.docx')
```

### Merging and unmerging cells

Cells are merged using **logical grid coordinates** (zero-based `row` / `col`). The content of a merged region is kept in its top-left *anchor* cell.

First, how to inspect a cell and its merge state:

```ruby
doc = Docx::Document.open('tables.docx')
table = doc.tables[0]

# Inspect a cell by logical coordinate
cell = table.cell_at(1, 2)
puts cell.text if cell

# Is this coordinate part of a merged region?
puts table.merged?(0, 0)

# Walk the logical grid, including merge metadata
table.each_cell do |cell, row, col|
  puts "#{row},#{col}: #{cell.text} (colspan=#{cell.colspan}, rowspan=#{cell.rowspan})"
  puts "  anchor: #{cell.merge_anchor?}, continuation: #{cell.merge_continuation?}"
end
```

Then, how to merge and unmerge a rectangular region (bounds are inclusive):

```ruby
# Merge the rectangle from (row0, col0) to (row1, col1); content stays in the top-left anchor
table.merge_cells(0, 0, 1, 1)

# Unmerge from the anchor coordinate (a no-op if the anchor is not merged)
table.unmerge_cells(0, 0)

# You can also unmerge directly from the anchor cell
anchor = table.cell_at(0, 0)
anchor.unmerge! if anchor&.merge_anchor?

doc.save('merged.docx')
```

> [!IMPORTANT]
> Merging is purely a presentation operation: it does **not** change `row_count` or `column_count`, because the table's `tblGrid` is left untouched.

Useful cell predicates and accessors:

| Method | Description |
| --- | --- |
| `cell.colspan` | Number of grid columns the cell spans |
| `cell.rowspan` | Number of grid rows the cell spans |
| `cell.merge_anchor?` | `true` if the cell is the top-left anchor of a merged region |
| `cell.merge_continuation?` | `true` if the cell is covered by another cell's merge |
| `cell.unmerge!` | Unmerge the region anchored at this cell |

Errors raised while merging / unmerging:

| Error | Raised when |
| --- | --- |
| `Docx::Errors::InvalidMergeRange` | Coordinates are out of bounds, or `row0 > row1` / `col0 > col1` |
| `Docx::Errors::MergeConflict` | The requested range overlaps an existing merge |
| `Docx::Errors::InvalidMergeTarget` | Unmerge was called on a non-anchor or non-merged cell |

#### Example: merge the rightmost three columns of every row

A real-world scenario — for one table, starting at the second row, merge the rightmost three columns of each row into a single cell:

```ruby
doc = Docx::Document.open('report.docx')
table = doc.tables[1]

# Logical column indices 5, 6, 7 are the rightmost three columns here
(1...table.row_count).each do |row|
  table.merge_cells(row, 5, row, 7)
end

doc.save('report-merged.docx')
```

## Images

`doc.images` lists image relationship IDs mapped to archive paths, after which you can replace images in several ways:

```ruby
doc = Docx::Document.open('with-images.docx')

# Inspect image relationship IDs mapped to archive paths
# => { "rId5" => "word/media/image1.png", ... }
pp doc.images

# Replace by relationship ID using a file path
doc.replace_image('rId5', 'replacement.png')

# Replace by entry path using an IO object
File.open('replacement.png', 'rb') do |io|
  doc.replace_image('word/media/image1.png', io)
end
```

If image slots are marked with placeholder text (e.g. `{{photo_a}}`) inside table cells, you can replace by placeholder and control output size and fit:

```ruby
# Replace by placeholder text in a table cell
doc.replace_image_by_placeholder_in_table('{{photo_a}}', 'replacement.png', fit: :cover)

# Specify output size explicitly (5 cm x 3 cm)
doc.replace_image_by_placeholder_in_table(
  '{{photo_a}}', 'replacement.png',
  width: 5.0, height: 3.0, fit: :cover
)

# Only width given — height is auto-calculated from the source image's aspect ratio
doc.replace_image_by_placeholder_in_table(
  '{{photo_a}}', 'replacement.png',
  width: 5.0
)
```

You can also place **multiple** images via a single placeholder:

```ruby
# Default max_images_per_row is 2 in the same cell; overflows append duplicated rows.
# width / height also apply to every image slot.
doc.replace_images_by_placeholder_in_table(
  '{{photo_a}}',
  ['image-a.png', 'image-b.png', 'image-c.png'],
  width: 5.0, height: 3.0, fit: :cover
)

doc.save('with-images-edited.docx')
```

Image replacement options:

| Option | Applies to | Description |
| --- | --- | --- |
| `fit:` | placeholder replacements | `:stretch` (default), `:cover`, or `:contain`. Controls how the new image fills the target box |
| `width:` | placeholder replacements | Output width in centimetres. If only `width` is given, the height is derived from the source aspect ratio |
| `height:` | placeholder replacements | Output height in centimetres. If only `height` is given, the width is derived from the source aspect ratio |
| `max_images_per_row:` | `replace_images_by_placeholder_in_table` | Maximum images placed in one cell before a duplicated row is appended (default `2`) |

Image-related errors:

| Error | Raised when |
| --- | --- |
| `Docx::Errors::ImageNotFound` | No image matches the given relationship id / archive path |
| `Docx::Errors::ImagePlaceholderNotFound` | The placeholder text was not found in any table cell |

## Writing and substituting text

The following demonstrates bookmark insertion, paragraph removal, and two substitution styles (plain and regex-capture based):

```ruby
doc = Docx::Document.open('example.docx')

# Insert a single line of text after one of our bookmarks
doc.bookmarks['example_bookmark'].insert_text_after('Hello world.')

# Insert multiple lines of text at a bookmark
doc.bookmarks['example_bookmark_2'].insert_multiple_lines_after(['Hello', 'World', 'foo'])

# Remove paragraphs
doc.paragraphs.each do |p|
  p.remove! if p.to_s =~ /TODO/
end

# Substitute text, preserving formatting
doc.paragraphs.each do |p|
  p.each_text_run do |tr|
    tr.substitute('_placeholder_', 'replacement value')
  end
end

# Substitute text with access to captures. The block arg is a MatchData,
# which behaves a bit differently than String#gsub.
# https://ruby-doc.org/3.3.7/MatchData.html
doc.paragraphs.each do |p|
  p.each_text_run do |tr|
    tr.substitute_with_block(/total: (\d+)/) { |match_data| "total: #{match_data[1].to_i * 10}" }
  end
end

# Save document to the specified path
doc.save('example-edited.docx')
```

## Styles

Through `styles_configuration` you can read, modify, add, and remove styles, then apply a style to paragraphs:

```ruby
d = Docx::Document.open('example.docx')

# Modify an existing style
existing_style = d.styles_configuration.style_of('Heading 1')
existing_style.font_color = '000000'

# Add a new style (attributes listed below)
new_style = d.styles_configuration.add_style('Red', name: 'Red', font_color: 'FF0000', font_size: 20)
new_style.bold = true

# Apply styles to paragraphs
d.paragraphs.each { |p| p.style = 'Red' }
d.paragraphs.each { |p| p.style = 'Heading 1' }

# Remove a style
d.styles_configuration.remove_style('Red')
```

### Style attributes

The attributes are grouped by category below. Attributes marked **bool** accept `true` / `false`; those marked **color** accept hex color codes (e.g. `FF0000`).

**Basics**

| Attribute | Description |
| --- | --- |
| `id` | The unique identifier of the style (required) |
| `name` | The human-readable name of the style (required) |
| `type` | The type of the style (e.g. paragraph, character) |

**Paragraph formatting**

| Attribute | Type | Description |
| --- | --- | --- |
| `keep_next` | bool | Keep a paragraph and the next one on the same page |
| `keep_lines` | bool | Keep all lines of a paragraph together on one page |
| `page_break_before` | bool | Insert a page break before the paragraph |
| `widow_control` | bool | Control widow and orphan lines in a paragraph |
| `suppress_auto_hyphens` | bool | Control automatic hyphenation |
| `bidirectional_text` | bool | Whether the paragraph contains bidirectional text |
| `spacing_before` | — | Spacing before a paragraph |
| `spacing_after` | — | Spacing after a paragraph |
| `line_spacing` | — | Line spacing of a paragraph |
| `line_rule` | — | How line spacing is calculated |
| `indent_left` | — | Left indentation of a paragraph |
| `indent_right` | — | Right indentation of a paragraph |
| `indent_first_line` | — | First-line indentation of a paragraph |
| `align` | — | Text alignment within a paragraph |
| `outline_level` | — | Outline level in the document's hierarchy |

**Fonts and characters**

| Attribute | Type | Description |
| --- | --- | --- |
| `font` | — | Set the font for all scripts (ASCII, complex script, East Asian, etc.) |
| `font_ascii` | — | Font for ASCII characters |
| `font_cs` | — | Font for complex script characters |
| `font_hAnsi` | — | Font for high ANSI characters |
| `font_eastAsia` | — | Font for East Asian characters |
| `font_color` | color | Text color |
| `font_size` | — | Font size |
| `font_size_cs` | — | Font size for complex script characters |
| `bold` | bool | Bold formatting |
| `italic` | bool | Italic formatting |
| `caps` | bool | All capitals |
| `small_caps` | bool | Small capital letters |
| `strike` | bool | Strikethrough |
| `double_strike` | bool | Double strikethrough |
| `outline` | bool | Outline effect |
| `underline_style` | — | Underline style |
| `underline_color` | color | Underline color |
| `spacing` | — | Character spacing |
| `kerning` | — | Space between characters |
| `position` | — | Character position (superscript / subscript) |
| `text_fill_color` | color | Fill color of text |
| `vertical_alignment` | — | Vertical alignment of text within a line |
| `lang` | — | Language tag for the text |

**Shading**

| Attribute | Type | Description |
| --- | --- | --- |
| `shading_style` | — | Shading pattern style |
| `shading_color` | color | Color of the shading pattern |
| `shading_fill` | — | Background fill color of shading |

## Advanced: raw node access

When the wrapped API isn't enough, you can drop down to the underlying `Nokogiri` nodes:

```ruby
d = Docx::Document.open('example.docx')

# The Nokogiri::XML::Node on which an element is based can be accessed using #node
d.paragraphs.each do |p|
  puts p.node.inspect
end

# The #xpath and #at_xpath methods are delegated to the node from the element, saving a step
p_element = d.paragraphs.first
p_children = p_element.xpath('//child::*') # selects all children
p_child = p_element.at_xpath('//child::*') # selects the first child
```

## Error reference

All errors live under the `Docx::Errors` namespace.

| Error | Raised when |
| --- | --- |
| `StyleNotFound` | A requested style name does not exist |
| `StyleInvalidPropertyValue` | A style property is set to an invalid value |
| `StyleRequiredPropertyValue` | A required style property is missing |
| `InvalidMergeRange` | Merge coordinates are out of bounds or reversed |
| `MergeConflict` | A merge range overlaps an existing merge |
| `InvalidMergeTarget` | Unmerge was called on a non-anchor or non-merged cell |
| `ImageNotFound` | No image matches the given relationship id / archive path |
| `ImagePlaceholderNotFound` | The image placeholder text was not found |

## Development

After checking out the repo, install dependencies and run the test suite:

```shell
bundle install
bundle exec rspec
```

### todo

- Calculate element formatting based on values present in element properties as well as properties inherited from parents
- Default formatting of inserted elements to inherited values
- Implement formattable elements.
- Easier multi-line text insertion at a single bookmark (inserting paragraph nodes after the one containing the bookmark)

## Contributing

Bug reports and pull requests are welcome on GitHub at <https://github.com/ruby-docx/docx>. Please make sure the test suite passes (`bundle exec rspec`) before opening a pull request.

## License

This gem is available as open source under the terms described in [LICENSE.md](LICENSE.md).
