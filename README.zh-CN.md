# docx

[![Gem Version](https://badge.fury.io/rb/docx.svg)](https://badge.fury.io/rb/docx)
[![Ruby](https://github.com/ruby-docx/docx/workflows/Ruby/badge.svg)](https://github.com/ruby-docx/docx/actions?query=workflow%3ARuby)
[![Coverage Status](https://coveralls.io/repos/github/ruby-docx/docx/badge.svg?branch=master)](https://coveralls.io/github/ruby-docx/docx?branch=master)
[![Gitter](https://badges.gitter.im/ruby-docx/community.svg)](https://gitter.im/ruby-docx/community?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)

[English](README.md) | **简体中文**

> 一个用于读写 Microsoft Word `.docx` 文件的 Ruby 库（gem）。

它让你通过友好的对象模型来操作文档内容（段落、书签、表格、图片、样式），而**不必手动修改底层的 Office Open XML**。

## 功能特性

| 能力 | 说明 |
| --- | --- |
| 📖 读取内容 | 遍历段落与书签，并可将段落渲染为 HTML |
| 📂 多种打开方式 | 支持从文件路径打开，也支持从内存缓冲区 / IO 对象打开 |
| 📊 表格操作 | 读取行 / 列 / 单元格，复制整行，替换占位符文本 |
| 🔗 单元格合并 | 在逻辑网格上合并 / 拆分矩形区域，并安全处理 `gridSpan` / `vMerge` |
| 🖼️ 图片替换 | 按关系 ID、压缩包内路径或占位符文本替换图片，支持表格内批量替换 |
| ✏️ 文本替换 | 保留原有格式替换文本，并可使用正则捕获组 |
| 🎨 样式管理 | 新增、修改、删除段落 / 字符样式 |
| 🔧 底层访问 | 需要更精细控制时，可直接访问底层的 `Nokogiri` 节点 |

## 目录

- [环境要求](#环境要求)
- [安装](#安装)
- [快速上手](#快速上手)
- [读取](#读取)
  - [段落与书签](#段落与书签)
  - [从缓冲区打开](#从缓冲区打开)
  - [渲染 HTML](#渲染-html)
- [表格](#表格)
  - [读取表格](#读取表格)
  - [写入表格](#写入表格)
  - [合并与拆分单元格](#合并与拆分单元格)
- [图片](#图片)
- [写入与替换文本](#写入与替换文本)
- [样式](#样式)
  - [样式属性](#样式属性)
- [进阶：访问底层节点](#进阶访问底层节点)
- [错误参考](#错误参考)
- [开发](#开发)
- [贡献](#贡献)
- [许可证](#许可证)

## 环境要求

- Ruby 2.6 或更高版本

## 安装

在应用的 Gemfile 中加入：

```ruby
gem 'docx'
```

然后执行 `bundle install`。或者直接手动安装：

```shell
gem install docx
```

## 快速上手

下面这段代码展示了最常见的「打开 → 读取 → 编辑 → 保存」完整流程：

```ruby
require 'docx'

# 打开一个已有文档
doc = Docx::Document.open('example.docx')

# 读取每个段落
doc.paragraphs.each { |p| puts p.to_s }

# 编辑后另存为新文件
doc.paragraphs.each do |p|
  p.each_text_run { |run| run.substitute('{{name}}', 'Alice') }
end
doc.save('example-edited.docx')
```

> [!NOTE]
> 下面所有示例都假设你已经通过 `require 'docx'` 引入了该 gem。

## 读取

### 段落与书签

打开文档后，可以分别遍历其中的段落和书签：

```ruby
# 为已有的 docx 文件创建一个 Docx::Document 对象
doc = Docx::Document.open('example.docx')

# 读取并打印段落
doc.paragraphs.each do |p|
  puts p
end

# 读取并打印书签，返回值是以书签名为键的哈希
doc.bookmarks.each_pair do |bookmark_name, bookmark_object|
  puts bookmark_name
end
```

段落对象支持 `to_s`（纯文本）和 `to_html`，并可通过 `each_text_run` 访问其中的文本片段（text run）。

### 从缓冲区打开

你不一定需要磁盘上的文件——内存缓冲区或任意类 IO 对象同样可用。这在处理通过 HTTP 获取的文档或 Web 上传文件时非常方便：

```ruby
# 从远程文件 / StringIO / 上传文件创建 Docx::Document 对象
doc = Docx::Document.open(buffer)

# 读取方式与上文完全相同
```

### 渲染 HTML

把段落转换成 HTML 字符串，便于在网页中展示：

```ruby
doc = Docx::Document.open('example.docx')
doc.paragraphs.each do |p|
  puts p.to_html
end
```

## 表格

### 读取表格

通过 `doc.tables` 拿到所有表格，再按行或按列访问单元格：

```ruby
doc = Docx::Document.open('tables.docx')

first_table = doc.tables[0]
puts first_table.row_count
puts first_table.column_count
puts first_table.rows[0].cells[0].text
puts first_table.columns[0].cells[0].text

# 遍历所有表格
doc.tables.each do |table|
  table.rows.each do |row| # 按行遍历
    row.cells.each do |cell|
      puts cell.text
    end
  end

  table.columns.each do |column| # 按列遍历
    column.cells.each do |cell|
      puts cell.text
    end
  end
end
```

定位单元格有两种方式，理解它们的区别对处理合并单元格很重要：

| 方式 | 写法 | 特点 |
| --- | --- | --- |
| **物理访问** | `table.rows[i].cells[j]` | 对应每行中实际存在的 `w:tc` 元素；含合并单元格的行，物理单元格数量更少 |
| **逻辑访问** | `table.cell_at(row, col)` | 基于*逻辑网格*；合并区域内任意坐标都返回该区域的锚点单元格，网格始终是规整矩形 |

### 写入表格

下面的示例复制表格最后一行，插入为新行，并替换其中的占位符文本：

```ruby
doc = Docx::Document.open('tables.docx')

# 遍历每个表格
doc.tables.each do |table|
  last_row = table.rows.last

  # 复制最后一行，并在其前插入一行新行
  new_row = last_row.copy
  new_row.insert_before(last_row)

  # 替换新行中每个单元格的文本
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

### 合并与拆分单元格

合并使用**逻辑网格坐标**（`row` / `col` 从 0 开始），合并区域的内容会保留在左上角的*锚点*单元格中。

先看如何查看单元格及其合并状态：

```ruby
doc = Docx::Document.open('tables.docx')
table = doc.tables[0]

# 按逻辑坐标查看某个单元格
cell = table.cell_at(1, 2)
puts cell.text if cell

# 该坐标是否属于某个合并区域？
puts table.merged?(0, 0)

# 遍历逻辑网格，并查看合并相关信息
table.each_cell do |cell, row, col|
  puts "#{row},#{col}: #{cell.text} (colspan=#{cell.colspan}, rowspan=#{cell.rowspan})"
  puts "  锚点: #{cell.merge_anchor?}, 延续: #{cell.merge_continuation?}"
end
```

再看如何合并与拆分一个矩形区域（边界为闭区间，含两端）：

```ruby
# 合并从 (row0, col0) 到 (row1, col1) 的矩形；内容保留在左上角锚点
table.merge_cells(0, 0, 1, 1)

# 从锚点坐标拆分（若锚点本身未合并，则为空操作）
table.unmerge_cells(0, 0)

# 也可以直接从锚点单元格拆分
anchor = table.cell_at(0, 0)
anchor.unmerge! if anchor&.merge_anchor?

doc.save('merged.docx')
```

> [!IMPORTANT]
> 合并只是一种展示层面的操作：它**不会**改变 `row_count` 或 `column_count`，因为表格的 `tblGrid` 保持不变。

常用的单元格判断方法与访问器：

| 方法 | 说明 |
| --- | --- |
| `cell.colspan` | 单元格横跨的网格列数 |
| `cell.rowspan` | 单元格纵跨的网格行数 |
| `cell.merge_anchor?` | 是否为合并区域左上角的锚点单元格 |
| `cell.merge_continuation?` | 是否被其他单元格的合并所覆盖 |
| `cell.unmerge!` | 拆分以该单元格为锚点的合并区域 |

合并 / 拆分时可能抛出的错误：

| 错误 | 触发场景 |
| --- | --- |
| `Docx::Errors::InvalidMergeRange` | 坐标越界，或 `row0 > row1` / `col0 > col1` |
| `Docx::Errors::MergeConflict` | 请求的范围与已有合并区域重叠 |
| `Docx::Errors::InvalidMergeTarget` | 对非锚点或未合并的单元格调用了拆分 |

#### 示例：把每一行最右边的三列合并

一个实战场景——对某张表，从第 2 行起，把每行最右边的三列各自合并成一格：

```ruby
doc = Docx::Document.open('report.docx')
table = doc.tables[1]

# 这里逻辑列索引 5、6、7 是最右边的三列
(1...table.row_count).each do |row|
  table.merge_cells(row, 5, row, 7)
end

doc.save('report-merged.docx')
```

## 图片

`doc.images` 会列出图片关系 ID 与压缩包内路径的映射，随后你可以按多种方式替换图片：

```ruby
doc = Docx::Document.open('with-images.docx')

# 查看图片关系 ID 与压缩包内路径的映射
# => { "rId5" => "word/media/image1.png", ... }
pp doc.images

# 按关系 ID 用文件路径替换
doc.replace_image('rId5', 'replacement.png')

# 按压缩包内路径用 IO 对象替换
File.open('replacement.png', 'rb') do |io|
  doc.replace_image('word/media/image1.png', io)
end
```

如果图片位置是用占位符文本（如 `{{photo_a}}`）标记在表格单元格里，可以按占位符替换，并控制输出尺寸与填充方式：

```ruby
# 按表格单元格中的占位符文本替换
doc.replace_image_by_placeholder_in_table('{{photo_a}}', 'replacement.png', fit: :cover)

# 明确指定输出尺寸（5 厘米 x 3 厘米）
doc.replace_image_by_placeholder_in_table(
  '{{photo_a}}', 'replacement.png',
  width: 5.0, height: 3.0, fit: :cover
)

# 只给定宽度——高度会按源图片的宽高比自动计算
doc.replace_image_by_placeholder_in_table(
  '{{photo_a}}', 'replacement.png',
  width: 5.0
)
```

还可以用一个占位符**批量**放入多张图片：

```ruby
# 同一单元格默认每行最多放 2 张图，超出部分会追加复制的行。
# width / height 同样适用，会应用到每个图片位。
doc.replace_images_by_placeholder_in_table(
  '{{photo_a}}',
  ['image-a.png', 'image-b.png', 'image-c.png'],
  width: 5.0, height: 3.0, fit: :cover
)

doc.save('with-images-edited.docx')
```

图片替换选项：

| 选项 | 适用范围 | 说明 |
| --- | --- | --- |
| `fit:` | 占位符替换 | `:stretch`（默认）、`:cover` 或 `:contain`，控制新图片如何填充目标框 |
| `width:` | 占位符替换 | 输出宽度（厘米）。若只给定 `width`，高度按源图宽高比推算 |
| `height:` | 占位符替换 | 输出高度（厘米）。若只给定 `height`，宽度按源图宽高比推算 |
| `max_images_per_row:` | `replace_images_by_placeholder_in_table` | 一个单元格中放置图片的上限，超出后追加复制行（默认为 `2`） |

图片相关错误：

| 错误 | 触发场景 |
| --- | --- |
| `Docx::Errors::ImageNotFound` | 没有图片匹配给定的关系 ID / 压缩包内路径 |
| `Docx::Errors::ImagePlaceholderNotFound` | 在任何表格单元格中都未找到该占位符文本 |

## 写入与替换文本

下面集中演示书签插入、段落删除，以及两种文本替换方式（普通替换与带正则捕获组的替换）：

```ruby
doc = Docx::Document.open('example.docx')

# 在某个书签后插入一行文本
doc.bookmarks['example_bookmark'].insert_text_after('Hello world.')

# 在书签处插入多行文本
doc.bookmarks['example_bookmark_2'].insert_multiple_lines_after(['Hello', 'World', 'foo'])

# 删除段落
doc.paragraphs.each do |p|
  p.remove! if p.to_s =~ /TODO/
end

# 替换文本，同时保留格式
doc.paragraphs.each do |p|
  p.each_text_run do |tr|
    tr.substitute('_placeholder_', 'replacement value')
  end
end

# 借助捕获组替换文本。块参数是 MatchData，
# 其行为与 String#gsub 略有不同。
# https://ruby-doc.org/3.3.7/MatchData.html
doc.paragraphs.each do |p|
  p.each_text_run do |tr|
    tr.substitute_with_block(/total: (\d+)/) { |match_data| "total: #{match_data[1].to_i * 10}" }
  end
end

# 将文档保存到指定路径
doc.save('example-edited.docx')
```

## 样式

通过 `styles_configuration` 可以读取、修改、新增和删除样式，再把样式应用到段落上：

```ruby
d = Docx::Document.open('example.docx')

# 修改已有样式
existing_style = d.styles_configuration.style_of('Heading 1')
existing_style.font_color = '000000'

# 新增样式（属性见下文）
new_style = d.styles_configuration.add_style('Red', name: 'Red', font_color: 'FF0000', font_size: 20)
new_style.bold = true

# 应用样式到段落
d.paragraphs.each { |p| p.style = 'Red' }
d.paragraphs.each { |p| p.style = 'Heading 1' }

# 删除样式
d.styles_configuration.remove_style('Red')
```

### 样式属性

下面按类别列出可设置的属性。标注「布尔」的属性取值为 `true` / `false`，标注「颜色」的取值为十六进制颜色码（如 `FF0000`）。

**基础**

| 属性 | 说明 |
| --- | --- |
| `id` | 样式的唯一标识符（必填） |
| `name` | 样式的可读名称（必填） |
| `type` | 样式类型（如 paragraph 段落、character 字符） |

**段落格式**

| 属性 | 类型 | 说明 |
| --- | --- | --- |
| `keep_next` | 布尔 | 是否让段落与下一段保持在同一页 |
| `keep_lines` | 布尔 | 是否让段落所有行保持在同一页 |
| `page_break_before` | 布尔 | 是否在段落前插入分页符 |
| `widow_control` | 布尔 | 控制段落的孤行 / 寡行 |
| `suppress_auto_hyphens` | 布尔 | 控制自动断字 |
| `bidirectional_text` | 布尔 | 段落是否包含双向文本 |
| `spacing_before` | — | 段落前间距 |
| `spacing_after` | — | 段落后间距 |
| `line_spacing` | — | 段落行距 |
| `line_rule` | — | 行距的计算方式 |
| `indent_left` | — | 段落左缩进 |
| `indent_right` | — | 段落右缩进 |
| `indent_first_line` | — | 段落首行缩进 |
| `align` | — | 段落内文本对齐方式 |
| `outline_level` | — | 文档层级中的大纲级别 |

**字体与字符**

| 属性 | 类型 | 说明 |
| --- | --- | --- |
| `font` | — | 为不同脚本统一设置字体（ASCII、复杂脚本、东亚字符等） |
| `font_ascii` | — | ASCII 字符字体 |
| `font_cs` | — | 复杂脚本字符字体 |
| `font_hAnsi` | — | 高位 ANSI 字符字体 |
| `font_eastAsia` | — | 东亚字符字体 |
| `font_color` | 颜色 | 文本颜色 |
| `font_size` | — | 字号 |
| `font_size_cs` | — | 复杂脚本字符的字号 |
| `bold` | 布尔 | 加粗 |
| `italic` | 布尔 | 斜体 |
| `caps` | 布尔 | 全部大写 |
| `small_caps` | 布尔 | 小型大写字母 |
| `strike` | 布尔 | 删除线 |
| `double_strike` | 布尔 | 双删除线 |
| `outline` | 布尔 | 轮廓效果 |
| `underline_style` | — | 下划线样式 |
| `underline_color` | 颜色 | 下划线颜色 |
| `spacing` | — | 字符间距 |
| `kerning` | — | 字距调整 |
| `position` | — | 字符位置（上标 / 下标） |
| `text_fill_color` | 颜色 | 文本填充色 |
| `vertical_alignment` | — | 行内文本的垂直对齐方式 |
| `lang` | — | 文本的语言标记 |

**底纹**

| 属性 | 类型 | 说明 |
| --- | --- | --- |
| `shading_style` | — | 底纹图案样式 |
| `shading_color` | 颜色 | 底纹图案颜色 |
| `shading_fill` | — | 底纹背景填充色 |

## 进阶：访问底层节点

当封装好的 API 无法满足需求时，可以直接操作底层的 `Nokogiri` 节点：

```ruby
d = Docx::Document.open('example.docx')

# 可通过 #node 访问元素所基于的 Nokogiri::XML::Node
d.paragraphs.each do |p|
  puts p.node.inspect
end

# 元素的 #xpath 和 #at_xpath 方法会委托给底层节点，省去一步
p_element = d.paragraphs.first
p_children = p_element.xpath('//child::*') # 选取所有子节点
p_child = p_element.at_xpath('//child::*') # 选取第一个子节点
```

## 错误参考

所有错误都位于 `Docx::Errors` 命名空间下。

| 错误 | 触发场景 |
| --- | --- |
| `StyleNotFound` | 请求的样式名不存在 |
| `StyleInvalidPropertyValue` | 样式属性被设置为非法值 |
| `StyleRequiredPropertyValue` | 缺少必填的样式属性 |
| `InvalidMergeRange` | 合并坐标越界或首尾颠倒 |
| `MergeConflict` | 合并范围与已有合并重叠 |
| `InvalidMergeTarget` | 对非锚点或未合并单元格调用了拆分 |
| `ImageNotFound` | 没有图片匹配给定的关系 ID / 压缩包内路径 |
| `ImagePlaceholderNotFound` | 未找到图片占位符文本 |

## 开发

克隆仓库后，安装依赖并运行测试套件：

```shell
bundle install
bundle exec rspec
```

### 待办

- 根据元素属性中的值以及从父级继承的属性来计算元素格式
- 让插入元素的默认格式取自继承值
- 实现可格式化的元素
- 让在单个书签处插入多行文本更便捷（在含书签的段落后插入段落节点）

## 贡献

欢迎在 GitHub 上提交问题与 Pull Request：<https://github.com/ruby-docx/docx>。提交 PR 前请确保测试套件通过（`bundle exec rspec`）。

## 许可证

本 gem 以开源形式提供，具体条款见 [LICENSE.md](LICENSE.md)。
