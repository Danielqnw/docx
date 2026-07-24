# docx

[![Gem Version](https://badge.fury.io/rb/docx.svg)](https://badge.fury.io/rb/docx)
[![Ruby](https://github.com/ruby-docx/docx/workflows/Ruby/badge.svg)](https://github.com/ruby-docx/docx/actions?query=workflow%3ARuby)
[![Coverage Status](https://coveralls.io/repos/github/ruby-docx/docx/badge.svg?branch=master)](https://coveralls.io/github/ruby-docx/docx?branch=master)
[![Gitter](https://badges.gitter.im/ruby-docx/community.svg)](https://gitter.im/ruby-docx/community?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)

[English](README.md) | **简体中文**

> 一个用于读写 Microsoft Word `.docx` 文件的 Ruby 库（gem）。

它让你通过友好的对象模型来操作文档内容（段落、书签、表格、图片、样式），而**不必手动修改底层的 Office Open XML**。

完整 API 说明、参数表与更多示例见 **[文档站](docs/index.html)**（支持中 / 英切换）。

## 功能特性

| 能力 | 说明 |
| --- | --- |
| 📖 读取内容 | 遍历段落与书签，并可将段落渲染为 HTML |
| 📂 多种打开方式 | 支持从文件路径打开，也支持从内存缓冲区 / IO 对象打开 |
| 📊 表格操作 | 读取行 / 列 / 单元格，复制整行，替换占位符文本 |
| 🔗 单元格合并 | 在逻辑网格上合并 / 拆分矩形区域，并安全处理 `gridSpan` / `vMerge` |
| 🖼️ 图片替换 | 按关系 ID、压缩包内路径或占位符文本替换图片，支持表格内批量替换 |
| ✏️ 文本替换 | 保留原有格式替换文本，并可使用正则捕获组 |
| 🧩 模板填充 | 跨 run 替换、向纯文本占位符插图、勾选框、换行，以及高层 `render` 入口 |
| 📎 跨文档导入 | 从另一份文档导入正文节点，并隔离样式 / 编号 / 媒体 / 书签 id |
| 🎨 样式管理 | 新增、修改、删除段落 / 字符样式 |
| 🔧 底层访问 | 需要更精细控制时，可直接访问底层的 `Nokogiri` 节点 |

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
> 下面示例都假设已 `require 'docx'`。更多读法（缓冲区、HTML 渲染等）见 [文档 · 读取](docs/index.html#reading)。

## 常用场景

### 表格：复制行并填值

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

逻辑网格合并 / 拆分（`cell_at`、`merge_cells`、`unmerge_cells`）见 [文档 · 表格](docs/index.html#tables)。

### 图片：按关系或占位符替换

```ruby
doc = Docx::Document.open('with-images.docx')

doc.replace_image('rId5', 'replacement.png')
doc.replace_image_by_placeholder_in_table('{{photo_a}}', 'replacement.png', fit: :cover)

doc.save('with-images-edited.docx')
```

`fit` / `width` / `height`、批量插图等选项见 [文档 · 图片](docs/index.html#images)。

### 模板填充：一次 `render`

面向「用运行时数据填 `.docx` 模板」；键即占位符名，不假设业务语义。

```ruby
doc = Docx::Document.open('template.docx')

doc.render(
  text:             { '{{name}}' => 'Alice', '{{bio}}' => "第一行\n第二行" },
  images:           { '{{photo}}' => 'avatar.png' },
  checkboxes:       { '选项A' => true },
  content_controls: { 'opt_a' => true },
  tables:           [{ placeholder_row: '{{row}}',
                       rows: [ { '{{city}}' => 'Paris' }, { '{{city}}' => 'Tokyo' } ] }],
  multiline:      true,
  strip_unfilled: true
)

doc.save('report.docx')
```

跨 run 替换、纯文本占位符插图、勾选框等拆分 API 见 [文档 · 模板填充](docs/index.html#template-filling)。

### 跨文档导入：带隔离地搬正文

普通 `dup` 只复制 `document.xml`，样式 / 编号 / 图片 / 书签 id 会错乱。用导入 API 会一并重映射并带上依赖定义：

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

也可用 `import_before` / `import_after` / `import_node`。隔离策略与 `fallback:` 说明见 [文档 · 跨文档导入](docs/index.html#cross-document-import)。

### 样式：改一处，套多段

```ruby
doc = Docx::Document.open('example.docx')

style = doc.styles_configuration.add_style('Red', name: 'Red', font_color: 'FF0000', font_size: 20)
style.bold = true
doc.paragraphs.each { |p| p.style = 'Red' }

doc.save('styled.docx')
```

可设置的样式属性一览见 [文档 · 样式](docs/index.html#styles)。

## 进阶

封装 API 不够用时，可通过 `#node` / `#xpath` 操作底层 `Nokogiri` 节点，见 [文档 · 进阶](docs/index.html#advanced)。

错误类型（`Docx::Errors::*`）汇总见 [文档 · 错误参考](docs/index.html#error-reference)。

## 开发

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
