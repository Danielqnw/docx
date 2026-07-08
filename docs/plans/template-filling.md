# docx gem 通用功能需求（模板填充方向）

> 版本：v0.1 ｜ 日期：2026-07-07
> 目标读者：docx gem 维护/开发者
> 定位：**面向通用 docx 模板填充场景**的能力需求。本文档**不含任何调用方业务上下文**，所有能力须以通用库形式设计，可被任意项目复用。
> 基线代码：`pzgz/docx`（rev `69a58359de0f`，version 0.11.0，依赖 nokogiri + rubyzip）

---

## 一、背景与目标

将一份 `.docx` 作为**模板**，通过程序把「占位符」替换为运行时数据，产出成品 `.docx`。典型模板填充需要四类操作：

1. **文本占位符替换**（保留原有字体/字号/颜色等格式）
2. **表格数据行动态生成**（按 N 条数据克隆模板行）
3. **图片按占位符插入/替换**（表格单元格内，以及普通段落中）
4. **勾选框/单选状态设置**（☐ → ☑）

本文档梳理**当前基线已支持的能力**，并定义**待补齐的通用能力（含 API 建议、边界、验收标准）**。开发时须保持库的通用性：不假设文档结构、不硬编码任何字段名、不引入调用方语义。

---

## 二、当前基线能力（rev 69a58359，已支持）

> 以下为现状盘点，供开发者判断改动面；**这些能力应保持向后兼容**。

| 能力 | 现有 API | 说明 |
| --- | --- | --- |
| 打开 / 保存 / 流式输出 | `Docx::Document.open(path)`、`#save(path)`、`#stream` | — |
| 段落 / 文本读取 | `#paragraphs`、`para.text`、`para.text_runs` | — |
| **文本替换（保留格式）** | `run.substitute(match, replacement)`、`run.substitute_with_block { ... }`（支持正则捕获） | **按单个 text run 操作** |
| 段落整体设置文本 | `para.text = "..."` | 会重置该段落 run |
| 表格读取 | `doc.tables`、`table.rows/columns/cell_at(r,c)`、`cell.text` | 逻辑网格 + 物理访问 |
| **表格行克隆/插入** | `row.copy`、`row.insert_before(x)` / `insert_after(x)` / `append_to` / `prepend_to` | 复制模板行→填值，动态行已可实现 |
| 单元格合并 | `table.merge_cells(r0,c0,r1,c1)`、`unmerge_cells`、`cell.colspan/rowspan` | — |
| **图片替换（按引用/路径）** | `#replace_image(ref, source)` | 按 relationship id 或归档路径整体替换字节流 |
| **单元格内图片按占位符替换** | `#replace_image_by_placeholder_in_table(ph, source, fit:, width:, height:, cleanup_placeholder:)` | fit 支持 `:cover/:contain/:stretch`，可设尺寸(cm) |
| **单元格内图片批量插入** | `#replace_images_by_placeholder_in_table(ph, [sources], max_images_per_row:, ...)` | 超出每行上限自动复制模板行 |
| 图片尺寸/裁剪 | 内部支持 png/jpeg 尺寸探测、cover/contain 计算、`srcRect` 裁剪 | — |
| 书签 / 超链接 / 样式 | `#bookmarks`、`#hyperlinks`、`styles_configuration` | — |

**结论**：文本替换、表格行动态生成、**「单元格内已有图片」的按占位符替换/批量插入**均已具备。缺口集中在下面第三节。

---

## 三、待补齐的通用能力（本次需求）

优先级：P0 = 阻塞模板填充；P1 = 强需要；P2 = 增强。

---

### FR-1【P0】跨 run 的文本占位符替换（段落级）

**问题**：Word 常把一个占位符（如 `{{name}}`）拆分到多个相邻 `w:r`/`w:t` 中（因拼写检查、格式痕迹等）。现有 `run.substitute` **按单个 run** 执行 `gsub`，占位符一旦跨 run 就**匹配不到**，导致替换静默失败——这是模板填充最常见的坑。

**需求**：提供**段落级**替换，能跨越同一段落内被拆分的多个 run 匹配并替换占位符，同时**保留占位符首个 run 的格式**。

**API 建议**：
```ruby
paragraph.substitute(match, replacement)          # match: String | Regexp
document.substitute_across_runs(match, replacement) # 便捷：全文所有段落 + 表格单元格
```

**行为要求**：
- 合并段落内连续 run 的文本进行匹配；命中后将替换文本写入首个命中 run，清空其余被占用 run 的对应文本。
- 保留命中区域**第一个 run 的格式**（字体/字号/颜色/粗斜体）。
- 支持一段落内**多次命中**；支持 `Regexp`（含捕获组 `\1`）。
- 未命中时不修改文档、不报错。

**验收**：模板中把 `{{title}}` 人为拆成 `{{ti` + `tle}}` 两个 run，替换后文本正确、格式不丢；Word 打开无修复提示。

---

### FR-2【P0】向「无图片的占位符位置」插入图片

**问题**：现有图片插入（`replace_image_by_placeholder_in_table` / `replace_images_by_placeholder_in_table`）**依赖单元格内已存在一张图片**（需要 `a:blip/@r:embed` 作为锚点，`duplicate_image_slot_in_cell` 也要求已有 `w:r[w:drawing]`）。若模板卡位是**纯文本占位符、无预置图片**，会抛 `ImageNotFound`。

**需求**：支持在**仅有文本占位符、无任何预置图片**的位置**从零插入图片**——自动创建 drawing/inline、注册 media 与 relationship、生成合法的 `w:drawing` XML。

**API 建议**：
```ruby
# 通用：在任意“文本占位符”处插入图片（段落或单元格均可）
document.insert_image_at_placeholder(placeholder, source, fit: :contain, width:, height:, cleanup_placeholder: true)

# 批量：一个占位符处按序插入多张（可选每行数量 / 网格）
document.insert_images_at_placeholder(placeholder, [sources], width:, height:, ...)
```

**行为要求**：
- `source` 支持：文件路径、`IO`/`StringIO`、原始字节 `String`（与现有 `read_replacement_source` 一致）。
- 支持 png / jpeg（最好含 gif/bmp 探测），自动写入 `word/media/`、追加 `[Content_Types].xml` 与 `document.xml.rels` 关系、分配唯一 `rId` 与媒体文件名。
- 尺寸：给定 `width`/`height`(cm) 即用之；未给定时按图片像素 + DPI 推断合理默认，避免超出页面。
- `fit`（`:cover/:contain/:stretch`）复用现有实现。
- `cleanup_placeholder: true` 时移除占位符文本。
- 位置命中同时适用于**普通段落**与**表格单元格**（见 FR-3 关系）。

**验收**：一个不含任何图片的空模板，仅有段落 `{{photo}}` 与单元格 `{{img}}`，各插入一张 png，产出文档在 Word/LibreOffice 正常显示、无修复提示；`.rels` 与 `[Content_Types]` 合法。

---

### FR-3【P1】普通段落（非表格）中的图片按占位符替换/插入

**问题**：现有按占位符的图片操作**仅限表格单元格**（方法名带 `_in_table`，实现走 `find_table_cell_by_placeholder`）。段落中的图片占位符无对应 API。

**需求**：把「按占位符定位 + 替换/插入图片」的能力**推广到普通段落**（不在表格内）。理想情况下 FR-2 的 `insert_image_at_placeholder` 统一处理段落与单元格两种宿主，`replace_image_by_placeholder`（去掉 `_in_table` 限制）统一替换已有图片。

**行为要求**：
- 占位符定位不限制其祖先是否为 `w:tc`。
- 替换已有图片、或从零插入（复用 FR-2），行为与表格内一致。
- 保留向后兼容：现有 `*_in_table` 方法保留或内部委托给通用实现。

**验收**：段落级占位符的替换与插入均通过，Word 打开正常。

---

### FR-4【P1】勾选框 / 单选状态设置

**问题**：模板常见「☐ 选项A / ☐ 选项B」需按数据勾选，目前无通用支持。

**需求**：提供通用方式把某个「选项占位符」的未选态改为已选态，覆盖两种常见模板写法：

1. **字符型**：把 `☐`（U+2610）替换为 `☑`（U+2611）/ `☒`；或把某标记占位符（如 `[ ]`→`[x]`）切换。
2. **内容控件型**（`w:sdt` + `w14:checkbox`）：设置 checkbox 选中状态并同步显示字形。

**API 建议**：
```ruby
document.set_checkbox(placeholder, checked: true)     # 字符型：按占位符切换字形
document.check_content_control(tag_or_alias, checked: true)  # sdt checkbox 型
```

**行为要求**：
- 字符型：仅切换目标占位符处的字形，不影响其它同字符。
- 内容控件型：正确修改 `w14:checkbox/w14:checked` 并同步 run 内显示字形。
- 幂等：重复调用结果一致。

**验收**：两种模板各一份 fixture，勾选后 Word 显示正确的 ☑ 状态。

---

### FR-5【P1】多行文本 / 换行替换

**问题**：`substitute` 的替换值含 `\n` 时，Word 不会渲染为换行（OOXML 需 `<w:br/>`），多行文本会挤成一行。

**需求**：替换文本中的 `\n` 能生成软换行 `<w:br/>`；可选按 `\n\n` 生成新段落。

**API 建议**：
```ruby
paragraph.substitute(match, "line1\nline2", multiline: true)  # \n → <w:br/>
```

**行为要求**：
- `multiline: true`（或独立方法）时把 `\n` 转为 `w:br`；保留原格式。
- 默认关闭，保持向后兼容。

**验收**：替换含多行字符串后，Word 中按行显示。

---

### FR-6【P2】高层模板渲染入口（便捷封装）

**问题**：调用方需自行编排「文本 + 表格行 + 图片 + 勾选」多步操作。

**需求**（可选增强）：提供一个数据驱动的高层入口，一次性完成常见替换。**纯通用、无业务语义**——键即占位符名，值的类型决定处理方式。

**API 建议**：
```ruby
document.render(
  text:      { "{{name}}" => "Alice", "{{date}}" => "2026-07-07" },
  images:    { "{{photo}}" => "a.png", "{{gallery}}" => ["1.png","2.png"] },
  checkboxes:{ "{{opt_a}}" => true },
  tables:    [{ placeholder_row: "{{row}}", rows: [ {"{{c1}}"=>"x"}, ... ] }]
)
```

**行为要求**：内部委托 FR-1~FR-5；未命中的占位符可选保留或清空（参数控制）。

**验收**：一份综合模板经单次 `render` 产出正确成品。

---

### FR-7【P2】未使用占位符 / 空槽清理

**需求**：提供统一清理能力——把模板中**未被赋值的占位符**批量清空（或删除所在空行/空图槽），避免成品残留 `{{...}}`。现有 `cleanup_placeholder` 仅局部，需提供全局：
```ruby
document.strip_unfilled_placeholders(pattern: /\{\{.*?\}\}/)
```

**验收**：产出文档不残留未赋值占位符；不误删已赋值内容。

---

## 四、通用性与非功能要求

- **无宿主上下文**：所有 API 仅围绕 docx 结构与占位符，不得引入任何特定业务字段/命名。
- **向后兼容**：不破坏现有公共 API（`substitute`、`replace_image*`、表格/合并等）；新能力以新方法或可选参数提供。
- **输入源统一**：图片/替换源统一支持 路径 / `IO` / 字节串。
- **产物合法性**：生成的 `document.xml` / `.rels` / `[Content_Types].xml` 必须合法，Word 与 LibreOffice 打开**无"需要修复"提示**。
- **编码**：正确处理 UTF-8（含中文、全角、Emoji 字形如 ☑）。
- **幂等/无副作用**：操作在内存 DOM 上进行，`save`/`stream` 前不落盘中间态。
- **性能**：单文档数十张图片、数百行表格量级下可用（避免每次全文档重复解析）。
- **错误处理**：定位失败给明确异常（复用 `ImagePlaceholderNotFound` / `ImageNotFound` 等），可选「未命中不报错」开关。
- **依赖**：不新增重量级依赖；沿用 nokogiri + rubyzip。

---

## 五、验收与测试

- 为 FR-1~FR-7 各提供 **fixture 模板（.docx）+ 单元测试**，断言：目标内容正确、格式保留、产物可被 Word/LibreOffice 无修复打开。
- 关键回归：跨 run 占位符、空位置插图、段落 vs 单元格、多图行复制、勾选字形、多行换行。
- 建议附一个「综合模板」端到端用例（文本 + 动态表格行 + 单元格图 + 段落图 + 勾选 + 多行）验证 FR-6。

---

## 六、交付物

1. 上述 API 的实现（保持向后兼容）。
2. 单元测试 + fixtures。
3. README 更新：新增能力的用法示例。
4. 变更说明：新增/调整的公共方法清单与兼容性说明。

---

## 七、优先级汇总

| 优先级 | 条目 |
| --- | --- |
| **P0** | FR-1 跨 run 文本替换；FR-2 无图占位符插图 |
| **P1** | FR-3 段落图片占位符；FR-4 勾选框；FR-5 多行换行 |
| **P2** | FR-6 高层 render 入口；FR-7 未用占位符清理 |
