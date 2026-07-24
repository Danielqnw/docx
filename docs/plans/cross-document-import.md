# docx gem 改造需求：跨文档正文导入（ID / 部件隔离）

> 目标 gem：基于 `ruby-docx` / `pzgz/docx` 的自有 fork（`Docx::Document`，当前基线版本 `0.11.0`，可自由修改）。
>
> 本文件是**面向 gem 维护者的独立需求**，不含任何调用方业务上下文；术语统一为「目标文档（target）」= 接收内容的文档，「源文档（source）」= 提供正文片段的文档。

---

## 1. 背景与问题

gem 目前提供把一个文档的正文节点复制到另一个文档的能力：`Docx::Elements::Element#copy`（`lib/docx/elements/element.rb`）本质是 `Nokogiri::XML::Node#dup`，配合 `insert_before` / `insert_after` 把节点 reparent 到目标文档。

问题：`copy` **只复制 `word/document.xml` 里的正文 XML**，不处理任何“靠 ID 去其它部件解析”的引用。把源文档的表格/段落插入目标文档后，这些引用会按**目标文档**的部件解析，而两个文档里同号 ID 往往语义不同（甚至目标文档根本没有该 ID），造成整类格式错乱：

| 正文里的引用 | 解析目标部件 | 跨文档后果 |
|---|---|---|
| `w:pStyle` / `w:rStyle` / `w:tblStyle` | `word/styles.xml` | 表格边框、字体、斜体、行距、缩进按目标样式渲染 |
| `w:numPr/w:numId` → `abstractNumId` | `word/numbering.xml` | 列表/大纲编号错乱或凭空出现 |
| `r:embed` / `r:id` / `r:link` | `word/_rels/document.xml.rels` + `word/media/*` | 内嵌图片/超链接指向错误或丢失 |
| `w:asciiTheme` 等主题字体 | `word/theme/theme1.xml` | 字体回退 |
| `w:bookmarkStart@w:id` / 批注引用 id | 文档内需全局唯一 | 与目标已有 id 撞车，书签/批注错乱 |

逐属性把源样式“固化成直接格式再删 styleId”只能治标：每类属性补一刀，无法穷尽（`cnfStyle`、theme、`docDefaults` 等仍无法覆盖）。

**根治方向**（与 Word 自身、`docxcompose` 等做法一致）：把正文片段导入目标文档时，对 `styleId` / `numId` / `rId` / 书签 id 做**命名空间隔离**，并把它们依赖的**样式、编号、媒体等定义一并导入**目标文档对应部件。此能力应内建于 gem。

---

## 2. 需求总览

新增“跨文档正文导入”能力：对外暴露稳定 API，内部完成四类隔离 + 定义导入 + 存盘。拆成 6 组需求：

- **R1 存盘链路补全**：让 gem 能新增/回写 `styles` / `numbering` / `rels` / `media` / `[Content_Types].xml` 等部件。
- **R2 样式导入器**：导入并重映射 `styles.xml`。
- **R3 编号导入器**：导入并重映射 `numbering.xml`。
- **R4 媒体/关系导入器**：导入并重映射 `rels` + `media`。
- **R5 正文引用改写器**：把复制出的子树里的引用改写到映射后的新 ID。
- **R6 导入器编排与去重**：组合成绑定 (target, source) 的导入器，多次导入复用映射。

---

## 3. 公开 API（建议）

```ruby
module Docx
  class Document
    # 从 source_doc 导入一个正文节点（w:tbl / w:p 等），完成 styleId/numId/rId/书签id 隔离
    # 与相关定义导入，返回一个已在“当前(目标)文档上下文”、可直接 insert_before/insert_after 的节点。
    # importer: 传入以复用映射（同一 source 的多次导入应复用），不传则内部按 source 缓存。
    def import_node(source_doc, node, importer: nil) -> Nokogiri::XML::Node

    # 便捷方法：导入并插入到锚点前/后
    def import_before(source_doc, node, anchor_node) -> Nokogiri::XML::Node
    def import_after(source_doc, node, anchor_node) -> Nokogiri::XML::Node

    # 供导入器使用的底层部件操作（见 R1）
    def add_part(zip_path, content_type, bytes)             # 新增一个部件 + 注册 content-type
    def add_relationship(type, target, mode: nil) -> String # 新增关系，返回新 rId
    def ensure_default_content_type(ext, content_type)      # 保证扩展名有 Default
  end

  module Merge
    # 绑定 (target_doc, source_doc)，持有四张映射表，跨多次 import 复用
    class Importer
      def initialize(target_doc, source_doc)
      def import(node) -> Nokogiri::XML::Node   # Document#import_node 的实际实现
      # 内部状态：
      #   style_id_map        : source styleId => target styleId
      #   num_id_map          : source numId => target numId
      #   abstract_num_id_map : source abstractNumId => target abstractNumId
      #   rid_map             : source rId => target rId
      #   bookmark_id_offset  : 书签/批注 id 偏移量
    end
  end
end
```

典型用法（把 source 的所有 body 级表格导入 target）：

```ruby
importer = Docx::Merge::Importer.new(target_doc, source_doc)
source_doc.tables.each do |table|
  next if table.parent.name != "body"
  imported = importer.import(table.node)
  imported.add_previous_sibling(anchor_node)   # 或用 target_doc.import_before(...)
end
target_doc.save(path)
```

---

## 4. 详细需求

### R1 存盘链路补全（前置，必须先做）

现状（`lib/docx/document.rb`）：

- `save` / `stream` 仅遍历**原 zip 的 entries**，用 `@replace[entry.name]` 覆盖内容 → **无法写出原 zip 中不存在的新条目**（新增 media、新建的 numbering.xml 等）。
- `update` 只回写 `word/document.xml`、`word/styles.xml`（且来自 `styles_configuration.serialize`，即 `@styles.dup` 的 memo）、以及 `DOCUMENT_PATHS`（headers/footers/numbering，**仅当加载时已存在**）。
- `@rels`（`word/_rels/document.xml.rels`）加载了但**从不回写**；`[Content_Types].xml`、`word/media/*` 均不参与回写。

需求：

1. **支持新增条目**：在 `@replace` 之外引入“新增部件集合”（zip 路径 → bytes）。`save` / `stream` 写出时 = 原 entries（可被 `@replace` 覆盖）∪ 新增部件集合，去重且顺序稳定。
2. **styles.xml 回写走 `@styles`**：`update` 直接序列化 `@styles`（或让 `styles_configuration` 与 `@styles` 保持同源，避免 memo 导致改动丢失）。
3. **numbering.xml 可新建**：目标无 `word/numbering.xml` 时允许创建 `@numbering`，`update` 需写出，并在 `[Content_Types].xml` 与 `word/_rels/document.xml.rels` 注册（content-type override + `numbering` 关系）。
4. **rels 回写**：`@rels` 变更（新增 Relationship）必须写回 `word/_rels/document.xml.rels`。
5. **提供底层 API**：`add_part`、`add_relationship`、`ensure_default_content_type`（见 §3），供 R3/R4 使用。
6. `add_relationship` 生成的 `rId` 必须在该 rels 内唯一（扫描现有最大 `rIdN` 后自增）。

> 注：若 gem 现有任何“向文档写入图片/媒体”的方法（存在自己的私有写盘路径），需一并统一到本节的“新增部件”机制，避免两套并存导致产物不一致。

### R2 样式导入器（styles.xml）

输入：`source_doc.styles`、`target_doc.styles`；输出：`style_id_map`，并把需要的 style 追加进 target styles。

规则：

1. **按需闭包导入**：只导入“正文实际引用到的 styleId”及其 `basedOn` / `next` / `link` 链上的依赖（最多 10 层，防环）。
2. **去重优先**：对每个待导入 styleId：
   - target 存在**同 styleId 且规范化后 XML 等价** → 复用，不导入，`map[id]=id`；
   - target 存在同 styleId 但**定义不同** → 生成新 styleId（如 `m{seq}_{原id}` 或全局自增），`map[id]=新id`，导入改名后的 style；
   - target **无**该 styleId → 直接导入，保留原 id，`map[id]=id`。
3. **递归改写依赖**：导入的 style 内 `w:basedOn/@w:val`、`w:next/@w:val`、`w:link/@w:val` 一律按 `style_id_map` 改写（依赖未导入时先递归导入）。
4. **latentStyles**：以 `w:name` 为键，target 已有同名则保留 target 的、不导入 source 的 `lsdException`（避免覆盖目标外观）。
5. **子元素顺序**：遵守 OOXML `w:styles` 顺序 `docDefaults → latentStyles → style*`。**不改动 target 的 docDefaults**；导入的 `w:style` 追加到 `style*` 末尾。
6. `link`（段落样式↔字符样式配对）两端要么都导入、要么都复用，保持配对一致。

### R3 编号导入器（numbering.xml）

输入：source / target 的 numbering（可能不存在）；输出：`num_id_map`、`abstract_num_id_map`。

规则：

1. target 无 numbering.xml → 按 R1 新建骨架并注册。
2. **abstractNum**：source 每个 `w:abstractNum/@w:abstractNumId` 偏移到 `target_max_abstract_id + 1` 起，登记 `abstract_num_id_map`；改写其内部 `w:numStyleLink` / `w:styleLink` 指向的 styleId（用 `style_id_map`）。
3. **num**：source 每个 `w:num/@w:numId` 偏移到 `target_max_num_id + 1` 起，登记 `num_id_map`；改写其 `w:abstractNumId/@w:val`（用 `abstract_num_id_map`）。
4. 可选优化：仅导入被正文 `numId` 实际引用到的 num 及其 abstractNum 链；先全量导入亦可接受（正确性优先）。
5. `w:numbering` 子元素顺序需合法（`w:abstractNum*` 在前，`w:num*` 在后，`w:numIdMacAtCleanup` 等在最后）。

### R4 媒体 / 关系导入器（rels + media）

对被导入子树中引用到的每个关系型属性（`r:embed` / `r:id` / `r:link`）：

1. 在 **source rels** 中按 rId 找到 `Target`（如 `media/image3.png`）与 `Type`。
2. **内部关系**（如图片，`TargetMode` 非 External）：
   - 从 source zip 读出该 media 二进制；以**不与 target 现有 media 重名**的新文件名（如 `image{n}.ext`，n 取 target 现有最大值 +1）通过 `add_part` 加入 target；
   - `ensure_default_content_type(ext, ...)`（png / jpeg / jpg / gif / emf / wmf 等）；
   - `add_relationship(type, "media/新名")` 得到新 rId，登记 `rid_map`。
3. **外部关系**（`TargetMode="External"`，如超链接）：只 `add_relationship(type, 原Target, mode: :external)` 拿新 rId，不复制 media。
4. 同一 media 在同一 source 内被多个 rId 指向时，按 source rId 去重，避免重复导入同一图片。

### R5 正文引用改写器

在 `node.dup` 之后、插入 target 之前，对整棵子树按映射改写（未命中映射的保持原值）：

| 节点 / 属性 | 用的映射 |
|---|---|
| `w:pStyle/@w:val`、`w:rStyle/@w:val`、`w:tblStyle/@w:val` | `style_id_map` |
| `w:numPr/w:numId/@w:val` | `num_id_map` |
| `@r:embed`、`@r:id`、`@r:link`（含 `a:blip`、`v:imagedata`、超链接、OLE 等） | `rid_map` |
| `w:bookmarkStart/@w:id`、`w:bookmarkEnd/@w:id`、批注/尾注引用 id | `bookmark_id_offset` 偏移 |

- 命名空间：改写属性时使用正确前缀（`w:`、`r:`），确保 `Nokogiri` 序列化后前缀与文档根声明一致。
- 采用“改写引用”而非“删除 styleId + 固化直接格式”，因此不再需要任何逐属性固化逻辑。

### R6 导入器编排与去重

- 一个 `Importer` 绑定 (target, source)，四张映射表跨多次 `import(node)` 复用：同一 source 的 N 个节点只会导入一次样式/编号/媒体。
- 多份不同 source：各自独立 `Importer`（或 `Document#import_node` 内部以 source 对象为 key 缓存 importer）。
- 幂等：对同一 source 重复导入同一 style / num / media，必须命中已有映射，不产生重复定义。

---

## 5. 边界与约定

- **不改动 target 既有内容**：目标文档自身的 styles / numbering / docDefaults / media 一律不动，只做“追加”。冲突一律通过给 source 侧改名 / 偏移解决。
- **规范化比较**：R2 的“定义等价”判断需忽略无意义空白 / 属性顺序（建议对 style 节点做规范化字符串后比较；可先用严格相等，后续再放宽）。
- **命名空间安全**：跨文档 `node.dup` 进 target 后，凡新建节点均以 target 的 `document` 作为上下文构造，避免默认命名空间/前缀错乱。
- **失败回退**：`import_node` 内部异常时应记录日志并回退到旧 `copy`（保守不阻断），由开关控制是否启用回退。
- **性能**：映射构建与 media 导入按 source 缓存；styles / numbering 的 xpath 查询在导入器内建立 id→node 索引，避免 O(n²)。

---

## 6. 验收标准（gem spec，使用自带 fixtures docx）

1. **styleId 冲突**：source `tblStyle=X`（无边框表格样式）、target `X`（有实线边框的表格样式）→ 导入后正文 `tblStyle` 指向新 id，新 id 样式=无边框；target 的 `X` 不变。
2. **styleId 缺失**：source 有、target 无 → 保留原 id 导入，正文引用可解析。
3. **styleId 同名等价**：source / target 同 id 定义等价 → 复用，不新增 style。
4. **numbering**：带大纲编号的 source 导入 target（含 target 无 numbering.xml 的情况）→ numId / abstractNumId 已偏移、互不冲突、正文编号正确。
5. **图片**：内嵌图片的 source 导入 → media 落盘、rels / content-type 注册、正文 `r:embed` 指向新 rId，Word 能正常显示。
6. **书签**：source / target 同名或同 id 书签 → 导入后 id 无冲突。
7. **多 source / 多次导入**：同一 source 多节点只导入一次定义（映射复用）；产出文档无重复 styleId / numId。
8. **产物合法性**：合并结果解压后 `document.xml` 无悬空引用（所有 pStyle / rStyle / tblStyle / numId / rId 都能在对应部件解析到），且能被 Word / LibreOffice 正常打开、无“需修复”提示。

---

## 7. 交付物 / 改动文件（gem 仓库）

| 文件 | 改动 |
|---|---|
| `lib/docx/document.rb` | R1：`save` / `stream` / `update` 支持新增部件；`@styles` / `@numbering` / `@rels` 回写；新增 `add_part` / `add_relationship` / `ensure_default_content_type` / `import_node` / `import_before` / `import_after` |
| `lib/docx/merge/importer.rb`（新增） | R2–R6：`Docx::Merge::Importer`，样式 / 编号 / 媒体 / 正文四类隔离与导入 |
| `lib/docx/merge/*`（可选拆分） | `styles_importer.rb` / `numbering_importer.rb` / `media_importer.rb` / `node_rewriter.rb` |
| `lib/docx.rb` / `lib/docx/*.rb` | require 新文件 |
| `spec/docx/merge/*_spec.rb`（新增） | §6 单元测试 + 测试用 fixtures docx |

---

## 8. 分阶段落地建议

| 阶段 | 范围 | 价值 |
|---|---|---|
| P1 | R1 + R2 + R3 + R5(样式/编号引用) | 消灭 styleId / numId 冲突（边框、编号、斜体、行距、字体等整类问题）|
| P2 | R4 + R5(rId) | 支持含内嵌图片的正文片段安全合并 |
| P3 | R5(书签/批注) + R2 等价比较放宽 + 性能索引 | 收尾与健壮性 |

每阶段可独立发版；P1 完成即可覆盖绝大多数跨文档样式冲突场景。
