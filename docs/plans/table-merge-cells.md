# Plan: 表格合并单元格扩展

> 状态：待实施  
> 目标版本：0.11.0（建议）  
> 关联模块：`lib/docx/containers/table*.rb`

## 0. 前置确认

- 现有 `Document#save` / `save_to` 支持 round-trip（写回 zip），Phase 3/4 验收依赖；开工前先跑一个最小 save→reopen 验证现状。
- Fixture 一律由脚本生成，不依赖本地 Word 环境。

## 1. 目标

为 `docx` gem 增加表格合并单元格的**读取、合并、拆分**能力，并修正现有 `Table#columns` 在合并表格上的错位问题。

### 成功标准

- [ ] 能正确解析 Word 制作的横/纵/矩形合并表格
- [ ] 提供稳定的逻辑坐标 API `(row, col)`
- [ ] `merge_cells` / `unmerge_cells` 可 round-trip（save → reopen 一致）
- [ ] Word / LibreOffice 能正常打开修改后的文档
- [ ] 现有 `spec/docx/document_spec.rb` 表格测试无回归

### 非目标（本期不做）

- 表格样式（边框、列宽、对齐）
- 从零创建表格
- 嵌套表格
- 表格 HTML 导出

---

## 2. 现状问题

| 问题 | 位置 | 影响 |
|------|------|------|
| `w:tr//w:tc[i+1]` 用后代轴 `//` + 谓词取列 | `Table#columns` | ① 合并表列错位；② 嵌套表会抓到内层 `w:tc` |
| 无逻辑网格 | 全局 | 无法实现 merge/unmerge |
| `tcPr` 可能不存在 | `Container#properties` | 写入 merge 属性失败 |
| 物理/逻辑坐标混用 | API 设计 | 调用方易误用 |

> 修正方向：`columns` 走 grid 构建，定位时用子轴 `w:tr/w:tc`（非 `//`），同时规避合并错位与嵌套表误抓。
> 两类问题都在 CHANGELOG 的 bugfix 描述中列出。

---

## 3. API 设计

### 3.1 Table

```ruby
table.cell_at(row, col)                    # 逻辑坐标 → TableCell
table.each_cell { |cell, row, col| ... }   # 遍历 anchor 单元格
table.merged?(row, col)                    # 是否属于某合并区域
table.merge_cells(row0, col0, row1, col1)  # 合并矩形（含边界）
table.unmerge_cells(row, col)              # 从 anchor 拆分
```

### 3.2 TableCell

```ruby
cell.colspan                 # Integer, 默认 1
cell.rowspan                 # Integer, 默认 1
cell.merged?                 # 参与合并（含延续格）
cell.merge_anchor?           # 合并区域左上角
cell.merge_continuation?     # vMerge 延续格
cell.unmerge!                # 委托 table.unmerge_cells
```

### 3.3 错误类型（`lib/docx/errors.rb`）

现有文件已是 `Docx::Errors` 模块，新异常须放进同一模块，调用处写 `Docx::Errors::MergeConflict`：

```ruby
module Docx
  module Errors
    # ... 现有 StyleNotFound 等 ...
    InvalidMergeRange   = Class.new(ArgumentError)
    MergeConflict       = Class.new(StandardError)
    InvalidMergeTarget  = Class.new(ArgumentError)
  end
end
```

### 3.4 行为约定

- 合并区域内容保留在**左上角 anchor**，其余格清空
- 单格 `(r,c,r,c)` → no-op
- 与已有合并重叠 → `MergeConflict`
- 对非 anchor 执行 unmerge → `InvalidMergeTarget`
- `rows[i].cells[j]` 保持为**物理**访问，文档中明确区分

---

## 4. 架构

### 4.1 新增文件

```
lib/docx/containers/
├── table_grid.rb      # 逻辑网格解析与缓存
└── table_merge.rb     # merge / unmerge XML 写操作
```

### 4.2 修改文件

```
lib/docx/containers/table.rb         # 委托 grid，新增公共 API
lib/docx/containers/table_cell.rb    # 合并属性读取
lib/docx/containers/table_column.rb    # 基于 grid 构建
lib/docx/containers.rb               # require 新模块
lib/docx/errors.rb                   # 新异常
README.md                            # 用法示例
CHANGELOG.md                         # 变更记录
```

### 4.3 核心数据结构

```ruby
CellSlot = Struct.new(
  :node, :row, :col, :colspan, :rowspan,
  :anchor,        # 该 slot 所属合并区域的 anchor CellSlot（anchor 自身指向自己）
  :continuation,  # 是否为 vMerge 延续格
  keyword_init: true
)
```

`TableGrid` 维护 `slots[row][col]`：

- anchor 位置 → 一个 `CellSlot`，其 `anchor` 指向自身
- 被 anchor 覆盖的占位位置 → **同一 anchor 的 `CellSlot` 引用**（不使用裸 `:occupied` 符号，
  否则 `cell_at(延续/被吞坐标)` 无法反查回 anchor）
- 网格外/空位 → `nil`

`cell_at(row, col)` 对合并区域内任意坐标都返回 anchor 对应的 `TableCell`。

### 4.4 columns 语义（合并表）

改造后 `columns[col].cells[row]` 一律返回 `cell_at(row, col)`，即该逻辑坐标的 anchor cell。
含义：

- `cells` 长度恒等于 `row_count`
- 同一个 anchor cell 会在多个 `(row, col)` 位置**重复出现**（这是逻辑坐标语义，调用方需知悉）
- 越界坐标返回 `nil`

### 4.5 缓存

结构变更后调用 `table.invalidate_grid!` 清除 `@grid`。`merge_cells` / `unmerge_cells` 内部完成写操作后自动调用。

---

## 5. OOXML 规范摘要

| 操作 | XML |
|------|-----|
| 横向 span N | `w:tcPr/w:gridSpan @w:val=N`，删除被吞并的 `w:tc` |
| 纵向起点 | `w:tcPr/w:vMerge @w:val=restart` |
| 纵向延续 | `w:tcPr/w:vMerge`（无 val 或 continue） |
| 拆分 | 移除上述属性，补回空 `w:tc` |

空单元格模板：`<w:tc><w:tcPr/><w:p/></w:tc>`

**约束**

- `ensure_tc_pr!` 新建的 `w:tcPr` 必须 **prepend 到 `w:tc` 首位**（排在 `w:p` 之前），否则 Word 视为损坏。
- 横向 `gridSpan` / 纵向 `vMerge` **均不改动 `w:tblGrid`**，故 `column_count`（取 `gridCol` 数）与 `row_count`（取 `w:tr` 数）在 merge/unmerge 前后保持不变。
- grid 宽度以**行内 `gridSpan` 累加**为准；若某行累加与 `tblGrid/gridCol` 数不一致（少数文档存在），取两者较大值并 `warn`，不盲信 `tblGrid`。

---

## 6. 实施阶段

### Phase 1 — 逻辑网格（P0）

**任务**

- [ ] **前置**：新建 fixture 生成脚本（如 `spec/support/build_table_fixtures.rb`），通过写 `document.xml` + 打 zip 构造 docx，使 fixture 可在 CI 复现、`gridSpan`/`vMerge` 标注可控
- [ ] 新建 `table_grid.rb`，实现 `build_slots`（占位存 anchor 引用，见 §4.3） / `cell_at`（越界返回 `nil`）
- [ ] `Table#cell_at`、`Table#each_cell`、`invalidate_grid!`
- [ ] 修正 `Table#columns` 使用 grid（子轴 `w:tr/w:tc`，语义见 §4.4）
- [ ] 生成 fixture：`horizontal_merge.docx`、`vertical_merge.docx`、`rect_merge.docx`
- [ ] 新增 `spec/docx/containers/table_grid_spec.rb`（含 `columns[col].cells[row] == cell_at(row,col)` 断言）

**验收**

- 合并 fixture 的逻辑坐标与 Word 一致
- `tables.docx` 现有测试通过

---

### Phase 2 — 读取 API（P1）

**任务**

- [ ] `TableCell` 增加 `colspan` / `rowspan` / `merged?` / `merge_anchor?` / `merge_continuation?`
- [ ] `Table#merged?(row, col)`
- [ ] 新增 `spec/docx/containers/table_merge_read_spec.rb`

**验收**

- 读属性与 fixture 标注一致

---

### Phase 3 — 横向合并（P2-a）

**任务**

- [ ] 新建 `table_merge.rb`
- [ ] 实现 `ensure_tc_pr!`、`set_grid_span`
- [ ] `merge_cells` 支持单行（`row0 == row1`）
- [ ] 新增 `spec/fixtures/tables/plain_3x3.docx`
- [ ] 新增 `spec/docx/containers/table_merge_write_spec.rb`（横向部分）

**验收**

- merge → save → reopen，`colspan` 正确
- Word 可打开

---

### Phase 4 — 矩形合并（P2-b）

**任务**

- [ ] `merge_cells` 支持多行多列
- [ ] `set_vmerge`、`ensure_continuation_cells`
- [ ] 重叠检测 → `MergeConflict`
- [ ] 扩展 write spec（2×2、3×1、1×3）

**验收**

- 矩形 merge round-trip 通过

---

### Phase 5 — 拆分（P3）

**任务**

- [ ] `unmerge_cells` / `TableCell#unmerge!`
- [ ] 横向补 `w:tc`、纵向恢复独立格
- [ ] 新增 `spec/docx/containers/table_unmerge_spec.rb`

**验收**

- 合并 → 拆分后逻辑列数 == `column_count`
- anchor 文本保留

---

### Phase 6 — 文档与发布（P4）

**任务**

- [ ] README「合并单元格」章节
- [ ] CHANGELOG（含 `columns` 行为修正说明）
- [ ] 版本号 bump（建议 0.11.0）

---

## 7. 测试清单

### Fixture

均由生成脚本产出（见 Phase 1 前置），不手工 Word 制作，保证 CI 可复现。

| 文件 | 用途 |
|------|------|
| `spec/fixtures/tables/horizontal_merge.docx` | 横向合并读 |
| `spec/fixtures/tables/vertical_merge.docx` | 纵向合并读 |
| `spec/fixtures/tables/rect_merge.docx` | 矩形合并读 |
| `spec/fixtures/tables/plain_3x3.docx` | 写操作基础表 |

### 必测场景

- [ ] `cell_at` vs 物理 `cells` 索引
- [ ] `columns` 与 `cell_at(r,c)` 一致
- [ ] merge 后 `row_count` / `column_count` 不变
- [ ] 越界 → `InvalidMergeRange`
- [ ] 重叠 → `MergeConflict`
- [ ] save → reopen 内容一致
- [ ] `document_spec.rb` 回归

---

## 8. 风险与缓解

| 风险 | 缓解 |
|------|------|
| `tblGrid` 与行内占位不一致 | merge 前后校验每行网格占用 |
| 插入 `w:tc` 位置错误 | 仅通过 `TableGrid` 定位 |
| `columns` 行为变更 | CHANGELOG 标注 bugfix |
| 延续格被误写 | 文档说明；可选对写操作 warn |

---

## 9. 工作量估算

| 阶段 | 估时 |
|------|------|
| Phase 1 | 2–3 天 |
| Phase 2 | 1 天 |
| Phase 3 | 2 天 |
| Phase 4 | 3 天 |
| Phase 5 | 2–3 天 |
| Phase 6 | 0.5 天 |
| **合计** | **约 10–12 天**（fixture 调试 + Word/LibreOffice 双端 round-trip 验证易超预期，建议预留 2–3 天 buffer，实际按 13–15 天排期） |

代码量约 **700–1000 行**（含 spec 与 fixture 生成脚本）。

---

## 10. 评审确认项

- [ ] 接受 `Table#columns` 行为修正（bugfix）
- [ ] 同意 Phase 1→2→3 先读后写顺序
- [ ] Phase 5 unmerge 纳入同版本或后续版本
- [ ] 目标版本 0.11.0

---

## 11. 后续扩展（不在本期）

- `anchor.merge_with(other_cell)` 对象式 API
- 合并时内容迁移策略可配置
- 与 `feature/api-table-image-replace` 分支协同
