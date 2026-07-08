# 跨文档正文导入 — 实施计划（Implementation Plan）

> 配套需求文档：[`cross-document-import.md`](./cross-document-import.md)
>
> 本文件是把该需求转成**可执行的落地计划**：任务拆分、依赖顺序、每步验收、风险前置验证。
> 术语沿用需求文档：目标文档 = target，源文档 = source。

---

## 0. 现状复核结论（决定工作量的关键）

核对 `lib/docx/document.rb` 后，需求文档 §4「R1 现状」有几处已过时，实际工作量比预估小。R1 主要是**泛化现有能力 + 补两个缺口**，而非从零搭建：

| 需求点 | 真实现状 | 结论 |
|---|---|---|
| R1.1 写出新增条目 | `save`/`stream` 末尾已 flush `@replace` 中原 zip 不存在的条目（`document.rb:244-249, 272-277`） | **已实现** |
| R1.4 rels 回写 | `update` 已 `replace_entry @rels_path, @rels.serialize`（`document.rb:672`） | **已实现** |
| R1.6 rId 唯一 | `next_relationship_id` 扫描最大 rIdN 自增（`document.rb:1235`） | **已实现**，需泛化 |
| R1.5 底层 API | `add_image_relationship` / `ensure_content_type_default` / `next_generated_media_path` 存在，但私有且图片专用 | **需泛化** |
| R1.2 styles 回写同源 | `update` 走 `styles_configuration`（`@styles.dup` 的 memo，`document.rb:520`），直接改 `@styles` 会丢 | **需修** |
| R1.3 numbering 新建 | 仅加载时存在才回写；无法新建并注册 | **需补** |

---

## 1. 阶段总览与顺序

```
P0 Spike（风险前置，~0.5d）
        │
        ▼
P1 R1 存盘链路 + R2 样式 + R3 编号 + R5(样式/编号引用)  ← 覆盖绝大多数冲突，可独立发版
        │
        ▼
P2 R4 媒体/关系 + R5(rId)                              ← 支持内嵌图片片段
        │
        ▼
P3 R5(书签/批注) + R2 等价比较放宽 + 性能索引 + R6 收尾
```

每个 P 结束都应绿测且能被 Word/LibreOffice 正常打开。

---

## 2. P0 — Spike（✅ 已完成 2026-07-08）

**目标：验证 source zip 生命周期。** `Document#initialize` 在 `ensure` 里 `@zip.close`（`document.rb:73-75`），R4 需要在导入时从 source 读 media 二进制。

**验证方式**：用 rubyzip 造含 `word/media/image1.png` 的临时 docx，分别以路径 / buffer(StringIO) 打开后（此时 `@zip` 已 close）尝试 `doc.zip.read(...)`，并测「open 后、read 前删除源文件」的边界。环境：ruby 3.4.6 + rubyzip **3.4.1**。

**结论**：

- [x] S0.1 构造后 `zip.read('word/media/image1.png')` **可成功**（路径与 buffer 两种都 OK，各读到 89838 bytes）。`@zip.close` 后 `Zip::File#read` 仍可用。
- [x] S0.2 路径 / buffer 均可读普通条目与 media。
- [x] S0.3 **关键边界**：路径打开的文档，`read` 是**惰性按路径重开文件**——若在 `read` 前删除/移动源文件，会 `Errno::ENOENT`。buffer 打开则数据在内存、不受影响。

**定稿策略（据此实现 R4）**：
> `Importer.new(target, source)` 在**构造时即遍历 source rels、预读并缓存所有被引用的 media bytes**（`source.zip.read` 存进内存 Hash），import 阶段只用缓存。这样：①与 source 文件生命周期彻底解耦（删文件/内存流都安全）；②同一 media 多 rId 引用天然去重；③避免惰性重开文件的 I/O。**不采用** import 时才惰性读的方案。

---

## 3. P1 — 样式/编号隔离（核心价值）✅ 已完成 2026-07-08

> 落地记录：R1a→R2→R3→R5(part)→R6(part) 逐任务实现并各自 commit（分支 feature/cross-document-import）。
> 全量测试 **277 examples, 0 failures**，多 seed（1/42/60631/12345）稳定。
> 新增文件：`lib/docx/merge/{importer,styles_importer,numbering_importer,node_rewriter}.rb`；
> `Docx::Document` 新增 `add_part`/`add_relationship`/`ensure_default_content_type`/`ensure_numbering!`/`import_node`/`import_before`/`import_after`。
>
> **P1 期间发现、留待后续的已知缺口（不影响 P1 验收）**：
> - numbering 的 `w:abstractNum/w:lvl/w:pStyle` 引用的段落样式未随 style_id_map 改写（R3 仅改 styleLink/numStyleLink）。若该 pStyle 在 target 被改名会悬空 → 归入 P3。
> - numbering 的 numStyleLink/styleLink 指向的 styleId，若未被正文直接引用则不会被 StylesImporter 导入 → 归入 P3（编号→样式反向依赖闭包）。
> - R2 latentStyles 合并未做（保留 target 的）→ 需求已允许，P3 视需要再补。

### R1a 存盘链路补全（前置）

改 `lib/docx/document.rb`：

- [x] R1a.1 新增通用 `add_part(zip_path, content_type, bytes)`：写入 `@replace[zip_path] = bytes`；若给了 `content_type` 则登记 `[Content_Types].xml` 的 `Override`（区别于按扩展名的 `Default`）。
- [x] R1a.2 新增通用 `add_relationship(type, target, mode: nil) -> rId`：复用 `next_relationship_id` 的自增逻辑，支持 `TargetMode="External"`；返回新 rId。把现有 `add_image_relationship` 改为它的薄封装。
- [x] R1a.3 新增 `ensure_default_content_type(ext, content_type)`：泛化现有 `ensure_content_type_default`，扩展名表补 `emf/wmf/tiff` 等。
- [x] R1a.4 **styles 回写同源**：让 `update` 直接序列化 `@styles`，或让 `styles_configuration` 持有 `@styles`（非 dup）。目标：importer 改 `@styles` DOM 后能被写出。加一条回归测试防丢改。
- [x] R1a.5 **numbering 新建**：新增 `ensure_numbering!`——目标无 `@numbering` 时创建骨架 `<w:numbering>`，通过 `add_part` 注册 content-type Override + 通过 `add_relationship` 加 `numbering` 关系；`update` 需能写出新建的 `@numbering`。

**验收**：单元测试——空白目标文档 `add_part` 一个新条目后 `save`，解压能读到该条目；`@styles`/`@numbering` 改动能落盘。

### R2 样式导入器

新增 `lib/docx/merge/styles_importer.rb`：

- [x] R2.1 建 target styles 的 `id→node` 索引（避免 O(n²)）。（`name→node` 索引未做——本实现不依赖 name 查找，故省略。）
- [x] R2.2 按需闭包导入：入口是正文实际引用的 styleId，沿 `basedOn`/`next`/`link` 递归（最多 10 层防环）。
- [x] R2.3 去重三分支：同 id 且等价→复用；同 id 不等价→改名（`m{seq}_{原id}`）导入；无此 id→原样导入。产出 `style_id_map`。
- [x] R2.4 等价比较：P1 用规范化（canonicalize）比较，P3-3 进一步放宽（剥空白 + 排序属性）。
- [x] R2.5 依赖改写：导入的 style 内 `basedOn/next/link` 按 `style_id_map` 改写；`link` 两端同进退。
- [x] R2.6 顺序合法：`docDefaults → latentStyles → style*`，导入的 style 追加到 `style*` 末尾；不动 target 的 `docDefaults`；`latentStyles` 保留 target 的。

### R3 编号导入器

新增 `lib/docx/merge/numbering_importer.rb`：

- [x] R3.1 target 无 numbering → 调 `ensure_numbering!`。
- [x] R3.2 `abstractNum` 偏移到 `target_max_abstract_id + 1`，产出 `abstract_num_id_map`；改写内部 `numStyleLink`/`styleLink`（P3-2 起并含 `lvl/pStyle`，且触发样式导入闭包）。
- [x] R3.3 `num` 偏移到 `target_max_num_id + 1`，产出 `num_id_map`；改写 `abstractNumId/@w:val`（用 `abstract_num_id_map`）。
- [x] R3.4 实现为按需导入被引用的 num 及其 abstractNum 链（`import(num_id)`），另提供 `import_all` 全量导入。
- [x] R3.5 `w:numbering` 子元素顺序合法（`abstractNum*` 在前、`num*` 在后）。

### R5(part) 正文引用改写（样式/编号）

新增 `lib/docx/merge/node_rewriter.rb`：

- [x] R5.1 在 `node.dup` 之后、插入 target 之前，遍历子树。
- [x] R5.2 改写 `pStyle/rStyle/tblStyle @w:val`（用 `style_id_map`）、`numPr/numId @w:val`（用 `num_id_map`）；未命中映射保持原值。
- [x] R5.3 命名空间安全：`dup(1)` 后前缀 `w:` 正确，序列化后无重复 `xmlns`（有断言）。

### R6(part) + 公开 API（P1 版）

- [x] R6.1 新增 `lib/docx/merge/importer.rb`：`Importer.new(target, source)` 持有各映射表（P2/P3 起含 rid_map、bookmark_id_offset）。
- [x] R6.2 `Importer#import(node)`：调 styles→numbering→media importer 建映射 → `dup` → node_rewriter 改写 → 返回在 target 上下文、可 `insert_before/after` 的节点。
- [x] R6.3 `Document#import_node/import_before/import_after`：按 source 对象缓存 importer。
- [x] R6.4 `lib/docx.rb` + autoload 新文件。
- [x] R6.5 失败回退：`import_node` 异常时 `warn` + 回退到 `dup`，由 `fallback:` 开关控制。

**P1 验收（对应需求 §6 的 1/2/3/4/7/8）**：styleId 冲突改名、styleId 缺失原样导入、同名等价复用、numbering 偏移正确、多次导入映射复用不重复、产物无悬空 style/num 引用且 Word 可开。

---

## 4. P2 — 媒体/关系（内嵌图片）✅ 已完成 2026-07-08

> 落地记录：P2-1（`lib/docx/merge/media_importer.rb` + NodeRewriter rId 改写）→ P2-2（Importer 接入 MediaImporter + 含内嵌图片端到端集成测试），各自 commit。
> 全量 **291 examples, 0 failures**，默认 + seed 60631 稳定。
> MediaImporter 构造时预读缓存 source media（落实 P0 定稿策略）；内部图片不重名落盘 + content-type + 新 rId，外部关系走 External 不复制，media 按 source 路径去重。
> 端到端断言：save→reopen 后 media 字节一致、无悬空 rId（`reopened.images` 可解析）、r:/a: 命名空间正确。
>
> **遗留（非阻断）**：P2-1/P2-2 的两个 merge spec 在顶层定义了同名常量（FIXTURES_PATH/RELS_NS 等）产生 `already initialized constant` 告警（与仓库既有 XML_NS 同类噪音），可在 P3 一并清理。

新增 `lib/docx/merge/media_importer.rb`：

- [x] R4.1 收集被导入子树引用的 `r:embed`/`r:id`/`r:link`（在 Importer 编排里收集）。
- [x] R4.2 在 source rels 按 rId 找 `Target`/`Type`。
- [x] R4.3 内部关系（图片）：构造时预读缓存 bytes → 以不重名文件名 `add_part` → `ensure_default_content_type` → `add_relationship` 得新 rId → 登记 `rid_map`。
- [x] R4.4 外部关系（超链接 `TargetMode=External`）：只 `add_relationship(type, 原Target, mode: :external)`，不复制 media。
- [x] R4.5 同一 media 多 rId 引用按 source 路径去重（映射到同一新 rId）。
- [x] R5.4 node_rewriter 增加 `@r:embed/@r:id/@r:link` 改写（r 命名空间属性），保持前缀正确、有断言。

**P2 验收（需求 §6.5）**：内嵌图片 source 导入后 media 落盘、rels/content-type 注册、正文 embed 指向新 rId，Word 正常显示。

---

## 5. P3 — 收尾与健壮性 ✅ 已完成 2026-07-08

> 落地记录：P3-1（书签 id 隔离）→ 既有 spec Tempfile flaky 稳定化 → P3-2（编号→样式依赖闭包）→ P3-3（等价比较放宽 + 幂等/latentStyles 测试 + 常量告警清零），各自 commit。
> 全量 **307 examples, 0 failures**，默认 + seed 1 + seed 60631 稳定，且 `already initialized constant` 告警清零。

- [x] R5.5 书签 id：`bookmarkStart/bookmarkEnd @w:id` 按 `bookmark_id_offset`（= target 现有最大 id + 1）偏移，成对一致。（批注/尾注引用 id 涉及独立 comments/footnotes 部件导入，未在本轮实现——见下方遗留。）
- [x] R2.7 等价比较放宽：`normalize_style_xml` 比较前剥离空白 + 排序属性，忽略属性顺序/无意义空白。
- [x] Perf.1 styles/numbering/media 的 id→node 索引与映射均在各 importer 内建立并按 source 缓存（P1/P2 即已具备）。
- [x] R6.6 幂等复测：同一 importer 重复导入同 node → 无重复 style/num/media 定义。
- [x] 编号→样式反向依赖闭包（P3-2）：abstractNum 的 styleLink/numStyleLink/lvl·pStyle 触发样式导入并改写。
- [x] latentStyles：保留 target 的、不导入 source 的（需求允许），有测试固化。
- [x] 测试卫生：既有 Tempfile GC flaky 修复 + 常量告警清零。

**P3 验收（需求 §6.6/7）**：书签 id 无冲突 ✅；产物无重复 styleId/numId ✅。

### 未实现 / 后续可选（超出 R1–R6 原始范围）
- **批注/尾注（comments.xml / footnotes.xml）整体导入**：正文里的 `w:commentReference` 等 id 偏移需连同独立部件一起导入，属新部件类型，本轮未做。
- **R2 latentStyles 深度合并**（当前仅保留 target）。
- Word/LibreOffice 真机打开校验（需求 §6.8 的人工/GUI 验证部分）：已用 save→reopen + 无悬空引用断言覆盖结构合法性，真机视觉校验需人工。

---

## 6. 交付物（文件清单）

| 文件 | 阶段 | 改动 |
|---|---|---|
| `lib/docx/document.rb` | P1/P2 | `add_part`/`add_relationship`/`ensure_default_content_type`/`ensure_numbering!`；styles 回写同源；`import_node`/`import_before`/`import_after` |
| `lib/docx/merge/importer.rb`（新增） | P1 | 编排 + 四映射表 |
| `lib/docx/merge/styles_importer.rb`（新增） | P1 | R2 |
| `lib/docx/merge/numbering_importer.rb`（新增） | P1 | R3 |
| `lib/docx/merge/node_rewriter.rb`（新增） | P1/P2/P3 | R5 |
| `lib/docx/merge/media_importer.rb`（新增） | P2 | R4 |
| `lib/docx.rb` | P1 | require/autoload 新文件 |
| `spec/docx/merge/*_spec.rb`（新增） | 各阶段 | 需求 §6 用例（document_parts / styles_importer / numbering_importer / node_rewriter / importer / importer_image / importer_bookmark） |
| ~~`spec/fixtures/merge/*.docx`~~ | — | 改为 spec 内基于 basic.docx 骨架动态构造，未新增二进制 fixtures（见 §7） |

---

## 7. 测试 fixtures（实现方式已调整）

> **调整说明**：最终**未**新增二进制 `.docx` fixtures，而是在各 merge spec 里用「基于 `spec/fixtures/basic.docx` 骨架、按需覆盖/追加 `document.xml`/`styles.xml`/`numbering.xml`/`rels`/`media` 条目、经 `StringIO` open」的方式**动态构造** target/source。好处：场景可精确控制、可读、无需维护二进制样本。下列验收场景均已由该方式覆盖。

- [x] 同名 `tblStyle` 一无边框一有边框（验收 §6.1）——`importer_spec.rb` / `styles_importer_spec.rb`。
- [x] 带大纲编号 source + 无 `numbering.xml` 的 target（验收 §6.4）——`numbering_importer_spec.rb` / `importer_spec.rb`。
- [x] 内嵌图片 source（验收 §6.5）——`importer_image_spec.rb` / `media_importer_spec.rb`（图片二进制用 `spec/fixtures/replacement.png`）。
- [x] 同 id 书签的 source+target（验收 §6.6）——`importer_bookmark_spec.rb`。

---

## 8. 风险登记（按优先级）

1. ~~**source zip 生命周期**（唯一真正未知数）~~ —— ✅ **已 spike 解决**：`@zip.close` 后仍可 `read`；路径打开为惰性重开文件，故 importer 构造时预读缓存所有 media bytes（见 §2）。
2. **Nokogiri 带前缀属性改写/序列化**（P2 R5.4）—— 已知有坑，重点测 `r:embed`。
3. **styles_configuration memo 一致性**（P1 R1a.4）—— 不处理会静默丢改。
4. **样式等价判断**（P1 简化 / P3 放宽）—— 已有降级路径，风险可控。
5. **跨文档 dup 的命名空间**（P1 R5.3）—— 以 target `@doc` 为上下文构造。

---

## 9. 立即可执行的第一步

先做 **P0 Spike**（第 2 节）。结论出来后，若确认 media 读取策略，即按 P1 顺序：R1a → R2 → R3 → R5(part) → R6(part)。
