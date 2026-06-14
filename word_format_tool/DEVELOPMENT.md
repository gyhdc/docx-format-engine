# 开发接手说明

## 定位

`word_format_tool` 是独立包，只处理本地 DOCX、PDF、JSON 和 Markdown。底层模块
不得调用网络、Agent 或 SDAU-agent 业务代码。Microsoft Word、PyMuPDF 和
pywin32 仅属于可选视觉验证路径。

## 依赖方向

```text
models / constants / exceptions
        ↓
document_io / ooxml_utils / layout_safety
        ↓
document_analyzer / template_profiler / reference_links / toc_repair
        ↓
format_inspector / structure_tools / paragraph_locator / structure_inspector
        ↓
format_fixer / structure_operations / layout_validator / optional_word_backend
        ↓
public API (__init__) / CLI / report_writer
```

禁止底层模块反向依赖 CLI。

## 0.6 模块状态

| 模块 | 职责 |
|---|---|
| `models.py` | 细分角色、规则、issue、coverage 和 report schema |
| `rule_loader.py` | JSON/正则校验及旧角色兼容回退 |
| `document_io.py` | 防覆盖、原子发布、内容/导航/布局对象指纹 |
| `ooxml_utils.py` | 格式读取、字体写入、段落和页面设置 |
| `document_analyzer.py` | 正文、表格、页眉页脚、标题和文献语言识别 |
| `layout_safety.py` | 封面表格、浮动对象、文本框和 VML 风险 |
| `template_profiler.py` | 模板样本、结构清单和布局风险事实 |
| `reference_links.py` | 文献编号、引用、书签、上标和双向跳转 |
| `toc_repair.py` | 唯一匹配的 `PAGEREF` 修复及全量刷新门禁 |
| `format_inspector.py` | 格式、页面、题注、目录、文献和 coverage 检查 |
| `format_fixer.py` | 支持范围内格式、文献和目录安全修复 |
| `structure_tools.py` | 模板/初稿结构计划及显式缺失章节补建 |
| `paragraph_locator.py` | 基于文本、角色、上下文和序号的稳定段落重定位 |
| `structure_inspector.py` | 分节、目录边界、字段、页眉页脚和段落分页事实 |
| `structure_operations.py` | 原子、严格、幂等的分节/页眉页脚/页码/分页操作 |
| `layout_validator.py` | 无图像依赖的结构布局断言与修复建议 |
| `optional_word_backend.py` | 可注入的 Word 字段刷新和真实分页适配器 |
| `visual_compare.py` | 可选 Word 导出、PDF 渲染和逐页视觉比较 |
| `report_writer.py` | JSON/Markdown 报告和未覆盖清单 |
| `cli.py` | `wordfmt` 细粒度命令入口 |
| `__init__.py` | Agent 稳定 API |

## 角色和定位

正文段落继续使用原 `paragraph_index`。新增：

- `story`: `body`、`table_cell`、`header`、`footer`；
- `location`: 例如 `body.paragraph[3]` 或
  `table[0].row[1].cell[0].paragraph[0]`。

细分角色：

- `title_zh`、`title_en`；
- `reference_entry_zh`、`reference_entry_en`；
- `cover_label`、`cover_value`、`table_text`；
- `header`、`footer`、`table_of_contents`；
- `acknowledgements`、`appendix`。

表格、页眉页脚不回退到 `body`，避免误改。

## 安全不变量

1. 输出路径不得与输入路径相同。
2. 普通 `fix` 不增加、删除或重排可见正文。
3. 不删除图片、表格、公式或原有导航对象。
4. 已有浮动/内联对象的位置、尺寸和环绕指纹必须保持。
5. `unknown` 和 `protected_front_matter` 不应用正文格式。
6. 封面表格、浮动对象或文本框存在时，页面几何默认不可修复。
7. 页面风险覆盖必须显式设置 `allow_unsafe_page_geometry=true`。
8. 目录目标只在唯一标题匹配时修复。
9. 目录存在任何无法修复目标时，不设置全局 `updateFields`。
10. 结构补建只执行 `rules.structure.required_sections`。
11. 不自动复制模板正文、文本框、浮动对象或任意块。
12. 无可修复项时原样复制输入包，避免无意义重序列化。
13. 所有本机路径由调用方传入，源码、示例和文档不得硬编码。
14. 结构操作每一步执行前重新解析选择器，不保存易失效的 Word Range。
15. 选择器匹配缺失或歧义时整批失败，输出文件不得发布。
16. Word 自动化只负责最终字段和分页事实，核心规则操作保持纯 OOXML 可测试。
17. 正文分节不得位于未闭合的 TOC 域内；先把外层 TOC `end` 移到分节符之前。
18. Word 字段刷新前后必须比较受保护结构指纹，变化时拒绝发布。

## 兼容策略

- schema 版本继续为 `0.1`，新增字段均有安全默认值。
- `title_zh/title_en -> title`。
- `reference_entry_zh/reference_entry_en -> reference_entry -> reference`。
- `reference_heading -> heading_1`。
- 旧 CLI 和五个原公开 API 保持可用。

## 测试重点

- 无样式中英文封面标题和元数据保护。
- 旧 Title 样式按语言拆分。
- 中英文文献差异格式和引用跳转。
- 表格单元格、页眉页脚显式格式修复。
- 普通正文表格不会误判为封面。
- 表格清单跨重复分析保持稳定。
- 页面几何风险阻断和显式覆盖。
- 浮动对象位置变化阻止原子发布。
- 唯一目录目标修复、部分修复禁止全局刷新。
- 结构计划和显式章节补建。
- 目录边界后的稳定段落定位、缺失和歧义选择器。
- 分节、页眉页脚、页码字段和段落分页操作的幂等性。
- 操作失败时不发布输出文件。
- 无图像布局断言和可注入 Word 字段刷新后端。
- TOC 外层域跨正文标题时的域结束迁移，以及 Word 刷新结构指纹安全门。
- 模板异常倍数行距过滤。
- coverage JSON/Markdown 输出。
- 视觉核心的页数、尺寸和像素差异。
- 全部既有测试兼容。

## 当前验证状态

验证日期：2026-06-14。

- Python：默认解释器和项目环境。
- pytest：83 项通过。
- CLI：模板、规则、检查、修复、结构、文献和视觉命令已验证。
- 真实初稿：识别 1 个中文题目、1 个英文题目、5 条中文文献、
  7 条英文文献和 23 个表格单元格段落。
- 真实修复：301 个问题中 279 个确定性问题全部修复；22 个风险或未覆盖项保留。
- 页面安全：2 个封面前置表格阻止页边距自动修改。
- 内容安全：正文、表格结构、媒体、公式、块顺序和布局对象指纹通过。
- Word/PDF 视觉回归：前后均为 9 页 A4；封面页像素差异为 0。
- 人工视觉抽查：中文/英文题目正常居中，封面未竖排，目录未因部分刷新丢项。
- 真实 Word 字段刷新：23 个字段、1 个目录，刷新后仍为 2 节和 9 页。
- TOC 域安全：修复外层 TOC 域结束标记跨入正文导致 Word 吞分节的问题。

## 仍然明确不自动完成

- 文本框和浮动对象内部格式编辑。
- 无映射信息时复制复杂封面或任意模板块。
- 无法唯一匹配的目录书签。
- Word 实际分页语义的纯 OOXML 推断。
- 学术内容补写、文献事实检索和引用编号猜测。

这些内容必须继续出现在 coverage/结构/视觉报告中，不得由上层 Agent 隐藏。
