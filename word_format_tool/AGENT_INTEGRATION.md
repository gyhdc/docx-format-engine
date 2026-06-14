# Agent / Programmer Integration Guide

本文档是 `word-format-tool` 的权威接入说明，面向调用该包的程序员和 Agent。

- 工具版本：`0.6.1`
- 规则版本：`0.1`
- Python：`>=3.10`
- 核心原则：优先组合细粒度原语，不默认对全文执行一键套版

## 能力边界总览

工具负责确定性的格式、导航和结构操作，不负责猜测未知模板语义。它可以检查和修改
已识别段落、节、页眉页脚、页码字段和分页控制，但不自动复制模板正文、复杂封面、
文本框或浮动对象。需要 Word 最终分页事实时使用可选 `refresh_fields()`。

`refresh_fields()` 在临时副本上运行，并比较刷新前后的分节数、正文首标题所属节、
页眉、页码字段和页码起始值。Word 更新目录若吞掉分节符，工具会抛出
`FieldRefreshIntegrityError`，且不会发布输出文件。

## 按能力调用

格式偏差使用 `inspect_document()` / `fix_document()`；模块局部对齐使用
`align_modules_from_template()`；结构收尾使用 `inspect_structure()`、
`apply_structure_operations()` 和 `validate_layout()`；字段和真实分页最后使用
`refresh_fields()`；视觉比较只作为最终风险抽查。

### 1. 提取模板事实

- API / CLI：`profile_template()` / `wordfmt profile-template`
- 能做：提取页面、样式、标题、目录、页眉页脚等可确定的模板事实。
- 不能做：不自动生成可靠的 `rules.json`，Agent 必须审核候选事实和异常值。
- 验收：检查 `TemplateProfile` 的样本数、置信度和未识别项，再决定是否生成规则。

### 2. 校验规则

- API / CLI：`validate_rules()` / `wordfmt validate-rules`
- 能做：按严格模型校验规则字段、类型、范围和枚举值。
- 不能做：不能证明规则符合用户真实意图，也不会修正语义错误。
- 验收：返回通过且不存在未知字段，再进入检查或修改阶段。

### 3. 检查与预演

- API / CLI：`inspect_document()` / `wordfmt dry-run`
- 能做：只读分析初稿并报告候选改动、未覆盖项和风险。
- 不能做：不会修改文档，也不能替代 Word 的真实分页结果。
- 验收：确认 `FormatReport` 的改动范围和未覆盖项符合本次任务边界。

### 4. 执行格式修复

- API / CLI：`fix_document()` / `wordfmt fix`
- 能做：按已审核规则修复受支持的页面、段落、字体、表格和标题格式。
- 不能做：不自动复制模板正文，不编辑文本框或浮动对象，也不覆盖原始输入文件。
- 验收：输出文件存在，报告无致命错误，并再次执行检查确认主要偏差已消除。

### 5. 仅处理文献跳转

- API / CLI：`link_references()` / `wordfmt link-references`
- 能做：在正文引用与参考文献条目之间建立可验证的内部跳转。
- 不能做：不会改写引用内容，也不猜测无法唯一匹配的文献。
- 验收：检查已链接、未匹配和歧义引用清单，歧义项必须人工处理。

### 6. 规划文档结构

- API / CLI：`plan_structure()` / `wordfmt plan-structure`
- 能做：识别缺失的摘要、关键词、目录、正文和参考文献等结构角色。
- 不能做：不会凭空生成可靠的学术内容，也不会直接改变原文。
- 验收：审核 `StructurePlan` 的角色识别和缺失项，排除目录条目造成的误判。

### 7. 补建缺失章节

- API / CLI：`complete_structure()` / `wordfmt complete-structure`
- 能做：按明确计划补建受支持的缺失章节占位结构。
- 不能做：不能替代作者写作，复杂封面、文本框和未知版式仍需单独处理。
- 验收：检查 `StructureCompletionResult`，并重新运行结构检查确认没有重复章节。

### 8. 视觉回归检查

- API / CLI：`compare_visual()` / `wordfmt visual-compare`
- 能做：比较渲染页的尺寸、像素差异和高风险页面，辅助发现串页和明显错位。
- 不能做：不能仅凭视觉相似度证明语义、字段或引用正确。
- 验收：查看 `VisualComparisonReport` 的差异页，并人工复核封面、目录和正文首页。

## profile → rules 映射

`TemplateProfile` 提供模板事实，`Rules` 表达 Agent 审核后的确定性格式要求。
`build_template_rules()` 可按模块生成候选规则，但异常倍数行距会被过滤，复杂任务仍应
由 Agent 审核后再执行。

## API 返回契约

格式接口返回 `FormatReport`，结构规划返回 `StructurePlan` 和
`StructureCompletionResult`，视觉比较返回 `VisualComparisonReport`。结构增强使用
`ParagraphSelector`、`StructureOperationPlan` 和 `LayoutExpectations`，所有公开返回
均可直接 JSON 序列化。

## 错误合约

预期失败均继承 `WordFormatToolError`。结构定位或后置条件失败抛出
`StructureOperationError`；Microsoft Word 不可用或刷新失败抛出
`WordAutomationUnavailableError`；刷新改变受保护结构时抛出
`FieldRefreshIntegrityError`。失败的结构操作和字段刷新不发布半成品。

## Agent 端到端流程

```text
convert / compare-modules
  -> derive-rules / fix
  -> inspect-structure
  -> apply-operations
  -> refresh-fields（可选）
  -> validate-layout（刷新后 update_fields_enabled 通常为 false）
  -> visual-compare（按风险选择）
```

## 1. 设计定位

Word 模板与初稿的差异通常同时涉及内容结构、段落格式、页面设置、目录域、表格和
浮动对象。单次自动操作无法可靠覆盖所有情况。

本工具把工作拆成五类能力：

1. **识别**：转换文件、识别模块边界、比较模板与初稿。
2. **提取**：从模板提取指定模块或页面的格式规则。
3. **执行**：只对齐指定模块，或按显式 `rules.json` 修复。
4. **验证**：输出检查报告、内容完整性校验和可选视觉比较。
5. **结构收尾**：稳定定位正文段落，执行分节、页眉页脚、页码和分页控制，
   并用机器可读断言验证结果。

Agent 负责：

- 判断学校规范和模板说明文字的真实含义。
- 审核模块识别和自动提取的规则。
- 决定本轮只处理哪些模块。
- 处理工具未覆盖的封面布局、文本框、浮动对象和复杂分页。

工具负责：

- 确定性识别、格式应用、原子保存和报告生成。
- 避免覆盖输入文件。
- 保存前验证文字、图片、表格、公式和对象顺序未被意外改变。

## 2. 安装与调用入口

```powershell
python -m pip install -e "<WORD_FORMAT_TOOL_ROOT>"
```

已经安装过依赖、只需要刷新当前可编辑包时，可加 `--no-deps`。

开发和测试环境：

```powershell
python -m pip install -e "<WORD_FORMAT_TOOL_ROOT>[test]"
```

Word/PDF 视觉比较，以及 Windows 上通过 Microsoft Word 转换旧 `.doc`：

```powershell
python -m pip install -e "<WORD_FORMAT_TOOL_ROOT>[visual]"
```

如果 `wordfmt` 不在 `PATH`，统一使用以下稳定入口：

```powershell
python -m word_format_tool.cli --help
```

Python API：

```python
from word_format_tool import (
    align_modules_from_template,
    apply_structure_operations,
    build_template_rules,
    compare_modules,
    compare_visual,
    complete_structure,
    convert_to_docx,
    detect_modules,
    fix_document,
    format_from_template,
    inspect_document,
    inspect_structure,
    link_references,
    plan_structure,
    profile_template,
    refresh_fields,
    rules_schema,
    validate_layout,
    validate_rules,
)
```

所有路径接受 `str` 或 `pathlib.Path`。公开函数返回可直接 JSON 序列化的 `dict`
或 `bool`。

## 3. 输入类型

| 能力 | `.doc` | `.docx` |
|---|---:|---:|
| `convert_to_docx()` / `convert` | 支持 | 支持，复制为新文件 |
| `format_from_template()` / `format-template` | 支持，内部临时转换 | 支持 |
| 其他识别、提取、修复 API | 不支持 | 支持 |

`.doc` 转换后端按以下顺序尝试：

1. Microsoft Word COM，需要 Windows、Microsoft Word 和 `pywin32`。
2. LibreOffice 命令行，需要 `soffice` 或 `libreoffice` 可执行文件位于 `PATH`。

两个后端都不可用时抛出 `DocumentReadError`。

## 4. 能力矩阵

| 任务 | Python API | CLI | 修改文档 |
|---|---|---|---:|
| 转换 `.doc/.docx` | `convert_to_docx()` | `convert` | 写新 DOCX |
| 识别单个文档模块 | `detect_modules()` | `detect-modules` | 否 |
| 比较模板与初稿模块 | `compare_modules()` | `compare-modules` | 否 |
| 按模块提取模板规则 | `build_template_rules()` | `derive-rules` | 否 |
| 只对齐选定模块 | `align_modules_from_template()` | `align-modules` | 写新 DOCX |
| 一键模板套版 | `format_from_template()` | `format-template` | 写新 DOCX |
| 提取完整模板事实 | `profile_template()` | `profile-template` | 否 |
| 导出模板段落样本 | 无单独公开 API | `extract-samples` | 否 |
| 获取规则 Schema | `rules_schema()` | 无 | 否 |
| 校验 `rules.json` | `validate_rules()` | `validate-rules` | 否 |
| 检查规则偏差 | `inspect_document()` | `inspect` / `dry-run` | 否 |
| 按显式规则修复 | `fix_document()` | `fix` | 写新 DOCX |
| 比较传统结构角色 | `plan_structure()` | `plan-structure` | 否 |
| 补建显式章节 | `complete_structure()` | `complete-structure` | 写新 DOCX |
| 只处理文献导航 | `link_references()` | `link-references` | 写新 DOCX |
| 检查节、目录边界和字段 | `inspect_structure()` | `inspect-structure` | 否 |
| 执行原子结构操作 | `apply_structure_operations()` | `apply-operations` | 写新 DOCX |
| 验证结构布局断言 | `validate_layout()` | `validate-layout` | 否 |
| 用 Word 刷新字段和分页 | `refresh_fields()` | `refresh-fields` | 写新 DOCX |
| 视觉比较 | `compare_visual()` | `visual-compare` | 生成 PDF/图片/报告 |

## 5. 模块名称

`--module` 和 `modules=` 使用下列模块名称：

| 模块 | 包含的角色 |
|---|---|
| `cover` | 中英文标题、封面标签、封面值 |
| `abstract_zh` | 中文摘要标题、正文、关键词 |
| `abstract_en` | 英文摘要标题、正文、关键词 |
| `toc` | 目录标题、一级至三级目录项、通用目录段落 |
| `headings` | 一级至三级正文标题 |
| `body` | 普通正文 |
| `captions` | 图题、表题 |
| `references` | 参考文献标题、中英文条目 |
| `acknowledgements` | 致谢 |
| `appendix` | 附录 |
| `tables` | 普通表格单元格文字 |
| `headers_footers` | 页眉、页脚 |

模块识别不是内容理解模型。它主要使用：

- 标题和关键词文本。
- 编号结构。
- 目录点线、制表符、TOC 样式和域。
- 摘要标题到关键词之间的上下文。
- 参考文献标题到致谢或附录之间的上下文。
- 表格、页眉、页脚所在 story。

## 6. 推荐 Agent 工作流

### 6.1 标准细粒度流程

这是默认推荐流程。

```text
convert
  -> detect-modules / compare-modules
  -> Agent 审核模块边界和缺失模块
  -> derive-rules --module ...
  -> Agent 审核或调整规则
  -> align-modules --module ...
  -> 分模块重复
  -> inspect / visual-compare / 人工复核
```

CLI 示例：

```powershell
# 1. 旧模板转换
python -m word_format_tool.cli convert `
  "<INPUT_DIR>/template.doc" `
  -o "<WORK_DIR>/template.docx"

# 2. 比较模块，不修改文档
python -m word_format_tool.cli compare-modules `
  "<WORK_DIR>/template.docx" `
  "<INPUT_DIR>/draft.docx" `
  -o "<WORK_DIR>/module-diff.json"

# 3. 只提取目录和参考文献规则
python -m word_format_tool.cli derive-rules `
  "<WORK_DIR>/template.docx" `
  -o "<WORK_DIR>/toc-reference-rules.json" `
  --module toc `
  --module references

# 4. 只对齐目录和参考文献
python -m word_format_tool.cli align-modules `
  "<WORK_DIR>/template.docx" `
  "<INPUT_DIR>/draft.docx" `
  -o "<OUTPUT_DIR>/stage-1.docx" `
  --module toc `
  --module references `
  --report "<WORK_DIR>/stage-1-report.json"

# 5. 下一阶段只处理标题和正文
python -m word_format_tool.cli align-modules `
  "<WORK_DIR>/template.docx" `
  "<OUTPUT_DIR>/stage-1.docx" `
  -o "<OUTPUT_DIR>/stage-2.docx" `
  --module headings `
  --module body
```

分阶段输出便于 Agent：

- 每轮只审查有限模块。
- 发现误识别时回退到上一阶段。
- 对封面、目录和正文使用不同处理策略。
- 避免页面设置或正文规则顺带影响其他模块。

### 6.2 Python 编排示例

```python
from pathlib import Path

from word_format_tool import (
    align_modules_from_template,
    build_template_rules,
    compare_modules,
    convert_to_docx,
    detect_modules,
)

template_source = Path("<INPUT_DIR>/template.doc")
template = Path("<WORK_DIR>/template.docx")
draft = Path("<INPUT_DIR>/draft.docx")
stage_1 = Path("<OUTPUT_DIR>/stage-1.docx")
stage_2 = Path("<OUTPUT_DIR>/stage-2.docx")

convert_to_docx(template_source, template)

module_diff = compare_modules(template, draft)
draft_modules = detect_modules(draft)

# Agent 在这里审核 module_diff 和 draft_modules。
toc_reference_rules = build_template_rules(
    template,
    modules=["toc", "references"],
    include_page=False,
)

# Agent 在这里审核或调整提取结果。
stage_1_report = align_modules_from_template(
    template,
    draft,
    stage_1,
    modules=["toc", "references"],
)

stage_2_report = align_modules_from_template(
    template,
    stage_1,
    stage_2,
    modules=["headings", "body"],
)
```

`build_template_rules()` 返回规则数据，但
`align_modules_from_template()` 会再次从模板提取所选模块规则。若 Agent 需要修改自动
规则，应把审核后的数据写为 `rules.json`，再调用 `validate_rules()`、
`inspect_document()` 和 `fix_document()`。

### 6.3 显式规则流程

以下情况使用传统 `rules.json`：

- 学校书面规范优先于模板实际样本。
- 模板包含大量括号说明文字或错误示例。
- 同一角色需要人工指定字体、字号或缩进。
- 需要自定义检测正则。
- 需要显式补建章节。
- 需要控制页面设置。

```python
from word_format_tool import fix_document, inspect_document, validate_rules

validate_rules("<WORK_DIR>/rules.json")
preflight = inspect_document("<INPUT_DIR>/draft.docx", "<WORK_DIR>/rules.json")
result = fix_document(
    "<INPUT_DIR>/draft.docx",
    "<WORK_DIR>/rules.json",
    "<OUTPUT_DIR>/formatted.docx",
)
postflight = inspect_document(
    "<OUTPUT_DIR>/formatted.docx",
    "<WORK_DIR>/rules.json",
)
```

### 6.4 一键流程

```powershell
python -m word_format_tool.cli format-template `
  "<INPUT_DIR>/template.doc" `
  "<INPUT_DIR>/draft.docx" `
  -o "<OUTPUT_DIR>/formatted.docx" `
  --report "<WORK_DIR>/format-report.json"
```

`format-template` 会自动转换输入、提取全部角色、包含模板第一节页面设置并执行修复。

只建议用于：

- 模板与初稿结构高度一致。
- 模板样本文字干净，没有大量格式说明。
- 页面分节简单。
- 允许完成后进行人工复核。

复杂论文任务不应把它作为默认路径。

## 7. 细粒度 API 契约

### 7.1 `convert_to_docx`

```python
result = convert_to_docx("template.doc", "template.docx")
```

返回：

```json
{
  "source": "template.doc",
  "output": "template.docx",
  "backend": "microsoft-word"
}
```

`backend` 可能为 `microsoft-word`、`libreoffice` 或 `copy`。

### 7.2 `detect_modules`

```python
result = detect_modules("draft.docx")
```

重点字段：

- `modules[]`：连续模块区段。
- `modules[].name`：模块名称。
- `modules[].start` / `end`：稳定位置。
- `modules[].roles`：区段内识别出的角色。
- `modules[].text_preview`：最多三条预览。
- `module_counts`：每类模块区段数量。

Agent 应检查：

- 摘要是否同时存在标题、正文和关键词角色。
- 目录是否被拆成异常多个区段。
- 参考文献标题是否被识别为 `reference_heading`。
- 封面表格是否被归入 `cover`。

### 7.3 `compare_modules`

```python
result = compare_modules("template.docx", "draft.docx")
```

重点字段：

- `missing_in_draft`：模板存在、初稿未识别到的模块类型。
- `extra_in_draft`：初稿额外模块类型。
- `shared_modules`：双方共有模块。
- `template_module_counts` / `draft_module_counts`：区段数量。

`missing_in_draft` 只表示“未识别到”，不等于必须自动创建。

### 7.4 `build_template_rules`

```python
rules = build_template_rules(
    "template.docx",
    modules=["abstract_zh", "abstract_en", "toc", "references"],
    include_page=False,
)
```

规则提取方式：

- 按语义角色收集模板实际段落格式。
- 同一角色存在多个样本时使用最常见的格式签名。
- 只返回选定模块包含的角色。
- `include_page=False` 时不生成页面规则。

注意：

- 模板中的说明段、示例段也可能成为样本。
- 自动规则是候选值，不是学校规范的权威解释。
- `include_page=True` 只读取模板第一节页面设置。
- 多分节模板应使用 `profile_template()` 审核全部 section，再手工构造页面规则。

### 7.5 `align_modules_from_template`

```python
report = align_modules_from_template(
    "template.docx",
    "draft.docx",
    "aligned.docx",
    modules=["toc", "references"],
    include_page=False,
)
```

行为：

- 只生成所选模块角色的规则。
- 默认不应用页面设置。
- 选择 `toc` 时才尝试目录导航修复。
- 选择 `references` 时才处理参考文献导航。
- 不移动段落、图片、表格、公式或题注。
- 输出路径必须与输入不同。

返回为 `FormatReport` 扩展数据，额外包含：

- `template`
- `selected_modules`
- `include_page`
- `template_rule_roles`

## 8. 传统规则和角色回退

细分角色在旧规则缺失时使用以下兼容回退：

- `title_zh` / `title_en` -> `title`
- `abstract_heading_zh` / `abstract_heading_en` -> `abstract` -> `heading_1`
- `abstract_body_zh` / `abstract_body_en` -> `abstract` -> `body`
- `keywords_zh` / `keywords_en` -> `keywords` -> `body`
- `toc_heading` / `toc_entry_1/2/3` -> `table_of_contents`
- `reference_entry_zh/en` -> `reference_entry` -> `reference`
- `reference_heading` -> `heading_1`

以下角色不会回退到 `body`：

- `cover_label`
- `cover_value`
- `table_text`
- `header`
- `footer`

这可以避免普通正文规则误改封面、表格和页眉页脚。

## 9. 能力边界

### 9.1 摘要

工具可识别：

- `摘要`、`摘 要`、带括号说明的摘要标题。
- `ABSTRACT` 及带括号说明的英文摘要标题。
- 摘要标题之后、关键词之前的摘要正文。
- `关键词`、`关键字`、`Keywords`、`Key words`。

工具不会：

- 在没有摘要标题和关键词线索时猜测任意正文是摘要。
- 生成真实摘要内容。
- 自动把模板摘要内容复制到初稿。

### 9.2 目录

工具可识别：

- 目录标题。
- TOC 样式和目录域。
- 带制表符、点线和页码的一级至三级目录项。
- 能唯一匹配正文标题的失效 `PAGEREF` 目标。

工具不会：

- 猜测歧义目录目标。
- 保证 Microsoft Word 打开后所有域结果立即刷新。
- 为缺失正文标题伪造目录目标。

`coverage.toc_unrepairable > 0` 时必须人工检查目录。

### 9.3 参考文献

工具可识别参考文献标题后的条目，并区分中英文条目。选择 `references` 模块时可处理
书签、正文上标链接和返回首次引用的链接。

工具不会猜测：

- 重复编号。
- 缺失编号。
- 一个引用对应多个歧义条目。
- 非标准正文引用语法。

模板参考文献区中的说明文字可能被视为条目，Agent 应审核自动规则。

### 9.4 封面、表格和浮动对象

工具可格式化封面表格中的文字角色，但不会复制模板表格结构，也不会编辑：

- 文本框。
- VML 图形。
- 浮动图片位置。
- 任意复杂封面对象。

存在封面表格或浮动对象时，页面设置默认受风险门控。

### 9.5 页面设置

`align-modules` 默认 `include_page=False`。只有确认模板和初稿分节兼容时才启用
`--include-page`。

自动模板规则只使用模板第一节页面数据。需要不同分节页边距、页眉距离或页脚距离时，
应使用 `profile_template()` 获取全部 section，由 Agent 构造显式规则或做专门收尾。

## 10. 报告判读

公开数据契约包括 `TemplateProfile`、`Rules`、`FormatReport`、
`StructurePlan`、`StructureCompletionResult` 和 `VisualComparisonReport`。
0.6 新增 `ParagraphSelector`、`StructureOperationPlan` 和
`LayoutExpectations`，用于稳定定位、原子结构操作和机器可读布局验收。

结构收尾推荐顺序：

```text
inspect_structure
  -> Agent 生成 StructureOperationPlan
  -> apply_structure_operations
  -> refresh_fields（仅 Word 可用时）
  -> validate_layout
  -> 必要时执行一次 visual-compare
```

`ParagraphSelector` 每次操作前都会重新解析，不应缓存 Word COM Range 或字符偏移。
目录刷新会改变字段结果长度，因此必须在刷新后重新调用 `inspect_structure()`。

`fix_document()` 和 `align_modules_from_template()` 返回 `FormatReport`：

- `summary.total_issues`：检查到的问题总数。
- `summary.fixable_issues`：工具能确定性处理的问题数。
- `summary.fixed_issues`：修复后复检确认消失的问题数。
- `summary.unfixed_issues`：仍存在的问题数。
- `issues[].fixable`：问题能否由工具确定性处理。
- `issues[].fixed`：修复后是否消失。
- `issues[].location`：正文、表格、页眉或页脚位置。
- `coverage.role_counts`：识别出的角色数量。
- `coverage.uncovered_areas`：未覆盖或需人工判断的区域。
- `coverage.layout_risk`：封面表格、浮动对象、文本框和分节风险。
- `coverage.toc_unrepairable`：无法唯一修复的目录目标数。

`fixed_issues > 0` 不代表整篇文档已经完全符合模板。

## 11. 错误处理

所有预期业务异常继承：

```python
from word_format_tool.exceptions import WordFormatToolError
```

主要异常：

| 异常 | 含义 |
|---|---|
| `InputFileError` | 输入不存在、扩展名错误、模块名称错误或未指定模块 |
| `RuleValidationError` | `rules.json`、Schema 或正则无效 |
| `UnsafeOutputPathError` | 输出路径会覆盖输入 |
| `DocumentReadError` | DOCX 读写、转换或完整性校验失败 |
| `ReportWriteError` | JSON/Markdown 报告写入失败 |
| `VisualValidationUnavailableError` | 视觉比较依赖或后端不可用 |

```python
from word_format_tool.exceptions import WordFormatToolError

try:
    result = align_modules_from_template(
        template,
        draft,
        output,
        modules=["toc"],
    )
except WordFormatToolError as exc:
    return {
        "status": "word_format_error",
        "message": str(exc),
    }
```

CLI 遇到业务异常时返回退出码 `1`。

## 12. 视觉与最终验收

```python
visual = compare_visual(
    "before.docx",
    "after.docx",
    "<WORK_DIR>/visual",
    dpi=120,
    changed_pixel_threshold=0.005,
)
```

重点检查：

- `page_count_changed`
- `page_size_changed`
- 各页 `changed_pixel_ratio`
- 封面、目录、参考文献和高变化页面

视觉比较只能指出变化，不能判断变化是否符合学校规范。

最终交付前至少确认：

- 输出文件可重新打开。
- 输入文件未被覆盖。
- 本轮只修改了预期模块。
- 模块缺失和额外模块已由 Agent 判断。
- `coverage.uncovered_areas` 已披露。
- 目录不可修复目标已披露。
- 封面、目录、参考文献和明显分页问题已复核。
- 未执行视觉比较时明确记录原因。

只有必要结构完整、未覆盖项为空且视觉或人工检查通过时，才应描述为“完整套版”；
否则应描述为“指定模块对齐”或“支持范围内格式修复”。
