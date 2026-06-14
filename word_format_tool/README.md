# word-format-tool

`word-format-tool` 是一个独立、本地、确定性的 Python 工具包，用于从
DOCX 模板提取事实、检查论文格式、执行支持范围内的安全修复、补建显式声明的
缺失章节、修复可确定匹配的目录/文献跳转，并输出覆盖范围与视觉比较报告。

工具不调用大模型、不联网，也不依赖 SDAU-agent 的其他业务模块。外部 Agent
负责理解学校规范、审核模块边界和模板规则；本工具负责识别、提取、局部执行和验证。

## 能力范围

- 抽取页面设置、样式、正文/表格/页眉/页脚样本、结构清单和布局风险。
- 区分 `title_zh` / `title_en`，并保护无法可靠识别的封面正文段落。
- 区分 `reference_entry_zh` / `reference_entry_en`，支持不同对齐和字体规则。
- 识别摘要、关键词、三级标题、正文、图表题注、致谢、附录和参考文献。
- 按显式规则处理封面表格单元格、普通表格、页眉和页脚。
- 检查字体、字号、粗斜体、对齐、行距、缩进、段间距、keep-with-next 和页面设置。
- 检测封面表格、浮动对象、文本框和 VML 图形；高风险时默认禁止修改页边距。
- 修复唯一匹配的目录 `PAGEREF` 书签；部分无法匹配时不触发全局目录刷新。
- 检查并建立参考文献书签、正文上标跳转和返回首次引用的链接。
- 显式补建 `rules.structure.required_sections` 声明且当前缺失的章节。
- 校验正文、表格、媒体、公式、块顺序、导航对象和浮动对象位置指纹。
- 可选通过 Microsoft Word 导出 PDF，比较页数、页面尺寸和逐页像素变化。
- 输出 JSON/Markdown 报告，其中包含机器可读的未覆盖区域清单。
- 内建 `.doc` / `.docx` 到 `.docx` 转换，不再要求 Agent 临时编写转换脚本。
- 识别中英文摘要、关键词、目录层级和参考文献等模块边界，即使原样式错误也可按文本与上下文判定。
- 支持比较模板与初稿模块、按模块提取规则、只对齐显式选择的模块。
- 检查分节、目录边界、页眉页脚、页码字段和段落分页属性，输出稳定定位事实。
- 按严格 JSON 计划执行幂等结构操作；每一步重新定位目标，歧义或失败时不发布输出。
- 在不依赖图像模型的情况下校验正文起始节、目录顺序、页眉、页码和字段刷新标记。
- 可选调用 Microsoft Word 刷新目录、字段和真实分页，避免 Agent 临时编写 COM 脚本。
- Word 刷新前后校验受保护结构指纹；目录更新若吞掉分节符，拒绝发布坏输出。

## 安全边界

普通 `fix` 是“支持范围内的格式修复”，不是任意模板的完整克隆。

- 不自动复制模板正文、任意表格、文本框、浮动对象或章节内容。
- 不猜测无法唯一匹配的目录目标、文献编号或引用关系。
- 不移动图题、表题、公式、图片锚点、表格或段落。
- 文本框和浮动对象布局只检测、指纹保护和视觉比较，不直接编辑。
- 页面几何命中封面/浮动对象风险时默认 `fixable=false`。
- 结构补建必须使用独立命令，并在规则中逐项声明。
- 所有输出必须使用不同路径，原文件不会被覆盖。

只有当报告中的 `coverage.uncovered_areas` 为空、必要结构完整且视觉检查通过时，
上层 Agent 才应把结果描述为“完整套版”；否则应使用“支持范围内格式修复”。

## 安装

要求 Python 3.10+。

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[test]"
```

视觉比较是可选能力：

```powershell
.\.venv\Scripts\python.exe -m pip install -e ".[visual]"
```

DOCX 到 PDF 的视觉比较还要求 Windows 上安装 Microsoft Word。

## CLI

复杂任务优先使用细粒度模块流程：

```text
wordfmt convert legacy-template.doc -o template.docx
wordfmt compare-modules template.docx draft.docx -o module-diff.json
wordfmt derive-rules template.docx -o selected-rules.json --module toc --module references
wordfmt align-modules template.docx draft.docx -o aligned.docx --module toc --module references
```

完整命令：

```text
wordfmt profile-template template.docx -o template_profile.json
wordfmt extract-samples template.docx -o template_samples.json
wordfmt plan-structure template.docx draft.docx -o structure_plan.json
wordfmt convert legacy-template.doc -o template.docx
wordfmt detect-modules draft.docx -o draft-modules.json
wordfmt compare-modules template.docx draft.docx -o module-diff.json
wordfmt derive-rules template.docx -o abstract-rules.json --module abstract_zh --module abstract_en
wordfmt align-modules template.docx draft.docx -o aligned.docx --module toc --module references

wordfmt validate-rules rules.json
wordfmt inspect draft.docx --rules rules.json -o report.json --md report.md
wordfmt dry-run draft.docx --rules rules.json --md preview.md --report preview.json
wordfmt fix draft.docx --rules rules.json -o fixed.docx --report fix.json --md fix.md

wordfmt complete-structure draft.docx --rules rules.json -o completed.docx --report structure.json
wordfmt link-references draft.docx --rules rules.json -o linked.docx --report refs.json --md refs.md

wordfmt inspect-structure draft.docx -o structure-facts.json
wordfmt apply-operations draft.docx --plan operations.json -o structured.docx --report operations-report.json
wordfmt validate-layout structured.docx --expect layout-expectations.json -o layout-report.json
wordfmt refresh-fields structured.docx -o refreshed.docx --report refresh-report.json

wordfmt visual-compare draft.docx fixed.docx --artifacts visual -o visual.json
```

命令职责：

- `profile-template`：输出模板事实、结构清单和布局风险，不自动生成规则。
- `convert`：使用 Microsoft Word，失败时回退 LibreOffice，将旧 `.doc` 转为 `.docx`。
- `detect-modules`：只报告模块边界、角色和段落范围，不修改文档。
- `compare-modules`：比较模板与初稿已有、缺失和额外模块，供 Agent 决策。
- `derive-rules`：只从模板提取指定模块的规则；默认不包含页面设置。
- `align-modules`：只修改显式选择的模块；默认不改页面设置。
- `plan-structure`：列出模板有而初稿缺失的语义章节，不修改文档。
- `inspect` / `dry-run`：检查并输出预计修改项和未覆盖区域。
- `fix`：格式修复、文献导航及安全目录书签修复。
- `complete-structure`：只补建规则显式声明的缺失章节。
- `link-references`：只处理文献导航，不改普通格式。
- `inspect-structure`：报告分节、目录边界、字段和段落定位事实，不修改文档。
- `apply-operations`：按严格操作计划处理分节、页眉页脚、页码和分页属性。
- `validate-layout`：以机器可读断言验收结构布局，适合无视觉能力的 Agent。
- `refresh-fields`：通过 Microsoft Word 刷新字段、目录和真实分页。
- `visual-compare`：用 Word/PDF 渲染结果发现分页或视觉布局变化。
- `format-template`：自动转换、提取全部角色和执行修复的便捷层，不建议作为复杂任务默认入口。

## Python API

Agent 或程序员接入时，请优先阅读
[`AGENT_INTEGRATION.md`](AGENT_INTEGRATION.md)。其中包含模块名称、细粒度工作流、
逐项 API 契约、能力边界、错误处理、报告判读和最终验收要求。

```python
from word_format_tool import (
    align_modules_from_template,
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
    apply_structure_operations,
    validate_layout,
    validate_rules,
)

convert_to_docx("template.doc", "template.docx")
module_diff = compare_modules("template.docx", "draft.docx")
draft_modules = detect_modules("draft.docx")
selected_rules = build_template_rules(
    "template.docx",
    modules=["toc", "references"],
    include_page=False,
)
module_report = align_modules_from_template(
    "template.docx",
    "draft.docx",
    "aligned.docx",
    modules=["toc", "references"],
)

profile = profile_template("template.docx")
structure_plan = plan_structure("template.docx", "draft.docx")
validate_rules("rules.json")

inspection = inspect_document("draft.docx", "rules.json")
fix_report = fix_document("draft.docx", "rules.json", "fixed.docx")
structure_result = complete_structure(
    "fixed.docx", "rules.json", "completed.docx"
)
reference_report = link_references(
    "draft.docx", "rules.json", "linked.docx"
)
structure_facts = inspect_structure("linked.docx")
operation_report = apply_structure_operations(
    "linked.docx", "operations.json", "structured.docx"
)
layout_report = validate_layout(
    "structured.docx", "layout-expectations.json"
)
refresh_report = refresh_fields("structured.docx", "refreshed.docx")
visual_report = compare_visual(
    "draft.docx", "fixed.docx", "visual-artifacts"
)
```

公开函数接受 `str` 或 `pathlib.Path`，返回可直接 JSON 序列化的 `dict`。

`refresh_fields()` 在临时副本上运行 Word，并比较刷新前后的分节拓扑、正文首标题、
页眉和页码字段。返回中的 `structure_preserved=true` 表示结构安全门通过。刷新完成后
Word 通常会清除 `updateFields` 请求标记，因此刷新后的布局期望应设置
`update_fields_enabled=false`。若目录更新删除分节，工具抛出
`FieldRefreshIntegrityError` 且不发布输出。

程序员和 Agent 应优先阅读
[`AGENT_INTEGRATION.md`](AGENT_INTEGRATION.md)。其中记录模块名称、API 契约、返回字段、
错误处理、推荐工作流和能力边界。

## rules.json

完整示例见 `examples/rules.example.json`。关键安全设置：

```json
{
  "version": "0.1",
  "priority": {
    "protect_front_matter": true,
    "allow_unsafe_page_geometry": false
  },
  "roles": {
    "title_zh": {
      "font_east_asia": "SimHei",
      "font_size_pt": 22,
      "alignment": "center"
    },
    "title_en": {
      "font_ascii": "Times New Roman",
      "font_size_pt": 22,
      "bold": true,
      "alignment": "center"
    },
    "reference_entry_zh": {
      "alignment": "both"
    },
    "reference_entry_en": {
      "alignment": "left"
    },
    "body": {
      "font_east_asia": "SimSun",
      "font_ascii": "Times New Roman",
      "font_size_pt": 12,
      "line_spacing": 1.5,
      "first_line_indent_chars": 2,
      "alignment": "both"
    }
  }
}
```

兼容回退：

- `title_zh` / `title_en` 未配置时回退到旧 `title`。
- 摘要标题、摘要正文和中英文关键词优先使用细分角色，缺失时回退到旧
  `abstract` / `keywords` / `body`。
- `toc_heading` / `toc_entry_1/2/3` 回退到 `table_of_contents`。
- `reference_entry_zh` / `reference_entry_en` 回退到
  `reference_entry`，再回退到旧 `reference`。
- `reference_heading` 未配置时回退到 `heading_1`。
- `cover_label`、`cover_value`、`table_text`、`header`、`footer` 不回退到
  `body`，避免误改封面和表格。

显式结构补建示例：

```json
{
  "structure": {
    "required_sections": [
      {
        "role": "acknowledgements",
        "heading_text": "致谢",
        "placeholder_text": "请在此补充致谢内容。",
        "insert_before_role": "reference_heading"
      }
    ]
  }
}
```

## 目录处理

目录段落可以通过 `table_of_contents` 角色配置格式。失效跳转只在目录可见标题
能唯一匹配正文 `heading_1/2/3` 时补建书签。

- 全部失效目标都可修复：设置 Word 打开时更新域。
- 只有部分目标可修复：改写确定目标，但不设置全局更新，防止 Word 删除无法识别
  的旧目录项。
- 无法唯一匹配：保持原样并记录到 `coverage.uncovered_areas`。

最终交付前仍建议在 Word 中人工检查并更新目录。

## 报告字段

- `fixable=true`：工具能确定性修改。
- `fixed=true`：修复后复检时原问题已消失。
- `location`：正文、表格单元格、页眉或页脚的稳定位置描述。
- `coverage.story_counts`：各文档区域的段落数量。
- `coverage.role_counts`：角色识别数量。
- `coverage.layout_risk`：封面表格、浮动对象、文本框和分节风险。
- `coverage.uncovered_areas`：未处理或需人工判断的区域。
- `coverage.visual_comparison`：视觉比较是否执行；普通报告默认为 `not_run`。

## Agent 推荐流程

1. 旧 `.doc` 先调用 `convert`，保留原模板副本。
2. 调用 `compare-modules`，先确认初稿缺失和已有模块。
3. 对需要处理的模块调用 `derive-rules --module ...`，不要默认提取全文规则。
4. Agent 审核模块边界和规则后，调用 `align-modules --module ...` 局部对齐。
5. 页面设置、封面、正文、图表、目录和文献应按任务需要分批处理。
6. 用 `inspect-structure` 获取正文起点和目录边界，再提交小型 `operations.json`。
7. 用 `validate-layout` 验收分节、页眉、页码和目录顺序；需要真实页码时再调用 `refresh-fields`。
8. `refresh-fields` 后再次执行 `validate-layout`；此时 `update_fields_enabled` 通常应为 `false`。
9. 只有明确需要全篇规则时才使用传统 `inspect` / `fix`。
10. 最终按风险选择 `visual-compare` 或人工检查封面、目录和正文首页。

`format-template` 是方便快速处理的组合命令，但不是推荐的 Agent 默认路径。模板和
初稿差异较大时，优先使用上述细粒度命令。

## 测试

```powershell
python -m pytest
```

测试动态创建 DOCX，覆盖正文、封面表格、页眉页脚、目录、文献跳转、结构补建、
稳定段落定位、幂等结构操作、布局断言、字段刷新适配器、页面风险、布局对象指纹和
视觉比较核心。
