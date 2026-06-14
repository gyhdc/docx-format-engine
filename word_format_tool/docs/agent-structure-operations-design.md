# Word 格式工具 Agent 结构操作增强设计

## 背景

真实论文格式修复任务表明，现有工具适合批量检查和修复字体、字号、段落及页面规则，
但 Agent 在处理目录、正文分节、页眉页码和分页边界时仍需临时编写 Word COM 脚本。
这类脚本通常同时承担定位、修改、刷新和验证，容易产生以下问题：

- 目录字段更新后字符位置变化，旧 Range 或段落索引失效；
- 多项结构操作互相影响，单步失败后仍继续修改；
- Agent 无法从现有报告确认正文是否真正开始于新节；
- 纯像素比较无法区分内容差异和格式错误；
- 没有多模态能力的模型难以判断封面、目录混页等明显版式问题；
- 同一操作重复执行可能增加分节符、页码字段或空段落。

工具仍定位为确定性的规则修复器，不尝试一键理解或克隆任意模板。本设计增加一组
细粒度、可组合、可验证的结构操作，供 Agent 根据具体文档选择。

## 目标

- 提供不依赖易漂移字符偏移量的段落定位器。
- 输出节、目录、标题、页眉页脚和分页控制的结构事实。
- 支持在指定段落前安全插入下一页分节符。
- 支持为目标节设置页眉、页码字段和页码重启值。
- 支持请求 Word 打开时更新字段，并提供可选 Microsoft Word 即时刷新。
- 每个操作执行前后都验证条件，失败时不发布半成品。
- 提供机器可读布局验证，使文本模型也能发现目录混页、正文未分节等问题。
- 修复模板规则抽取中的异常行距数值，避免明显页面溢出。
- 保持现有 `fix`、模块对齐和目录书签修复行为兼容。

## 非目标

- 不自动复制复杂封面、文本框、浮动对象或模板正文。
- 不根据视觉相似度自行决定内容块应移动到哪里。
- 不在没有唯一定位条件时猜测目标段落。
- 不用纯 OOXML 伪造 Word 的最终页码或分页结果。
- 不把 Microsoft Word 或 pywin32 变为核心强制依赖。

## 架构

新增三个核心模块：

```text
document_analyzer
        |
        v
paragraph_locator ----> structure_inspector
        |                       |
        v                       v
structure_operations ----> layout_validator
        |
        v
optional_word_backend
```

### `paragraph_locator.py`

定义稳定的正文段落选择器：

- `text`：规范化后的完整文本；
- `role`：可选语义角色；
- `occurrence`：重复文本中的第几个匹配，默认要求唯一；
- `after_role`：可选边界，例如只在目录之后查找；
- `style_name`：可选样式约束。

每次操作前重新分析当前文档并解析选择器，不保存跨操作的字符偏移。匹配为零或多项且
未显式提供 `occurrence` 时返回明确错误和候选位置。

结构报告为每个正文段落生成 `locator` 建议和上下文指纹。指纹只用于诊断和验证，
不作为唯一写入地址。

### `structure_inspector.py`

输出：

- 文档节数量及每节页面设置；
- 每节的首末正文段落；
- 页眉、页脚文本及是否链接前节；
- PAGE、TOC 等字段代码；
- 页码重启值；
- 正文段落的角色、样式、分页属性和分节边界；
- 目录标题、目录项范围及候选正文首标题；
- `updateFields` 状态。

纯 OOXML 报告不声称提供最终页码。可选 Word 后端报告可补充真实页码和总页数。

### `structure_operations.py`

公开批量操作入口 `apply_structure_operations`。操作计划为严格 JSON，按顺序执行：

```json
{
  "version": "0.1",
  "operations": [
    {
      "type": "insert_section_before",
      "target": {
        "text": "1 任务描述",
        "role": "heading_1",
        "after_role": "table_of_contents"
      },
      "break_type": "next_page"
    },
    {
      "type": "set_header",
      "section_start": {
        "text": "1 任务描述",
        "role": "heading_1"
      },
      "text": "山东农业大学学士学位论文",
      "alignment": "center",
      "bottom_border": true
    },
    {
      "type": "set_page_number",
      "section_start": {
        "text": "1 任务描述",
        "role": "heading_1"
      },
      "start": 1,
      "alignment": "center"
    },
    {
      "type": "request_field_update"
    }
  ]
}
```

首期支持：

- `insert_section_before`
- `set_header`
- `set_footer_text`
- `set_page_number`
- `set_paragraph_pagination`
- `request_field_update`

目录内容重建继续由 Word 字段系统负责。工具可设置或修复 TOC 字段、请求刷新，但不在
纯 OOXML 路径中伪造目录页码。

每个操作返回：

- 解析到的目标；
- 修改前事实；
- 修改后事实；
- `changed`；
- `postcondition_passed`；
- 警告或失败原因。

所有操作必须幂等。操作失败时不保存输出，成功后沿用现有原子写入和内容保护机制。

### `layout_validator.py`

验证规则与操作计划分离。首期支持以下断言：

- `section_count_at_least`
- `body_starts_new_section`
- `section_header_equals`
- `section_has_page_field`
- `section_page_number_starts_at`
- `toc_before_body`
- `body_first_heading_equals`
- `update_fields_enabled`

报告状态为 `passed`、`warning` 或 `failed`，每个断言包含错误代码、实际值、期望值、
位置和建议操作。结构事实足够时直接判定；必须依赖 Word 最终分页时标记
`requires_word_backend`，而不是猜测。

### `optional_word_backend.py`

仅在 Windows 且 Microsoft Word 可用时提供：

- 更新全部字段；
- 重新分页；
- 保存到不同输出路径；
- 返回总页数及指定定位器所在页码。

模块延迟导入 pywin32。不可用时返回明确能力错误，不影响其他命令。

## CLI 与 API

新增 CLI：

```text
wordfmt inspect-structure document.docx -o structure.json
wordfmt apply-operations document.docx --plan operations.json -o output.docx --report result.json
wordfmt validate-layout document.docx --expect expectations.json -o validation.json
wordfmt refresh-fields document.docx -o refreshed.docx --report refresh.json
```

`inspect` 的 Markdown 输出调整为可选，JSON 仍为必需输出。现有调用保持兼容。

新增 Python API：

- `inspect_structure`
- `apply_structure_operations`
- `validate_layout`
- `refresh_fields`

## 异常行距防护

模板抽取只接受能够转换为合理 Word 行距的值：

- 倍数行距限制在合理区间；
- 固定值和最小值统一转换为磅；
- 检测到疑似 EMU、twip 或内部整数值时先按单位转换；
- 无法确定单位或转换后超出范围时不生成该字段，并写入提取警告。

本防护优先保证不产生明显页面溢出，不猜测高风险异常值。

## 操作顺序建议

Agent 应按以下顺序处理：

1. 普通格式规则检查和修复；
2. 检查结构并生成操作计划；
3. 调整封面或其他人工映射区域；
4. 插入分节符；
5. 设置页眉、页脚和页码；
6. 修复或创建目录字段；
7. 最后请求字段刷新；
8. 执行布局验证；
9. 仅在结构验证通过后进行一次视觉抽查。

工具报告必须提醒 Agent：目录刷新前不要缓存字符位置，视觉比较不能替代结构验证。

## 测试计划

### 段落定位

- 唯一文本和角色可以稳定定位。
- 目录项与同名正文标题并存时，可通过 `after_role` 定位正文。
- 未指定 occurrence 的重复匹配必须失败。
- 格式修改后定位器仍可重新解析。

### 结构操作

- 在正文标题前插入下一页分节符，保存重载后节数量增加。
- 重复执行同一分节操作不新增节。
- 设置页眉会解除与前节链接且不修改前置封面节。
- PAGE 字段和页码起始值正确写入，重复执行不重复字段。
- 任一操作后置条件失败时不发布输出。

### 布局验证

- 发现正文未从新节开始。
- 发现正文节缺少页眉或 PAGE 字段。
- 发现正文首标题与期望不一致。
- 目录字段位于正文标题之前时通过。
- 需要真实页码的断言在无 Word 后端时明确标记。

### Word 后端

- 无 Word 或 pywin32 时返回能力错误。
- 在 Windows 集成环境中验证字段刷新、重新分页和页码报告。

### 回归

- 现有全部单元测试通过。
- 使用真实模板和初稿验证：目录独立、正文首标题正确、正文新节、页眉和页码存在。
- 重复运行操作计划后结构不继续变化。

## 分阶段实现

1. 段落定位、结构检查和纯 OOXML 布局验证。
2. 分节、页眉、页码和分页属性原子操作。
3. CLI、公开 API、报告和 Agent 文档。
4. 可选 Word 字段刷新后端。
5. 异常行距防护和真实文档回归。
