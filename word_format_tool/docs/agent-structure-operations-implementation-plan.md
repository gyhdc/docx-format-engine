# Agent 结构操作增强实现计划

## 目标

在保留现有规则修复能力和安全边界的前提下，为 Agent 提供稳定定位、结构修改、
布局验证和可选 Word 字段刷新能力，减少临时 COM 脚本和反复视觉试错。

## 阶段一：定位与结构检查

新增：

- `paragraph_locator.py`
- `structure_inspector.py`
- 对应 Pydantic 输入模型和公开 API

测试先覆盖：

1. 目录项和正文标题同名时，`after_role=table_of_contents` 只定位正文。
2. 重复匹配且未声明 `occurrence` 时返回可诊断错误。
3. 结构报告列出节、页眉页脚、字段、页码起始值和段落分节边界。
4. 格式变化不影响基于文本、角色和上下文的重新定位。

## 阶段二：原子结构操作

新增：

- `structure_operations.py`
- 操作计划模型和 JSON 加载

测试先覆盖：

1. 在正文首标题前插入下一页分节符。
2. 同一计划重复执行不增加额外节。
3. 只修改正文节页眉，不污染封面节。
4. 页脚只保留一个 PAGE 字段，页码从指定值开始。
5. 操作失败时不发布输出。

## 阶段三：布局验证

新增：

- `layout_validator.py`
- 期望模型和报告模型

测试先覆盖：

1. 正文是否从新节开始。
2. 正文首标题是否正确。
3. 目录是否位于正文之前。
4. 正文节页眉、PAGE 字段和页码起始值是否正确。
5. 需要 Word 分页事实的检查明确标记后端要求。

## 阶段四：CLI 和公开 API

新增命令：

- `inspect-structure`
- `apply-operations`
- `validate-layout`
- `refresh-fields`

调整：

- `inspect --md` 改为可选；
- `README.md`、`AGENT_INTEGRATION.md`、`DEVELOPMENT.md` 增加推荐工作流；
- 帮助文本强调“先结构检查、后操作、最后刷新和视觉抽查”。

## 阶段五：Word 后端与规则防护

新增：

- `optional_word_backend.py`
- 字段刷新和重新分页报告

修复：

- 倍数行距超出合理范围时不生成模板规则；
- 行距提取异常写入模板规则警告；
- 真实模板回归不得再次产生超大行距。

## 验证

1. 新增测试逐项完成红、绿、重构循环。
2. 运行工具包全部 pytest。
3. 使用真实模板与初稿执行：
   - `inspect-structure`
   - `apply-operations`
   - `refresh-fields`
   - `validate-layout`
4. 重复执行同一操作计划并确认结构不再变化。
5. 最后只渲染一次关键页面，确认结构报告与视觉结果一致。
