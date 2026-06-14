# DOCX Format Engine

本目录包含 SDAU-agent 的独立 Word 文档格式引擎。

- 核心包：`word_format_tool/`
- 接入说明：`word_format_tool/AGENT_INTEGRATION.md`
- 开发说明：`word_format_tool/DEVELOPMENT.md`
- 示例：`word_format_tool/examples/`
- 测试：`word_format_tool/tests/`

工具采用“模板事实提取 + Agent 审核规则 + 确定性局部修复 + 结构/视觉验收”
流程。`work_temp/` 只用于本机临时产物，不属于源码和发布内容。
