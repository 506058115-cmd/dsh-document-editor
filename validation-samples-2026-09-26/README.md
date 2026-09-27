# 文档工具验证样例

以下文件由 MCP stdio 工具实际创建、读取、修改并通过 `verify_document` 校验。编辑版与原件分开保存。

| 原件 | 编辑版 | 验证内容 |
| --- | --- | --- |
| `project-summary.pdf` | `project-summary-edited.pdf` | 2 页 A4/Letter 原件 → 3 页编辑版；追加页与中文文本读取通过 |
| `chinese-multipage.pdf` | `chinese-multipage-edited.pdf` | 3 页中文原件 → 5 页编辑版；自动分页、读取与追加长内容通过 |
| `weekly-notes.docx` | `weekly-notes-edited.docx` | 创建段落/表格，按索引改段落和表格单元格 |
| `budget-tracker.xlsx` | `budget-tracker-edited.xlsx` | 改现有单元格、写公式、追加行 |
| `project-update.pptx` | `project-update-edited.pptx` | 创建两页并分别修改既有幻灯片文本 |

验证结果：10 个文件结构校验通过。
说明：openpyxl 会保存公式表达式，但不会计算公式结果。
