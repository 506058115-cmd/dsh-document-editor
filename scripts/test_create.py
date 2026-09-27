#!/usr/bin/env python3
"""Quick self-test: try creating all 4 document types."""
import sys, os, json

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# Ensure src path can be found
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

print("=== Creating Word doc ===")
try:
    from server import _tool_create_word
    result = _tool_create_word(
        title="测试文档",
        paragraphs=[
            {"text": "第一段内容", "style": "heading1"},
            {"text": "这是正文段落。", "style": "Normal"},
            {"text": "第二标题", "style": "heading2"},
            {"text": "更多内容。", "style": "Normal"},
        ],
        output_path=os.path.join(os.path.dirname(__file__), "..", "output.docx")
    )
    print(json.dumps(result, ensure_ascii=False))
except Exception as e:
    print(f"Error: {e}")

print("\n=== Creating Excel workbook ===")
try:
    from server import _tool_create_excel
    result = _tool_create_excel(
        sheets=[
            {"name": "成绩表", "headers": ["姓名", "数学", "语文", "英语"],
             "rows": [["张三", 95, 88, 92], ["李四", 87, 91, 85]]},
        ],
        output_path=os.path.join(os.path.dirname(__file__), "..", "output.xlsx")
    )
    print(json.dumps(result, ensure_ascii=False))
except Exception as e:
    print(f"Error: {e}")

print("\n=== Creating PowerPoint ===")
try:
    from server import _tool_create_pptx
    result = _tool_create_pptx(
        title="演示文稿",
        slides=[
            {"layout": "title", "title": "项目汇报"},
            {"layout": "blank", "content": "已完成:\n- 用户模块\n- API接口"},
        ],
        output_path=os.path.join(os.path.dirname(__file__), "..", "output.pptx")
    )
    print(json.dumps(result, ensure_ascii=False))
except Exception as e:
    print(f"Error: {e}")

print("\nAll tests done.")
