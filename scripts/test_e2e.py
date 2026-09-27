#!/usr/bin/env python3
"""MCP round-trip smoke check for all supported document formats."""
import asyncio
import json
import os
import shutil
import sys
import tempfile

from mcp.client.stdio import stdio_client, StdioServerParameters
from mcp.client import ClientSession

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
PLUGIN_DIR = os.path.dirname(HERE)
SERVER = os.path.join(PLUGIN_DIR, "src", "server.py")
WORKDIR = tempfile.mkdtemp(prefix="dsh-document-editor-", dir=PLUGIN_DIR)


def result_data(result):
    if getattr(result, "is_error", False):
        raise AssertionError(result.content[0].text)
    data = json.loads(result.content[0].text)
    if data.get("error"):
        raise AssertionError(data["error"])
    return data


async def main():
    params = StdioServerParameters(
        command=sys.executable,
        args=[SERVER],
        cwd=PLUGIN_DIR,
        env={**os.environ, "PYTHONIOENCODING": "utf-8"},
    )

    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            init = await session.initialize()
            assert init.capabilities.tools is not None
            tools = await session.list_tools()
            names = {tool.name for tool in tools.tools}
            assert len(names) == 13, f"expected 13 tools, got {sorted(names)}"
            assert {"read_pdf", "verify_document"} <= names
            print(f"MCP handshake OK; {len(names)} tools registered")

            word = os.path.join(WORKDIR, "roundtrip.docx")
            result_data(await session.call_tool("create_word_document", {
                "title": "文档测试",
                "paragraphs": [{"text": "原始正文", "style": "Normal"}],
                "output_path": word,
            }))
            edited_word = result_data(await session.call_tool("modify_word_document", {
                "input_path": word,
                "paragraphs": [{"index": 1, "text": "修改后的正文"}],
                "tables": [{"headers": ["字段", "值"], "rows": [["状态", "待修改"]]}],
                "table_updates": [{"table_index": 0, "row_index": 1, "column_index": 1, "text": "完成"}],
            }))["output_path"]
            assert edited_word != word and os.path.isfile(word)
            word_data = result_data(await session.call_tool("read_word_document", {"input_path": edited_word}))
            assert word_data["paragraphs"][1]["text"] == "修改后的正文"
            assert word_data["tables"][0]["data"][1][1] == "完成"

            xlsx = os.path.join(WORKDIR, "roundtrip.xlsx")
            result_data(await session.call_tool("create_excel_workbook", {
                "sheets": [{"name": "数据", "headers": ["名称", "值"], "rows": [["a", 1], ["b", 2]]}],
                "output_path": xlsx,
            }))
            edited_xlsx = result_data(await session.call_tool("modify_excel_worksheet", {
                "input_path": xlsx,
                "cells": {"B2": 42, "C2": "=SUM(A2:B2)"},
                "rows": [["c", 3]],
            }))["output_path"]
            assert edited_xlsx != xlsx and os.path.isfile(xlsx)
            xlsx_data = result_data(await session.call_tool("read_excel_workbook", {"input_path": edited_xlsx}))
            assert xlsx_data["sheets"]["数据"]["rows"][0][1] == 42
            assert xlsx_data["sheets"]["数据"]["formulas"]["C2"] == "=SUM(A2:B2)"
            assert xlsx_data["sheets"]["数据"]["rows"][2][0] == "c"

            pdf = os.path.join(WORKDIR, "roundtrip.pdf")
            result_data(await session.call_tool("create_pdf", {
                "title": "PDF 测试",
                "pages": [{"content": "中文内容，PDF 自动换页检查。" * 250}],
                "output_path": pdf,
            }))
            pdf_data = result_data(await session.call_tool("read_pdf", {"input_path": pdf}))
            assert pdf_data["page_count"] > 1
            assert "中文内容" in pdf_data["pages"][0]["text"]
            edited_pdf = result_data(await session.call_tool("edit_pdf", {
                "input_path": pdf,
                "content": "追加页面。" * 300,
            }))["output_path"]
            assert edited_pdf != pdf
            assert result_data(await session.call_tool("read_pdf", {"input_path": edited_pdf}))["page_count"] > pdf_data["page_count"]

            pptx = os.path.join(WORKDIR, "roundtrip.pptx")
            result_data(await session.call_tool("create_pptx_presentation", {
                "title": "演示测试",
                "slides": [
                    {"layout": "title_and_content", "title": "第一页", "content": "初始正文"},
                    {"layout": "blank", "title": "空白标题", "content": "空白正文"},
                ],
                "output_path": pptx,
            }))
            ppt_data = result_data(await session.call_tool("read_pptx_presentation", {"input_path": pptx}))
            assert any("空白正文" in shape["text"] for shape in ppt_data["slides"][1]["shapes"])
            edited_pptx = result_data(await session.call_tool("modify_pptx_presentation", {
                "input_path": pptx,
                "slide_data": {"index": 0, "title": "更新标题", "content": "更新正文"},
            }))["output_path"]
            ppt_data = result_data(await session.call_tool("read_pptx_presentation", {"input_path": edited_pptx}))
            assert any("更新标题" in shape["text"] for shape in ppt_data["slides"][0]["shapes"])
            assert any("更新正文" in shape["text"] for shape in ppt_data["slides"][0]["shapes"])

            for path in (pdf, edited_pdf, edited_word, edited_xlsx, edited_pptx):
                check = result_data(await session.call_tool("verify_document", {"input_path": path}))
                assert check["valid"], f"verification failed for {path}: {check}"
            print("Create, read, edit-copy, CJK PDF pagination, and verification passed for all formats")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    finally:
        shutil.rmtree(WORKDIR, ignore_errors=True)
