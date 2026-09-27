#!/usr/bin/env python3
"""Verify Chinese text survives the stdio MCP round-trip WITHOUT PYTHONIOENCODING."""
import asyncio, json, os, sys
from mcp.client.stdio import stdio_client, StdioServerParameters
from mcp.client import ClientSession

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

HERE = os.path.dirname(os.path.abspath(__file__))
PLUGIN = os.path.dirname(HERE)
SERVER = os.path.join(PLUGIN, "src", "server.py")
OUT = os.path.join(PLUGIN, "test-output", "enc.docx")
os.makedirs(os.path.dirname(OUT), exist_ok=True)

CN = "中文测试：项目进度报告（含标点、全角符号）"

async def main():
    # Deliberately do NOT set PYTHONIOENCODING
    env = {k: v for k, v in os.environ.items() if k != "PYTHONIOENCODING"}
    params = StdioServerParameters(command=sys.executable, args=[SERVER], cwd=PLUGIN, env=env)

    async with stdio_client(params) as (r, w):
        async with ClientSession(r, w) as s:
            await s.initialize()
            await s.call_tool("create_word_document", {
                "title": CN,
                "paragraphs": [{"text": CN, "style": "Normal"}],
                "output_path": OUT,
            })
            res = await s.call_tool("read_word_document", {"input_path": OUT})
            data = json.loads(res.content[0].text)
            got = data["paragraphs"][0]["text"]
            print(f"sent: {CN}")
            print(f"got : {got}")
            print("RESULT:", "ENCODING OK" if got == CN else "ENCODING CORRUPTED")

asyncio.run(main())
