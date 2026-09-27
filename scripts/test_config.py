#!/usr/bin/env python3
"""Verify deployment-style MCP launch from outside the plugin directory.

Set DOCUMENT_EDITOR_PYTHON, DOCUMENT_EDITOR_SERVER, and DOCUMENT_EDITOR_CWD
to check the exact command configured in cordis.patch.yml.
"""
import asyncio
import json
import os
import shutil
import sys
import tempfile

try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

from mcp.client.stdio import stdio_client, StdioServerParameters
from mcp.client import ClientSession

PLUGIN_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COMMAND = os.environ.get("DOCUMENT_EDITOR_PYTHON", sys.executable)
SERVER = os.environ.get("DOCUMENT_EDITOR_SERVER", os.path.join(PLUGIN_DIR, "src", "server.py"))
# DSH does not guarantee the plugin dir as cwd — default to its parent.
CWD = os.environ.get("DOCUMENT_EDITOR_CWD", os.path.dirname(PLUGIN_DIR))


async def main():
    assert os.path.isfile(COMMAND), f"interpreter not found: {COMMAND}"
    assert os.path.isfile(SERVER), f"server script not found: {SERVER}"
    outdir = tempfile.mkdtemp(prefix="dsh-document-editor-config-")
    try:
        params = StdioServerParameters(command=COMMAND, args=[SERVER], cwd=CWD)
        async with stdio_client(params) as (r, w):
            async with ClientSession(r, w) as s:
                await s.initialize()
                tools = (await s.list_tools()).tools
                print(f"handshake OK, tools={len(tools)}")

                target = os.path.join(outdir, "cfg.docx")
                res = await s.call_tool("create_word_document", {
                    "title": "配置校验",
                    "paragraphs": [{"text": "由 cordis.patch.yml 指定的命令创建。", "style": "Normal"}],
                    "output_path": target,
                })
                print("create:", res.content[0].text)

                res = await s.call_tool("read_word_document", {"input_path": target})
                data = json.loads(res.content[0].text)
                body = "由 cordis.patch.yml 指定的命令创建。"
                assert any(p["text"] == body for p in data["paragraphs"])
                print("read  :", body)
                print("bytes :", os.path.getsize(target))

        print("*** CONFIG VERIFIED ***")
    finally:
        shutil.rmtree(outdir, ignore_errors=True)


if __name__ == "__main__":
    asyncio.run(main())
