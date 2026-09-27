# dsh-document-editor

DSH MCP server for reading, creating, editing, and verifying PDF, Word, Excel, and PowerPoint files. The tools use the existing Python libraries and return structured results.

## Tools

| Format | Tools |
| --- | --- |
| PDF | `create_pdf`, `read_pdf`, `edit_pdf` |
| Word (`.docx`) | `create_word_document`, `read_word_document`, `modify_word_document` |
| Excel (`.xlsx`) | `create_excel_workbook`, `read_excel_workbook`, `modify_excel_worksheet` |
| PowerPoint (`.pptx`) | `create_pptx_presentation`, `read_pptx_presentation`, `modify_pptx_presentation` |
| All four | `verify_document` |

Edit tools create a sibling `*_edited` file by default, preserving the source. Pass `output_path` to choose another destination. Writes replace the destination only after the complete file has been saved successfully.

## Install

Python 3.10 or newer is required.

```powershell
pip install -r src\requirements.txt
```

Add `cordis.patch.yml` to the DSH profile and set `command` to the absolute path of the Python interpreter that has these dependencies installed. Restart the profile after changing its MCP configuration. Do not use `uv run` unless its environment also has the dependencies.

## Use

Paths may be absolute or relative to the MCP server's current working directory. Tools create parent folders when writing output files.

### Security and privacy

The server runs locally over MCP stdio and makes no network requests. It can read any path passed to a tool, subject to the operating system permissions of the user running it. Read tools return extracted document contents to the connected MCP client, which handles them under its own data policy.

## Formats

### PDF

`create_pdf` takes `pages: [{content, pagesize?}]`, with optional `title`, `author`, and `output_path`. Long text flows across as many pages as needed; `pagesize` accepts `A4` or `Letter`. Chinese text uses ReportLab's built-in STSong CID font.

`read_pdf` returns metadata and extracted text by page. `edit_pdf` appends text pages and writes a copy by default. `verify_document` strictly parses PDFs and reports page and readable-text counts.

```json
{
  "title": "项目报告",
  "author": "张三",
  "pages": [{"content": "中文报告正文。", "pagesize": "A4"}],
  "output_path": "report.pdf"
}
```

### Word

`create_word_document` accepts `paragraphs: [{text, style?}]` and `tables: [{headers?, rows}]`. Heading styles such as `heading1` and `heading2` are supported.

`modify_word_document` accepts `paragraphs`, `tables`, and `table_updates`. A paragraph with a zero-based `index` replaces that paragraph; one without an index is appended. `table_updates` entries use `{table_index, row_index, column_index, text}` with zero-based indexes. New tables use `{headers?, rows}`.

`read_word_document` returns paragraphs and table cell data.

### Excel

`create_excel_workbook` takes `sheets: [{name, headers?, rows?}]`.

`modify_excel_worksheet` edits the named sheet (creating it if needed); if omitted, it edits the first sheet. `cells` maps A1-style addresses to values or formulas, for example `{"B2": 12, "C2": "=SUM(A2:B2)"}`. `rows` are appended after existing data; `headers`, when supplied, update row 1. `read_excel_workbook` returns cached cell values in `rows` and formula expressions in a separate `formulas` map; formulas are not recalculated by openpyxl.

### PowerPoint

`create_pptx_presentation` takes `slides: [{layout, title?, content?}]`. Layouts include `title`, `title_and_content`, `section_header`, `two_content`, and `blank`. Text is added to placeholders when available and to text boxes otherwise.

`modify_pptx_presentation` takes `slide_data`. Include a zero-based `index` to replace that slide's title/content; omit it to append a slide. `read_pptx_presentation` returns all text shapes without truncation.

### Verify

`verify_document(input_path)` checks PDF parsing or Office package CRC and parser readability. A supported but damaged file returns `success: true, valid: false` with the error; unsupported extensions and missing files return `error`.

## Check

Run from the plugin directory:

```powershell
python scripts\test_e2e.py
python scripts\test_encoding.py
```

`test_e2e.py` performs an MCP handshake and round-trips the four formats, edits generated files, and calls verification. `test_config.py` checks a deployment-style launch from outside the plugin directory. To check a specific install, set `DOCUMENT_EDITOR_PYTHON`, `DOCUMENT_EDITOR_SERVER`, and `DOCUMENT_EDITOR_CWD` before running it.

## Examples

The [validation samples](validation-samples-2026-09-26/README.md) include original and edited PDF, DOCX, XLSX, and PPTX files created and verified through MCP.

## Limits

- PDF creation and editing handle text pages; editing existing page text, forms, annotations, images, and layout is not implemented.
- Word edits target paragraphs and table cells; advanced Word features such as tracked changes and embedded objects are preserved by python-docx only where its format support allows.
- Excel formulas are written and read but not calculated. Existing workbook macros are not supported by `.xlsx` tools.
- PowerPoint tools edit slide text and create basic text-first slides; they do not render slides or edit charts, media, or complex shape geometry.
- Office rendering can vary between Microsoft Office, WPS, LibreOffice, and Google Docs.

## Files

```text
src/server.py          MCP server
src/requirements.txt   Python dependencies
scripts/test_e2e.py    MCP and document round-trip check
scripts/test_encoding.py Chinese stdio round-trip check
cordis.patch.yml       DSH profile entry
```

MIT License.
