#!/usr/bin/env python3
"""
dsh-document-editor MCP Server (SDK v2)
========================================

Provides tools for creating, reading, editing, and verifying:
  PDF, Word/DOCX, Excel/XLSX, PowerPoint/PPTX.

Usage:
    python src/server.py

Dependencies:
    pip install mcp python-docx openpyxl python-pptx pypdf reportlab
"""

import asyncio
import json
import os
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional

# ── MCP SDK v2 imports ───────────────────────────────────────────
try:
    from mcp.server.mcpserver import MCPServer
except ImportError:
    print("ERROR: 'mcp' SDK not installed. Run: pip install mcp", file=sys.stderr)
    sys.exit(1)

# ── Library availability detection ───────────────────────────────

DOCX_OK = False
try:
    from docx import Document as DOCX_Document
    DOCX_OK = True
except ImportError:
    pass

EXCEL_OK = False
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    EXCEL_OK = True
except ImportError:
    pass

PPT_OK = False
try:
    from pptx import Presentation as PPT_Presentation
    from pptx.util import Inches
    PPT_OK = True
except ImportError:
    pass

PYPDF_OK = False
RL_OK = False
try:
    from pypdf import PdfReader, PdfWriter as PDF_Writer_cls
    PYPDF_OK = True
except ImportError:
    pass

try:
    from reportlab.pdfgen import canvas as RL_canvas
    from reportlab.lib.pagesizes import A4 as RL_A4, LETTER as RL_LETTER
    from reportlab.pdfbase import pdfmetrics as RL_pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    RL_OK = True
    RL_pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    RL_FONT = "STSong-Light"
except ImportError:
    pass

PDF_OK = PYPDF_OK and RL_OK


# === Helpers =====================================================

def _resolve(s: str) -> Path:
    p = Path(s)
    return p if p.is_absolute() else Path(os.getcwd()) / p


def _edited_path(input_path: str) -> Path:
    inp = _resolve(input_path)
    return inp.with_name(f"{inp.stem}_edited{inp.suffix}")


def _save_atomic(out: Path, save):
    """Write beside the destination, then replace it only after a successful save."""
    out.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(dir=out.parent, suffix=out.suffix)
    os.close(fd)
    try:
        save(temp_path)
        os.replace(temp_path, out)
    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)


def _wrap(text: str, width: float, font_size: int = 10):
    """Wrap at PDF font width so CJK text and long unspaced lines fit."""
    out = []
    for line in (text or "").split('\n'):
        if not line.strip():
            out.append('')
            continue
        buf = ''
        for char in line:
            cand = buf + char
            if buf and RL_pdfmetrics.stringWidth(cand, RL_FONT, font_size) > width:
                out.append(buf)
                buf = char
            else:
                buf = cand
        if buf:
            out.append(buf)
    return '\n'.join(out)


def _pdf_page_size(value: str):
    sizes = {"a4": RL_A4, "letter": RL_LETTER}
    try:
        return sizes[str(value).strip().lower()]
    except KeyError:
        raise ValueError("pagesize must be A4 or Letter")


def _render_pdf_text(pages, output_path: str, title: str = "", author: str = "") -> int:
    """Render text pages; each requested page may flow onto additional pages."""
    out = _resolve(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    fd, temp_path = tempfile.mkstemp(dir=out.parent, suffix=".pdf")
    os.close(fd)
    page_count = 0
    try:
        page_data = pages or [{"content": title or ""}]
        first_size = _pdf_page_size(page_data[0].get("pagesize", "A4"))
        c = RL_canvas.Canvas(temp_path, pagesize=first_size)
        if title:
            c.setTitle(title)
        if author:
            c.setAuthor(author)

        for pd in page_data:
            page_size = _pdf_page_size(pd.get("pagesize", "A4"))
            c.setPageSize(page_size)
            width, height = page_size
            x, y, margin = 72, height - 72, 72
            c.setFont(RL_FONT, 10)
            for line in _wrap(pd.get("content", ""), width - 2 * margin).split('\n'):
                if y < margin:
                    c.showPage()
                    page_count += 1
                    c.setPageSize(page_size)
                    c.setFont(RL_FONT, 10)
                    y = height - margin
                if line:
                    c.drawString(x, y, line)
                y -= 14
            c.showPage()
            page_count += 1
        c.save()

        # Keep this helper reusable for temporary output and final files.
        if out != Path(temp_path):
            os.replace(temp_path, out)
        return page_count
    finally:
        if os.path.exists(temp_path):
            os.unlink(temp_path)


# ---------------------------------------------------------------
#  PDF handlers
# ---------------------------------------------------------------

def _tool_create_pdf(title: str, author: str, output_path: str, pages):
    """Tool handler for create_pdf."""
    if not PDF_OK:
        return {"error": "Need reportlab + pypdf (pip install reportlab pypdf)"}
    try:
        pg_list = pages if pages else [{"content": title or ""}]
        out = _resolve(output_path)
        page_count = _render_pdf_text(pg_list, str(out), title, author)
        return {"success": True, "output_path": str(out), "page_count": page_count}
    except Exception as exc:
        return {"error": f"PDF create failed: {exc}"}


def _tool_edit_pdf(input_path: str, output_path: str, content: str):
    """Tool handler for edit_pdf."""
    if not PDF_OK:
        return {"error": "Need reportlab + pypdf (pip install reportlab pypdf)"}
    try:
        inp = _resolve(input_path)
        if not inp.exists():
            return {"error": f"File not found: {input_path}"}

        out = _resolve(output_path) if output_path else _edited_path(input_path)
        out.parent.mkdir(parents=True, exist_ok=True)

        reader = PdfReader(str(inp))
        writer = PDF_Writer_cls()
        for page in reader.pages:
            writer.add_page(page)
        if reader.metadata:
            writer.add_metadata({str(k): str(v) for k, v in reader.metadata.items() if str(k).startswith("/") and v is not None})

        if content:
            fd, tmp = tempfile.mkstemp(suffix=".pdf")
            os.close(fd)
            try:
                _render_pdf_text([{"content": content}], tmp)
                for page in PdfReader(tmp).pages:
                    writer.add_page(page)
            finally:
                if os.path.exists(tmp):
                    os.unlink(tmp)

        _save_atomic(out, lambda temp: writer.write(temp))
        return {"success": True, "output_path": str(out), "page_count": len(writer.pages)}
    except Exception as exc:
        return {"error": f"PDF edit failed: {exc}"}


def _tool_read_pdf(input_path: str):
    if not PYPDF_OK:
        return {"error": "Need pypdf (pip install pypdf)"}
    try:
        inp = _resolve(input_path)
        if not inp.is_file():
            return {"error": f"File not found: {input_path}"}
        reader = PdfReader(str(inp))
        return {
            "success": True,
            "page_count": len(reader.pages),
            "metadata": {str(k): str(v) for k, v in (reader.metadata or {}).items() if v is not None},
            "pages": [{"index": i, "text": page.extract_text() or ""} for i, page in enumerate(reader.pages)],
        }
    except Exception as exc:
        return {"error": f"PDF read failed: {exc}"}


def _tool_verify_document(input_path: str):
    """Check that a supported file opens and its container/structure can be read."""
    inp = _resolve(input_path)
    if not inp.is_file():
        return {"error": f"File not found: {input_path}"}
    ext = inp.suffix.lower()
    try:
        if ext == ".pdf":
            if not PYPDF_OK:
                return {"error": "Need pypdf (pip install pypdf)"}
            reader = PdfReader(str(inp), strict=True)
            page_text = [page.extract_text() or "" for page in reader.pages]
            return {
                "success": True,
                "valid": True,
                "format": "pdf",
                "page_count": len(reader.pages),
                "text_pages": sum(bool(text.strip()) for text in page_text),
                "checks": ["strict_parse", "pages_read"],
            }

        parsers = {".docx": ("docx", DOCX_OK), ".xlsx": ("xlsx", EXCEL_OK), ".pptx": ("pptx", PPT_OK)}
        if ext not in parsers:
            return {"error": f"Unsupported document format: {ext or '(no extension)'}"}
        kind, available = parsers[ext]
        if not available:
            package = {".docx": "python-docx", ".xlsx": "openpyxl", ".pptx": "python-pptx"}[ext]
            return {"error": f"Need {package} to verify this format"}

        with zipfile.ZipFile(inp) as archive:
            corrupt_member = archive.testzip()
            if corrupt_member:
                raise ValueError(f"Corrupt package member: {corrupt_member}")

        if kind == "docx":
            doc = DOCX_Document(str(inp))
            summary = {"paragraph_count": len(doc.paragraphs), "table_count": len(doc.tables)}
        elif kind == "xlsx":
            wb = openpyxl.load_workbook(str(inp), read_only=True, data_only=False)
            summary = {"sheet_names": list(wb.sheetnames), "sheet_count": len(wb.sheetnames)}
            wb.close()
        else:
            prs = PPT_Presentation(str(inp))
            summary = {"slide_count": len(prs.slides), "shape_count": sum(len(slide.shapes) for slide in prs.slides)}
        return {"success": True, "valid": True, "format": kind, **summary, "checks": ["zip_crc", "parser_open"]}
    except Exception as exc:
        return {"success": True, "valid": False, "format": ext.lstrip("."), "error": str(exc)}


# ---------------------------------------------------------------
#  Word / DOCX handlers
# ---------------------------------------------------------------

def _tool_create_word(title: str, paragraphs, output_path: str, tables=None):
    if not DOCX_OK:
        return {"error": "Need python-docx (pip install python-docx)"}
    try:
        doc = DOCX_Document()
        if title:
            doc.add_heading(title, level=0)

        for pd in (paragraphs or []):
            txt = pd.get("text", "")
            sty = pd.get("style", "Normal")
            if isinstance(sty, str) and sty.lower().startswith("heading"):
                lvl = int(sty[-1]) if sty[-1].isdigit() else 1
                doc.add_heading(txt, level=lvl)
            elif not txt:
                continue
            else:
                doc.add_paragraph(txt)

        for td in (tables or []):
            data = ([td.get("headers", [])] if td.get("headers") else []) + (td.get("rows", []) or [])
            col_count = max((len(row) for row in data), default=0)
            if col_count:
                table = doc.add_table(rows=0, cols=col_count)
                table.style = "Table Grid"
                for row_data in data:
                    cells = table.add_row().cells
                    for ci, value in enumerate(row_data):
                        cells[ci].text = str(value) if value is not None else ""

        out = _resolve(output_path)
        _save_atomic(out, doc.save)
        return {"success": True, "output_path": str(out), "paragraph_count": len(paragraphs or [])}
    except Exception as exc:
        return {"error": f"Word create failed: {exc}"}


def _tool_read_word(input_path: str):
    if not DOCX_OK:
        return {"error": "Need python-docx (pip install python-docx)"}
    try:
        inp = _resolve(input_path)
        if not inp.exists():
            return {"error": f"File not found: {input_path}"}

        doc = DOCX_Document(str(inp))
        paras = [{"text": p.text, "style": p.style.name if hasattr(p, 'style') else "?"} for p in doc.paragraphs]
        tables = []
        for ti, tbl in enumerate(doc.tables):
            tr = [[c.text for c in row.cells] for row in tbl.rows]
            tables.append({"index": ti, "rows": len(tbl.rows), "cols": len(tbl.columns), "data": tr})
        return {"success": True, "paragraphs": paras, "tables": tables}
    except Exception as exc:
        return {"error": f"Word read failed: {exc}"}


def _tool_modify_word(input_path: str, output_path: str, paragraphs, tables=None, table_updates=None):
    if not DOCX_OK:
        return {"error": "Need python-docx (pip install python-docx)"}
    try:
        inp = _resolve(input_path)
        if not inp.is_file():
            return {"error": f"File not found: {input_path}"}
        out = _resolve(output_path) if output_path else _edited_path(input_path)

        doc = DOCX_Document(str(inp))
        for pd in (paragraphs or []):
            txt = pd.get("text", "")
            sty = pd.get("style", "Normal")
            if "index" in pd:
                index = int(pd["index"])
                if index < 0 or index >= len(doc.paragraphs):
                    raise ValueError(f"Paragraph index out of range: {index}")
                doc.paragraphs[index].text = txt
            elif txt:
                if isinstance(sty, str) and sty.lower().startswith("heading"):
                    lvl = int(sty[-1]) if sty[-1].isdigit() else 1
                    doc.add_heading(txt, level=lvl)
                else:
                    doc.add_paragraph(txt, style=sty)

        for td in (tables or []):
            data = ([td.get("headers", [])] if td.get("headers") else []) + (td.get("rows", []) or [])
            col_count = max((len(row) for row in data), default=0)
            if col_count:
                table = doc.add_table(rows=0, cols=col_count)
                table.style = "Table Grid"
                for row_data in data:
                    cells = table.add_row().cells
                    for ci, value in enumerate(row_data):
                        cells[ci].text = str(value) if value is not None else ""

        for update in (table_updates or []):
            ti, ri, ci = (int(update[key]) for key in ("table_index", "row_index", "column_index"))
            if ti < 0 or ti >= len(doc.tables) or ri < 0 or ri >= len(doc.tables[ti].rows) or ci < 0 or ci >= len(doc.tables[ti].columns):
                raise ValueError(f"Table cell index out of range: table={ti}, row={ri}, column={ci}")
            value = update.get("text", "")
            doc.tables[ti].cell(ri, ci).text = str(value) if value is not None else ""

        _save_atomic(out, doc.save)

        return {"success": True, "output_path": str(out)}
    except Exception as exc:
        return {"error": f"Word modify failed: {exc}"}


# ---------------------------------------------------------------
#  Excel / XLSX handlers
# ---------------------------------------------------------------

def _tool_create_excel(sheets, output_path: str):
    if not EXCEL_OK:
        return {"error": "Need openpyxl (pip install openpyxl)"}
    try:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # remove default empty sheet

        for sd in (sheets or [{"name": "Sheet1", "headers": [], "rows": []}]):
            ws = wb.create_sheet(title=sd.get("name", "Sheet"))
            for ci, hdr in enumerate(sd.get("headers", []), 1):
                cell = ws.cell(row=1, column=ci, value=hdr)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                cell.alignment = Alignment(horizontal="center")

            for ri, row_data in enumerate(sd.get("rows", []), 2):
                for ci, val in enumerate(row_data, 1):
                    ws.cell(row=ri, column=ci, value=val)

        out = _resolve(output_path)
        _save_atomic(out, wb.save)
        return {"success": True, "output_path": str(out), "sheet_count": len(sheets or [1])}
    except Exception as exc:
        return {"error": f"Excel create failed: {exc}"}


def _tool_read_excel(input_path: str):  # sheet_name is ignored (reads all)
    if not EXCEL_OK:
        return {"error": "Need openpyxl (pip install openpyxl)"}
    try:
        inp = _resolve(input_path)
        if not inp.exists():
            return {"error": f"File not found: {input_path}"}

        wb = openpyxl.load_workbook(str(inp), data_only=True)
        formula_wb = openpyxl.load_workbook(str(inp), data_only=False)
        result_sheets = {}
        for name in wb.sheetnames:
            ws = wb[name]
            formula_ws = formula_wb[name]
            rd = []
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=max(ws.max_column, 1), values_only=True):
                rd.append([v for v in row])
            result_sheets[name] = {
                "headers": rd[0] if rd else [],
                "rows": rd[1:] if len(rd) > 1 else [],
                "formulas": {
                    cell.coordinate: cell.value
                    for row in formula_ws.iter_rows()
                    for cell in row
                    if cell.data_type == "f"
                },
            }

        wb.close()
        formula_wb.close()
        return {"success": True, "sheets": result_sheets, "sheet_names": list(wb.sheetnames)}
    except Exception as exc:
        return {"error": f"Excel read failed: {exc}"}


def _tool_modify_excel(input_path: str, sheet_name: str, headers, rows, output_path: str, cells=None):
    if not EXCEL_OK:
        return {"error": "Need openpyxl (pip install openpyxl)"}
    try:
        inp = _resolve(input_path)
        if not inp.is_file():
            return {"error": f"File not found: {input_path}"}
        wb = openpyxl.load_workbook(str(inp))

        ws = None
        if not sheet_name:
            ws = wb.active
        elif sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            for existing in wb.sheetnames:
                if existing.lower() == sheet_name.lower():
                    ws = wb[existing]; break

        if not ws:
            ws = wb.create_sheet(title=sheet_name or "Sheet")

        # headers
        for ci, hdr in enumerate(headers or [], 1):
            cell = ws.cell(row=1, column=ci, value=hdr)
            cell.font = Font(bold=True)
            cell.fill = PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid")

        # Add rows after current data; never replace existing rows implicitly.
        sr = max(ws.max_row + 1, 2 if headers else 1)
        for ri, row_data in enumerate(rows or [], sr):
            for ci, val in enumerate(row_data, 1):
                ws.cell(row=ri, column=ci, value=val)

        for address, value in (cells or {}).items():
            ws[str(address)] = value

        out = _resolve(output_path) if output_path else _edited_path(input_path)
        _save_atomic(out, wb.save)
        wb.close()
        return {"success": True, "output_path": str(out)}
    except Exception as exc:
        return {"error": f"Excel modify failed: {exc}"}


# ---------------------------------------------------------------
#  PowerPoint / PPTX handlers
# ---------------------------------------------------------------

def _ppt_text_shape(slide, exclude=None):
    exclude_id = getattr(exclude, "shape_id", None)
    for shape in slide.shapes:
        if shape.shape_id != exclude_id and shape.name == "Generated Body" and hasattr(shape, "text_frame"):
            return shape
    for shape in slide.shapes:
        if shape.shape_id == exclude_id or not hasattr(shape, "text_frame"):
            continue
        return shape
    return None


def _ppt_title_shape(slide):
    return slide.shapes.title or next((s for s in slide.shapes if s.name == "Generated Title"), None)


def _set_ppt_text(shape, text: str):
    shape.text_frame.clear()
    shape.text_frame.text = text


def _add_ppt_textbox(slide, text: str, *, title: bool = False):
    top = Inches(0.45) if title else Inches(1.5)
    height = Inches(0.7) if title else Inches(5.3)
    shape = slide.shapes.add_textbox(Inches(0.8), top, Inches(11.7), height)
    shape.name = "Generated Title" if title else "Generated Body"
    shape.text_frame.text = text
    return shape


def _ppt_add_slide(prs, data, fallback_title=""):
    layouts = {"title": 0, "title_and_content": 1, "content": 1, "section_header": 2, "two_content": 3, "blank": 6}
    requested = data.get("layout", "title_and_content")
    layout_index = layouts.get(requested, layouts["title_and_content"])
    slide = prs.slides.add_slide(prs.slide_layouts[min(layout_index, len(prs.slide_layouts) - 1)])
    title = data.get("title") or fallback_title
    content = data.get("content", "")
    title_shape = _ppt_title_shape(slide)
    if title:
        if title_shape:
            _set_ppt_text(title_shape, title)
        else:
            title_shape = _add_ppt_textbox(slide, title, title=True)
    if content:
        body_shape = _ppt_text_shape(slide, exclude=title_shape)
        if body_shape:
            _set_ppt_text(body_shape, content)
        else:
            _add_ppt_textbox(slide, content)
    return slide


def _tool_create_pptx(title: str, slides, output_path: str):
    if not PPT_OK:
        return {"error": "Need python-pptx (pip install python-pptx)"}
    try:
        prs = PPT_Presentation()  # default blank

        slide_list = slides or [{"layout": "title", "title": title, "content": ""}]
        for index, sd in enumerate(slide_list):
            _ppt_add_slide(prs, sd, title if index == 0 else "")

        out = _resolve(output_path)
        _save_atomic(out, prs.save)
        return {"success": True, "output_path": str(out), "slide_count": len(slide_list)}
    except Exception as exc:
        return {"error": f"PPT create failed: {exc}"}


def _tool_read_ppt(input_path: str):  # Note: this function name is just internal
    if not PPT_OK:
        return {"error": "Need python-pptx (pip install python-pptx)"}
    try:
        inp = _resolve(input_path)
        if not inp.exists():
            return {"error": f"File not found: {input_path}"}

        prs = PPT_Presentation(str(inp))
        ss = []
        for idx, sl in enumerate(prs.slides):
            shapes_info = []
            for shp in sl.shapes:
                if hasattr(shp, 'text') and shp.text.strip():
                    shapes_info.append({
                        "name": shp.name, "type": str(shp.shape_type),
                        "text": shp.text})

            ss.append({"index": idx, "layout": str(sl.slide_layout.name), "shapes": shapes_info})

        return {"success": True, "slide_count": len(prs.slides), "slides": ss}
    except Exception as exc:
        return {"error": f"PPT read failed: {exc}"}


def _tool_modify_ppt(input_path: str, slide_data: dict, output_path: str):
    if not PPT_OK:
        return {"error": "Need python-pptx (pip install python-pptx)"}
    try:
        inp = _resolve(input_path)
        if not inp.is_file():
            return {"error": f"File not found: {input_path}"}

        prs = PPT_Presentation(str(inp))

        sd = slide_data or {}
        if "index" in sd:
            index = int(sd["index"])
            if index < 0 or index >= len(prs.slides):
                raise ValueError(f"Slide index out of range: {index}")
            slide = prs.slides[index]
            title_shape = _ppt_title_shape(slide)
            if "title" in sd:
                if title_shape:
                    _set_ppt_text(title_shape, str(sd["title"]))
                elif sd["title"]:
                    title_shape = _add_ppt_textbox(slide, str(sd["title"]), title=True)
            if "content" in sd:
                body_shape = _ppt_text_shape(slide, exclude=title_shape)
                if body_shape:
                    _set_ppt_text(body_shape, str(sd["content"]))
                elif sd["content"]:
                    _add_ppt_textbox(slide, str(sd["content"]))
        else:
            _ppt_add_slide(prs, sd)

        out = _resolve(output_path) if output_path else _edited_path(input_path)
        _save_atomic(out, prs.save)
        return {"success": True, "output_path": str(out), "slide_count": len(prs.slides)}
    except Exception as exc:
        return {"error": f"PPT modify failed: {exc}"}


# ================================================================
#  Tool registration
# ================================================================

DOC_NOTE = (
    "Paths may be absolute or relative to CWD. Edit tools save a sibling *_edited copy unless output_path is supplied."
)

doc_editor = MCPServer("dsh-document-editor")


# ── PDF ─────────────────────────────────────────────────────────

@doc_editor.tool(
    name="create_pdf",
    description=f"Create a PDF document from one or more text pages.\n{DOC_NOTE}\nRequires: pip install reportlab pypdf",
)
def _handle_create_pdf(
    pages: List[Dict[str, Any]],
    title: str = "",
    author: str = "",
    output_path: str = "output.pdf",
) -> Dict[str, Any]:
    return _tool_create_pdf(
        title=title,
        author=author,
        output_path=output_path,
        pages=pages,
    )


@doc_editor.tool(
    name="edit_pdf",
    description=f"Append text to one or more pages in an existing PDF and save a copy by default.\n{DOC_NOTE}",
)
def _handle_edit_pdf(
    input_path: str,
    content: str = "",
    output_path: str = "",
) -> Dict[str, Any]:
    return _tool_edit_pdf(
        input_path=input_path,
        output_path=output_path,
        content=content,
    )


@doc_editor.tool(
    name="read_pdf",
    description=f"Read PDF metadata and extract text page by page.\n{DOC_NOTE}\nRequires: pip install pypdf",
)
def _handle_read_pdf(input_path: str) -> Dict[str, Any]:
    return _tool_read_pdf(input_path=input_path)


@doc_editor.tool(
    name="verify_document",
    description=f"Verify that a PDF, DOCX, XLSX, or PPTX file opens and its structure is readable.\n{DOC_NOTE}",
)
def _handle_verify_document(input_path: str) -> Dict[str, Any]:
    return _tool_verify_document(input_path=input_path)


# ── Word ────────────────────────────────────────────────────────

@doc_editor.tool(
    name="create_word_document",
    description=f"Create a Word (.docx) document from paragraphs and tables.\n{DOC_NOTE}\nRequires: pip install python-docx",
)
def _handle_create_word(
    title: str = "",
    paragraphs: Optional[List[Dict[str, Any]]] = None,
    tables: Optional[List[Dict[str, Any]]] = None,
    output_path: str = "output.docx",
) -> Dict[str, Any]:
    return _tool_create_word(
        title=title,
        paragraphs=paragraphs,
        output_path=output_path,
        tables=tables,
    )


@doc_editor.tool(
    name="read_word_document",
    description=f"Read text paragraphs and tables from a .docx document.\n{DOC_NOTE}",
)
def _handle_read_word(input_path: str) -> Dict[str, Any]:
    return _tool_read_word(input_path=input_path)


@doc_editor.tool(
    name="modify_word_document",
    description=f"Replace or append paragraphs, add tables, or edit existing table cells in a .docx copy. Paragraph indexes are zero-based.\n{DOC_NOTE}",
)
def _handle_modify_word(
    input_path: str,
    paragraphs: Optional[List[Dict[str, Any]]] = None,
    tables: Optional[List[Dict[str, Any]]] = None,
    table_updates: Optional[List[Dict[str, Any]]] = None,
    output_path: str = "",
) -> Dict[str, Any]:
    return _tool_modify_word(
        input_path=input_path,
        output_path=output_path,
        paragraphs=paragraphs,
        tables=tables,
        table_updates=table_updates,
    )


# ── Excel ───────────────────────────────────────────────────────

@doc_editor.tool(
    name="create_excel_workbook",
    description=f"Create an Excel (.xlsx) workbook with one or more formatted sheets.\n{DOC_NOTE}\nRequires: pip install openpyxl",
)
def _handle_create_excel(
    sheets: Optional[List[Dict[str, Any]]] = None,
    output_path: str = "output.xlsx",
) -> Dict[str, Any]:
    return _tool_create_excel(sheets=sheets, output_path=output_path)


@doc_editor.tool(
    name="read_excel_workbook",
    description=f"Read all sheets, cached values, and formula expressions from an .xlsx workbook.\n{DOC_NOTE}",
)
def _handle_read_excel(input_path: str) -> Dict[str, Any]:
    return _tool_read_excel(input_path=input_path)


@doc_editor.tool(
    name="modify_excel_worksheet",
    description=f"Set cells/formulas and append rows in a worksheet, saving a copy by default.\n{DOC_NOTE}",
)
def _handle_modify_excel(
    input_path: str,
    sheet_name: str = "",
    headers: Optional[List[str]] = None,
    rows: Optional[List[List[Any]]] = None,
    cells: Optional[Dict[str, Any]] = None,
    output_path: str = "",
) -> Dict[str, Any]:
    return _tool_modify_excel(
        input_path=input_path,
        sheet_name=sheet_name,
        headers=headers,
        rows=rows,
        cells=cells,
        output_path=output_path,
    )


# ── PowerPoint ──────────────────────────────────────────────────

@doc_editor.tool(
    name="create_pptx_presentation",
    description=f"Create a PowerPoint (.pptx) presentation from slides.\n{DOC_NOTE}\nRequires: pip install python-pptx",
)
def _handle_create_ppt(
    title: str = "",
    slides: Optional[List[Dict[str, Any]]] = None,
    output_path: str = "output.pptx",
) -> Dict[str, Any]:
    return _tool_create_pptx(
        title=title,
        slides=slides,
        output_path=output_path,
    )


@doc_editor.tool(
    name="read_pptx_presentation",
    description=f"Read all slides and text shapes from a .pptx presentation.\n{DOC_NOTE}",
)
def _handle_read_ppt(input_path: str) -> Dict[str, Any]:
    return _tool_read_ppt(input_path=input_path)


@doc_editor.tool(
    name="modify_pptx_presentation",
    description=f"Edit a slide by zero-based index, or append a slide when index is omitted; saves a copy by default.\n{DOC_NOTE}",
)
def _handle_modify_ppt(
    input_path: str,
    slide_data: Optional[Dict[str, Any]] = None,
    output_path: str = "",
) -> Dict[str, Any]:
    return _tool_modify_ppt(
        input_path=input_path,
        slide_data=slide_data,
        output_path=output_path,
    )


# ── Stdio entry point ───────────────────────────────────────────

async def main():
    await doc_editor.run_stdio_async()


if __name__ == "__main__":
    asyncio.run(main())
