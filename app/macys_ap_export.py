import html
import zipfile


def _clean_sheet_name(name, fallback):
    safe = str(name or "").strip()
    for char in ("/", "\\", "*", "?", "[", "]", ":"):
        safe = safe.replace(char, "-")
    return safe[:31] or fallback


def _column_letter(n):
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _cell_xml(value, row_idx, col_idx, style=""):
    ref = f"{_column_letter(col_idx)}{row_idx}"
    value = "" if value is None else str(value)
    style_attr = f' s="{style}"' if style else ""
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{html.escape(value)}</t></is></c>'


def write_xlsx_workbook(path, sheets):
    safe_sheets = []
    for idx, (name, headers, rows) in enumerate(sheets, start=1):
        safe_sheets.append((_clean_sheet_name(name, f"Sheet{idx}"), headers, rows))

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>"""
            + "".join(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                for i in range(1, len(safe_sheets) + 1)
            )
            + "</Types>",
        )
        z.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""",
        )
        z.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>"""
            + "".join(
                f'<sheet name="{html.escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
                for i, (name, _, _) in enumerate(safe_sheets, start=1)
            )
            + "</sheets></workbook>",
        )
        z.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"""
            + "".join(
                f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                for i in range(1, len(safe_sheets) + 1)
            )
            + f'<Relationship Id="rId{len(safe_sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        )
        z.writestr(
            "xl/styles.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Segoe UI"/></font><font><b/><sz val="11"/><name val="Segoe UI"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEFEFEF"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" xfId="0"/><xf fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>""",
        )
        for idx, (_, headers, rows) in enumerate(safe_sheets, start=1):
            sheet_rows = []
            sheet_rows.append(
                f'<row r="1">'
                + "".join(_cell_xml(h, 1, c, "1") for c, h in enumerate(headers, start=1))
                + "</row>"
            )
            for r_idx, row in enumerate(rows, start=2):
                sheet_rows.append(
                    f'<row r="{r_idx}">'
                    + "".join(_cell_xml(v, r_idx, c) for c, v in enumerate(row, start=1))
                    + "</row>"
                )
            widths = "".join(
                f'<col min="{c}" max="{c}" width="22" customWidth="1"/>'
                for c in range(1, len(headers) + 1)
            )
            z.writestr(
                f"xl/worksheets/sheet{idx}.xml",
                f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols>{widths}</cols><sheetData>{''.join(sheet_rows)}</sheetData><autoFilter ref="A1:{_column_letter(max(1, len(headers)))}{max(1, len(rows)+1)}"/></worksheet>""",
            )
