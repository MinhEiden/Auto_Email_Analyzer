from pathlib import Path

def _write_docx(path: Path, exec_summary: str, stats: dict) -> None:
    """Tạo file .docx chuyên nghiệp qua Node.js + thư viện docx."""
    import subprocess, tempfile, os, json as _json

    cat_rows_js = ""
    for cat, cnt in stats["category_counts"].items():
        pct = round(cnt / stats["total_emails"] * 100, 1)
        cat_rows_js += f"""
        new TableRow({{ children: [
            _cell(\"{cat}\",     3120, false),
            _cell(\"{cnt}\",     3120, false),
            _cell(\"{pct}%\",    3120, false),
        ]}}),"""

    urgent_rows_js = ""
    for item in stats["urgent_list"]:
        cat  = str(item.get("AI_Category", "")).replace('"', "'")
        frm  = str(item.get("From", "")).replace('"', "'")
        summ = str(item.get("AI_Summary", "")).replace('"', "'")
        urgent_rows_js += f"""
        new TableRow({{ children: [
            _cell(\"{cat}\",  2340, false),
            _cell(\"{frm}\",  2340, false),
            _cell(\"{summ}\", 4680, false),
        ]}}),"""

    exec_summary_escaped = exec_summary.replace('"', '\\"').replace('\n', ' ')

    js_script = f"""
const {{ Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
         AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
         LevelFormat, VerticalAlign }} = require('docx');
const fs = require('fs');

const border = {{ style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" }};
const borders = {{ top: border, bottom: border, left: border, right: border }};

function _cell(text, widthDxa, isHeader) {{
    return new TableCell({{
        borders,
        width: {{ size: widthDxa, type: WidthType.DXA }},
        shading: {{ fill: isHeader ? "2E75B6" : "FFFFFF", type: ShadingType.CLEAR }},
        margins: {{ top: 80, bottom: 80, left: 120, right: 120 }},
        children: [new Paragraph({{
            children: [new TextRun({{
                text: String(text),
                bold: isHeader,
                color: isHeader ? "FFFFFF" : "000000",
                font: "Arial",
                size: 20,
            }})]
        }})]
    }});
}}

function heading1(text) {{
    return new Paragraph({{
        heading: HeadingLevel.HEADING_1,
        spacing: {{ before: 320, after: 160 }},
        children: [new TextRun({{ text, bold: true, size: 32, font: "Arial", color: "2E75B6" }})]
    }});
}}

function heading2(text) {{
    return new Paragraph({{
        heading: HeadingLevel.HEADING_2,
        spacing: {{ before: 240, after: 120 }},
        children: [new TextRun({{ text, bold: true, size: 26, font: "Arial", color: "1F4E79" }})]
    }});
}}

function body(text) {{
    return new Paragraph({{
        spacing: {{ after: 160 }},
        children: [new TextRun({{ text, size: 22, font: "Arial" }})]
    }});
}}

function kpiRow(label, value, color) {{
    return new Paragraph({{
        spacing: {{ after: 120 }},
        children: [
            new TextRun({{ text: label + ": ", bold: true, size: 22, font: "Arial" }}),
            new TextRun({{ text: String(value), size: 22, font: "Arial", color: color || "000000" }}),
        ]
    }});
}}

const doc = new Document({{
    styles: {{
        default: {{ document: {{ run: {{ font: "Arial", size: 22 }} }} }},
    }},
    numbering: {{
        config: [{{
            reference: "bullets",
            levels: [{{ level: 0, format: LevelFormat.BULLET, text: "•",
                alignment: AlignmentType.LEFT,
                style: {{ paragraph: {{ indent: {{ left: 720, hanging: 360 }} }} }} }}]
        }}]
    }},
    sections: [{{
        properties: {{
            page: {{
                size: {{ width: 12240, height: 15840 }},
                margin: {{ top: 1440, right: 1440, bottom: 1440, left: 1440 }}
            }}
        }},
        children: [

            // ── TIÊU ĐỀ BÁO CÁO ────────────────────────────────────────────
            new Paragraph({{
                alignment: AlignmentType.CENTER,
                spacing: {{ before: 0, after: 240 }},
                children: [new TextRun({{
                    text: "BÁO CÁO PHÂN TÍCH EMAIL HÀNG NGÀY",
                    bold: true, size: 40, font: "Arial", color: "1F4E79",
                }})]
            }}),
            new Paragraph({{
                alignment: AlignmentType.CENTER,
                spacing: {{ after: 480 }},
                border: {{ bottom: {{ style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 }} }},
                children: [new TextRun({{
                    text: "Tổng hợp & Phân tích tự động bởi AI",
                    italics: true, size: 22, font: "Arial", color: "595959",
                }})]
            }}),

            // ── 1. EXECUTIVE SUMMARY ────────────────────────────────────────
            heading1("1. Tóm tắt Điều hành (Executive Summary)"),
            body("{exec_summary_escaped}"),

            // ── 2. SỐ LIỆU TỔNG QUAN ───────────────────────────────────────
            heading1("2. Số liệu Tổng quan"),
            kpiRow("Tổng số email",          {stats['total_emails']},    "000000"),
            kpiRow("Email mức Critical",      {stats['critical_count']},  "C00000"),
            kpiRow("Email mức High",          {stats['high_count']},      "E36C09"),
            kpiRow("Email cần xử lý gấp (Critical + High)",
                   {stats['critical_count'] + stats['high_count']}, "C00000"),

            // ── 3. PHÂN BỔ DANH MỤC ────────────────────────────────────────
            heading1("3. Phân bổ Danh mục"),
            new Table({{
                width: {{ size: 9360, type: WidthType.DXA }},
                columnWidths: [3120, 3120, 3120],
                rows: [
                    new TableRow({{ tableHeader: true, children: [
                        _cell("Danh mục",      3120, true),
                        _cell("Số lượng",      3120, true),
                        _cell("Tỷ lệ (%)",     3120, true),
                    ]}}),
                    {cat_rows_js}
                ]
            }}),

            // ── 4. HẠNG MỤC CẦN XỬ LÝ GẤP ────────────────────────────────
            heading1("4. Hạng mục Cần xử lý gấp (Critical & High)"),
            new Table({{
                width: {{ size: 9360, type: WidthType.DXA }},
                columnWidths: [2340, 2340, 4680],
                rows: [
                    new TableRow({{ tableHeader: true, children: [
                        _cell("Danh mục",  2340, true),
                        _cell("Từ",        2340, true),
                        _cell("Tóm tắt",   4680, true),
                    ]}}),
                    {urgent_rows_js}
                ]
            }}),
        ]
    }}]
}});

Packer.toBuffer(doc).then(buf => {{
    fs.writeFileSync("{str(path).replace(chr(92), '/')}", buf);
    console.log("OK");
}}).catch(err => {{ console.error(err); process.exit(1); }});
"""

    with tempfile.NamedTemporaryFile("w", suffix=".js", delete=False, encoding="utf-8") as f:
        f.write(js_script)
        f.flush()
        subprocess.run(["node", f.name], check=True)
        os.unlink(f.name)