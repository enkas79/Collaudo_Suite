from __future__ import annotations

import ast
from html.parser import HTMLParser
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
)
from reportlab.lib.utils import ImageReader
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "collaudo_suite" / "help_dialog.py"
OUTPUT = ROOT / "output" / "pdf" / "Guida_operativa_Collaudo_Suite.pdf"
CHAPTERS = [
    ("Processo completo", "_workflow_html"),
    ("Analyzer anomalie", "_analyzer_html"),
    ("Creazione dei controlli", "_controls_html"),
    ("Checklist e MAP/JARVIS", "_checklist_html"),
    ("Salvataggio, PDF e risoluzione problemi", "_save_output_html"),
    ("Metodo, interpretazione e limiti", "_method_html"),
]


def source_sections() -> list[tuple[str, str]]:
    tree = ast.parse(SOURCE.read_text(encoding="utf-8"))
    cls = next(n for n in tree.body if isinstance(n, ast.ClassDef) and n.name == "SuiteHelpDialog")
    methods = {n.name: n for n in cls.body if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef))}
    result = []
    for title, method_name in CHAPTERS:
        method = methods[method_name]
        raw = next(
            n.value
            for n in ast.walk(method)
            if isinstance(n, ast.Call)
            and isinstance(n.func, ast.Attribute)
            and n.func.attr == "_base"
            for n in n.args[:1]
            if isinstance(n, ast.Constant) and isinstance(n.value, str)
        )
        result.append((title, raw))
    return result


class SectionParser(HTMLParser):
    def __init__(self, styles, assets_dir: Path):
        super().__init__(convert_charrefs=True)
        self.styles = styles
        self.assets_dir = assets_dir
        self.story = []
        self.capture = None
        self.table_rows = None
        self.table_row = None
        self.table_cell = None
        self.list_stack = []
        self.list_counts = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "img" and self.capture is None:
            source = str(attrs.get("src", "")).strip()
            image_path = self.assets_dir / Path(source).name
            if image_path.exists():
                reader = ImageReader(str(image_path))
                width, height = reader.getSize()
                max_width = 170 * mm
                max_height = 105 * mm
                scale = min(max_width / width, max_height / height, 1.0)
                self.story.append(Image(str(image_path), width=width * scale, height=height * scale, hAlign="CENTER"))
                self.story.append(Spacer(1, 6))
            return
        if tag in {"h1", "h2", "h3", "p", "li"} and self.capture is None:
            self.capture = [tag, [], attrs.get("class", "")]
            if tag == "li":
                if self.list_counts:
                    self.list_counts[-1] += 1
                number = self.list_counts[-1] if self.list_counts else 1
                self.capture.append(str(number))
        elif tag == "div" and self.capture is None and attrs.get("class") in {"flow", "warn", "ok"}:
            self.capture = ["div", [], attrs["class"]]
        elif tag == "table":
            self.table_rows = []
        elif tag == "tr" and self.table_rows is not None:
            self.table_row = []
        elif tag in {"td", "th"} and self.table_row is not None:
            self.table_cell = [tag, []]
        elif tag in {"ol", "ul"}:
            self.list_stack.append(tag)
            self.list_counts.append(0)
        elif tag == "br":
            self._append("<br/>")
        elif tag in {"b", "strong", "i", "em", "code"}:
            markup = {"b": "b", "strong": "b", "i": "i", "em": "i", "code": "font"}[tag]
            opening = '<font name="Courier" backColor="#edf1f4">' if tag == "code" else f"<{markup}>"
            self._append(opening)

    def handle_endtag(self, tag):
        if tag in {"b", "strong", "i", "em", "code"}:
            self._append("</font>" if tag == "code" else f"</{'b' if tag in {'b', 'strong'} else 'i'}>")
        if tag in {"h1", "h2", "h3", "p", "li", "div"} and self.capture and self.capture[0] == tag:
            self._finish_capture()
        elif tag in {"td", "th"} and self.table_cell is not None:
            cell_tag, parts = self.table_cell
            self.table_row.append((cell_tag, "".join(parts)))
            self.table_cell = None
        elif tag == "tr" and self.table_row is not None:
            self.table_rows.append(self.table_row)
            self.table_row = None
        elif tag == "table" and self.table_rows is not None:
            self._finish_table()
        elif tag in {"ol", "ul"} and self.list_stack:
            self.list_stack.pop()
            self.list_counts.pop()

    def handle_data(self, data):
        data = (data.replace("→", "->").replace("←", "<-")
                .replace("–", "-").replace("—", "-").replace("’", "'"))
        if self.capture is not None:
            self.capture[1].append(escape(data))
        elif self.table_cell is not None:
            self.table_cell[1].append(escape(data))

    def _append(self, value):
        if self.capture is not None:
            self.capture[1].append(value)
        elif self.table_cell is not None:
            self.table_cell[1].append(value)

    def _finish_capture(self):
        tag, parts, css_class, *extra = self.capture
        markup = "".join(parts).strip()
        self.capture = None
        if not markup:
            return
        style_name = {"h1": "h1", "h2": "h2", "h3": "h3", "p": "body", "li": "body", "div": "callout"}.get(tag, "body")
        paragraph = Paragraph(markup, self.styles[style_name])
        if tag == "li":
            bullet = "-" if self.list_stack and self.list_stack[-1] == "ul" else f"{extra[0] if extra else '1'}."
            self.story.append(Paragraph(f"{bullet} &nbsp; {markup}", self.styles["body"]))
        elif tag == "div":
            self.story.append(KeepTogether([paragraph, Spacer(1, 5)]))
        else:
            self.story.append(paragraph)
            if not tag.startswith("h"):
                self.story.append(Spacer(1, 3))

    def _finish_table(self):
        rows = []
        for row in self.table_rows:
            rows.append([Paragraph(text or " ", self.styles["th" if tag == "th" else "td"]) for tag, text in row])
        self.table_rows = None
        if not rows:
            return
        columns = max(map(len, rows))
        width = 170 * mm / columns
        table = Table(rows, colWidths=[width] * columns, repeatRows=1, hAlign="LEFT")
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eaf2f7")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#19364a")),
            ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#cbd6de")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]))
        heading = None
        if self.story and isinstance(self.story[-1], Paragraph) and self.story[-1].style.name in {"GuideH2", "GuideH3"}:
            heading = self.story.pop()
        if heading is not None:
            self.story.append(KeepTogether([heading, table]))
        else:
            self.story.append(table)
        self.story.append(Spacer(1, 9))


def main():
    regular = Path("C:/Windows/Fonts/arial.ttf")
    bold = Path("C:/Windows/Fonts/arialbd.ttf")
    pdfmetrics.registerFont(TTFont("GuideArial", str(regular)))
    pdfmetrics.registerFont(TTFont("GuideArial-Bold", str(bold)))
    sample = getSampleStyleSheet()
    styles = {
        "h1": ParagraphStyle("GuideH1", parent=sample["Heading1"], fontName="GuideArial-Bold", fontSize=19, leading=23, textColor=colors.HexColor("#19364a"), spaceAfter=8),
        "h2": ParagraphStyle("GuideH2", parent=sample["Heading2"], fontName="GuideArial-Bold", fontSize=13, leading=16, textColor=colors.HexColor("#2878a8"), spaceBefore=10, spaceAfter=5, keepWithNext=True),
        "h3": ParagraphStyle("GuideH3", parent=sample["Heading3"], fontName="GuideArial-Bold", fontSize=10.5, leading=13, textColor=colors.HexColor("#334e5c"), spaceBefore=7, spaceAfter=3, keepWithNext=True),
        "body": ParagraphStyle("GuideBody", parent=sample["BodyText"], fontName="GuideArial", fontSize=9, leading=12.2, textColor=colors.HexColor("#243442"), spaceAfter=3),
        "callout": ParagraphStyle("GuideCallout", parent=sample["BodyText"], fontName="GuideArial", fontSize=9, leading=12.2, textColor=colors.HexColor("#243442"), backColor=colors.HexColor("#eef6fb"), borderColor=colors.HexColor("#b9d9ec"), borderWidth=0.6, borderPadding=7, spaceBefore=5, spaceAfter=8),
        "th": ParagraphStyle("GuideTH", fontName="GuideArial-Bold", fontSize=8, leading=10, textColor=colors.HexColor("#19364a")),
        "td": ParagraphStyle("GuideTD", fontName="GuideArial", fontSize=7.7, leading=10, textColor=colors.HexColor("#243442")),
    }
    sections = source_sections()
    story = [Spacer(1, 28 * mm), Paragraph("COLLAUDO SUITE", ParagraphStyle("Cover", fontName="GuideArial-Bold", fontSize=27, leading=32, alignment=TA_CENTER, textColor=colors.HexColor("#19364a"))), Spacer(1, 5 * mm), Paragraph("Guida operativa completa", ParagraphStyle("Sub", fontName="GuideArial", fontSize=16, leading=20, alignment=TA_CENTER, textColor=colors.HexColor("#2878a8"))), Spacer(1, 9 * mm), Paragraph("Manuale unico per Analyzer anomalie, preparazione dei controlli, Checklist, MAP/JARVIS, salvataggio e PDF.", ParagraphStyle("Intro", fontName="GuideArial", fontSize=11, leading=16, alignment=TA_CENTER, textColor=colors.HexColor("#334e5c"))), Spacer(1, 18 * mm), Paragraph("Indice", styles["h1"])]
    for i, (title, _) in enumerate(sections, 1):
        story.append(Paragraph(f"{i}. &nbsp; {escape(title)}", styles["body"]))
    story.append(Spacer(1, 12 * mm))
    story.append(Paragraph("Versione della guida: 7 ottobre 2026", ParagraphStyle("Date", fontName="GuideArial", fontSize=9, alignment=TA_CENTER, textColor=colors.HexColor("#667784"))))
    story.append(PageBreak())
    for index, (title, html) in enumerate(sections):
        parser = SectionParser(styles, ROOT / "collaudo_suite" / "assets")
        parser.feed(html.replace("&rarr;", "-&gt;"))
        if index:
            story.append(PageBreak())
        story.extend(parser.story)
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)

    def page(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#d7e0e6"))
        canvas.line(20 * mm, 15 * mm, 190 * mm, 15 * mm)
        canvas.setFont("GuideArial", 8)
        canvas.setFillColor(colors.HexColor("#667784"))
        canvas.drawString(20 * mm, 10 * mm, "Collaudo Suite - Guida operativa")
        canvas.drawRightString(190 * mm, 10 * mm, f"Pagina {doc.page}")
        canvas.restoreState()

    doc = BaseDocTemplate(str(OUTPUT), pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm, topMargin=18 * mm, bottomMargin=21 * mm, title="Guida operativa Collaudo Suite", author="Collaudo Suite")
    doc.addPageTemplates(PageTemplate(id="guide", frames=[Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")], onPage=page))
    doc.build(story)
    bundled_copy = ROOT / "collaudo_suite" / "assets" / OUTPUT.name
    bundled_copy.parent.mkdir(parents=True, exist_ok=True)
    bundled_copy.write_bytes(OUTPUT.read_bytes())
    print(OUTPUT)


if __name__ == "__main__":
    main()
