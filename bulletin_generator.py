from __future__ import annotations

import html
import re
import sys
import urllib.error
import urllib.request
import zipfile
from collections import OrderedDict
from dataclasses import dataclass
from datetime import date
from html.parser import HTMLParser
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from openpyxl import load_workbook


DEFAULT_SECTIONS = [
    "SBP Related News / Press Release",
    "Domestic News",
    "Editorials / Opinion / Analysis",
    "Foreign Media Updates",
]

HEADER_ALIASES = {
    "include": ["include", "use", "selected"],
    "section": ["section", "category"],
    "story_group": ["storygroup", "storytopic", "story", "topic"],
    "headline": ["headline", "articleheadline", "newstitle", "title"],
    "source": ["source", "newspaper", "publication"],
    "url": ["url", "link", "articlelink"],
    "article_text": ["articletext", "manualtext", "content", "body"],
    "language": ["language", "lang"],
    "priority": ["priority"],
    "notes": ["notes", "printreference", "reference"],
}

BLUE = RGBColor(31, 111, 104)
DARK_BLUE = RGBColor(20, 61, 58)
NAVY = RGBColor(23, 50, 77)
GRAY = RGBColor(83, 93, 104)


@dataclass
class NewsItem:
    row_number: int
    section: str
    story_group: str
    headline: str
    source: str
    url: str
    article_text: str
    language: str
    priority: str
    notes: str


class ParagraphCollector(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.in_script = False
        self.in_style = False
        self.in_paragraph = False
        self.current: list[str] = []
        self.paragraphs: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        tag = tag.lower()
        if tag == "script":
            self.in_script = True
        elif tag == "style":
            self.in_style = True
        elif tag == "p":
            self.in_paragraph = True
            self.current = []

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag == "script":
            self.in_script = False
        elif tag == "style":
            self.in_style = False
        elif tag == "p" and self.in_paragraph:
            text = clean_text(" ".join(self.current))
            if len(text) >= 45:
                self.paragraphs.append(text)
            self.in_paragraph = False
            self.current = []

    def handle_data(self, data: str) -> None:
        if self.in_script or self.in_style:
            return
        if self.in_paragraph:
            self.current.append(data)


def clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value)
    text = html.unescape(text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def normalize_header(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def first_alias_match(header_map: dict[str, int], field: str) -> int | None:
    for alias in HEADER_ALIASES[field]:
        if alias in header_map:
            return header_map[alias]
    return None


def find_header_row(sheet) -> tuple[int, dict[str, int]]:
    for row in range(1, min(sheet.max_row, 12) + 1):
        normalized = {
            normalize_header(sheet.cell(row=row, column=col).value): col
            for col in range(1, sheet.max_column + 1)
            if clean_text(sheet.cell(row=row, column=col).value)
        }
        matches = {
            field: first_alias_match(normalized, field)
            for field in HEADER_ALIASES
        }
        required_found = [matches["section"], matches["headline"], matches["source"]]
        if all(required_found):
            return row, {key: value for key, value in matches.items() if value is not None}
    raise ValueError("Could not find the input header row. Please use the official empty template.")


def read_news_items(excel_path: Path) -> tuple[list[NewsItem], list[str]]:
    workbook = load_workbook(excel_path, data_only=True)
    sheet = workbook["News Input"] if "News Input" in workbook.sheetnames else workbook.active
    header_row, columns = find_header_row(sheet)
    warnings: list[str] = []
    items: list[NewsItem] = []

    def cell_text(row: int, field: str) -> str:
        col = columns.get(field)
        if not col:
            return ""
        return clean_text(sheet.cell(row=row, column=col).value)

    for row in range(header_row + 1, sheet.max_row + 1):
        raw_values = [clean_text(sheet.cell(row=row, column=col).value) for col in range(1, sheet.max_column + 1)]
        if not any(raw_values):
            continue

        include = cell_text(row, "include").lower()
        if include in {"no", "n", "false", "0", "skip", "skipped"}:
            continue

        section = cell_text(row, "section") or "Domestic News"
        headline = cell_text(row, "headline")
        source = cell_text(row, "source") or "Unspecified source"
        story_group = cell_text(row, "story_group") or headline
        url = cell_text(row, "url")
        article_text = cell_text(row, "article_text")
        language = cell_text(row, "language") or "English"
        priority = cell_text(row, "priority") or "Medium"
        notes = cell_text(row, "notes")

        if not headline and story_group:
            headline = story_group
        if not headline:
            warnings.append(f"Row {row}: skipped because Headline is empty.")
            continue

        if not article_text and url:
            article_text, extraction_note = extract_article_from_url(url)
            if extraction_note:
                warnings.append(f"Row {row}: {extraction_note}")
        elif not article_text:
            article_text = "[Article text was not provided. Add article text or a reachable URL before final circulation.]"
            warnings.append(f"Row {row}: no article text or URL was provided.")

        items.append(
            NewsItem(
                row_number=row,
                section=section,
                story_group=story_group or headline,
                headline=headline,
                source=source,
                url=url,
                article_text=article_text,
                language=language,
                priority=priority,
                notes=notes,
            )
        )

    if not items:
        raise ValueError("No included news rows were found. Add at least one row with Include = Yes.")
    return items, warnings


def extract_article_from_url(url: str) -> tuple[str, str]:
    try:
        request = urllib.request.Request(
            url,
            headers={
                "User-Agent": "Mozilla/5.0 DailyNewsBulletinMVP/1.0",
                "Accept": "text/html,application/xhtml+xml",
            },
        )
        with urllib.request.urlopen(request, timeout=12) as response:
            content_type = response.headers.get_content_charset() or "utf-8"
            raw = response.read(1_500_000)
        markup = raw.decode(content_type, errors="replace")
    except (urllib.error.URLError, TimeoutError, ValueError) as exc:
        return (
            "[The app could not extract this article automatically. Please paste the article text in the Excel template.]",
            f"URL extraction failed for {url}: {exc}",
        )

    collector = ParagraphCollector()
    collector.feed(markup)
    paragraphs = [p for p in collector.paragraphs if not is_boilerplate(p)]
    if paragraphs:
        return "\n\n".join(paragraphs[:24]), ""

    text = re.sub(r"(?is)<(script|style).*?</\1>", " ", markup)
    text = re.sub(r"(?s)<[^>]+>", " ", text)
    text = clean_text(text)
    if len(text) < 120:
        return (
            "[The app could not detect the article body. Please paste the article text in the Excel template.]",
            f"URL extraction returned too little text for {url}.",
        )
    return text[:6000], ""


def is_boilerplate(text: str) -> bool:
    lowered = text.lower()
    blocked = [
        "subscribe",
        "advertisement",
        "all rights reserved",
        "sign in",
        "accept cookies",
        "privacy policy",
        "related stories",
        "share this",
    ]
    return any(term in lowered for term in blocked)


def set_run_font(run, name="Calibri", size=None, color=None, bold=None, italic=None):
    run.font.name = name
    r_pr = run._element.get_or_add_rPr()
    r_pr.rFonts.set(qn("w:ascii"), name)
    r_pr.rFonts.set(qn("w:hAnsi"), name)
    r_pr.rFonts.set(qn("w:cs"), "Jameel Noori Nastaleeq")
    if size is not None:
        run.font.size = Pt(size)
    if color is not None:
        run.font.color.rgb = color
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def set_para_spacing(paragraph, before=0, after=6, line=1.1):
    paragraph.paragraph_format.space_before = Pt(before)
    paragraph.paragraph_format.space_after = Pt(after)
    paragraph.paragraph_format.line_spacing = line


def shade_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def set_cell_text(cell, text, bold=False, color=None, size=9.5, align=WD_ALIGN_PARAGRAPH.CENTER):
    cell.text = ""
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    set_para_spacing(paragraph, after=0, line=1.0)
    run = paragraph.add_run(text)
    set_run_font(run, size=size, color=color, bold=bold)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


class LinkBuilder:
    def __init__(self):
        self.bookmark_counter = 1

    def bookmark(self, paragraph, name: str) -> None:
        start = OxmlElement("w:bookmarkStart")
        start.set(qn("w:id"), str(self.bookmark_counter))
        start.set(qn("w:name"), name)
        end = OxmlElement("w:bookmarkEnd")
        end.set(qn("w:id"), str(self.bookmark_counter))
        paragraph._p.insert(0, start)
        paragraph._p.append(end)
        self.bookmark_counter += 1

    def internal_link(self, paragraph, text: str, anchor: str) -> None:
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("w:anchor"), anchor)
        run = OxmlElement("w:r")
        r_pr = OxmlElement("w:rPr")
        r_style = OxmlElement("w:rStyle")
        r_style.set(qn("w:val"), "Hyperlink")
        r_pr.append(r_style)
        run.append(r_pr)
        node = OxmlElement("w:t")
        node.text = text
        run.append(node)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)

    def external_link(self, paragraph, text: str, url: str) -> None:
        if not url:
            return
        r_id = paragraph.part.relate_to(
            url,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        hyperlink = OxmlElement("w:hyperlink")
        hyperlink.set(qn("r:id"), r_id)
        run = OxmlElement("w:r")
        r_pr = OxmlElement("w:rPr")
        r_style = OxmlElement("w:rStyle")
        r_style.set(qn("w:val"), "Hyperlink")
        r_pr.append(r_style)
        run.append(r_pr)
        node = OxmlElement("w:t")
        node.text = text
        run.append(node)
        hyperlink.append(run)
        paragraph._p.append(hyperlink)


def slug(text: str, fallback: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_")
    return (cleaned[:36] or fallback)


def apply_core_styles(doc: Document) -> None:
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11)
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.85)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    section.header_distance = Inches(0.42)
    section.footer_distance = Inches(0.42)

    normal = doc.styles["Normal"]
    normal.font.name = "Calibri"
    normal._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
    normal._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
    normal._element.rPr.rFonts.set(qn("w:cs"), "Jameel Noori Nastaleeq")
    normal.font.size = Pt(10.5)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.08

    style_tokens = [
        ("Heading 1", 15, BLUE, 14, 7),
        ("Heading 2", 12.5, BLUE, 10, 5),
        ("Heading 3", 11.5, NAVY, 7, 3),
    ]
    for style_name, size, color, before, after in style_tokens:
        style = doc.styles[style_name]
        style.font.name = "Calibri"
        style._element.rPr.rFonts.set(qn("w:ascii"), "Calibri")
        style._element.rPr.rFonts.set(qn("w:hAnsi"), "Calibri")
        style._element.rPr.rFonts.set(qn("w:cs"), "Jameel Noori Nastaleeq")
        style.font.size = Pt(size)
        style.font.color.rgb = color
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)


def set_header_footer(doc: Document) -> None:
    section = doc.sections[0]
    header = section.header.paragraphs[0]
    header.text = ""
    run = header.add_run("Daily News Bulletin")
    set_run_font(run, size=8.5, color=GRAY, bold=True)
    run = header.add_run("    External Communications Department")
    set_run_font(run, size=8.5, color=GRAY)

    footer = section.footer.paragraphs[0]
    footer.text = ""
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Generated by Daily Bulletin System MVP")
    set_run_font(run, size=8, color=GRAY)


def add_masthead(doc: Document, report_date: str, prepared_for: str, links: LinkBuilder, sections: list[str]) -> None:
    title = doc.add_paragraph()
    set_para_spacing(title, before=2, after=1)
    run = title.add_run("Daily News Bulletin")
    set_run_font(run, size=20, bold=True, color=DARK_BLUE)
    links.bookmark(title, "top")

    subtitle = doc.add_paragraph()
    set_para_spacing(subtitle, after=10)
    run = subtitle.add_run("External Communications Department")
    set_run_font(run, size=11, color=GRAY, bold=True)

    meta = doc.add_paragraph()
    set_para_spacing(meta, after=12)
    for label, value in [
        ("Date: ", report_date),
        ("Prepared for: ", prepared_for),
    ]:
        run = meta.add_run(label)
        set_run_font(run, size=10.5, bold=True, color=DARK_BLUE)
        run = meta.add_run(value + "    ")
        set_run_font(run, size=10.5, color=RGBColor(30, 41, 59))

    table = doc.add_table(rows=1, cols=len(sections))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    for index, section_name in enumerate(sections):
        cell = table.rows[0].cells[index]
        shade_cell(cell, "E8F2F0")
        set_cell_text(cell, section_name, bold=True, color=DARK_BLUE, size=8.8)
        link_para = cell.paragraphs[0]
        links.bookmark(link_para, f"nav_{index + 1}")


def section_order(items: Iterable[NewsItem]) -> list[str]:
    seen = list(DEFAULT_SECTIONS)
    for item in items:
        if item.section not in seen:
            seen.append(item.section)
    return seen


def group_items(items: list[NewsItem]) -> OrderedDict[str, OrderedDict[str, list[NewsItem]]]:
    grouped: OrderedDict[str, OrderedDict[str, list[NewsItem]]] = OrderedDict()
    for section in section_order(items):
        grouped[section] = OrderedDict()
    for item in items:
        grouped.setdefault(item.section, OrderedDict())
        grouped[item.section].setdefault(item.story_group, [])
        grouped[item.section][item.story_group].append(item)
    return grouped


def add_headline_index(doc: Document, grouped, links: LinkBuilder) -> dict[tuple[str, str], str]:
    anchors: dict[tuple[str, str], str] = {}
    paragraph = doc.add_paragraph(style="Heading 1")
    paragraph.add_run("Headlines")
    links.bookmark(paragraph, "headlines")

    group_count = 1
    for section_name, stories in grouped.items():
        if not stories:
            continue
        heading = doc.add_paragraph(style="Heading 2")
        heading.add_run(section_name)
        links.bookmark(heading, f"section_{slug(section_name, 'section')}")

        for story_group, items in stories.items():
            anchor = f"article_{group_count}_{slug(story_group, 'story')}"
            anchors[(section_name, story_group)] = anchor
            group_count += 1

            p = doc.add_paragraph()
            set_para_spacing(p, after=2, line=1.05)
            p.paragraph_format.left_indent = Inches(0.16)
            links.internal_link(p, story_group, anchor)

            source_text = " / ".join(item.source for item in items)
            p = doc.add_paragraph()
            set_para_spacing(p, after=7, line=1.0)
            p.paragraph_format.left_indent = Inches(0.34)
            run = p.add_run(source_text)
            set_run_font(run, size=9.3, color=GRAY)

    return anchors


def add_separator(doc: Document) -> None:
    paragraph = doc.add_paragraph()
    set_para_spacing(paragraph, after=8)
    p_pr = paragraph._p.get_or_add_pPr()
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "4")
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), "D7DBE2")
    p_bdr.append(bottom)
    p_pr.append(p_bdr)


def add_article_body(doc: Document, grouped, anchors, links: LinkBuilder) -> None:
    doc.add_page_break()
    for section_name, stories in grouped.items():
        if not stories:
            continue
        section_heading = doc.add_paragraph(style="Heading 1")
        section_heading.add_run(section_name)

        for story_group, items in stories.items():
            story_heading = doc.add_paragraph(style="Heading 2")
            story_heading.add_run(story_group)
            links.bookmark(story_heading, anchors[(section_name, story_group)])

            for item in items:
                article_heading = doc.add_paragraph(style="Heading 3")
                article_heading.add_run(item.headline)
                if item.language.lower() == "urdu":
                    article_heading.alignment = WD_ALIGN_PARAGRAPH.RIGHT

                source = doc.add_paragraph()
                set_para_spacing(source, after=5)
                run = source.add_run(item.source)
                set_run_font(run, size=10, bold=True, color=DARK_BLUE)
                if item.priority:
                    run = source.add_run(f" | Priority: {item.priority}")
                    set_run_font(run, size=9.2, color=GRAY)
                if item.url:
                    run = source.add_run(" | ")
                    set_run_font(run, size=9.2, color=GRAY)
                    links.external_link(source, "Original URL", item.url)

                if item.notes:
                    notes = doc.add_paragraph()
                    set_para_spacing(notes, after=5)
                    run = notes.add_run(f"Note: {item.notes}")
                    set_run_font(run, size=9.2, italic=True, color=GRAY)

                for body_paragraph in split_body(item.article_text):
                    p = doc.add_paragraph()
                    set_para_spacing(p, after=5, line=1.08)
                    if item.language.lower() == "urdu":
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    run = p.add_run(body_paragraph)
                    set_run_font(run, size=10.3, color=RGBColor(17, 24, 39))

                back = doc.add_paragraph()
                set_para_spacing(back, before=1, after=9)
                links.internal_link(back, "Back to Headlines", "headlines")

            add_separator(doc)


def split_body(text: str) -> list[str]:
    parts = [clean_text(part) for part in re.split(r"\n\s*\n", text or "")]
    return [part for part in parts if part] or ["[No article body available.]"]


def scrub_metadata(docx_path: Path) -> None:
    tmp_path = docx_path.with_suffix(".tmp.docx")
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                text = data.decode("utf-8")
                text = re.sub(r"<dc:creator>.*?</dc:creator>", "<dc:creator>Daily Bulletin System</dc:creator>", text)
                text = re.sub(
                    r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>",
                    "<cp:lastModifiedBy>Daily Bulletin System</cp:lastModifiedBy>",
                    text,
                )
                data = text.encode("utf-8")
            if item.filename.endswith(".xml"):
                text = data.decode("utf-8", errors="ignore")
                text = re.sub(r'\s+w:rsid\w+="[^"]+"', "", text)
                data = text.encode("utf-8")
            zout.writestr(item, data)
    tmp_path.replace(docx_path)


def resolve_report_date(report_date: str | None) -> str:
    return report_date or date.today().strftime("%A, %B %d, %Y")


def filename_date() -> str:
    return date.today().strftime("%Y%m%d")


def build_metadata(items: list[NewsItem], grouped, warnings: list[str]) -> dict:
    return {
        "article_count": len(items),
        "story_count": sum(len(stories) for stories in grouped.values()),
        "section_count": sum(1 for stories in grouped.values() if stories),
        "warnings": warnings,
    }


def write_bulletin_docx(
    items: list[NewsItem],
    grouped,
    output_dir: str | Path,
    report_date: str,
    prepared_for: str,
) -> Path:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = output_dir / f"Daily_News_Bulletin_{filename_date()}.docx"

    doc = Document()
    links = LinkBuilder()
    apply_core_styles(doc)
    set_header_footer(doc)
    add_masthead(doc, report_date, prepared_for, links, section_order(items))
    anchors = add_headline_index(doc, grouped, links)
    add_article_body(doc, grouped, anchors, links)
    doc.save(output_path)
    scrub_metadata(output_path)
    return output_path


def render_preview_html(items: list[NewsItem], grouped, report_date: str, prepared_for: str) -> str:
    def esc(value: str) -> str:
        return html.escape(value or "", quote=True)

    active_sections = [(section, stories) for section, stories in grouped.items() if stories]
    anchors: dict[tuple[str, str], str] = {}
    story_number = 1
    for section_name, stories in active_sections:
        for story_group in stories:
            anchors[(section_name, story_group)] = f"article-{story_number}-{slug(story_group, 'story').lower()}"
            story_number += 1

    parts = [
        "<!doctype html>",
        '<html lang="en">',
        "<head>",
        '<meta charset="utf-8" />',
        '<meta name="viewport" content="width=device-width, initial-scale=1" />',
        "<title>Daily News Bulletin Preview</title>",
        "<style>",
        """
        :root {
          --ink: #111827;
          --muted: #5f6c70;
          --line: #d5dddf;
          --green: #1f6f68;
          --green-dark: #143d3a;
          --gold: #b98416;
          --paper: #ffffff;
          --page: #f3f6f5;
        }
        * { box-sizing: border-box; }
        body {
          margin: 0;
          background: var(--page);
          color: var(--ink);
          font-family: "Segoe UI", Arial, sans-serif;
          line-height: 1.58;
        }
        main {
          width: min(980px, calc(100vw - 28px));
          margin: 24px auto;
          background: var(--paper);
          border: 1px solid var(--line);
          border-radius: 8px;
          padding: 34px;
        }
        header {
          border-bottom: 4px solid var(--green);
          padding-bottom: 18px;
          margin-bottom: 24px;
        }
        h1, h2, h3, h4, p { letter-spacing: 0; }
        h1 {
          margin: 0;
          color: var(--green-dark);
          font-size: 34px;
          line-height: 1.1;
        }
        h2 {
          margin: 30px 0 12px;
          color: var(--green);
          font-size: 23px;
          border-bottom: 1px solid var(--line);
          padding-bottom: 7px;
        }
        h3 {
          margin: 22px 0 6px;
          color: var(--green-dark);
          font-size: 20px;
        }
        h4 {
          margin: 15px 0 5px;
          color: #17324d;
          font-size: 16px;
        }
        a { color: var(--green); font-weight: 700; }
        .subtitle, .meta, .source, .notes, .sources {
          color: var(--muted);
        }
        .subtitle {
          margin: 6px 0 12px;
          font-weight: 700;
        }
        .meta {
          display: flex;
          flex-wrap: wrap;
          gap: 16px;
          margin: 0;
          font-size: 14px;
        }
        .nav {
          display: flex;
          flex-wrap: wrap;
          gap: 8px;
          margin: 18px 0 0;
          padding: 0;
          list-style: none;
        }
        .nav a {
          display: inline-flex;
          border: 1px solid var(--line);
          border-radius: 999px;
          padding: 6px 10px;
          text-decoration: none;
          background: #f8fbfa;
          font-size: 13px;
        }
        .headlines {
          margin: 0 0 24px;
          padding-left: 22px;
        }
        .headlines li { margin: 11px 0; }
        .sources {
          display: block;
          margin-top: 2px;
          font-size: 13px;
          font-weight: 600;
        }
        .article {
          border-bottom: 1px solid var(--line);
          padding-bottom: 18px;
          margin-bottom: 18px;
        }
        .source, .notes {
          margin: 0 0 9px;
          font-size: 13px;
        }
        .notes { font-style: italic; }
        .body p { margin: 0 0 11px; }
        .urdu {
          direction: rtl;
          text-align: right;
          font-family: "Jameel Noori Nastaleeq", "Noto Nastaliq Urdu", serif;
        }
        .back {
          display: inline-flex;
          margin-top: 8px;
          font-size: 13px;
        }
        @media print {
          body { background: white; }
          main { width: auto; margin: 0; border: 0; padding: 0; }
          .nav, .back { display: none; }
        }
        @media (max-width: 640px) {
          main { padding: 22px; }
          h1 { font-size: 28px; }
        }
        """,
        "</style>",
        "</head>",
        "<body>",
        "<main>",
        '<header id="top">',
        "<h1>Daily News Bulletin</h1>",
        '<p class="subtitle">External Communications Department</p>',
        f'<p class="meta"><span><strong>Date:</strong> {esc(report_date)}</span><span><strong>Prepared for:</strong> {esc(prepared_for)}</span></p>',
        '<ul class="nav">',
    ]

    for section_name, _ in active_sections:
        parts.append(f'<li><a href="#section-{esc(slug(section_name, "section").lower())}">{esc(section_name)}</a></li>')
    parts.extend(["</ul>", "</header>", '<section id="headlines">', "<h2>Headlines</h2>"])

    for section_name, stories in active_sections:
        parts.append(f'<h3 id="section-{esc(slug(section_name, "section").lower())}">{esc(section_name)}</h3>')
        parts.append('<ol class="headlines">')
        for story_group, story_items in stories.items():
            source_text = " / ".join(item.source for item in story_items)
            anchor = anchors[(section_name, story_group)]
            parts.append(
                f'<li><a href="#{esc(anchor)}">{esc(story_group)}</a><span class="sources">{esc(source_text)}</span></li>'
            )
        parts.append("</ol>")

    parts.append("</section>")
    parts.append('<section id="articles">')
    for section_name, stories in active_sections:
        parts.append(f"<h2>{esc(section_name)}</h2>")
        for story_group, story_items in stories.items():
            anchor = anchors[(section_name, story_group)]
            parts.append(f'<article id="{esc(anchor)}" class="article">')
            parts.append(f"<h3>{esc(story_group)}</h3>")
            for item in story_items:
                article_class = ' class="urdu"' if item.language.lower() == "urdu" else ""
                parts.append(f"<h4{article_class}>{esc(item.headline)}</h4>")
                source_line = esc(item.source)
                if item.priority:
                    source_line += f" | Priority: {esc(item.priority)}"
                if item.url:
                    source_line += f' | <a href="{esc(item.url)}" target="_blank" rel="noopener">Original URL</a>'
                parts.append(f'<p class="source">{source_line}</p>')
                if item.notes:
                    parts.append(f'<p class="notes">Note: {esc(item.notes)}</p>')
                body_class = ' class="body urdu"' if item.language.lower() == "urdu" else ' class="body"'
                parts.append(f"<div{body_class}>")
                for paragraph in split_body(item.article_text):
                    parts.append(f"<p>{esc(paragraph)}</p>")
                parts.append("</div>")
            parts.append('<a class="back" href="#headlines">Back to Headlines</a>')
            parts.append("</article>")

    parts.extend(["</section>", "</main>", "</body>", "</html>"])
    return "\n".join(parts)


def write_preview_html(
    items: list[NewsItem],
    grouped,
    output_dir: str | Path,
    report_date: str,
    prepared_for: str,
) -> Path:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / f"Daily_News_Bulletin_{filename_date()}_Preview.html"
    output_path.write_text(render_preview_html(items, grouped, report_date, prepared_for), encoding="utf-8")
    return output_path


def generate_bulletin(
    excel_path: str | Path,
    output_dir: str | Path,
    report_date: str | None = None,
    prepared_for: str = "SBP Management Forum",
) -> tuple[Path, dict]:
    excel_path = Path(excel_path)
    items, warnings = read_news_items(excel_path)
    grouped = group_items(items)
    resolved_date = resolve_report_date(report_date)
    output_path = write_bulletin_docx(items, grouped, output_dir, resolved_date, prepared_for)
    return output_path, build_metadata(items, grouped, warnings)


def generate_preview_html(
    excel_path: str | Path,
    output_dir: str | Path,
    report_date: str | None = None,
    prepared_for: str = "SBP Management Forum",
) -> tuple[Path, dict]:
    excel_path = Path(excel_path)
    items, warnings = read_news_items(excel_path)
    grouped = group_items(items)
    resolved_date = resolve_report_date(report_date)
    output_path = write_preview_html(items, grouped, output_dir, resolved_date, prepared_for)
    return output_path, build_metadata(items, grouped, warnings)


def generate_bulletin_artifacts(
    excel_path: str | Path,
    report_dir: str | Path,
    preview_dir: str | Path,
    report_date: str | None = None,
    prepared_for: str = "SBP Management Forum",
) -> tuple[Path, Path, dict]:
    excel_path = Path(excel_path)
    items, warnings = read_news_items(excel_path)
    grouped = group_items(items)
    resolved_date = resolve_report_date(report_date)
    report_path = write_bulletin_docx(items, grouped, report_dir, resolved_date, prepared_for)
    preview_path = write_preview_html(items, grouped, preview_dir, resolved_date, prepared_for)
    return report_path, preview_path, build_metadata(items, grouped, warnings)


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print("Usage: python bulletin_generator.py <input.xlsx> <output_dir> [report_date] [prepared_for]")
        return 2
    report_date = argv[3] if len(argv) >= 4 else None
    prepared_for = argv[4] if len(argv) >= 5 else "SBP Management Forum"
    output, metadata = generate_bulletin(argv[1], argv[2], report_date, prepared_for)
    print(output)
    print(metadata)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
