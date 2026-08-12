"""インストールマニュアルのMarkdownからPDFを生成するスクリプト。

本文は `docs/install_manual.md` にテキストで置き、配布用のPDFはここで組版する。
PDFはバイナリのためリポジトリへは含めず、必要になったときに生成する。

    python -m pip install reportlab
    python tools/make_manual.py

日本語が豆腐（□）にならないよう、PCにあるゴシック体のTrueTypeフォントを
探して埋め込む。Windowsではメイリオ、游ゴシック、MS ゴシックの順に探し、
見つからない場合はIPAゴシックやNoto Sans CJKも候補にする。

Markdownは本文で使う記法だけを解釈する。

    # 見出し1 / ## 見出し2 / ### 見出し3
    - 箇条書き
    1. 番号付き箇条書き
    ```
    コードブロック
    ```
    **強調** と `コード`
    ---            水平線は改ページとして扱う
    <!-- toc -->   目次を差し込む位置
"""

from __future__ import annotations

import html
import re
import sys
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    PageBreak,
    PageTemplate,
    Paragraph,
    Preformatted,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SOURCE_PATH = PROJECT_ROOT / "docs" / "install_manual.md"
OUTPUT_PATH = PROJECT_ROOT / "docs" / "インストールマニュアル.pdf"

BODY_FONT = "ManualGothic"
BOLD_FONT = "ManualGothic-Bold"

# 埋め込むフォントの候補。(標準の書体, 太字の書体) の順に並べ、先に見つかった
# 組み合わせを使う。TrueTypeコレクション(.ttc)は書体の番号も指定する。
FONT_CANDIDATES: tuple[tuple[tuple[str, int], tuple[str, int] | None], ...] = (
    ((r"C:\Windows\Fonts\meiryo.ttc", 0), (r"C:\Windows\Fonts\meiryob.ttc", 0)),
    ((r"C:\Windows\Fonts\YuGothM.ttc", 0), (r"C:\Windows\Fonts\YuGothB.ttc", 0)),
    ((r"C:\Windows\Fonts\yugothm.ttc", 0), (r"C:\Windows\Fonts\yugothb.ttc", 0)),
    ((r"C:\Windows\Fonts\msgothic.ttc", 0), None),
    (("/usr/share/fonts/opentype/ipafont-gothic/ipagp.ttf", 0), None),
    (("/usr/share/fonts/truetype/fonts-japanese-gothic.ttf", 0), None),
    (
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc", 0),
        ("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc", 0),
    ),
    (
        ("/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc", 0),
        ("/System/Library/Fonts/ヒラギノ角ゴシック W6.ttc", 0),
    ),
)

PAGE_MARGIN = 20 * mm
CONTENT_WIDTH = A4[0] - 2 * PAGE_MARGIN
INK = colors.HexColor("#243b53")
ACCENT = colors.HexColor("#35679b")
MUTED = colors.HexColor("#627d98")
CODE_BACKGROUND = colors.HexColor("#f0f4f8")
CODE_BORDER = colors.HexColor("#d3dce6")


def _register_fonts() -> None:
    """日本語を表示できるフォントを探して登録する。"""

    for regular, bold in FONT_CANDIDATES:
        regular_path = Path(regular[0])
        if not regular_path.is_file():
            continue
        try:
            pdfmetrics.registerFont(
                TTFont(BODY_FONT, str(regular_path), subfontIndex=regular[1])
            )
            bold_path = Path(bold[0]) if bold is not None else None
            if bold_path is not None and bold_path.is_file():
                pdfmetrics.registerFont(
                    TTFont(BOLD_FONT, str(bold_path), subfontIndex=bold[1])
                )
            else:
                # 太字の書体が無い場合は同じ書体を割り当て、文字化けを避ける。
                pdfmetrics.registerFont(
                    TTFont(BOLD_FONT, str(regular_path), subfontIndex=regular[1])
                )
        except Exception:
            continue

        pdfmetrics.registerFontFamily(
            BODY_FONT,
            normal=BODY_FONT,
            bold=BOLD_FONT,
            italic=BODY_FONT,
            boldItalic=BOLD_FONT,
        )
        return

    searched = "\n".join(f"  {regular[0]}" for regular, _ in FONT_CANDIDATES)
    raise SystemExit(
        "日本語のフォントが見つかりませんでした。次の場所を探しました。\n"
        f"{searched}\n"
        "別のフォントを使う場合は、tools/make_manual.py の FONT_CANDIDATES へ"
        "パスを追加してください。"
    )


def _build_styles() -> dict[str, ParagraphStyle]:
    """日本語の折り返し（wordWrap="CJK"）を有効にした書式をまとめて作る。"""

    base = ParagraphStyle(
        "Body",
        fontName=BODY_FONT,
        fontSize=10,
        leading=16.5,
        textColor=INK,
        alignment=TA_LEFT,
        wordWrap="CJK",
        spaceAfter=7,
    )
    return {
        "Body": base,
        "Title": ParagraphStyle(
            "Title",
            parent=base,
            fontName=BOLD_FONT,
            fontSize=20,
            leading=28,
            textColor=ACCENT,
            spaceAfter=14,
        ),
        "Heading2": ParagraphStyle(
            "Heading2",
            parent=base,
            fontName=BOLD_FONT,
            fontSize=14.5,
            leading=22,
            textColor=ACCENT,
            spaceBefore=10,
            spaceAfter=9,
            borderPadding=(0, 0, 4, 0),
        ),
        "Heading3": ParagraphStyle(
            "Heading3",
            parent=base,
            fontName=BOLD_FONT,
            fontSize=11.5,
            leading=18,
            spaceBefore=11,
            spaceAfter=5,
        ),
        "Heading4": ParagraphStyle(
            "Heading4",
            parent=base,
            fontName=BOLD_FONT,
            fontSize=10,
            leading=16,
            textColor=ACCENT,
            spaceBefore=9,
            spaceAfter=2,
        ),
        # 箇条書きの記号は本文と同じフォントで描く。既定のHelveticaには
        # 「・」が無く、豆腐（□）になってしまうため。
        "Bullet": ParagraphStyle(
            "Bullet",
            parent=base,
            bulletFontName=BODY_FONT,
            bulletFontSize=10,
            leftIndent=13,
            bulletIndent=2,
            spaceAfter=3,
        ),
        "Ordered": ParagraphStyle(
            "Ordered",
            parent=base,
            bulletFontName=BODY_FONT,
            bulletFontSize=10,
            leftIndent=17,
            bulletIndent=2,
            spaceAfter=3,
        ),
        "Code": ParagraphStyle(
            "Code",
            parent=base,
            fontSize=9,
            leading=14,
            spaceAfter=0,
        ),
        "TOCTitle": ParagraphStyle(
            "TOCTitle",
            parent=base,
            fontName=BOLD_FONT,
            fontSize=14.5,
            leading=22,
            textColor=ACCENT,
            spaceBefore=10,
            spaceAfter=9,
        ),
        "TOC0": ParagraphStyle(
            "TOC0",
            parent=base,
            fontName=BOLD_FONT,
            leftIndent=0,
            spaceBefore=5,
            spaceAfter=0,
        ),
        "TOC1": ParagraphStyle(
            "TOC1",
            parent=base,
            fontSize=9.5,
            leftIndent=14,
            textColor=MUTED,
            spaceAfter=0,
        ),
    }


def _to_markup(text: str) -> str:
    """行内の記法をreportlabのタグへ変換する。"""

    escaped = html.escape(text, quote=False)
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", escaped)
    # 行内コードは等幅にすると日本語が崩れるため、背景色だけで区別する。
    return re.sub(
        r"`(.+?)`",
        r'<font backColor="#eef2f7">\1</font>',
        escaped,
    )


def _parse_blocks(markdown: str) -> list[tuple[str, object]]:
    """本文で使う記法だけを解釈し、ブロックの並びへ変換する。"""

    blocks: list[tuple[str, object]] = []
    lines = markdown.splitlines()
    index = 0
    while index < len(lines):
        line = lines[index]
        stripped = line.strip()

        if not stripped:
            index += 1
            continue

        if stripped.startswith("```"):
            index += 1
            code_lines: list[str] = []
            while index < len(lines) and not lines[index].strip().startswith("```"):
                code_lines.append(lines[index].rstrip())
                index += 1
            index += 1
            blocks.append(("code", code_lines))
            continue

        if stripped == "<!-- toc -->":
            blocks.append(("toc", None))
            index += 1
            continue

        if set(stripped) == {"-"} and len(stripped) >= 3:
            blocks.append(("pagebreak", None))
            index += 1
            continue

        heading = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if heading is not None:
            blocks.append((f"h{len(heading.group(1))}", heading.group(2).strip()))
            index += 1
            continue

        if stripped.startswith("- "):
            items: list[str] = []
            while index < len(lines) and lines[index].strip().startswith("- "):
                items.append(lines[index].strip()[2:].strip())
                index += 1
            blocks.append(("ul", items))
            continue

        if re.match(r"^\d+\.\s", stripped):
            numbered: list[str] = []
            while index < len(lines) and re.match(r"^\d+\.\s", lines[index].strip()):
                numbered.append(re.sub(r"^\d+\.\s*", "", lines[index].strip()))
                index += 1
            blocks.append(("ol", numbered))
            continue

        # 段落。日本語は単語の区切りに空白を使わないため、連結時も空白を入れない。
        paragraph: list[str] = []
        while index < len(lines) and lines[index].strip():
            current = lines[index].strip()
            if current.startswith(("#", "- ", "```", "<!--")) or re.match(
                r"^\d+\.\s", current
            ):
                break
            if set(current) == {"-"} and len(current) >= 3:
                break
            paragraph.append(current)
            index += 1
        blocks.append(("p", "".join(paragraph)))

    return blocks


class ManualDocTemplate(BaseDocTemplate):
    """見出しを目次へ登録しながら組版するテンプレート。"""

    def __init__(self, path: Path, styles: dict[str, ParagraphStyle]) -> None:
        super().__init__(
            str(path),
            pagesize=A4,
            leftMargin=PAGE_MARGIN,
            rightMargin=PAGE_MARGIN,
            topMargin=PAGE_MARGIN,
            bottomMargin=PAGE_MARGIN,
            title="Office PDF コンバーター インストールマニュアル",
            author="Office PDF コンバーター",
        )
        self.styles = styles
        frame = Frame(
            self.leftMargin,
            self.bottomMargin,
            self.width,
            self.height - 8 * mm,
            id="body",
        )
        self.addPageTemplates(
            [PageTemplate(id="manual", frames=[frame], onPage=self._draw_footer)]
        )

    def _draw_footer(self, canvas, document) -> None:
        canvas.saveState()
        canvas.setFont(BODY_FONT, 8)
        canvas.setFillColor(MUTED)
        canvas.setStrokeColor(CODE_BORDER)
        baseline = self.bottomMargin - 6 * mm
        canvas.line(
            self.leftMargin, baseline + 4 * mm, A4[0] - self.rightMargin, baseline + 4 * mm
        )
        canvas.drawString(
            self.leftMargin, baseline, "Office PDF コンバーター インストールマニュアル"
        )
        canvas.drawRightString(A4[0] - self.rightMargin, baseline, str(document.page))
        canvas.restoreState()

    def afterFlowable(self, flowable) -> None:
        if not isinstance(flowable, Paragraph):
            return
        level = {"Heading2": 0, "Heading3": 1}.get(flowable.style.name)
        if level is not None:
            self.notify("TOCEntry", (level, flowable.getPlainText(), self.page))


def _code_block(lines: list[str], styles: dict[str, ParagraphStyle]) -> Table:
    """コードブロックを、背景と枠線を付けた表として組む。"""

    table = Table(
        [[Preformatted("\n".join(lines), styles["Code"])]],
        colWidths=[CONTENT_WIDTH],
        style=TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), CODE_BACKGROUND),
                ("BOX", (0, 0), (-1, -1), 0.6, CODE_BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        ),
        spaceBefore=3,
        spaceAfter=9,
    )
    return table


def _build_story(
    blocks: list[tuple[str, object]], styles: dict[str, ParagraphStyle]
) -> list[object]:
    story: list[object] = []
    for kind, payload in blocks:
        if kind == "h1":
            story.append(Paragraph(_to_markup(str(payload)), styles["Title"]))
        elif kind == "h2":
            story.append(Paragraph(_to_markup(str(payload)), styles["Heading2"]))
        elif kind == "h3":
            story.append(Paragraph(_to_markup(str(payload)), styles["Heading3"]))
        elif kind == "h4":
            story.append(Paragraph(_to_markup(str(payload)), styles["Heading4"]))
        elif kind == "p":
            story.append(Paragraph(_to_markup(str(payload)), styles["Body"]))
        elif kind == "ul":
            for item in payload:  # type: ignore[union-attr]
                story.append(
                    Paragraph(_to_markup(item), styles["Bullet"], bulletText="・")
                )
            story.append(Spacer(0, 5))
        elif kind == "ol":
            for number, item in enumerate(payload, start=1):  # type: ignore[arg-type]
                story.append(
                    Paragraph(
                        _to_markup(item), styles["Ordered"], bulletText=f"{number}."
                    )
                )
            story.append(Spacer(0, 5))
        elif kind == "code":
            story.append(_code_block(payload, styles))  # type: ignore[arg-type]
        elif kind == "pagebreak":
            story.append(PageBreak())
        elif kind == "toc":
            table_of_contents = TableOfContents()
            table_of_contents.levelStyles = [styles["TOC0"], styles["TOC1"]]
            # 見出しではなく専用の書式で描き、目次自身を目次へ載せない。
            story.append(Paragraph("目次", styles["TOCTitle"]))
            story.append(table_of_contents)
    return story


def main() -> None:
    if not SOURCE_PATH.is_file():
        raise SystemExit(f"本文が見つかりません: {SOURCE_PATH}")

    _register_fonts()
    styles = _build_styles()
    blocks = _parse_blocks(SOURCE_PATH.read_text(encoding="utf-8"))

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    document = ManualDocTemplate(OUTPUT_PATH, styles)
    # 目次のページ番号を確定させるため、2回組版する。
    document.multiBuild(_build_story(blocks, styles))
    print(f"生成しました: {OUTPUT_PATH}", file=sys.stderr)


if __name__ == "__main__":
    main()
