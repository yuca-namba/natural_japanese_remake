# /// script
# requires-python = ">=3.10"
# dependencies = [
#     "python-pptx>=0.6.21",
# ]
# ///
"""pptx_to_md.py — パワーポイントのスライドの文字を、スキルに渡す Markdown にする。

スキルは文字を相手にするので、スライドを直すときは、まず文字を取り出して
見出しと本文の形に並べる。並べた Markdown に lint.py などをそのまま当てられる。
パワーポイントには書き戻さない。直した案は、スライドごとの一覧でユーザーに返す。

並べ方は次のとおり。
    - 表紙・背表紙(枠が2つ未満のスライドを含む): 最初の枠を ##、残りを段落
    - スライド上部のラベル(一番上の枠): ## 番号. ラベル
    - メッセージ(2番目の枠。スライドの主見出し): ###
    - リード文(そのほかの枠): 段落
    - 図の中の文字: 段落ごとに、本文より大きい文字(17pt 以上)は ####、それ以外は箇条書き
    - 表: Markdown の表
    - 注記(スライドの下から15%にある枠): 段落
枠の役割は、プレースホルダーかどうかと位置で決める。図の中の文字は位置(上から、左から)の
順に並べるので、左右に列のある図では列が交互に混ざる。図の中のどこにある文字かは、
スライドを画像にして確かめる。

使い方:
    uv run scripts/pptx_to_md.py <deck.pptx> > slides.md
    uv run scripts/pptx_to_md.py <deck.pptx> --notes > slides.md   # スピーカーノートも含める

入力エラー(ファイル不在、パワーポイントとして読めないなど)は exit code 1、それ以外は 0。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# 図の中で、これ以上の大きさの文字を小見出しとみなす(本文はふつう14pt前後)
HEADING_MIN_PT = 17
# これ以上の大きさは数字の強調(「7割超」など)なので、小見出しにしない
HEADING_MAX_PT = 40
# スライドの下からこの割合にある枠を注記とみなす
FOOTNOTE_FROM = 0.85
COVER_LAYOUTS = {"表紙", "背表紙", "Title Slide"}


def _walk(shapes):
    for sh in shapes:
        if sh.shape_type == 6:  # グループ
            yield from _walk(sh.shapes)
        else:
            yield sh


def _text_of(shape) -> str:
    return " ".join(p.text.strip() for p in shape.text_frame.paragraphs if p.text.strip())


def _box_lines(shape) -> list[str]:
    """図の中の枠を、文字の大きさのそろった段落ごとに見出しか箇条書きにする。"""
    groups: list[tuple[bool, list[str]]] = []
    for para in shape.text_frame.paragraphs:
        if not para.text.strip():
            continue
        sizes = [r.font.size.pt for r in para.runs if r.font.size]
        big = bool(sizes) and HEADING_MIN_PT <= max(sizes) < HEADING_MAX_PT
        if groups and groups[-1][0] == big:
            groups[-1][1].append(para.text.strip())
        else:
            groups.append((big, [para.text.strip()]))
    return [("#### " if big else "- ") + " ".join(texts) for big, texts in groups]


def convert(path: Path, notes: bool = False) -> str:
    from pptx import Presentation

    prs = Presentation(str(path))
    slide_height = prs.slide_height
    out: list[str] = []
    for n, slide in enumerate(prs.slides, 1):
        placeholders, boxes, tables = [], [], []
        for sh in _walk(slide.shapes):
            if getattr(sh, "has_table", False) and sh.has_table:
                tables.append(sh)
            elif sh.has_text_frame and sh.text_frame.text.strip():
                (placeholders if sh.is_placeholder else boxes).append(sh)
        placeholders.sort(key=lambda s: s.top or 0)
        out.append(f"<!-- スライド{n} -->")
        if slide.slide_layout.name in COVER_LAYOUTS or len(placeholders) < 2:
            for i, sh in enumerate(placeholders + boxes):
                out += [("## " if i == 0 else "") + _text_of(sh), ""]
        else:
            label, message, *rest = placeholders
            out += [f"## {n}. {_text_of(label)}", "", f"### {_text_of(message)}", ""]
            footnotes = [s for s in rest if (s.top or 0) > slide_height * FOOTNOTE_FROM]
            out += [line for sh in rest if sh not in footnotes for line in (_text_of(sh), "")]
            boxes.sort(key=lambda s: (s.top or 0, s.left or 0))
            for sh in boxes:
                out += _box_lines(sh)
            if boxes:
                out.append("")
            for tb in tables:
                rows = list(tb.table.rows)
                for i, row in enumerate(rows):
                    out.append("| " + " | ".join(c.text.replace("\n", " ") for c in row.cells) + " |")
                    if i == 0:
                        out.append("|" + "---|" * len(row.cells))
                out.append("")
            out += [line for sh in footnotes for line in (_text_of(sh), "")]
        if notes and slide.has_notes_slide:
            text = slide.notes_slide.notes_text_frame.text.strip()
            if text:
                out += ["<!-- スピーカーノート -->", "", text, ""]
    return "\n".join(out)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("file")
    ap.add_argument("--notes", action="store_true", help="スピーカーノートも含める")
    a = ap.parse_args()
    p = Path(a.file)
    if not p.is_file():
        print(f"ファイルが見つからない: {p}", file=sys.stderr)
        return 1
    try:
        print(convert(p, notes=a.notes))
    except Exception as e:  # noqa: BLE001 — 壊れたファイルやパワーポイント以外のファイル
        print(f"パワーポイントとして読めない: {p}: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
