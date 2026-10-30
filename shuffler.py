# -*- coding: utf-8 -*-
"""
Module lõi Trộn Đề Trắc Nghiệm Chuẩn GDPT 2018 (3 Phần: TNKQ 4 lựa chọn, Đúng/Sai, Trả lời ngắn).
Phiên bản v3.0:
1. Cài đặt kích thước Header và Footer chuẩn:
   - section.header_distance = 6 mm
   - section.footer_distance = 6 mm
2. Tự động đồng bộ Mã đề dưới Footer ("Trang X/Y - Mã đề thi <mã_đề>") theo đúng từng mã đề xuất ra
   (xóa bỏ hoàn toàn mã đề cũ như "1111" còn lưu trong footer của file đề gốc).
3. Giữ nguyên toàn bộ định dạng chuẩn của bản v2.6:
   - Bảng tiêu đề 2 cột không viền + Ô đóng khung [MÃ ĐỀ: ...]
   - Lề trang: top = 11 mm, bottom = 11 mm, left = 10 mm, right = 8 mm
   - Tự động chia cột thông minh cho 4 phương án A. B. C. D. (1 dòng / 2 dòng / 4 dòng)
   - Tô màu xanh dương (Blue #0000FF) cho các nhãn A. B. C. D. và a) b) c) d).
"""

from __future__ import annotations
import copy
import io
import random
import re
import zipfile
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import docx
from docx.document import Document as DocxDocument
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
BLUE_HEX = "0000FF"

IGNORED_PARAGRAPH_MARKER_TAGS = {
    qn("w:proofErr"),
    qn("w:bookmarkStart"),
    qn("w:bookmarkEnd"),
    qn("w:permStart"),
    qn("w:permEnd"),
    qn("w:commentRangeStart"),
    qn("w:commentRangeEnd"),
    qn("w:commentReference"),
    qn("w:moveFromRangeStart"),
    qn("w:moveFromRangeEnd"),
    qn("w:moveToRangeStart"),
    qn("w:moveToRangeEnd"),
}

MEDIA_OR_FIELD_RUN_CHILD_TAGS = {
    qn("w:object"),
    qn("w:pict"),
    qn("w:drawing"),
    qn("w:sym"),
    qn("w:fldChar"),
    qn("w:instrText"),
    qn("m:oMath"),
    qn("m:oMathPara"),
}


@dataclass
class OptionBlock:
    """Đại diện cho 1 phương án lựa chọn (A/B/C/D ở Phần I hoặc a/b/c/d ở Phần II)."""
    original_label: str
    is_correct: bool
    xml_elements: List[OxmlElement] = field(default_factory=list)
    fixed_order: bool = False


@dataclass
class QuestionBlock:
    """Đại diện cho 1 câu hỏi trong đề thi."""
    part: int
    original_index: int
    stem_elements: List[OxmlElement] = field(default_factory=list)
    options: List[OptionBlock] = field(default_factory=list)
    short_answer: str = ""


@dataclass
class ParsedExam:
    """Cấu trúc đề thi sau khi phân tích từ file .docx gốc."""
    source_doc: DocxDocument
    part1_questions: List[QuestionBlock] = field(default_factory=list)
    part2_questions: List[QuestionBlock] = field(default_factory=list)
    part3_questions: List[QuestionBlock] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Các hàm tiện ích thao tác XML của Word (<w:p>, <w:r>, <w:t>, <m:oMath>, ...)
# ---------------------------------------------------------------------------

def get_paragraph_text(p_elem: OxmlElement) -> str:
    """Lấy toàn bộ văn bản thuần túy trong thẻ <w:p> để nhận diện từ khóa."""
    texts = []
    for elem in p_elem.iter():
        if elem.tag == qn("w:t") and elem.text:
            texts.append(elem.text)
        elif elem.tag in (qn("w:tab"), qn("w:ptab")):
            texts.append("\t")
        elif elem.tag in (qn("w:br"), qn("w:cr")):
            texts.append("\n")
    return "".join(texts)


def run_has_non_text_media(r_elem: OxmlElement) -> bool:
    """Kiểm tra xem thẻ <w:r> có chứa công thức MathType (<w:object>), hình vẽ (<w:drawing>, <w:pict>), <w:sym>, <w:fldChar> không."""
    for sub in r_elem:
        if sub.tag in MEDIA_OR_FIELD_RUN_CHILD_TAGS or sub.tag.endswith("AlternateContent"):
            return True
    return False


def is_rPr_underlined(rPr: Optional[OxmlElement]) -> bool:
    if rPr is None:
        return False
    u_elem = rPr.find(qn("w:u"))
    if u_elem is not None:
        val = u_elem.get(qn("w:val")) or u_elem.get("val")
        if val is None or val.lower() not in ("none", "0", "false"):
            return True
    return False


def is_rPr_bold(rPr: Optional[OxmlElement]) -> bool:
    if rPr is None:
        return False
    for tag_name in ("w:b", "w:bCs"):
        b_elem = rPr.find(qn(tag_name))
        if b_elem is not None:
            val = b_elem.get(qn("w:val")) or b_elem.get("val")
            if val is None or val.lower() not in ("none", "0", "false"):
                return True
    return False


def is_rPr_caps(rPr: Optional[OxmlElement]) -> bool:
    if rPr is None:
        return False
    for tag_name in ("w:caps", "w:smallCaps"):
        c_elem = rPr.find(qn(tag_name))
        if c_elem is not None:
            val = c_elem.get(qn("w:val")) or c_elem.get("val")
            if val is None or val.lower() not in ("none", "0", "false"):
                return True
    return False


def is_run_underlined(r_elem: OxmlElement, p_elem: Optional[OxmlElement] = None) -> bool:
    rPr = r_elem.find(qn("w:rPr"))
    if is_rPr_underlined(rPr):
        return True
    if p_elem is not None:
        pPr = p_elem.find(qn("w:pPr"))
        if pPr is not None and is_rPr_underlined(pPr.find(qn("w:rPr"))):
            return True
    return False


def is_run_bold(r_elem: OxmlElement, p_elem: Optional[OxmlElement] = None) -> bool:
    rPr = r_elem.find(qn("w:rPr"))
    if is_rPr_bold(rPr):
        return True
    if p_elem is not None:
        pPr = p_elem.find(qn("w:pPr"))
        if pPr is not None and is_rPr_bold(pPr.find(qn("w:rPr"))):
            return True
    return False


def remove_underline_from_rPr(rPr: Optional[OxmlElement]) -> None:
    if rPr is None:
        return
    for u_elem in list(rPr.findall(qn("w:u"))):
        rPr.remove(u_elem)


def remove_underline_from_run(r_elem: OxmlElement) -> None:
    remove_underline_from_rPr(r_elem.find(qn("w:rPr")))


def remove_underline_from_paragraph_pPr(p_elem: OxmlElement) -> None:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is not None:
        remove_underline_from_rPr(pPr.find(qn("w:rPr")))
        for bad_tag in ("w:framePr", "w:pageBreakBefore"):
            for el in list(pPr.findall(qn(bad_tag))):
                pPr.remove(el)


def clean_option_paragraph_underlines(p_elem: OxmlElement) -> None:
    remove_underline_from_paragraph_pPr(p_elem)
    direct_runs = [c for c in p_elem if c.tag == qn("w:r")]
    if not direct_runs:
        return

    non_empty_runs = [
        r for r in direct_runs
        if "".join((t.text or "") for t in r.findall(qn("w:t"))).strip()
    ]
    if non_empty_runs and all(is_run_underlined(r) for r in non_empty_runs):
        for r in direct_runs:
            remove_underline_from_run(r)
        return

    seen_meaningful_content = False
    for child in p_elem:
        if child.tag == qn("w:pPr") or child.tag in IGNORED_PARAGRAPH_MARKER_TAGS:
            continue
        if child.tag != qn("w:r"):
            seen_meaningful_content = True
            continue
        if run_has_non_text_media(child):
            remove_underline_from_run(child)
            seen_meaningful_content = True
            continue
        txt = "".join((t.text or "") for t in child.findall(qn("w:t")))
        stripped = txt.strip()
        if not stripped:
            remove_underline_from_run(child)
            continue
        if not seen_meaningful_content:
            if re.fullmatch(r"[\.\:\)\-\s\u00a0]+", stripped):
                remove_underline_from_run(child)
                continue
            if re.match(r"^[\.\:\)]\s*", stripped) and is_run_underlined(child):
                remove_underline_from_run(child)
            seen_meaningful_content = True


def _iter_paragraph_effective_children(p_elem: OxmlElement) -> List[OxmlElement]:
    out: List[OxmlElement] = []
    for child in p_elem:
        if child.tag == qn("w:pPr") or child.tag in IGNORED_PARAGRAPH_MARKER_TAGS:
            continue
        if child.tag in (qn("w:ins"), qn("w:smartTag"), qn("w:customXml"), qn("w:fldSimple")):
            out.extend(_iter_paragraph_effective_children(child))
        elif child.tag == qn("w:sdt"):
            sdt_content = child.find(qn("w:sdtContent"))
            if sdt_content is not None:
                out.extend(_iter_paragraph_effective_children(sdt_content))
        else:
            out.append(child)
    return out


def strip_prefix_from_paragraph(p_elem: OxmlElement, regex_pattern: str) -> Tuple[OxmlElement, bool, str]:
    new_p = copy.deepcopy(p_elem)
    remove_underline_from_paragraph_pPr(new_p)

    t_nodes: List[Tuple[OxmlElement, OxmlElement, str]] = []
    for child in _iter_paragraph_effective_children(new_p):
        if child.tag != qn("w:r"):
            break
        rPr = child.find(qn("w:rPr"))
        has_caps = is_rPr_caps(rPr)
        stop_here = False
        for sub in child:
            if sub.tag in (qn("w:rPr"), qn("w:lastRenderedPageBreak")) or sub.tag in IGNORED_PARAGRAPH_MARKER_TAGS:
                continue
            if sub.tag == qn("w:t") and sub.text:
                txt_val = sub.text.upper() if has_caps else sub.text
                t_nodes.append((child, sub, txt_val))
            elif sub.tag in (qn("w:tab"), qn("w:ptab")):
                t_nodes.append((child, sub, "\t"))
            elif sub.tag in (qn("w:br"), qn("w:cr")):
                t_nodes.append((child, sub, "\n"))
            elif sub.tag in MEDIA_OR_FIELD_RUN_CHILD_TAGS or sub.tag.endswith("AlternateContent"):
                stop_here = True
                break
        if stop_here:
            break

    full_text = "".join(item[2] for item in t_nodes)
    match = re.match(regex_pattern, full_text, flags=re.IGNORECASE)
    if not match:
        return new_p, False, ""

    matched_str = match.group(0)
    label_char = match.group("label") if "label" in match.groupdict() else ""
    has_asterisk = "*" in matched_str
    cut_len = len(matched_str)

    is_underlined = False
    curr_pos = 0

    for r_elem, t_elem, txt in t_nodes:
        node_start = curr_pos
        node_end = curr_pos + len(txt)
        curr_pos = node_end

        if node_start < cut_len and node_end > 0 and txt.strip():
            if is_run_underlined(r_elem, p_elem):
                is_underlined = True

        if t_elem.tag in (qn("w:tab"), qn("w:ptab"), qn("w:br"), qn("w:cr")):
            if node_end <= cut_len and t_elem.getparent() is not None:
                r_elem.remove(t_elem)
            continue

        if node_end <= cut_len:
            t_elem.text = ""
            remove_underline_from_run(r_elem)
        elif node_start < cut_len < node_end:
            offset = cut_len - node_start
            orig_sub_text = t_elem.text or txt
            t_elem.text = orig_sub_text[offset:].lstrip(" \t\u00a0")
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            remove_underline_from_run(r_elem)

    clean_option_paragraph_underlines(new_p)
    if label_char:
        _trim_option_paragraph_in_place(new_p)
    return new_p, (is_underlined or has_asterisk), label_char


def _split_text_into_option_pieces(raw_txt: str, part: int, run_styled: bool) -> List[Tuple[str, bool]]:
    if part == 1:
        if run_styled:
            inline_re = re.compile(
                r"([\s\u00a0]+)(\*?\s*[A-D]\s*[\.\:\)\uff0e\u2024]\s*)"
                r"|([\s\u00a0]+)(\*?\s*[A-D])$"
                r"|(?<=[\.\;\,])([\s\u00a0]*)(\*?\s*[A-D]\s*[\.\:\)\uff0e\u2024]\s*)"
                r"|(?<=[\.\;\,])([\s\u00a0]*)(\*?\s*[A-D])$"
            )
        else:
            inline_re = re.compile(
                r"([\s\u00a0]{2,}|[\s\u00a0]+(?=\*)|(?<=[\.\;\,])[\s\u00a0]*)(\*?\s*[A-D]\s*[\.\:\)\uff0e\u2024]\s*)"
                r"|([\s\u00a0]{2,}|(?<=[\.\;\,])[\s\u00a0]*)(\*?\s*[A-D])$"
            )
    else:
        if run_styled:
            inline_re = re.compile(
                r"([\s\u00a0]+)(\*?\s*[a-d]\s*[\)\.\uff0e\u2024]\s*)"
                r"|([\s\u00a0]+)(\*?\s*[a-d])$"
            )
        else:
            inline_re = re.compile(
                r"([\s\u00a0]+)(\*?\s*[a-d]\s*\)\s*)"
                r"|([\s\u00a0]{2,}|(?<=[\.\;\,])[\s\u00a0]+)(\*?\s*[a-d]\s*[\.\uff0e\u2024]\s*)"
                r"|([\s\u00a0]+)(\*?\s*[a-d])$"
            )

    pieces: List[Tuple[str, bool]] = []
    pos = 0
    for m in inline_re.finditer(raw_txt):
        sep_space = ""
        marker_part = ""
        start_sep = m.start()
        for g_idx in range(1, len(m.groups()), 2):
            if m.group(g_idx + 1) is not None:
                sep_space = m.group(g_idx) or ""
                marker_part = m.group(g_idx + 1) or ""
                start_sep = m.start(g_idx) if m.group(g_idx) is not None else m.start(g_idx + 1)
                break

        before_part = raw_txt[pos:start_sep]
        if before_part:
            pieces.append((before_part, False))
        if sep_space:
            pieces.append((sep_space, True))
        if marker_part:
            pieces.append((marker_part, False))
        pos = m.end()

    rem_txt = raw_txt[pos:]
    if rem_txt:
        pieces.append((rem_txt, False))

    return pieces


def _split_runs_at_char_offsets(p_elem: OxmlElement, part: int = 1) -> List[OxmlElement]:
    expanded: List[OxmlElement] = []

    for child in _iter_paragraph_effective_children(p_elem):
        if child.tag != qn("w:r"):
            expanded.append(copy.deepcopy(child))
            continue

        has_t = any(sub.tag == qn("w:t") for sub in child)
        has_tab = any(sub.tag in (qn("w:tab"), qn("w:ptab")) for sub in child)
        has_br = any(sub.tag in (qn("w:br"), qn("w:cr")) for sub in child)

        if not has_t and not has_tab and not has_br:
            expanded.append(copy.deepcopy(child))
            continue

        rPr = child.find(qn("w:rPr"))
        has_caps = is_rPr_caps(rPr)
        run_styled = is_rPr_bold(rPr) or is_rPr_underlined(rPr) or is_run_bold(child, p_elem) or is_run_underlined(child, p_elem)

        for sub in child:
            if sub.tag in (qn("w:rPr"), qn("w:lastRenderedPageBreak")) or sub.tag in IGNORED_PARAGRAPH_MARKER_TAGS:
                continue

            if sub.tag in (qn("w:tab"), qn("w:ptab")):
                new_r = OxmlElement("w:r")
                if rPr is not None:
                    new_r.append(copy.deepcopy(rPr))
                tab_el = OxmlElement("w:tab")
                new_r.append(tab_el)
                expanded.append(new_r)
                continue

            if sub.tag in (qn("w:br"), qn("w:cr")):
                new_r = OxmlElement("w:r")
                if rPr is not None:
                    new_r.append(copy.deepcopy(rPr))
                br_el = OxmlElement("w:br")
                new_r.append(br_el)
                expanded.append(new_r)
                continue

            if sub.tag == qn("w:t"):
                raw_txt = (sub.text or "")
                if not raw_txt:
                    continue
                if has_caps:
                    raw_txt = raw_txt.upper()
                pieces = _split_text_into_option_pieces(raw_txt, part=part, run_styled=run_styled)
                for piece_txt, is_sep in pieces:
                    r_piece = OxmlElement("w:r")
                    if not is_sep and rPr is not None:
                        r_piece.append(copy.deepcopy(rPr))
                    t_piece = OxmlElement("w:t")
                    t_piece.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t_piece.text = piece_txt
                    r_piece.append(t_piece)
                    expanded.append(r_piece)
            else:
                new_r = OxmlElement("w:r")
                if rPr is not None:
                    new_r.append(copy.deepcopy(rPr))
                new_r.append(copy.deepcopy(sub))
                expanded.append(new_r)

    return expanded


def split_inline_options_in_paragraph(
    p_elem: OxmlElement,
    part: int = 1,
    allow_start_at_any_valid_label: bool = False,
) -> Optional[List[Tuple[str, bool, OxmlElement]]]:
    if part == 1:
        opt_prefix_re = re.compile(r"^([\s\u00a0]*\*?[\s\u00a0]*([A-D])[\s\u00a0]*([\.\:\)\uff0e\u2024])[\s\u00a0]*)")
        single_strip_re = r"^[\s\u00a0]*\*?[\s\u00a0]*(?P<label>[A-D])[\s\u00a0]*[\.\:\)\uff0e\u2024][\s\u00a0]*"
        valid_seq = ["A", "B", "C", "D"]
    else:
        opt_prefix_re = re.compile(r"^([\s\u00a0]*\*?[\s\u00a0]*([a-d])[\s\u00a0]*([\)\.\uff0e\u2024])[\s\u00a0]*)")
        single_strip_re = r"^[\s\u00a0]*\*?[\s\u00a0]*(?P<label>[a-d])[\s\u00a0]*[\)\.\uff0e\u2024][\s\u00a0]*"
        valid_seq = ["a", "b", "c", "d"]

    test_p, test_corr, test_lbl = strip_prefix_from_paragraph(p_elem, single_strip_re)
    if not test_lbl and not allow_start_at_any_valid_label:
        return None

    pPr = p_elem.find(qn("w:pPr"))
    children = _split_runs_at_char_offsets(p_elem, part=part)

    segments: List[Tuple[str, bool, List[OxmlElement]]] = []
    current_label: Optional[str] = None
    current_correct: bool = False
    current_nodes: List[OxmlElement] = []

    is_line_start = True
    after_strong_sep = True
    after_any_space = True
    after_dot_or_formula = False
    consecutive_spaces = 0

    i = 0
    n = len(children)
    while i < n:
        node = children[i]

        if node.tag in IGNORED_PARAGRAPH_MARKER_TAGS:
            i += 1
            continue

        if node.tag == qn("w:r") and not node.findall(qn("w:t")) and not run_has_non_text_media(node):
            has_tab = node.find(qn("w:tab")) is not None or node.find(qn("w:ptab")) is not None
            has_br = node.find(qn("w:br")) is not None or node.find(qn("w:cr")) is not None
            if has_tab or has_br:
                if has_br:
                    is_line_start = True
                after_strong_sep = True
                after_any_space = True
                i += 1
                continue

        if node.tag == qn("w:r") and run_has_non_text_media(node):
            node_copy = copy.deepcopy(node)
            remove_underline_from_run(node_copy)
            current_nodes.append(node_copy)
            is_line_start = False
            after_strong_sep = False
            after_any_space = True
            after_dot_or_formula = True
            consecutive_spaces = 0
            i += 1
            continue

        if node.tag == qn("w:r"):
            txt_curr = "".join((t.text or "") for t in node.findall(qn("w:t")))

            if not txt_curr.strip():
                consecutive_spaces += len(txt_curr)
                if consecutive_spaces >= 2 or "\t" in txt_curr or "\u00a0" in txt_curr:
                    after_strong_sep = True
                after_any_space = True
                if current_nodes:
                    node_copy = copy.deepcopy(node)
                    remove_underline_from_run(node_copy)
                    current_nodes.append(node_copy)
                i += 1
                continue

            lookahead_runs: List[Tuple[int, OxmlElement, str]] = []
            j = i
            while j < n and len(lookahead_runs) < 6:
                cand = children[j]
                if cand.tag in IGNORED_PARAGRAPH_MARKER_TAGS:
                    j += 1
                    continue
                if (
                    cand.tag != qn("w:r")
                    or cand.find(qn("w:tab")) is not None
                    or cand.find(qn("w:ptab")) is not None
                    or cand.find(qn("w:br")) is not None
                    or cand.find(qn("w:cr")) is not None
                    or run_has_non_text_media(cand)
                ):
                    break
                t_list = cand.findall(qn("w:t"))
                if not t_list:
                    break
                txt_j = "".join((t.text or "") for t in t_list)
                if not txt_j:
                    break
                lookahead_runs.append((j, cand, txt_j))
                j += 1

            combined_txt = "".join(item[2] for item in lookahead_runs)
            m_opt = opt_prefix_re.match(combined_txt) if lookahead_runs else None

            if m_opt:
                matched_prefix = m_opt.group(1)
                cand_raw_label = m_opt.group(2)
                punct_char = m_opt.group(3)
                cand_label = cand_raw_label.upper() if part == 1 else cand_raw_label.lower()

                is_expected_label = False
                if current_label is None:
                    is_expected_label = cand_label in valid_seq
                else:
                    try:
                        curr_idx = valid_seq.index(current_label)
                        cand_idx = valid_seq.index(cand_label)
                        is_expected_label = cand_idx > curr_idx
                    except ValueError:
                        is_expected_label = False

                prefix_len = len(matched_prefix)
                has_ast = "*" in matched_prefix
                label_is_bold = False
                label_is_ul = False
                c_pos = 0
                for _, r_cand, r_txt in lookahead_runs:
                    if c_pos >= prefix_len:
                        break
                    overlap = r_txt[: max(0, prefix_len - c_pos)]
                    c_pos += len(r_txt)
                    if overlap.strip():
                        if is_run_underlined(r_cand, p_elem):
                            label_is_ul = True
                        if is_run_bold(r_cand, p_elem):
                            label_is_bold = True

                leading_spaces = len(matched_prefix) - len(matched_prefix.lstrip(" \t\u00a0"))
                is_genuine_option_marker = (
                    is_line_start
                    or after_strong_sep
                    or (leading_spaces >= 2)
                    or label_is_bold
                    or label_is_ul
                    or has_ast
                    or (current_label is not None and after_dot_or_formula)
                    or ((after_any_space or leading_spaces >= 1) and punct_char == ")")
                )

                if is_expected_label and is_genuine_option_marker:
                    if current_label is not None:
                        segments.append((current_label, current_correct, current_nodes))

                    consumed_chars = 0
                    last_consumed_idx = i
                    leftover_runs: List[OxmlElement] = []

                    for run_idx, r_cand, r_txt in lookahead_runs:
                        if consumed_chars >= prefix_len:
                            break
                        r_start = consumed_chars
                        r_end = consumed_chars + len(r_txt)
                        consumed_chars = r_end
                        last_consumed_idx = run_idx

                        if r_end > prefix_len:
                            offset = prefix_len - r_start
                            rem_str = r_txt[offset:].lstrip(" \t\u00a0")
                            if rem_str or run_has_non_text_media(r_cand):
                                l_run = copy.deepcopy(r_cand)
                                remove_underline_from_run(l_run)
                                for idx_t, t_el in enumerate(l_run.findall(qn("w:t"))):
                                    t_el.text = rem_str if idx_t == 0 else ""
                                    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                                leftover_runs.append(l_run)
                        elif run_has_non_text_media(r_cand):
                            l_run = copy.deepcopy(r_cand)
                            remove_underline_from_run(l_run)
                            for t_el in l_run.findall(qn("w:t")):
                                t_el.text = ""
                            leftover_runs.append(l_run)

                    current_label = cand_label
                    current_correct = has_ast or label_is_ul
                    current_nodes = []
                    is_line_start = False
                    consecutive_spaces = 0

                    if leftover_runs:
                        current_nodes.extend(leftover_runs)
                        l_txt = "".join((t.text or "") for t in leftover_runs[-1].findall(qn("w:t")))
                        after_strong_sep = bool(re.search(r"(\t|[\s\u00a0]{2,})$", l_txt))
                        after_any_space = bool(re.search(r"[\s\u00a0\.\;\,]$", l_txt))
                        after_dot_or_formula = bool(re.search(r"[\.\;\,][\s\u00a0]*$", l_txt))
                    else:
                        after_strong_sep = False
                        after_any_space = True
                        after_dot_or_formula = False

                    i = last_consumed_idx + 1
                    continue

            node_copy = copy.deepcopy(node)
            if not current_nodes:
                if is_run_underlined(node_copy):
                    remove_underline_from_run(node_copy)
                t_els = node_copy.findall(qn("w:t"))
                if t_els and t_els[0].text:
                    t_els[0].text = t_els[0].text.lstrip(" \t\u00a0")
                    t_els[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            current_nodes.append(node_copy)
            is_line_start = False
            consecutive_spaces = 0
            after_strong_sep = bool(re.search(r"(\t|[\s\u00a0]{2,})$", txt_curr))
            after_any_space = bool(re.search(r"[\s\u00a0\.\;\,]$", txt_curr))
            after_dot_or_formula = bool(re.search(r"[\.\;\,][\s\u00a0]*$", txt_curr))
            i += 1
        else:
            current_nodes.append(copy.deepcopy(node))
            is_line_start = False
            consecutive_spaces = 0
            after_strong_sep = False
            after_any_space = True
            after_dot_or_formula = True
            i += 1

    if current_label is not None:
        segments.append((current_label, current_correct, current_nodes))

    if not segments:
        return None

    if len(segments) == 1 and test_lbl:
        lbl0 = test_lbl.upper() if part == 1 else test_lbl.lower()
        return [(lbl0, test_corr, test_p)]

    result: List[Tuple[str, bool, OxmlElement]] = []
    for lbl, is_corr, nodes in segments:
        opt_p = OxmlElement("w:p")
        if pPr is not None:
            pPr_copy = copy.deepcopy(pPr)
            remove_underline_from_rPr(pPr_copy.find(qn("w:rPr")))
            opt_p.append(pPr_copy)
        for nd in nodes:
            opt_p.append(nd)
        clean_option_paragraph_underlines(opt_p)
        _trim_option_paragraph_in_place(opt_p)
        result.append((lbl, is_corr, opt_p))

    return result


def split_paragraph_on_line_breaks_before_options(p_elem: OxmlElement, part: int = 1) -> List[OxmlElement]:
    br_elems = p_elem.findall(".//" + qn("w:br")) + p_elem.findall(".//" + qn("w:cr"))
    if not br_elems:
        return [p_elem]

    pPr = p_elem.find(qn("w:pPr"))
    lines: List[List[OxmlElement]] = [[]]

    for child in _iter_paragraph_effective_children(p_elem):
        if child.tag != qn("w:r"):
            lines[-1].append(copy.deepcopy(child))
            continue

        has_br = any(sub.tag in (qn("w:br"), qn("w:cr")) for sub in child)
        if not has_br:
            lines[-1].append(copy.deepcopy(child))
            continue

        rPr = child.find(qn("w:rPr"))
        curr_r_subs: List[OxmlElement] = []
        for sub in child:
            if sub.tag == qn("w:rPr"):
                continue
            if sub.tag in (qn("w:br"), qn("w:cr")):
                if curr_r_subs:
                    new_r = OxmlElement("w:r")
                    if rPr is not None:
                        new_r.append(copy.deepcopy(rPr))
                    for s in curr_r_subs:
                        new_r.append(s)
                    lines[-1].append(new_r)
                    curr_r_subs = []
                lines.append([])
            else:
                curr_r_subs.append(copy.deepcopy(sub))

        if curr_r_subs:
            new_r = OxmlElement("w:r")
            if rPr is not None:
                new_r.append(copy.deepcopy(rPr))
            for s in curr_r_subs:
                new_r.append(s)
            lines[-1].append(new_r)

    if len(lines) <= 1:
        return [p_elem]

    opt_line_re = (
        re.compile(r"^[\s\u00a0]*\*?[\s\u00a0]*[A-D][\s\u00a0]*[\.\:\)\uff0e\u2024]")
        if part == 1
        else re.compile(r"^[\s\u00a0]*\*?[\s\u00a0]*[a-d][\s\u00a0]*[\)\.\uff0e\u2024]")
    )

    has_opt_after_br = False
    for line_nodes in lines[1:]:
        temp_p = OxmlElement("w:p")
        for nd in line_nodes:
            temp_p.append(copy.deepcopy(nd))
        if opt_line_re.match(get_paragraph_text(temp_p)):
            has_opt_after_br = True
            break

    if not has_opt_after_br:
        return [p_elem]

    out_paragraphs: List[OxmlElement] = []
    for line_nodes in lines:
        if not line_nodes:
            continue
        new_p = OxmlElement("w:p")
        if pPr is not None:
            new_p.append(copy.deepcopy(pPr))
        for nd in line_nodes:
            new_p.append(nd)
        out_paragraphs.append(new_p)

    return out_paragraphs if out_paragraphs else [p_elem]


def split_stem_and_trailing_options(
    stem_p: OxmlElement,
    part: int = 1,
) -> Tuple[OxmlElement, Optional[List[Tuple[str, bool, OxmlElement]]]]:
    pPr = stem_p.find(qn("w:pPr"))
    children = _split_runs_at_char_offsets(stem_p, part=part)
    if len(children) < 2:
        return stem_p, None

    first_opt_re = (
        re.compile(r"^[\s\u00a0]*\*?[\s\u00a0]*A[\s\u00a0]*[\.\:\)\uff0e\u2024]")
        if part == 1
        else re.compile(r"^[\s\u00a0]*\*?[\s\u00a0]*a[\s\u00a0]*[\)\.\uff0e\u2024]")
    )

    for idx in range(1, len(children)):
        nd = children[idx]
        if nd.tag != qn("w:r") or run_has_non_text_media(nd):
            continue
        t_list = nd.findall(qn("w:t"))
        if not t_list:
            continue
        comb = ""
        for k in range(idx, min(len(children), idx + 3)):
            ck = children[k]
            if ck.tag != qn("w:r") or run_has_non_text_media(ck):
                break
            comb += "".join((t.text or "") for t in ck.findall(qn("w:t")))
        if not first_opt_re.match(comb):
            continue

        prev_nd = children[idx - 1]
        prev_is_tab = prev_nd.tag == qn("w:r") and (prev_nd.find(qn("w:tab")) is not None or prev_nd.find(qn("w:ptab")) is not None)
        prev_txt = "".join((t.text or "") for t in prev_nd.findall(qn("w:t"))) if prev_nd.tag == qn("w:r") else ""
        prev_is_space = bool(re.search(r"(\t|[\s\u00a0]{2,})$", prev_txt))
        is_styled = is_run_bold(nd, stem_p) or is_run_underlined(nd, stem_p) or ("*" in comb[:4])

        if not (is_styled or prev_is_tab or prev_is_space):
            continue

        cand_opts_p = OxmlElement("w:p")
        if pPr is not None:
            cand_opts_p.append(copy.deepcopy(pPr))
        for rem_nd in children[idx:]:
            cand_opts_p.append(copy.deepcopy(rem_nd))

        opts = split_inline_options_in_paragraph(cand_opts_p, part=part)
        if opts and len(opts) >= 2 and opts[0][0] == ("A" if part == 1 else "a"):
            new_stem_p = OxmlElement("w:p")
            if pPr is not None:
                new_stem_p.append(copy.deepcopy(pPr))
            stem_slice = children[:idx]
            while stem_slice and stem_slice[-1].tag == qn("w:r") and not run_has_non_text_media(stem_slice[-1]):
                last_r = stem_slice[-1]
                has_tb = last_r.find(qn("w:tab")) is not None or last_r.find(qn("w:ptab")) is not None
                l_txt = "".join((t.text or "") for t in last_r.findall(qn("w:t")))
                if has_tb and not l_txt.strip():
                    stem_slice.pop()
                elif not l_txt.strip() and not has_tb:
                    stem_slice.pop()
                else:
                    break
            for s_nd in stem_slice:
                new_stem_p.append(copy.deepcopy(s_nd))
            return new_stem_p, opts

    return stem_p, None


def create_bold_label_run(label_text: str, color_hex: Optional[str] = None) -> OxmlElement:
    r_label = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    b_el = OxmlElement("w:b")
    rPr.append(b_el)

    u_none = OxmlElement("w:u")
    u_none.set(qn("w:val"), "none")
    rPr.append(u_none)

    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:ascii"), "Times New Roman")
    rFonts.set(qn("w:hAnsi"), "Times New Roman")
    rFonts.set(qn("w:cs"), "Times New Roman")
    rPr.append(rFonts)

    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "24")  # 12pt
    rPr.append(sz)

    if color_hex:
        color_el = OxmlElement("w:color")
        color_el.set(qn("w:val"), color_hex)
        rPr.append(color_el)

    r_label.append(rPr)
    t_label = OxmlElement("w:t")
    t_label.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    t_label.text = label_text
    r_label.append(t_label)
    return r_label


def prepend_bold_label_to_paragraph(
    p_elem: OxmlElement,
    label_text: str,
    color_hex: Optional[str] = None,
    hanging_indent_dxa: Optional[int] = None,
    left_indent_dxa: Optional[int] = None,
) -> OxmlElement:
    new_p = copy.deepcopy(p_elem)
    remove_underline_from_paragraph_pPr(new_p)

    r_label = create_bold_label_run(label_text, color_hex=color_hex)

    pPr = new_p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        new_p.insert(0, pPr)

    if hanging_indent_dxa is not None or left_indent_dxa is not None:
        for old_ind in list(pPr.findall(qn("w:ind"))):
            pPr.remove(old_ind)
        ind = OxmlElement("w:ind")
        if left_indent_dxa is not None:
            ind.set(qn("w:left"), str(left_indent_dxa))
        if hanging_indent_dxa is not None:
            ind.set(qn("w:hanging"), str(hanging_indent_dxa))
        pPr.append(ind)

    pPr.addnext(r_label)
    return new_p


def _extract_trimmed_option_content_nodes(p_elem: OxmlElement) -> List[OxmlElement]:
    nodes: List[OxmlElement] = []
    for child in _iter_paragraph_effective_children(p_elem):
        nodes.append(copy.deepcopy(child))

    # Xóa các thẻ <w:tab>, <w:ptab>, <w:br>, <w:cr> bên trong các run của phương án
    # và chuẩn hóa khoảng trắng thừa (nhiều dấu cách liên tiếp -> 1 dấu cách)
    for nd in nodes:
        if nd.tag == qn("w:r"):
            for bad_sub in (
                list(nd.findall(qn("w:tab")))
                + list(nd.findall(qn("w:ptab")))
                + list(nd.findall(qn("w:br")))
                + list(nd.findall(qn("w:cr")))
            ):
                nd.remove(bad_sub)
            for t_el in nd.findall(qn("w:t")):
                if t_el.text:
                    t_el.text = re.sub(r"[ \t\u00a0]{2,}", " ", t_el.text)
                    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    # Cắt bỏ các run rỗng hoặc chỉ chứa khoảng trắng ở đầu phương án
    while nodes:
        first = nodes[0]
        if first.tag != qn("w:r") or run_has_non_text_media(first):
            break
        t_els = first.findall(qn("w:t"))
        txt = "".join((t.text or "") for t in t_els)
        if not txt.strip():
            nodes.pop(0)
        else:
            if t_els and t_els[0].text:
                t_els[0].text = t_els[0].text.lstrip(" \t\u00a0")
                t_els[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            break

    # Cắt bỏ các run rỗng hoặc chỉ chứa khoảng trắng ở cuối phương án (ví dụ khoảng trắng dùng để căn cột trong đề gốc)
    while nodes:
        last = nodes[-1]
        if last.tag != qn("w:r") or run_has_non_text_media(last):
            break
        t_els = last.findall(qn("w:t"))
        txt = "".join((t.text or "") for t in t_els)
        if not txt.strip():
            nodes.pop()
        else:
            if t_els and t_els[-1].text:
                t_els[-1].text = t_els[-1].text.rstrip(" \t\u00a0")
                t_els[-1].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            break

    return nodes


def _trim_option_paragraph_in_place(p_elem: OxmlElement) -> None:
    """Làm sạch khoảng trắng đầu/cuối và tab thừa trực tiếp trong thẻ <w:p> của phương án."""
    pPr = p_elem.find(qn("w:pPr"))
    pPr_copy = copy.deepcopy(pPr) if pPr is not None else None
    trimmed_nodes = _extract_trimmed_option_content_nodes(p_elem)
    for child in list(p_elem):
        p_elem.remove(child)
    if pPr_copy is not None:
        p_elem.append(pPr_copy)
    for nd in trimmed_nodes:
        p_elem.append(nd)


def _convert_css_dim_to_pt(val_str: str, unit_str: Optional[str]) -> float:
    val = float(val_str.replace(",", "."))
    unit = (unit_str or "pt").lower()
    if unit == "in":
        return val * 72.0
    if unit == "cm":
        return val * 28.35
    if unit == "mm":
        return val * 2.835
    if unit == "px":
        return val * 0.75
    return val


def _measure_xml_node_visual_chars(elem: OxmlElement) -> float:
    ln = elem.tag.split("}")[-1] if "}" in elem.tag else elem.tag

    if ln == "AlternateContent":
        for child in elem:
            c_ln = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if c_ln == "Choice":
                return sum(_measure_xml_node_visual_chars(sub) for sub in child)
        for child in elem:
            c_ln = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if c_ln == "Fallback":
                return sum(_measure_xml_node_visual_chars(sub) for sub in child)
        return 0.0

    if ln == "drawing":
        for sub in elem.iter():
            s_ln = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
            if s_ln == "extent":
                cx_str = sub.get("cx")
                cy_str = sub.get("cy")
                if cx_str and cy_str:
                    try:
                        w_pt = float(cx_str) / 12700.0
                        h_pt = float(cy_str) / 12700.0
                        if h_pt > 42.0:
                            return 999.0
                        return max(2.0, w_pt / 5.8)
                    except ValueError:
                        pass
        return 999.0

    if ln == "shape":
        style_str = elem.get("style") or ""
        m_h = re.search(
            r"(?:^|;|\s)height\s*:\s*([\d\.,]+)\s*(pt|in|cm|mm|px)\b",
            style_str,
            flags=re.IGNORECASE,
        )
        if m_h:
            h_pt = _convert_css_dim_to_pt(m_h.group(1), m_h.group(2))
            if h_pt > 42.0:
                return 999.0

        m_w = re.search(
            r"(?:^|;|\s)width\s*:\s*([\d\.,]+)\s*(pt|in|cm|mm|px)\b",
            style_str,
            flags=re.IGNORECASE,
        )
        if m_w:
            w_pt = _convert_css_dim_to_pt(m_w.group(1), m_w.group(2))
            return max(2.0, w_pt / 5.8)
        return 4.5

    if ln == "oMath":
        math_texts = []
        for sub in elem.iter():
            s_ln = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
            if s_ln == "t" and sub.text:
                math_texts.append(sub.text)
        combined = re.sub(r"\s+", " ", "".join(math_texts)).strip()
        return float(len(combined)) + 1.0

    if ln == "t":
        if not elem.text:
            return 0.0
        norm_txt = re.sub(r"[ \t\u00a0]+", " ", elem.text)
        return float(len(norm_txt))

    if ln == "sym":
        return 1.5

    return sum(_measure_xml_node_visual_chars(child) for child in elem)


def estimate_option_visual_length(opt: OptionBlock) -> int:
    if not opt.xml_elements:
        return 0
    if any(el.tag != qn("w:p") for el in opt.xml_elements):
        return 999
    if len(opt.xml_elements) > 1:
        return 999

    p_el = opt.xml_elements[0]
    trimmed_nodes = _extract_trimmed_option_content_nodes(p_el)
    total_chars = sum(_measure_xml_node_visual_chars(nd) for nd in trimmed_nodes)
    return int(round(total_chars))


def build_part1_options_paragraphs(
    opts: List[OptionBlock],
    letters: List[str],
    label_color_hex: str = BLUE_HEX,
) -> List[OxmlElement]:
    max_len = max((estimate_option_visual_length(o) for o in opts), default=999)

    if len(opts) == 4 and max_len <= 21:
        cols_per_row = 4
        tab_stops = [3040, 5655, 8270]
    elif len(opts) == 4 and max_len <= 48:
        cols_per_row = 2
        tab_stops = [5655]
    else:
        cols_per_row = 1
        tab_stops = []

    if cols_per_row == 1:
        out_paras: List[OxmlElement] = []
        for o_idx, opt in enumerate(opts):
            letter = letters[o_idx] if o_idx < len(letters) else chr(ord("A") + o_idx)
            for el_idx, el in enumerate(opt.xml_elements):
                if el_idx == 0 and el.tag == qn("w:p"):
                    labeled_opt = prepend_bold_label_to_paragraph(
                        el,
                        f"{letter}. ",
                        color_hex=label_color_hex,
                        left_indent_dxa=425,
                    )
                    out_paras.append(labeled_opt)
                else:
                    out_paras.append(copy.deepcopy(el))
        return out_paras

    out_paras = []
    for row_start in range(0, len(opts), cols_per_row):
        row_opts = opts[row_start : row_start + cols_per_row]
        row_p = OxmlElement("w:p")
        pPr = OxmlElement("w:pPr")

        tabs_el = OxmlElement("w:tabs")
        for pos in tab_stops:
            t_stop = OxmlElement("w:tab")
            t_stop.set(qn("w:val"), "left")
            t_stop.set(qn("w:pos"), str(pos))
            tabs_el.append(t_stop)
        pPr.append(tabs_el)

        spacing = OxmlElement("w:spacing")
        spacing.set(qn("w:before"), "20")
        spacing.set(qn("w:after"), "30")
        spacing.set(qn("w:line"), "260")
        spacing.set(qn("w:lineRule"), "auto")
        pPr.append(spacing)

        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "425")
        pPr.append(ind)

        row_p.append(pPr)

        for c_idx, opt in enumerate(row_opts):
            global_idx = row_start + c_idx
            letter = letters[global_idx] if global_idx < len(letters) else chr(ord("A") + global_idx)
            if c_idx > 0:
                r_tab = OxmlElement("w:r")
                r_tab.append(OxmlElement("w:tab"))
                row_p.append(r_tab)

            r_lbl = create_bold_label_run(f"{letter}. ", color_hex=label_color_hex)
            row_p.append(r_lbl)

            content_nodes = _extract_trimmed_option_content_nodes(opt.xml_elements[0])
            for nd in content_nodes:
                row_p.append(nd)

        out_paras.append(row_p)

    return out_paras


# ---------------------------------------------------------------------------
# Bộ phân tích cú pháp (Parser) file Word đề gốc (.docx)
# ---------------------------------------------------------------------------

PART1_HEADER_RE = re.compile(
    r"^\s*(?:PHẦN\s*(?:I|1)\b|(?:I|1)\s*[\.\:\)]\s*(?:PHẦN\b|CÂU\s+TRẮC\s+NGHIỆM|TRẮC\s+NGHIỆM))",
    re.IGNORECASE,
)
PART2_HEADER_RE = re.compile(
    r"^\s*(?:PHẦN\s*(?:II|2)\b|(?:II|2)\s*[\.\:\)]\s*(?:PHẦN\b|CÂU\s+TRẮC\s+NGHIỆM|TRẮC\s+NGHIỆM))",
    re.IGNORECASE,
)
PART3_HEADER_RE = re.compile(
    r"^\s*(?:PHẦN\s*(?:III|3)\b|(?:III|3)\s*[\.\:\)]\s*(?:PHẦN\b|CÂU\s+TRẮC\s+NGHIỆM|TRẮC\s+NGHIỆM))",
    re.IGNORECASE,
)
END_MARKER_RE = re.compile(r"^\s*[-—–_=\*]{2,}\s*HẾT\s*[-—–_=\*]{2,}", re.IGNORECASE)

QUESTION_START_RE = re.compile(
    r"^\s*(?P<prefix>(?:Câu|Question|Bài)\s*\d+\s*(?:\([^)]*\))?\s*[\.\:\)]\s*)",
    re.IGNORECASE,
)
SHORT_ANS_LINE_RE = re.compile(
    r"^\s*(?:Đáp\s*án|ĐS|Đ/A|ĐÁP\s*ÁN|Trả\s*lời|Key)\s*[\:\=]\s*(?P<ans>.+)$",
    re.IGNORECASE,
)
FIXED_OPTION_RE = re.compile(
    r"(?:cả\s+[a-d].*(?:đều|đúng|sai)|tất\s+cả\s+đều|không\s+có\s+đáp\s+án)",
    re.IGNORECASE,
)



def _int_to_roman(val: int) -> str:
    if val <= 0:
        return str(val)
    roman_pairs = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    res = []
    n = val
    for num, sym in roman_pairs:
        while n >= num:
            res.append(sym)
            n -= num
    return "".join(res)


def _format_numbering_counter(val: int, num_fmt: str) -> str:
    fmt = (num_fmt or "decimal").strip()
    if fmt == "decimal":
        return str(val)
    if fmt == "decimalZero":
        return f"{val:02d}"
    if fmt in ("upperLetter", "ideographTraditional", "taiwaneseCounting"):
        idx = max(1, val) - 1
        return chr(ord("A") + (idx % 26)) * ((idx // 26) + 1)
    if fmt == "lowerLetter":
        idx = max(1, val) - 1
        return chr(ord("a") + (idx % 26)) * ((idx // 26) + 1)
    if fmt == "upperRoman":
        return _int_to_roman(val)
    if fmt == "lowerRoman":
        return _int_to_roman(val).lower()
    if fmt in ("bullet", "none"):
        return ""
    return str(val)


def _flatten_numbering_fields_in_paragraph(p_elem: OxmlElement, seq_counters: Dict[str, int]) -> int:
    """
    Chuyển các trường đánh số tự động dạng Field (<w:fldSimple> hoặc <w:fldChar>/<w:instrText>
    như SEQ, AUTONUM, LISTNUM) ở đầu đoạn văn thành văn bản thuần <w:r><w:t>.
    """
    converted = 0

    # 1. Unwrap <w:fldSimple> nếu là trường đánh số SEQ / AUTONUM / LISTNUM
    for child in list(p_elem):
        if child.tag == qn("w:fldSimple"):
            instr = (child.get(qn("w:instr")) or child.get("instr") or "").strip()
            if re.match(r"^(?:SEQ|AUTONUM|LISTNUM)\b", instr, flags=re.IGNORECASE):
                idx = list(p_elem).index(child)
                runs = [copy.deepcopy(sub) for sub in child if sub.tag == qn("w:r")]
                if not runs:
                    seq_key = instr.upper()
                    seq_counters[seq_key] = seq_counters.get(seq_key, 0) + 1
                    r_new = OxmlElement("w:r")
                    t_new = OxmlElement("w:t")
                    t_new.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t_new.text = f"{seq_counters[seq_key]}. "
                    r_new.append(t_new)
                    runs = [r_new]
                p_elem.remove(child)
                for offset, r_el in enumerate(runs):
                    p_elem.insert(idx + offset, r_el)
                converted += 1

    # 2. Xử lý complex field (<w:fldChar w:fldCharType="begin"> ... <w:instrText>SEQ/AUTONUM/LISTNUM</w:instrText> ... <w:fldChar w:fldCharType="end">)
    runs = [c for c in p_elem if c.tag == qn("w:r")]
    in_field = False
    is_num_field = False
    in_result = False
    field_instr = ""
    runs_to_remove: List[OxmlElement] = []
    result_runs: List[OxmlElement] = []

    for r_el in runs:
        fld_chars = r_el.findall(qn("w:fldChar"))
        instr_els = r_el.findall(qn("w:instrText"))

        if fld_chars:
            for fc in fld_chars:
                f_type = fc.get(qn("w:fldCharType")) or fc.get("fldCharType")
                if f_type == "begin":
                    in_field = True
                    is_num_field = False
                    in_result = False
                    field_instr = ""
                    runs_to_remove = [r_el]
                    result_runs = []
                elif f_type == "separate" and in_field:
                    if re.match(r"^\s*(?:SEQ|AUTONUM|LISTNUM)\b", field_instr, flags=re.IGNORECASE):
                        is_num_field = True
                        runs_to_remove.append(r_el)
                    in_result = True
                elif f_type == "end" and in_field:
                    if is_num_field or re.match(r"^\s*(?:SEQ|AUTONUM|LISTNUM)\b", field_instr, flags=re.IGNORECASE):
                        is_num_field = True
                        runs_to_remove.append(r_el)
                        if not result_runs:
                            seq_key = re.sub(r"\s+", " ", field_instr.strip().upper())
                            seq_counters[seq_key] = seq_counters.get(seq_key, 0) + 1
                            t_new = OxmlElement("w:t")
                            t_new.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                            t_new.text = f"{seq_counters[seq_key]}."
                            r_el.remove(fc)
                            r_el.append(t_new)
                            runs_to_remove.remove(r_el)
                        for rem_r in runs_to_remove:
                            if rem_r.getparent() is p_elem:
                                p_elem.remove(rem_r)
                        converted += 1
                    in_field = False
                    is_num_field = False
                    in_result = False
            continue

        if in_field:
            if instr_els and not in_result:
                field_instr += "".join((ie.text or "") for ie in instr_els)
                runs_to_remove.append(r_el)
            elif in_result:
                result_runs.append(r_el)
            else:
                runs_to_remove.append(r_el)

    return converted


def convert_numbering_to_text_in_doc(doc: DocxDocument) -> int:
    """
    Chuyển toàn bộ đánh số tự động (Word Automatic Numbering `<w:numPr>`, Style Numbering
    và trường `SEQ`/`AUTONUM`/`LISTNUM`) trong tài liệu `doc` thành văn bản `<w:r><w:t>` thực sự.
    Trả về tổng số đoạn văn đã được chuyển từ Numbering sang Text.
    """
    body = doc.element.body
    converted_count = 0
    seq_counters: Dict[str, int] = {}

    # Bước 1: Làm phẳng các trường SEQ / AUTONUM / LISTNUM nếu có
    for p_el in body.iter(qn("w:p")):
        converted_count += _flatten_numbering_fields_in_paragraph(p_el, seq_counters)

    # Bước 2: Đọc cấu trúc đánh số trong styles.xml
    style_map: Dict[str, Dict] = {}
    styles_root = doc.styles.element if doc.styles is not None else None
    if styles_root is not None:
        for st_el in styles_root.findall(qn("w:style")):
            st_id = st_el.get(qn("w:styleId")) or st_el.get("styleId")
            if not st_id:
                continue
            based_el = st_el.find(qn("w:basedOn"))
            based_id = (based_el.get(qn("w:val")) or based_el.get("val")) if based_el is not None else None
            st_num_id: Optional[str] = None
            st_ilvl: Optional[int] = None
            st_pPr = st_el.find(qn("w:pPr"))
            if st_pPr is not None:
                st_numPr = st_pPr.find(qn("w:numPr"))
                if st_numPr is not None:
                    nid_el = st_numPr.find(qn("w:numId"))
                    ilvl_el = st_numPr.find(qn("w:ilvl"))
                    if nid_el is not None:
                        st_num_id = nid_el.get(qn("w:val")) or nid_el.get("val")
                    if ilvl_el is not None:
                        try:
                            st_ilvl = int(ilvl_el.get(qn("w:val")) or ilvl_el.get("val") or "0")
                        except ValueError:
                            st_ilvl = 0
            style_map[st_id] = {
                "basedOn": based_id,
                "numId": st_num_id,
                "ilvl": st_ilvl,
                "rPr": st_el.find(qn("w:rPr")),
            }

    def resolve_style_numbering(style_id: Optional[str]) -> Tuple[Optional[str], Optional[int]]:
        curr = style_id
        visited = set()
        found_num_id: Optional[str] = None
        found_ilvl: Optional[int] = None
        while curr and curr in style_map and curr not in visited:
            visited.add(curr)
            info = style_map[curr]
            if found_num_id is None and info["numId"] is not None:
                found_num_id = info["numId"]
            if found_ilvl is None and info["ilvl"] is not None:
                found_ilvl = info["ilvl"]
            if found_num_id is not None and found_ilvl is not None:
                break
            curr = info["basedOn"]
        return found_num_id, found_ilvl

    # Bước 3: Tìm numbering_part qua rels (tránh tạo part giả nếu file không có numbering.xml)
    numbering_root: Optional[OxmlElement] = None
    for rel in doc.part.rels.values():
        if rel.reltype.endswith("/numbering"):
            try:
                numbering_root = rel.target_part.element
            except Exception:
                numbering_root = None
            break

    if numbering_root is None:
        return converted_count

    def parse_lvl_element(lvl_el: OxmlElement, is_override: bool = False) -> Dict:
        ilvl_str = lvl_el.get(qn("w:ilvl")) or lvl_el.get("ilvl") or "0"
        try:
            ilvl_val = int(ilvl_str)
        except ValueError:
            ilvl_val = 0

        start_el = lvl_el.find(qn("w:start"))
        start_val = 1
        if start_el is not None:
            try:
                start_val = int(start_el.get(qn("w:val")) or start_el.get("val") or "1")
            except ValueError:
                start_val = 1

        fmt_el = lvl_el.find(qn("w:numFmt"))
        num_fmt = (fmt_el.get(qn("w:val")) or fmt_el.get("val") or "decimal") if fmt_el is not None else "decimal"

        txt_el = lvl_el.find(qn("w:lvlText"))
        lvl_text = (txt_el.get(qn("w:val")) or txt_el.get("val") or "%1.") if txt_el is not None else "%1."

        pstyle_el = lvl_el.find(qn("w:pStyle"))
        pstyle_val = (pstyle_el.get(qn("w:val")) or pstyle_el.get("val")) if pstyle_el is not None else None

        return {
            "ilvl": ilvl_val,
            "start": start_val,
            "numFmt": num_fmt,
            "lvlText": lvl_text,
            "pStyle": pstyle_val,
            "rPr": lvl_el.find(qn("w:rPr")),
            "isOverride": is_override,
        }

    abstract_map: Dict[str, Dict] = {}
    for abs_el in numbering_root.findall(qn("w:abstractNum")):
        abs_id = abs_el.get(qn("w:abstractNumId")) or abs_el.get("abstractNumId")
        if not abs_id:
            continue
        ns_link_el = abs_el.find(qn("w:numStyleLink"))
        ns_link = (ns_link_el.get(qn("w:val")) or ns_link_el.get("val")) if ns_link_el is not None else None
        levels: Dict[int, Dict] = {}
        for lvl_el in abs_el.findall(qn("w:lvl")):
            l_info = parse_lvl_element(lvl_el, is_override=False)
            levels[l_info["ilvl"]] = l_info
            if l_info["pStyle"] and l_info["pStyle"] in style_map:
                if style_map[l_info["pStyle"]]["ilvl"] is None:
                    style_map[l_info["pStyle"]]["ilvl"] = l_info["ilvl"]
        abstract_map[abs_id] = {
            "numStyleLink": ns_link,
            "levels": levels,
        }

    num_map: Dict[str, Dict] = {}
    for num_el in numbering_root.findall(qn("w:num")):
        num_id = num_el.get(qn("w:numId")) or num_el.get("numId")
        if not num_id:
            continue
        abs_id_el = num_el.find(qn("w:abstractNumId"))
        abs_id = (abs_id_el.get(qn("w:val")) or abs_id_el.get("val")) if abs_id_el is not None else None
        overrides: Dict[int, Dict] = {}
        for ovr_el in num_el.findall(qn("w:lvlOverride")):
            o_ilvl_str = ovr_el.get(qn("w:ilvl")) or ovr_el.get("ilvl") or "0"
            try:
                o_ilvl = int(o_ilvl_str)
            except ValueError:
                o_ilvl = 0
            st_ovr_el = ovr_el.find(qn("w:startOverride"))
            st_ovr: Optional[int] = None
            if st_ovr_el is not None:
                try:
                    st_ovr = int(st_ovr_el.get(qn("w:val")) or st_ovr_el.get("val") or "1")
                except ValueError:
                    st_ovr = None
            lvl_sub = ovr_el.find(qn("w:lvl"))
            lvl_ovr = parse_lvl_element(lvl_sub, is_override=True) if lvl_sub is not None else None
            overrides[o_ilvl] = {
                "startOverride": st_ovr,
                "lvl": lvl_ovr,
            }
        num_map[num_id] = {
            "abstractNumId": abs_id,
            "overrides": overrides,
        }

    # Giải quyết liên kết numStyleLink giữa abstractNum -> style -> num -> abstractNum
    for abs_id, abs_info in abstract_map.items():
        if abs_info["numStyleLink"] and not abs_info["levels"]:
            linked_num_id, _ = resolve_style_numbering(abs_info["numStyleLink"])
            if linked_num_id and linked_num_id in num_map:
                target_abs_id = num_map[linked_num_id]["abstractNumId"]
                if target_abs_id and target_abs_id in abstract_map:
                    abs_info["levels"] = copy.deepcopy(abstract_map[target_abs_id]["levels"])
                    for l_info in abs_info["levels"].values():
                        if l_info["pStyle"] and l_info["pStyle"] in style_map:
                            if style_map[l_info["pStyle"]]["ilvl"] is None:
                                style_map[l_info["pStyle"]]["ilvl"] = l_info["ilvl"]

    def get_effective_lvl(num_id: str, ilvl: int) -> Tuple[Optional[Dict], Optional[int]]:
        n_info = num_map.get(num_id)
        if not n_info:
            return None, None
        abs_info = abstract_map.get(n_info["abstractNumId"] or "")
        ovr = n_info["overrides"].get(ilvl)
        start_ovr = ovr["startOverride"] if ovr else None
        if ovr and ovr["lvl"] is not None:
            return ovr["lvl"], start_ovr
        if abs_info:
            lvl_d = abs_info["levels"].get(ilvl) or abs_info["levels"].get(0)
            return lvl_d, start_ovr
        return None, start_ovr

    counters: Dict[Tuple[str, int], int] = {}
    initialized: set = set()

    for p_el in body.iter(qn("w:p")):
        pPr = p_el.find(qn("w:pPr"))
        if pPr is None:
            continue

        direct_num_id: Optional[str] = None
        direct_ilvl: Optional[int] = None
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            nid_el = numPr.find(qn("w:numId"))
            ilvl_el = numPr.find(qn("w:ilvl"))
            if nid_el is not None:
                direct_num_id = nid_el.get(qn("w:val")) or nid_el.get("val")
            if ilvl_el is not None:
                try:
                    direct_ilvl = int(ilvl_el.get(qn("w:val")) or ilvl_el.get("val") or "0")
                except ValueError:
                    direct_ilvl = 0
            pPr.remove(numPr)

        pstyle_el = pPr.find(qn("w:pStyle"))
        pstyle_id = (pstyle_el.get(qn("w:val")) or pstyle_el.get("val")) if pstyle_el is not None else None
        style_num_id, style_ilvl = resolve_style_numbering(pstyle_id)

        eff_num_id = direct_num_id if direct_num_id is not None else style_num_id
        eff_ilvl = direct_ilvl if direct_ilvl is not None else (style_ilvl if style_ilvl is not None else 0)

        if not eff_num_id or eff_num_id == "0" or eff_num_id not in num_map:
            continue

        existing_raw_txt = get_paragraph_text(p_el)
        existing_txt = existing_raw_txt.lstrip()
        has_media = (
            len(p_el.findall(".//" + qn("w:drawing"))) > 0
            or len(p_el.findall(".//" + qn("m:oMath"))) > 0
            or len(p_el.findall(".//" + qn("w:object"))) > 0
            or len(p_el.findall(".//" + qn("w:pict"))) > 0
            or len(p_el.findall(".//" + qn("w:sym"))) > 0
        )
        lvl_info, start_ovr = get_effective_lvl(eff_num_id, eff_ilvl)
        if not lvl_info:
            continue

        num_fmt = lvl_info["numFmt"]
        if num_fmt in ("bullet", "none"):
            continue

        is_question_lvl = (
            bool(re.search(r"(?:Câu|Question|Bài)", lvl_info["lvlText"] or "", flags=re.IGNORECASE))
            or (num_fmt in ("decimal", "decimalZero"))
        )
        if not existing_txt and not has_media and not is_question_lvl:
            continue

        for k_num, k_lvl in list(initialized):
            if k_num == eff_num_id and k_lvl > eff_ilvl:
                initialized.remove((k_num, k_lvl))

        key = (eff_num_id, eff_ilvl)
        if key not in initialized:
            start_v = start_ovr if start_ovr is not None else lvl_info["start"]
            counters[key] = start_v
            initialized.add(key)
        else:
            counters[key] += 1

        curr_val = counters[key]
        raw_label = lvl_info["lvlText"] or "%1."
        for k in range(1, 10):
            token = f"%{k}"
            if token in raw_label:
                sub_lvl, _ = get_effective_lvl(eff_num_id, k - 1)
                sub_start = sub_lvl["start"] if sub_lvl else 1
                sub_fmt = sub_lvl["numFmt"] if sub_lvl else "decimal"
                sub_val = counters.get((eff_num_id, k - 1), sub_start)
                raw_label = raw_label.replace(token, _format_numbering_counter(sub_val, sub_fmt))

        label_str = raw_label.strip()
        if not label_str:
            continue

        prefix_str = ""
        if re.match(r"^(?:Câu|Question|Bài)\s*\d+$", label_str, flags=re.IGNORECASE):
            if QUESTION_START_RE.match(existing_txt):
                prefix_str = ""
            elif existing_txt[:1] in ".:)":
                prefix_str = label_str
            else:
                prefix_str = f"{label_str}. "
        elif re.match(r"^(?:Câu|Question|Bài)\s*\d+\s*(?:\([^)]*\))?\s*[\.\:\)]$", label_str, flags=re.IGNORECASE):
            if QUESTION_START_RE.match(existing_txt):
                prefix_str = ""
            elif existing_txt[:1] in ".:)":
                prefix_str = re.sub(r"[\.\:\)]\s*$", "", label_str)
            else:
                prefix_str = f"{label_str} "
        elif num_fmt in ("decimal", "decimalZero") and re.match(r"^\d+[\.\:\)]?$", label_str):
            if QUESTION_START_RE.match(existing_txt):
                prefix_str = ""
            elif re.match(r"^PHẦN\b", existing_txt, flags=re.IGNORECASE):
                prefix_str = f"{label_str} "
            elif re.match(r"^(?:Câu|Question|Bài)\s*$", existing_txt, flags=re.IGNORECASE):
                prefix_str = f"{label_str} "
            else:
                punct = label_str[-1] if label_str[-1] in ".:)" else "."
                if existing_txt[:1] in ".:)":
                    prefix_str = f"Câu {curr_val}"
                else:
                    prefix_str = f"Câu {curr_val}{punct} "
        elif num_fmt == "upperLetter" and re.match(r"^\*?[A-D][\.\:\)]?$", label_str, flags=re.IGNORECASE):
            if re.match(r"^\*?\s*[A-D]\s*[\.\:\)]", existing_txt):
                prefix_str = ""
            elif re.match(r"^\*?[A-D]$", label_str, flags=re.IGNORECASE):
                prefix_str = label_str if existing_txt[:1] in ".:)" else f"{label_str}. "
            else:
                prefix_str = f"{label_str} "
        elif num_fmt == "lowerLetter" and re.match(r"^\*?[a-d][\)\.\:]?$", label_str, flags=re.IGNORECASE):
            if re.match(r"^\*?\s*[a-d]\s*[\)\.\:]", existing_txt):
                prefix_str = ""
            elif re.match(r"^\*?[a-d]$", label_str, flags=re.IGNORECASE):
                prefix_str = label_str if existing_txt[:1] in ").:" else f"{label_str}) "
            else:
                prefix_str = f"{label_str} "
        elif num_fmt == "upperRoman" and re.match(r"^(?:I|II|III)[\.\:\)]?$", label_str, flags=re.IGNORECASE):
            if re.match(r"^PHẦN\b", existing_txt, flags=re.IGNORECASE):
                prefix_str = ""
            elif re.search(r"trắc\s*nghiệm|nhiều\s*phương\s*án|đúng\s*sai|trả\s*lời\s*ngắn", existing_txt, flags=re.IGNORECASE):
                punct = label_str[-1] if label_str[-1] in ".:)" else "."
                rom = re.sub(r"[\.\:\)]$", "", label_str).upper()
                prefix_str = f"PHẦN {rom}{punct} "
            else:
                prefix_str = f"{label_str} "
        else:
            prefix_str = f"{label_str} "

        pPr_rPr = pPr.find(qn("w:rPr"))
        lvl_rPr = lvl_info["rPr"]
        marker_is_ul = is_rPr_underlined(pPr_rPr) or (lvl_info["isOverride"] and is_rPr_underlined(lvl_rPr))
        marker_is_bold = (
            is_rPr_bold(pPr_rPr)
            or is_rPr_bold(lvl_rPr)
            or prefix_str.upper().startswith("CÂU")
            or prefix_str.upper().startswith("PHẦN")
        )

        # Nếu đoạn văn dùng gạch chân ở <w:pPr><w:rPr> để gạch chân nhãn tự động (A. hoặc a)),
        # xóa gạch chân khỏi <w:pPr><w:rPr> sau khi đã chuyển sang run nhãn để không làm gạch chân cả câu
        if pPr_rPr is not None and is_rPr_underlined(pPr_rPr):
            remove_underline_from_rPr(pPr_rPr)

        if prefix_str:
            r_num = OxmlElement("w:r")
            rPr_num = OxmlElement("w:rPr")
            if marker_is_bold:
                rPr_num.append(OxmlElement("w:b"))
            if marker_is_ul and not (prefix_str.upper().startswith("CÂU") or prefix_str.upper().startswith("PHẦN")):
                u_el = OxmlElement("w:u")
                u_el.set(qn("w:val"), "single")
                rPr_num.append(u_el)
            r_num.append(rPr_num)
            t_num = OxmlElement("w:t")
            t_num.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t_num.text = prefix_str
            r_num.append(t_num)
            pPr.addnext(r_num)
            converted_count += 1
        elif marker_is_ul:
            # Nếu đoạn văn đã có sẵn nhãn A. trong text nhưng gạch chân nằm ở <w:pPr><w:rPr>
            first_r = p_el.find(qn("w:r"))
            if first_r is not None:
                f_rPr = first_r.find(qn("w:rPr"))
                if f_rPr is None:
                    f_rPr = OxmlElement("w:rPr")
                    first_r.insert(0, f_rPr)
                if not is_rPr_underlined(f_rPr):
                    u_el = OxmlElement("w:u")
                    u_el.set(qn("w:val"), "single")
                    f_rPr.append(u_el)

    # Bước 4: Xóa toàn bộ <w:numPr> trong styles.xml để khi xuất đề Word không bị đánh số tự động chồng lên
    if styles_root is not None:
        for st_el in styles_root.findall(qn("w:style")):
            st_pPr = st_el.find(qn("w:pPr"))
            if st_pPr is not None:
                for st_numPr in list(st_pPr.findall(qn("w:numPr"))):
                    st_pPr.remove(st_numPr)

    return converted_count


def parse_docx_exam(docx_bytes: bytes) -> ParsedExam:
    doc = docx.Document(io.BytesIO(docx_bytes))
    converted_num_count = convert_numbering_to_text_in_doc(doc)
    parsed = ParsedExam(source_doc=doc)

    body = doc.element.body
    current_part = 1
    current_q: Optional[QuestionBlock] = None

    def finalize_question(q: Optional[QuestionBlock]) -> None:
        if q is None:
            return
        has_stem_content = any(
            (
                el.tag == qn("w:tbl")
                or bool(get_paragraph_text(el).strip())
                or len(el.findall(".//" + qn("w:drawing"))) > 0
                or len(el.findall(".//" + qn("m:oMath"))) > 0
                or len(el.findall(".//" + qn("w:object"))) > 0
                or len(el.findall(".//" + qn("w:pict"))) > 0
                or len(el.findall(".//" + qn("w:sym"))) > 0
            )
            for el in q.stem_elements
        )
        if not has_stem_content and not q.options and not q.short_answer:
            q_counters[q.part] = max(0, q_counters[q.part] - 1)
            return
        if q.part == 1:
            if q.options and not any(opt.is_correct for opt in q.options):
                q.options[0].is_correct = True
                parsed.warnings.append(
                    f"Phần I - Câu {q.original_index}: Chưa gạch chân đáp án đúng, hệ thống tạm lấy phương án A."
                )
            for opt in q.options:
                if len(opt.xml_elements) > 1 and all(el.tag == qn("w:p") for el in opt.xml_elements):
                    non_empty_paras = [
                        el for el in opt.xml_elements
                        if _extract_trimmed_option_content_nodes(el)
                    ]
                    if len(non_empty_paras) == 1:
                        opt.xml_elements = [non_empty_paras[0]]
                    elif len(non_empty_paras) > 1:
                        base_p = copy.deepcopy(non_empty_paras[0])
                        _trim_option_paragraph_in_place(base_p)
                        for extra_p in non_empty_paras[1:]:
                            extra_nodes = _extract_trimmed_option_content_nodes(extra_p)
                            if extra_nodes:
                                r_sp = OxmlElement("w:r")
                                t_sp = OxmlElement("w:t")
                                t_sp.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                                t_sp.text = " "
                                r_sp.append(t_sp)
                                base_p.append(r_sp)
                                for en in extra_nodes:
                                    base_p.append(en)
                        opt.xml_elements = [base_p]
                opt_txt = "".join(get_paragraph_text(el) for el in opt.xml_elements)
                if FIXED_OPTION_RE.search(opt_txt):
                    opt.fixed_order = True
            parsed.part1_questions.append(q)
        elif q.part == 2:
            for opt in q.options:
                if opt.xml_elements:
                    last_p = opt.xml_elements[-1]
                    txt = get_paragraph_text(last_p).strip()
                    m_ds = re.search(r"[\(\[](Đ|Đúng|S|Sai)[\)\]]\s*\.?\s*$", txt, flags=re.IGNORECASE)
                    if m_ds:
                        val = m_ds.group(1).upper()
                        opt.is_correct = val.startswith("Đ")
                        for r_el in reversed(last_p.findall(qn("w:r"))):
                            for t_el in reversed(r_el.findall(qn("w:t"))):
                                if t_el.text and re.search(r"[\(\[](Đ|Đúng|S|Sai)[\)\]]", t_el.text, flags=re.IGNORECASE):
                                    t_el.text = re.sub(
                                        r"\s*[\(\[](Đ|Đúng|S|Sai)[\)\]]\s*\.?\s*$",
                                        "",
                                        t_el.text,
                                        flags=re.IGNORECASE,
                                    )
                                    break
            parsed.part2_questions.append(q)
        elif q.part == 3:
            if not q.short_answer:
                underlined_chunks = []
                for el in q.stem_elements:
                    if el.tag == qn("w:p"):
                        for r_el in el.findall(qn("w:r")):
                            if is_run_underlined(r_el, el):
                                chunk = "".join((t.text or "") for t in r_el.findall(qn("w:t"))).strip()
                                if chunk:
                                    underlined_chunks.append(chunk)
                                remove_underline_from_run(r_el)
                if underlined_chunks:
                    q.short_answer = "".join(underlined_chunks)
            parsed.part3_questions.append(q)

    q_counters = {1: 0, 2: 0, 3: 0}

    for raw_child in body:
        if raw_child.tag == qn("w:sectPr"):
            continue

        if raw_child.tag == qn("w:p"):
            sub_paragraphs = split_paragraph_on_line_breaks_before_options(raw_child, part=current_part)
            for child in sub_paragraphs:
                text = get_paragraph_text(child).strip()
                has_media = (
                    len(child.findall(".//" + qn("w:drawing"))) > 0
                    or len(child.findall(".//" + qn("m:oMath"))) > 0
                    or len(child.findall(".//" + qn("w:object"))) > 0
                    or len(child.findall(".//" + qn("w:pict"))) > 0
                    or len(child.findall(".//" + qn("w:sym"))) > 0
                )
                if not text and not has_media:
                    continue

                if END_MARKER_RE.match(text):
                    finalize_question(current_q)
                    current_q = None
                    return parsed

                if PART1_HEADER_RE.match(text):
                    finalize_question(current_q)
                    current_q = None
                    current_part = 1
                    continue
                if PART2_HEADER_RE.match(text):
                    finalize_question(current_q)
                    current_q = None
                    current_part = 2
                    continue
                if PART3_HEADER_RE.match(text):
                    finalize_question(current_q)
                    current_q = None
                    current_part = 3
                    continue

                m_q = QUESTION_START_RE.match(text)
                if m_q:
                    finalize_question(current_q)
                    q_counters[current_part] += 1
                    stripped_p, _, _ = strip_prefix_from_paragraph(
                        child,
                        r"^\s*(?:Câu|Question|Bài)\s*\d+\s*(?:\([^)]*\))?\s*[\.\:\)]\s*",
                    )
                    if current_part in (1, 2):
                        clean_stem_p, trailing_opts = split_stem_and_trailing_options(stripped_p, part=current_part)
                        current_q = QuestionBlock(
                            part=current_part,
                            original_index=q_counters[current_part],
                            stem_elements=[clean_stem_p],
                        )
                        if trailing_opts:
                            for lbl, is_corr, opt_p in trailing_opts:
                                current_q.options.append(
                                    OptionBlock(
                                        original_label=lbl,
                                        is_correct=is_corr,
                                        xml_elements=[opt_p],
                                    )
                                )
                    else:
                        current_q = QuestionBlock(
                            part=current_part,
                            original_index=q_counters[current_part],
                            stem_elements=[stripped_p],
                        )
                    continue

                if current_q is None:
                    continue

                if current_part in (1, 2):
                    inline_opts = split_inline_options_in_paragraph(child, part=current_part)
                    if inline_opts:
                        first_lbl = inline_opts[0][0]
                        expected_first = "A" if current_part == 1 else "a"
                        has_complete_opts = len(current_q.options) >= 4 or any(
                            o.original_label.lower() == "d" for o in current_q.options
                        )
                        if first_lbl == expected_first and has_complete_opts:
                            finalize_question(current_q)
                            q_counters[current_part] += 1
                            empty_stem_p = OxmlElement("w:p")
                            orig_pPr = child.find(qn("w:pPr"))
                            if orig_pPr is not None:
                                empty_stem_p.append(copy.deepcopy(orig_pPr))
                            current_q = QuestionBlock(
                                part=current_part,
                                original_index=q_counters[current_part],
                                stem_elements=[empty_stem_p],
                            )
                        for lbl, is_corr, opt_p in inline_opts:
                            current_q.options.append(
                                OptionBlock(
                                    original_label=lbl,
                                    is_correct=is_corr,
                                    xml_elements=[opt_p],
                                )
                            )
                        continue

                    if not current_q.options:
                        clean_stem_p, trailing_opts = split_stem_and_trailing_options(child, part=current_part)
                        if (
                            len(current_q.stem_elements) == 1
                            and current_q.stem_elements[0].tag == qn("w:p")
                            and not get_paragraph_text(current_q.stem_elements[0]).strip()
                            and not _extract_trimmed_option_content_nodes(current_q.stem_elements[0])
                        ):
                            current_q.stem_elements[0] = copy.deepcopy(clean_stem_p)
                        else:
                            current_q.stem_elements.append(copy.deepcopy(clean_stem_p))
                        if trailing_opts:
                            for lbl, is_corr, opt_p in trailing_opts:
                                current_q.options.append(
                                    OptionBlock(
                                        original_label=lbl,
                                        is_correct=is_corr,
                                        xml_elements=[opt_p],
                                    )
                                )
                    else:
                        opt_child = copy.deepcopy(child)
                        clean_option_paragraph_underlines(opt_child)
                        current_q.options[-1].xml_elements.append(opt_child)

                elif current_part == 3:
                    m_ans = SHORT_ANS_LINE_RE.match(text)
                    if m_ans:
                        current_q.short_answer = m_ans.group("ans").strip().rstrip(".")
                    else:
                        current_q.stem_elements.append(copy.deepcopy(child))

        elif raw_child.tag == qn("w:tbl"):
            if current_q is not None:
                if current_q.part in (1, 2) and current_q.options:
                    current_q.options[-1].xml_elements.append(copy.deepcopy(raw_child))
                else:
                    current_q.stem_elements.append(copy.deepcopy(raw_child))

    finalize_question(current_q)
    return parsed


# ---------------------------------------------------------------------------
# Bộ sinh mã đề đã trộn (.docx) và Bảng đáp án chấm thi (.xlsx)
# ---------------------------------------------------------------------------

@dataclass
class ExamMetadata:
    """Thông tin tiêu đề đề thi in ở đầu mỗi mã đề."""
    governing_body: str = "SỞ GD & ĐT BẮC NINH"
    school_name: str = "TRƯỜNG THPT HIỆP HÒA SỐ 3"
    exam_name: str = "ĐỀ KIỂM TRA ĐỊNH KỲ NĂM HỌC 2026 - 2027"
    subject_name: str = "MÔN: TOÁN 12"
    duration_minutes: str = "90 phút; không kể thời gian phát đề"
    shuffle_part1_questions: bool = True
    shuffle_part1_options: bool = True
    shuffle_part2_questions: bool = True
    shuffle_part2_options: bool = True
    shuffle_part3_questions: bool = True


@dataclass
class ShuffledVariantResult:
    """Kết quả trộn của 1 mã đề."""
    exam_code: str
    docx_bytes: bytes
    part1_keys: List[str]
    part2_keys: List[str]
    part2_detail: List[Dict[str, str]]
    part3_keys: List[str]
    question_mapping: Dict[str, List[int]]


def _set_cell_margins(cell, top=20, bottom=20, left=40, right=40):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _remove_table_borders(table) -> None:
    """Xóa toàn bộ khung viền của bảng tiêu đề 2 cột."""
    tblPr = table._tbl.tblPr
    for old_b in list(tblPr.findall(qn("w:tblBorders"))):
        tblPr.remove(old_b)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b_el = OxmlElement(f"w:{side}")
        b_el.set(qn("w:val"), "none")
        b_el.set(qn("w:sz"), "0")
        b_el.set(qn("w:space"), "0")
        b_el.set(qn("w:color"), "auto")
        borders.append(b_el)
    tblPr.append(borders)


def _format_run(
    run,
    font_name="Times New Roman",
    size_pt=12,
    bold=False,
    italic=False,
    underline=False,
    boxed=False,
):
    run.font.name = font_name
    rPr = run._element.get_or_add_rPr()
    rPr.rFonts.set(qn("w:eastAsia"), font_name)
    rPr.rFonts.set(qn("w:cs"), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    if boxed:
        bdr = OxmlElement("w:bdr")
        bdr.set(qn("w:val"), "single")
        bdr.set(qn("w:sz"), "4")
        bdr.set(qn("w:space"), "1")
        bdr.set(qn("w:color"), "000000")
        rPr.append(bdr)


def _append_word_field_to_paragraph(
    p,
    field_name: str,
    default_val: str = "1",
    font_name: str = "Times New Roman",
    size_pt: float = 12,
    italic: bool = False,
    bold: bool = False,
) -> None:
    """Chèn một trường động của Word (như PAGE hoặc NUMPAGES) với cờ w:dirty=true để Word tự động cập nhật."""
    r_begin = p.add_run()
    _format_run(r_begin, font_name=font_name, size_pt=size_pt, italic=italic, bold=bold)
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    fld_begin.set(qn("w:dirty"), "true")
    r_begin._element.append(fld_begin)

    r_instr = p.add_run()
    _format_run(r_instr, font_name=font_name, size_pt=size_pt, italic=italic, bold=bold)
    instr = OxmlElement("w:instrText")
    instr.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    instr.text = f" {field_name} "
    r_instr._element.append(instr)

    r_sep = p.add_run()
    _format_run(r_sep, font_name=font_name, size_pt=size_pt, italic=italic, bold=bold)
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    r_sep._element.append(fld_sep)

    r_val = p.add_run(default_val)
    _format_run(r_val, font_name=font_name, size_pt=size_pt, italic=italic, bold=bold)

    r_end = p.add_run()
    _format_run(r_end, font_name=font_name, size_pt=size_pt, italic=italic, bold=bold)
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    r_end._element.append(fld_end)


def _append_numpages_field_to_paragraph(p, default_pages: str = "4", font_name="Times New Roman", size_pt=12, italic=True):
    """Chèn trường tự động đếm tổng số trang của Word: 'Đề gồm: <NUMPAGES> trang'."""
    r1 = p.add_run("Đề gồm: ")
    _format_run(r1, font_name=font_name, size_pt=size_pt, italic=italic)
    _append_word_field_to_paragraph(p, "NUMPAGES", default_val=default_pages, font_name=font_name, size_pt=size_pt, italic=italic)
    r2 = p.add_run(" trang")
    _format_run(r2, font_name=font_name, size_pt=size_pt, italic=italic)


def _enable_auto_update_fields_in_settings(doc: DocxDocument) -> None:
    """Bật cờ <w:updateFields w:val="true"/> trong word/settings.xml để Word luôn cập nhật số trang NUMPAGES."""
    try:
        settings_el = doc.settings.element
        if settings_el is not None:
            upd = settings_el.find(qn("w:updateFields"))
            if upd is None:
                upd = OxmlElement("w:updateFields")
                settings_el.append(upd)
            upd.set(qn("w:val"), "true")
    except Exception:
        pass


def _estimate_body_page_count(body_el: OxmlElement, source_docx_bytes: Optional[bytes] = None) -> int:
    """
    Ước lượng chính xác số trang của đề xuất ra dựa trên tổng chiều cao (pt) của các đoạn văn,
    công thức MathType/OMML, bảng biểu và hình vẽ trên khổ A4 (lề trên 11mm, lề dưới 11mm).
    """
    import math
    total_height_pt = 0.0
    for child in body_el:
        if child.tag == qn("w:sectPr"):
            continue
        if child.tag == qn("w:tbl"):
            rows = child.findall(".//" + qn("w:tr"))
            total_height_pt += max(1, len(rows)) * 52.0 + 10.0
            continue
        if child.tag == qn("w:p"):
            vis_chars = _measure_xml_node_visual_chars(child)
            br_count = len(child.findall(".//" + qn("w:br"))) + len(child.findall(".//" + qn("w:cr")))
            lines = max(1, int(math.ceil(vis_chars / 82.0)) + br_count)
            inline_extras: List[float] = []
            block_extra = 0.0
            for sub in child.iter():
                s_ln = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
                if s_ln == "extent":
                    cy_s = sub.get("cy")
                    if cy_s:
                        try:
                            h_pt = float(cy_s) / 12700.0
                            if h_pt > 45.0:
                                block_extra += h_pt + 8.0
                            elif h_pt > 15.0:
                                inline_extras.append(h_pt - 13.0)
                        except ValueError:
                            pass
                elif s_ln == "shape":
                    st_s = sub.get("style") or ""
                    m_h = re.search(r"(?:^|;|\s)height\s*:\s*([\d\.,]+)\s*(pt|in|cm|mm|px)\b", st_s, flags=re.IGNORECASE)
                    if m_h:
                        h_pt = _convert_css_dim_to_pt(m_h.group(1), m_h.group(2))
                        if h_pt > 45.0:
                            block_extra += h_pt + 8.0
                        elif h_pt > 15.0:
                            inline_extras.append(h_pt - 13.0)
                elif s_ln == "oMath":
                    has_tall_math = any(
                        (m_sub.tag.split("}")[-1] if "}" in m_sub.tag else m_sub.tag) in ("f", "rad", "nary", "eqArr")
                        for m_sub in sub.iter()
                    )
                    inline_extras.append(14.5 if has_tall_math else 4.5)
            inline_extras.sort(reverse=True)
            top_inline_extra = sum(inline_extras[: max(1, lines)])
            total_height_pt += lines * 16.5 + 4.0 + block_extra + min(lines * 18.0, top_inline_extra)

    est_pages = max(1, int(math.ceil(total_height_pt / 710.0)))
    return est_pages


def _configure_section_headers_and_footers(doc: DocxDocument, exam_code: str) -> None:
    """
    Đồng bộ toàn bộ Header và Footer của tài liệu với mã đề `exam_code`:
    - Xóa nội dung header cũ (nếu đề gốc có header thừa).
    - Đặt nội dung Footer chuẩn: 'Trang <PAGE>/<NUMPAGES> - Mã đề thi <exam_code>'
      cho tất cả các loại footer (footer chính, first_page_footer, even_page_footer)
      mà vẫn giữ nguyên căn lề `<w:pPr>` của footer gốc (hoặc căn giữa nếu chưa có).
    """
    for section in doc.sections:
        section.different_first_page_header_footer = False

        # Làm sạch các header cũ nếu có
        for hdr in (section.header, section.first_page_header, section.even_page_header):
            if hdr is not None and not hdr.is_linked_to_previous:
                for p in list(hdr.paragraphs):
                    for r in list(p._element.findall(qn("w:r"))):
                        p._element.remove(r)

        # Cập nhật tất cả các footer để luôn hiển thị đúng: "Trang X/Y - Mã đề thi <exam_code>"
        for ftr in (section.footer, section.first_page_footer, section.even_page_footer):
            if ftr is None:
                continue
            ftr.is_linked_to_previous = False

            # Lưu lại <w:pPr> của đoạn văn đầu tiên trong footer gốc (nếu có) để giữ nguyên vị trí căn lề
            saved_pPr: Optional[OxmlElement] = None
            if ftr.paragraphs:
                orig_pPr = ftr.paragraphs[0]._element.find(qn("w:pPr"))
                if orig_pPr is not None:
                    saved_pPr = copy.deepcopy(orig_pPr)

            # Xóa toàn bộ các phần tử cũ trong footer (kể cả bảng hay đoạn văn chứa mã đề cũ "1111")
            for child in list(ftr._element):
                ftr._element.remove(child)

            p_ftr = ftr.add_paragraph()
            if saved_pPr is not None:
                old_pPr = p_ftr._element.find(qn("w:pPr"))
                if old_pPr is not None:
                    p_ftr._element.remove(old_pPr)
                p_ftr._element.insert(0, saved_pPr)
            else:
                p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER

            p_ftr.paragraph_format.space_before = Pt(0)
            p_ftr.paragraph_format.space_after = Pt(0)

            r_t1 = p_ftr.add_run("Trang ")
            _format_run(r_t1, size_pt=12)
            _append_word_field_to_paragraph(p_ftr, "PAGE", default_val="1", size_pt=12)
            r_slash = p_ftr.add_run("/")
            _format_run(r_slash, size_pt=12)
            _append_word_field_to_paragraph(p_ftr, "NUMPAGES", default_val="4", size_pt=12)
            r_code = p_ftr.add_run(f" - Mã đề thi {exam_code}")
            _format_run(r_code, size_pt=12)


def generate_single_variant_docx(
    source_docx_bytes: bytes,
    parsed: ParsedExam,
    exam_code: str,
    meta: ExamMetadata,
    rng: random.Random,
) -> ShuffledVariantResult:
    doc = docx.Document(io.BytesIO(source_docx_bytes))
    convert_numbering_to_text_in_doc(doc)
    body = doc.element.body

    sectPr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not sectPr:
            body.remove(child)

    # Định dạng lề trang chuẩn theo yêu cầu:
    # top=11mm, bottom=11mm, left=10mm, right=8mm, header=6mm, footer=6mm (Khổ A4 210x297mm)
    for section in doc.sections:
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(11)
        section.bottom_margin = Mm(11)
        section.left_margin = Mm(10)
        section.right_margin = Mm(8)
        section.header_distance = Mm(6)
        section.footer_distance = Mm(6)

    # Cập nhật Footer theo đúng mã đề hiện tại (xóa bỏ mã đề cũ như "1111" trong đề gốc)
    _configure_section_headers_and_footers(doc, exam_code)

    # Bảng tiêu đề 2 cột không viền chuẩn mẫu
    header_tbl = doc.add_table(rows=1, cols=2)
    header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_tbl.autofit = False
    _remove_table_borders(header_tbl)

    cell_left = header_tbl.cell(0, 0)
    cell_right = header_tbl.cell(0, 1)
    cell_left.width = Mm(76)
    cell_right.width = Mm(116)
    _set_cell_margins(cell_left, 10, 10, 20, 20)
    _set_cell_margins(cell_right, 10, 10, 20, 20)

    p_left = cell_left.paragraphs[0]
    p_left.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left.paragraph_format.space_before = Pt(0)
    p_left.paragraph_format.space_after = Pt(0)
    p_left.paragraph_format.line_spacing = 1.15
    r_gov = p_left.add_run(meta.governing_body.upper() + "\n")
    _format_run(r_gov, size_pt=12, bold=False)
    r_sch = p_left.add_run(meta.school_name.upper() + "\n")
    _format_run(r_sch, size_pt=12, bold=True, underline=False)
    _append_numpages_field_to_paragraph(p_left, size_pt=12, italic=True)

    p_right = cell_right.paragraphs[0]
    p_right.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right.paragraph_format.space_before = Pt(0)
    p_right.paragraph_format.space_after = Pt(0)
    p_right.paragraph_format.line_spacing = 1.15
    r_exam = p_right.add_run(meta.exam_name.upper() + "\n")
    _format_run(r_exam, size_pt=12, bold=True)

    subj_display = meta.subject_name.strip()
    if subj_display and not subj_display.upper().startswith("MÔN"):
        subj_display = "MÔN: " + subj_display
    r_subj = p_right.add_run(subj_display.upper() + "\n")
    _format_run(r_subj, size_pt=12, bold=True)

    dur_display = meta.duration_minutes.strip()
    if not dur_display.lower().startswith("thời gian"):
        dur_display = f"Thời gian làm bài: {dur_display}"
    r_time = p_right.add_run(dur_display)
    _format_run(r_time, size_pt=12, italic=True)

    # Dòng Họ và tên thí sinh ... Số báo danh ... [MÃ ĐỀ: 1001] (đóng khung viền đơn)
    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(6)
    p_info.paragraph_format.space_after = Pt(2)
    p_info_pPr = p_info._element.get_or_add_pPr()
    tabs_info = OxmlElement("w:tabs")
    t_right = OxmlElement("w:tab")
    t_right.set(qn("w:val"), "right")
    t_right.set(qn("w:pos"), "10885")
    tabs_info.append(t_right)
    p_info_pPr.append(tabs_info)

    r_info = p_info.add_run(
        "Họ và tên thí sinh: ........................................................................ Số báo danh: ...........................\t"
    )
    _format_run(r_info, size_pt=12, bold=False)
    r_code_box = p_info.add_run(f"MÃ ĐỀ: {exam_code}")
    _format_run(r_code_box, size_pt=12, bold=True, boxed=True)

    def append_element_to_body(elem: OxmlElement) -> None:
        if sectPr is not None:
            sectPr.addprevious(elem)
        else:
            body.append(elem)

    part1_keys: List[str] = []
    part2_keys: List[str] = []
    part2_detail: List[Dict[str, str]] = []
    part3_keys: List[str] = []
    q_mapping: Dict[str, List[int]] = {"part1": [], "part2": [], "part3": []}

    # PHẦN I
    if parsed.part1_questions:
        p_h1 = doc.add_paragraph()
        p_h1.paragraph_format.space_before = Pt(2)
        p_h1.paragraph_format.space_after = Pt(2)
        r_h1 = p_h1.add_run(
            f"PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu {len(parsed.part1_questions)}. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
        )
        _format_run(r_h1, size_pt=12, bold=True)

        p1_list = list(parsed.part1_questions)
        if meta.shuffle_part1_questions:
            rng.shuffle(p1_list)

        letters_p1 = ["A", "B", "C", "D"]
        for idx, q in enumerate(p1_list, start=1):
            q_mapping["part1"].append(q.original_index)

            for s_idx, stem_el in enumerate(q.stem_elements):
                if s_idx == 0 and stem_el.tag == qn("w:p"):
                    new_stem = prepend_bold_label_to_paragraph(
                        stem_el,
                        f"Câu {idx}. ",
                        color_hex=None,
                        left_indent_dxa=425,
                        hanging_indent_dxa=425,
                    )
                    append_element_to_body(new_stem)
                else:
                    append_element_to_body(copy.deepcopy(stem_el))

            opts = list(q.options)
            if meta.shuffle_part1_options and len(opts) > 1:
                movable = [o for o in opts if not o.fixed_order]
                rng.shuffle(movable)
                it = iter(movable)
                opts = [o if o.fixed_order else next(it) for o in opts]

            correct_letter = "A"
            for o_idx, opt in enumerate(opts):
                letter = letters_p1[o_idx] if o_idx < len(letters_p1) else chr(ord("A") + o_idx)
                if opt.is_correct:
                    correct_letter = letter

            opt_paragraphs = build_part1_options_paragraphs(opts, letters_p1, label_color_hex=BLUE_HEX)
            for opt_p in opt_paragraphs:
                append_element_to_body(opt_p)

            part1_keys.append(correct_letter)

    # PHẦN II
    if parsed.part2_questions:
        p_h2 = doc.add_paragraph()
        p_h2.paragraph_format.space_before = Pt(6)
        p_h2.paragraph_format.space_after = Pt(2)
        r_h2 = p_h2.add_run(
            f"PHẦN II. Câu trắc nghiệm đúng sai. Thí sinh trả lời từ câu 1 đến câu {len(parsed.part2_questions)}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
        )
        _format_run(r_h2, size_pt=12, bold=True)

        p2_list = list(parsed.part2_questions)
        if meta.shuffle_part2_questions:
            rng.shuffle(p2_list)

        letters_p2 = ["a", "b", "c", "d"]
        for idx, q in enumerate(p2_list, start=1):
            q_mapping["part2"].append(q.original_index)

            for s_idx, stem_el in enumerate(q.stem_elements):
                if s_idx == 0 and stem_el.tag == qn("w:p"):
                    new_stem = prepend_bold_label_to_paragraph(
                        stem_el,
                        f"Câu {idx}. ",
                        color_hex=None,
                        left_indent_dxa=425,
                        hanging_indent_dxa=425,
                    )
                    append_element_to_body(new_stem)
                else:
                    append_element_to_body(copy.deepcopy(stem_el))

            opts = list(q.options)
            if meta.shuffle_part2_options and len(opts) > 1:
                rng.shuffle(opts)

            q_ds_list: List[str] = []
            q_ds_dict: Dict[str, str] = {}
            for o_idx, opt in enumerate(opts):
                letter = letters_p2[o_idx] if o_idx < len(letters_p2) else chr(ord("a") + o_idx)
                ds_val = "Đ" if opt.is_correct else "S"
                q_ds_list.append(ds_val)
                q_ds_dict[letter] = ds_val

                for el_idx, el in enumerate(opt.xml_elements):
                    if el_idx == 0 and el.tag == qn("w:p"):
                        labeled_opt = prepend_bold_label_to_paragraph(
                            el,
                            f"{letter}) ",
                            color_hex=BLUE_HEX,
                            left_indent_dxa=425,
                        )
                        append_element_to_body(labeled_opt)
                    else:
                        append_element_to_body(copy.deepcopy(el))

            part2_keys.append("-".join(q_ds_list))
            part2_detail.append(q_ds_dict)

    # PHẦN III
    if parsed.part3_questions:
        p_h3 = doc.add_paragraph()
        p_h3.paragraph_format.space_before = Pt(6)
        p_h3.paragraph_format.space_after = Pt(2)
        r_h3 = p_h3.add_run(
            f"PHẦN III. Câu trắc nghiệm trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu {len(parsed.part3_questions)}."
        )
        _format_run(r_h3, size_pt=12, bold=True)

        p3_list = list(parsed.part3_questions)
        if meta.shuffle_part3_questions:
            rng.shuffle(p3_list)

        for idx, q in enumerate(p3_list, start=1):
            q_mapping["part3"].append(q.original_index)
            for s_idx, stem_el in enumerate(q.stem_elements):
                if s_idx == 0 and stem_el.tag == qn("w:p"):
                    new_stem = prepend_bold_label_to_paragraph(
                        stem_el,
                        f"Câu {idx}. ",
                        color_hex=None,
                        left_indent_dxa=425,
                        hanging_indent_dxa=425,
                    )
                    append_element_to_body(new_stem)
                else:
                    append_element_to_body(copy.deepcopy(stem_el))
            part3_keys.append(q.short_answer)

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(8)
    r_end = p_end.add_run("-------------- HẾT --------------\n")
    _format_run(r_end, size_pt=11, bold=True)
    r_note = p_end.add_run("(Thí sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm)")
    _format_run(r_note, size_pt=10.5, italic=True)

    _enable_auto_update_fields_in_settings(doc)
    est_pages_str = str(_estimate_body_page_count(body, source_docx_bytes))
    # Cập nhật giá trị hiển thị mặc định của trường NUMPAGES trong bảng tiêu đề và footer
    for fld_instr in body.findall(".//" + qn("w:instrText")):
        if fld_instr.text and "NUMPAGES" in fld_instr.text:
            r_parent = fld_instr.getparent()
            if r_parent is not None:
                nxt = r_parent.getnext()
                while nxt is not None and nxt.tag == qn("w:r"):
                    if nxt.find(qn("w:fldChar")) is not None:
                        fc_el = nxt.find(qn("w:fldChar"))
                        if (fc_el.get(qn("w:fldCharType")) or fc_el.get("fldCharType")) == "end":
                            break
                    else:
                        t_el = nxt.find(qn("w:t"))
                        if t_el is not None:
                            t_el.text = est_pages_str
                    nxt = nxt.getnext()

    for section in doc.sections:
        for ftr in (section.footer, section.first_page_footer, section.even_page_footer):
            if ftr is not None:
                for fld_instr in ftr._element.findall(".//" + qn("w:instrText")):
                    if fld_instr.text and "NUMPAGES" in fld_instr.text:
                        r_parent = fld_instr.getparent()
                        if r_parent is not None:
                            nxt = r_parent.getnext()
                            while nxt is not None and nxt.tag == qn("w:r"):
                                if nxt.find(qn("w:fldChar")) is not None:
                                    fc_el = nxt.find(qn("w:fldChar"))
                                    if (fc_el.get(qn("w:fldCharType")) or fc_el.get("fldCharType")) == "end":
                                        break
                                else:
                                    t_el = nxt.find(qn("w:t"))
                                    if t_el is not None:
                                        t_el.text = est_pages_str
                                nxt = nxt.getnext()

    out_buf = io.BytesIO()
    doc.save(out_buf)

    return ShuffledVariantResult(
        exam_code=exam_code,
        docx_bytes=out_buf.getvalue(),
        part1_keys=part1_keys,
        part2_keys=part2_keys,
        part2_detail=part2_detail,
        part3_keys=part3_keys,
        question_mapping=q_mapping,
    )


def generate_answer_key_excel(
    variants: List[ShuffledVariantResult],
    meta: ExamMetadata,
) -> bytes:
    wb = openpyxl.Workbook()

    font_title = Font(name="Times New Roman", size=14, bold=True, color="1E3A8A")
    font_sub = Font(name="Times New Roman", size=11, italic=True, color="334155")
    font_part = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_header = Font(name="Times New Roman", size=11, bold=True, color="0F172A")
    font_cell = Font(name="Times New Roman", size=11, bold=False, color="0F172A")
    font_bold = Font(name="Times New Roman", size=11, bold=True, color="0F172A")

    fill_part = PatternFill("solid", fgColor="1E40AF")
    fill_header = PatternFill("solid", fgColor="DBEAFE")
    fill_zebra = PatternFill("solid", fgColor="F8FAFC")

    thin_side = Side(border_style="thin", color="CBD5E1")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    ws1 = wb.active
    ws1.title = "Bang_Dap_An_Tong_Hop"
    total_cols = max(2 + len(variants), 6)

    ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
    ws1.cell(1, 1, f"BẢNG ĐÁP ÁN CHẤM THI - {meta.exam_name.upper()}").font = font_title
    ws1.cell(1, 1).alignment = align_center

    ws1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_cols)
    ws1.cell(2, 1, f"{meta.governing_body} - {meta.school_name} | {meta.subject_name}").font = font_sub
    ws1.cell(2, 1).alignment = align_center

    row_cursor = 4

    def write_table_header(ws, row_idx: int, part_title: str) -> int:
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=1 + len(variants))
        c_part = ws.cell(row_idx, 1, part_title)
        c_part.font = font_part
        c_part.fill = fill_part
        c_part.alignment = align_left
        ws.row_dimensions[row_idx].height = 24

        row_idx += 1
        c_q = ws.cell(row_idx, 1, "Câu \\ Mã đề")
        c_q.font = font_header
        c_q.fill = fill_header
        c_q.border = border_all
        c_q.alignment = align_center

        for v_idx, v in enumerate(variants, start=2):
            c = ws.cell(row_idx, v_idx, f"Mã đề {v.exam_code}")
            c.font = font_header
            c.fill = fill_header
            c.border = border_all
            c.alignment = align_center
        ws.row_dimensions[row_idx].height = 22
        return row_idx + 1

    # Bảng 1: Phần I
    if variants and variants[0].part1_keys:
        row_cursor = write_table_header(
            ws1, row_cursor, "PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (A, B, C, D)"
        )
        num_q1 = len(variants[0].part1_keys)
        for q_i in range(num_q1):
            c_label = ws1.cell(row_cursor, 1, f"Câu {q_i + 1}")
            c_label.font = font_bold
            c_label.border = border_all
            c_label.alignment = align_center
            if q_i % 2 == 1:
                c_label.fill = fill_zebra

            for v_idx, v in enumerate(variants, start=2):
                ans_val = v.part1_keys[q_i] if q_i < len(v.part1_keys) else ""
                c_val = ws1.cell(row_cursor, v_idx, ans_val)
                c_val.font = font_bold
                c_val.border = border_all
                c_val.alignment = align_center
                if q_i % 2 == 1:
                    c_val.fill = fill_zebra
            ws1.row_dimensions[row_cursor].height = 20
            row_cursor += 1
        row_cursor += 1

    # Bảng 2: Phần II (Đúng/Sai)
    if variants and variants[0].part2_keys:
        row_cursor = write_table_header(
            ws1, row_cursor, "PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (Đ: Đúng, S: Sai theo thứ tự a - b - c - d)"
        )
        num_q2 = len(variants[0].part2_keys)
        for q_i in range(num_q2):
            for sub_idx, sub_letter in enumerate(["a", "b", "c", "d"]):
                c_label = ws1.cell(row_cursor, 1, f"Câu {q_i + 1}.{sub_letter})")
                c_label.font = font_cell
                c_label.border = border_all
                c_label.alignment = align_center
                if q_i % 2 == 1:
                    c_label.fill = fill_zebra

                for v_idx, v in enumerate(variants, start=2):
                    detail_map = v.part2_detail[q_i] if q_i < len(v.part2_detail) else {}
                    ans_val = detail_map.get(sub_letter, "")
                    c_val = ws1.cell(row_cursor, v_idx, ans_val)
                    c_val.font = font_bold
                    c_val.border = border_all
                    c_val.alignment = align_center
                    if q_i % 2 == 1:
                        c_val.fill = fill_zebra
                ws1.row_dimensions[row_cursor].height = 19
                row_cursor += 1
        row_cursor += 1

    # Bảng 3: Phần III (Trả lời ngắn)
    if variants and variants[0].part3_keys:
        row_cursor = write_table_header(
            ws1, row_cursor, "PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN"
        )
        num_q3 = len(variants[0].part3_keys)
        for q_i in range(num_q3):
            c_label = ws1.cell(row_cursor, 1, f"Câu {q_i + 1}")
            c_label.font = font_bold
            c_label.border = border_all
            c_label.alignment = align_center
            if q_i % 2 == 1:
                c_label.fill = fill_zebra

            for v_idx, v in enumerate(variants, start=2):
                ans_val = v.part3_keys[q_i] if q_i < len(v.part3_keys) else ""
                c_val = ws1.cell(row_cursor, v_idx, ans_val)
                c_val.font = font_bold
                c_val.border = border_all
                c_val.alignment = align_center
                if q_i % 2 == 1:
                    c_val.fill = fill_zebra
            ws1.row_dimensions[row_cursor].height = 20
            row_cursor += 1

    ws1.column_dimensions["A"].width = 16
    for col_i in range(2, 2 + len(variants)):
        ws1.column_dimensions[get_column_letter(col_i)].width = 15

    # Sheet 2: Dạng bảng phẳng phục vụ chấm máy
    ws2 = wb.create_sheet(title="Chi_Tiet_Cham_May")
    headers2 = ["Mã đề", "Phần", "Câu số", "Đáp án chuẩn", "Câu gốc tham chiếu"]
    for c_i, h_text in enumerate(headers2, start=1):
        cell = ws2.cell(1, c_i, h_text)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = border_all
        cell.alignment = align_center
    ws2.row_dimensions[1].height = 22

    r2_idx = 2
    for v in variants:
        for idx, ans in enumerate(v.part1_keys, start=1):
            orig_q = v.question_mapping["part1"][idx - 1]
            row_vals = [v.exam_code, "Phần I", f"Câu {idx}", ans, f"Câu {orig_q} (Phần I gốc)"]
            for c_i, val in enumerate(row_vals, start=1):
                c = ws2.cell(r2_idx, c_i, val)
                c.font = font_cell
                c.border = border_all
                c.alignment = align_center
            r2_idx += 1

        for idx, ans in enumerate(v.part2_keys, start=1):
            orig_q = v.question_mapping["part2"][idx - 1]
            row_vals = [v.exam_code, "Phần II", f"Câu {idx} (a-b-c-d)", ans, f"Câu {orig_q} (Phần II gốc)"]
            for c_i, val in enumerate(row_vals, start=1):
                c = ws2.cell(r2_idx, c_i, val)
                c.font = font_cell
                c.border = border_all
                c.alignment = align_center
            r2_idx += 1

        for idx, ans in enumerate(v.part3_keys, start=1):
            orig_q = v.question_mapping["part3"][idx - 1]
            row_vals = [v.exam_code, "Phần III", f"Câu {idx}", ans, f"Câu {orig_q} (Phần III gốc)"]
            for c_i, val in enumerate(row_vals, start=1):
                c = ws2.cell(r2_idx, c_i, val)
                c.font = font_cell
                c.border = border_all
                c.alignment = align_center
            r2_idx += 1

    for c_i, w in enumerate([12, 14, 18, 18, 24], start=1):
        ws2.column_dimensions[get_column_letter(c_i)].width = w

    out_excel = io.BytesIO()
    wb.save(out_excel)
    return out_excel.getvalue()


def create_shuffled_exam_package_zip(
    source_docx_bytes: bytes,
    exam_codes: List[str],
    meta: ExamMetadata,
    seed: Optional[int] = None,
) -> Tuple[bytes, ParsedExam, List[ShuffledVariantResult]]:
    parsed = parse_docx_exam(source_docx_bytes)
    total_q = len(parsed.part1_questions) + len(parsed.part2_questions) + len(parsed.part3_questions)
    if total_q == 0:
        raise ValueError(
            "Không tìm thấy câu hỏi nào trong file Word tải lên. "
            "Vui lòng đảm bảo các câu hỏi bắt đầu bằng 'Câu 1.', 'Câu 2.' hoặc 'Câu 1:'."
        )

    rng = random.Random(seed)
    variants: List[ShuffledVariantResult] = []
    for code in exam_codes:
        variant = generate_single_variant_docx(
            source_docx_bytes=source_docx_bytes,
            parsed=parsed,
            exam_code=str(code).strip(),
            meta=meta,
            rng=rng,
        )
        variants.append(variant)

    excel_bytes = generate_answer_key_excel(variants, meta)

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("00_De_Goc_Tham_Chieu.docx", source_docx_bytes)
        for v in variants:
            zf.writestr(f"Ma_De_{v.exam_code}.docx", v.docx_bytes)
        zf.writestr("Bang_Dap_An_Cham_Thi.xlsx", excel_bytes)

    return zip_buf.getvalue(), parsed, variants


def generate_exam_codes(
    num_variants: int = 4,
    starting_code: int = 1001,
    custom_codes: Optional[List[str]] = None,
) -> List[str]:
    if custom_codes:
        cleaned = [str(c).strip() for c in custom_codes if str(c).strip()]
        if cleaned:
            return cleaned[:24]
    count = max(1, min(int(num_variants or 4), 24))
    start = int(starting_code or 1001)
    return [str(start + i) for i in range(count)]


shuffle_and_package_zip = create_shuffled_exam_package_zip
