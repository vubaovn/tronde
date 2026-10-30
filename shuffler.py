# -*- coding: utf-8 -*-
"""
Module lõi Trộn Đề Trắc Nghiệm Chuẩn GDPT 2018 (3 Phần: TNKQ 4 lựa chọn, Đúng/Sai, Trả lời ngắn).
Phiên bản v2.4 (nâng cấp từ v2.2):
1. Bảo toàn tuyệt đối 100% công thức MathType OLE (<w:object>, <w:pict>, Word Field EMBED Equation.DSMT4),
   Word Equation (<m:oMath>, <m:oMathPara>), hình vẽ (<w:drawing>), ký hiệu (<w:sym>) ở cả thân câu hỏi
   lẫn trong tất cả các phương án A, B, C, D và a), b), c), d).
2. Khi một dòng <w:p> chỉ chứa 1 phương án đơn lẻ, sử dụng cơ chế cắt nhãn tại chỗ (in-place clone)
   giống hệt phần thân câu hỏi để không bao giờ tháo rời các thẻ <w:r>.
3. Khi một dòng <w:p> chứa nhiều phương án nối tiếp (A. ... B. ... C. ... D. ...), bảo vệ nguyên vẹn
   các trường Word phức hợp (<w:fldChar> begin..end) và không bao giờ bỏ qua bất kỳ thẻ <w:r> nào.
4. Xóa sạch 100% gạch chân ở nhãn phương án sau khi trộn.
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
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"


@dataclass
class OptionBlock:
    """Đại diện cho 1 phương án lựa chọn (A/B/C/D ở Phần I hoặc a/b/c/d ở Phần II)."""
    original_label: str  # 'A', 'B', 'C', 'D' hoặc 'a', 'b', 'c', 'd'
    is_correct: bool     # Phần I: True nếu là đáp án đúng; Phần II: True nếu là mệnh đề Đúng (Đ), False nếu Sai (S)
    xml_elements: List[OxmlElement] = field(default_factory=list)  # Các thẻ <w:p> hoặc <w:tbl> của phương án (đã bỏ nhãn đầu)
    fixed_order: bool = False  # Nếu phương án là "Cả A, B, C đều đúng" -> giữ cố định vị trí cuối


@dataclass
class QuestionBlock:
    """Đại diện cho 1 câu hỏi trong đề thi."""
    part: int  # 1: TN 4 lựa chọn, 2: TN Đúng/Sai, 3: Trả lời ngắn
    original_index: int
    stem_elements: List[OxmlElement] = field(default_factory=list)  # Các thẻ <w:p>, <w:tbl> của thân câu hỏi (đã bỏ chữ "Câu X:")
    options: List[OptionBlock] = field(default_factory=list)       # Danh sách phương án (Phần I: 4 phương án, Phần II: 4 ý a,b,c,d)
    short_answer: str = ""                                         # Đáp án trả lời ngắn cho Phần III


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
        elif elem.tag == qn("w:tab"):
            texts.append("\t")
        elif elem.tag in (qn("w:br"), qn("w:cr")):
            texts.append("\n")
    return "".join(texts)


def run_has_non_text_media(r_elem: OxmlElement) -> bool:
    """
    Kiểm tra xem thẻ <w:r> có chứa đối tượng công thức MathType (<w:object>),
    hình vẽ (<w:drawing>, <w:pict>), ký hiệu (<w:sym>), trường công thức (<w:fldChar>, <w:instrText>)
    hay phần tử phi văn bản nào khác không.
    """
    for sub in r_elem:
        if sub.tag not in (qn("w:rPr"), qn("w:t"), qn("w:tab")):
            return True
    return False


def get_run_fld_char_types(r_elem: OxmlElement) -> List[str]:
    """Lấy danh sách các giá trị w:fldCharType ('begin', 'separate', 'end') trong một <w:r>."""
    types: List[str] = []
    for sub in r_elem:
        if sub.tag == qn("w:fldChar"):
            val = sub.get(qn("w:fldCharType")) or sub.get("fldCharType") or ""
            if val:
                types.append(val.lower())
    return types


def is_rPr_underlined(rPr: Optional[OxmlElement]) -> bool:
    """Kiểm tra xem một thẻ <w:rPr> có thuộc tính gạch chân <w:u> hay không."""
    if rPr is None:
        return False
    u_elem = rPr.find(qn("w:u"))
    if u_elem is not None:
        val = u_elem.get(qn("w:val")) or u_elem.get("val")
        if val is None or val.lower() not in ("none", "0", "false"):
            return True
    return False


def is_rPr_bold(rPr: Optional[OxmlElement]) -> bool:
    """Kiểm tra xem một thẻ <w:rPr> có thuộc tính in đậm <w:b> hay không."""
    if rPr is None:
        return False
    b_elem = rPr.find(qn("w:b"))
    if b_elem is not None:
        val = b_elem.get(qn("w:val")) or b_elem.get("val")
        if val is None or val.lower() not in ("none", "0", "false"):
            return True
    return False


def is_run_underlined(r_elem: OxmlElement, p_elem: Optional[OxmlElement] = None) -> bool:
    """Kiểm tra xem thẻ <w:r> (hoặc <w:pPr>/<w:rPr> bao ngoài) có định dạng gạch chân hay không."""
    rPr = r_elem.find(qn("w:rPr"))
    if is_rPr_underlined(rPr):
        return True
    if p_elem is not None:
        pPr = p_elem.find(qn("w:pPr"))
        if pPr is not None and is_rPr_underlined(pPr.find(qn("w:rPr"))):
            return True
    return False


def is_run_bold(r_elem: OxmlElement, p_elem: Optional[OxmlElement] = None) -> bool:
    """Kiểm tra xem thẻ <w:r> có định dạng in đậm hay không."""
    rPr = r_elem.find(qn("w:rPr"))
    if is_rPr_bold(rPr):
        return True
    if p_elem is not None:
        pPr = p_elem.find(qn("w:pPr"))
        if pPr is not None and is_rPr_bold(pPr.find(qn("w:rPr"))):
            return True
    return False


def remove_underline_from_rPr(rPr: Optional[OxmlElement]) -> None:
    """Xóa mọi thẻ <w:u> trong <w:rPr>."""
    if rPr is None:
        return
    for u_elem in list(rPr.findall(qn("w:u"))):
        rPr.remove(u_elem)


def remove_underline_from_run(r_elem: OxmlElement) -> None:
    """Xóa thuộc tính gạch chân khỏi thẻ <w:r>."""
    remove_underline_from_rPr(r_elem.find(qn("w:rPr")))


def remove_underline_from_paragraph_pPr(p_elem: OxmlElement) -> None:
    """Xóa thuộc tính gạch chân mặc định ở cấp đoạn văn <w:pPr>/<w:rPr> nếu có."""
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is not None:
        remove_underline_from_rPr(pPr.find(qn("w:rPr")))


def clean_option_paragraph_underlines(p_elem: OxmlElement) -> None:
    """
    Xóa triệt để gạch chân còn sót lại ở đầu phương án (dấu chấm '.', dấu ')', khoảng trắng),
    đồng thời bảo toàn nguyên vẹn mọi đối tượng công thức MathType / Equation / hình ảnh.
    """
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
        if child.tag == qn("w:pPr"):
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
            if re.fullmatch(r"[\.\:\)\-\s]+", stripped):
                remove_underline_from_run(child)
                continue
            if re.match(r"^[\.\:\)]\s*", stripped) and is_run_underlined(child):
                remove_underline_from_run(child)
            seen_meaningful_content = True


def strip_prefix_from_paragraph(p_elem: OxmlElement, regex_pattern: str) -> Tuple[OxmlElement, bool, str]:
    """
    Sao chép thẻ <w:p>, kiểm tra và cắt bỏ phần nhãn đầu câu/đáp án khớp với `regex_pattern`
    (ví dụ: 'Câu 12: ', 'Câu 1. ', 'A. ', '*A. ', 'a) ') ngay tại chỗ (in-place),
    đồng thời kiểm tra xem phần nhãn có được gạch chân hay đánh dấu * hay không.
    Dừng thu thập văn bản đầu dòng ngay khi gặp công thức (<w:object>, <w:pict>, <m:oMath>)
    hoặc hình vẽ (<w:drawing>) để bảo toàn 100% mọi công thức và văn bản phía sau.
    """
    new_p = copy.deepcopy(p_elem)
    remove_underline_from_paragraph_pPr(new_p)

    t_nodes: List[Tuple[OxmlElement, OxmlElement, str]] = []
    for child in new_p:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag != qn("w:r"):
            # Gặp <m:oMath>, <m:oMathPara>, <mc:AlternateContent>, <w:fldSimple>... -> dừng ngay
            break
        stop_here = False
        for sub in child:
            if sub.tag == qn("w:rPr"):
                continue
            if sub.tag == qn("w:t") and sub.text:
                t_nodes.append((child, sub, sub.text))
            elif sub.tag == qn("w:tab"):
                t_nodes.append((child, sub, "\t"))
            elif sub.tag in (
                qn("w:object"),
                qn("w:pict"),
                qn("w:drawing"),
                qn("w:sym"),
                qn("w:fldChar"),
                qn("w:instrText"),
            ):
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

        if t_elem.tag == qn("w:tab"):
            if node_end <= cut_len:
                r_elem.remove(t_elem)
            continue

        if node_end <= cut_len:
            t_elem.text = ""
            remove_underline_from_run(r_elem)
        elif node_start < cut_len < node_end:
            offset = cut_len - node_start
            t_elem.text = txt[offset:].lstrip(" \t")
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            remove_underline_from_run(r_elem)

    clean_option_paragraph_underlines(new_p)
    return new_p, (is_underlined or has_asterisk), label_char


def _split_runs_at_char_offsets(p_elem: OxmlElement, part: int = 1) -> List[OxmlElement]:
    """
    Chuẩn hóa danh sách các phần tử con trong <w:p> để phục vụ tách nhiều phương án trên cùng 1 dòng:
    - Bảo vệ nguyên vẹn các trường Word phức hợp (<w:fldChar> begin ... end, như EMBED Equation.DSMT4).
    - Nếu một <w:r> chỉ chứa công thức MathType (<w:object>, <w:pict>), hình ảnh (<w:drawing>) mà không có
      <w:t> hay <w:tab>, giữ nguyên bản sao <w:r> gốc (copy.deepcopy(child)).
    - Chỉ tách <w:t> khi bên trong <w:t> thực sự có ranh giới nhãn phương án nối tiếp (A., B., C., D. / a), b), c), d)).
    """
    expanded: List[OxmlElement] = []
    field_depth = 0

    for child in p_elem:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag != qn("w:r"):
            expanded.append(copy.deepcopy(child))
            continue

        fld_types = get_run_fld_char_types(child)
        was_in_field = field_depth > 0
        for ft in fld_types:
            if ft == "begin":
                field_depth += 1
            elif ft == "end":
                field_depth = max(0, field_depth - 1)

        # Nếu đang ở trong một Word Field (ví dụ EMBED Equation.DSMT4 hoặc EQ), giữ nguyên <w:r> 100%
        if was_in_field or field_depth > 0 or fld_types or child.find(qn("w:instrText")) is not None:
            expanded.append(copy.deepcopy(child))
            continue

        has_t = any(sub.tag == qn("w:t") for sub in child)
        has_tab = any(sub.tag == qn("w:tab") for sub in child)

        # Nếu <w:r> không có <w:t> và không có <w:tab> (ví dụ chỉ chứa <w:object>, <w:pict>, <w:drawing>, <w:sym>)
        # -> Giữ nguyên toàn bộ thẻ <w:r> gốc mà không tháo rời!
        if not has_t and not has_tab:
            expanded.append(copy.deepcopy(child))
            continue

        rPr = child.find(qn("w:rPr"))
        run_styled = is_rPr_bold(rPr) or is_rPr_underlined(rPr)

        for sub in child:
            if sub.tag == qn("w:rPr"):
                continue
            if sub.tag == qn("w:t"):
                raw_txt = sub.text or ""
                if not raw_txt:
                    continue
                if part == 1:
                    if run_styled:
                        inline_re = re.compile(r"(\s+)(\*?\s*[A-D]\s*[\.\:\)]\s*)")
                    else:
                        inline_re = re.compile(r"(\s{2,}|\s+(?=\*))(\*?\s*[A-D]\s*[\.\:\)]\s*)")
                else:
                    if run_styled:
                        inline_re = re.compile(r"(\s+)(\*?\s*[a-d]\s*[\)\.]\s*)")
                    else:
                        inline_re = re.compile(r"(\s+)(\*?\s*[a-d]\s*\)\s*)|(\s{2,})(\*?\s*[a-d]\s*\.\s*)")

                pos = 0
                for m in inline_re.finditer(raw_txt):
                    sep_space = m.group(1) or m.group(3)
                    marker_part = m.group(2) or m.group(4)
                    start_sep = m.start(1) if m.group(1) is not None else m.start(3)

                    before_part = raw_txt[pos:start_sep]
                    if before_part:
                        r1 = OxmlElement("w:r")
                        if rPr is not None:
                            r1.append(copy.deepcopy(rPr))
                        t1 = OxmlElement("w:t")
                        t1.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        t1.text = before_part
                        r1.append(t1)
                        expanded.append(r1)

                    r_sep = OxmlElement("w:r")
                    t_sep = OxmlElement("w:t")
                    t_sep.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t_sep.text = sep_space
                    r_sep.append(t_sep)
                    expanded.append(r_sep)

                    r_mk = OxmlElement("w:r")
                    if rPr is not None:
                        r_mk.append(copy.deepcopy(rPr))
                    t_mk = OxmlElement("w:t")
                    t_mk.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t_mk.text = marker_part
                    r_mk.append(t_mk)
                    expanded.append(r_mk)

                    pos = m.end()

                rem_txt = raw_txt[pos:]
                if rem_txt:
                    r_rem = OxmlElement("w:r")
                    if rPr is not None:
                        r_rem.append(copy.deepcopy(rPr))
                    t_rem = OxmlElement("w:t")
                    t_rem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    t_rem.text = rem_txt
                    r_rem.append(t_rem)
                    expanded.append(r_rem)
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
) -> Optional[List[Tuple[str, bool, OxmlElement]]]:
    """
    Tách các phương án A., B., C., D. (Phần I) hoặc a), b), c), d) (Phần II) trong một đoạn văn `<w:p>`:
    - Nếu đoạn văn chỉ chứa 1 phương án duy nhất: dùng ngay `strip_prefix_from_paragraph` (giống hệt cách
      xử lý thân câu hỏi) để giữ nguyên 100% cấu trúc `<w:p>` và mọi công thức MathType/Equation.
    - Nếu đoạn văn chứa nhiều phương án trên cùng 1 dòng: tách chính xác tại các nhãn phương án,
      bảo vệ các trường `<w:fldChar>` và không bao giờ làm rơi bất kỳ `<w:r>` hay `<w:object>` nào.
    """
    if part == 1:
        start_check_re = re.compile(r"^\s*(\*?)\s*([A-D])\s*[\.\:\)]\s*", re.DOTALL)
        opt_prefix_re = re.compile(r"^(\s*\*?\s*([A-D])\s*([\.\:\)])\s*)")
        single_strip_re = r"^\s*\*?\s*(?P<label>[A-D])\s*[\.\:\)]\s*"
        valid_seq = ["A", "B", "C", "D"]
    else:
        start_check_re = re.compile(r"^\s*(\*?)\s*([a-d])\s*[\)\.]\s*", re.DOTALL)
        opt_prefix_re = re.compile(r"^(\s*\*?\s*([a-d])\s*([\)\.])\s*)")
        single_strip_re = r"^\s*\*?\s*(?P<label>[a-d])\s*[\)\.]\s*"
        valid_seq = ["a", "b", "c", "d"]

    # Kiểm tra bằng strip_prefix_from_paragraph để đảm bảo nhãn nằm trước công thức đầu tiên
    test_p, test_corr, test_lbl = strip_prefix_from_paragraph(p_elem, single_strip_re)
    if not test_lbl:
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
    field_depth = 0

    i = 0
    n = len(children)
    while i < n:
        node = children[i]

        # Theo dõi trạng thái bên trong Word Complex Field (<w:fldChar> begin ... end)
        if node.tag == qn("w:r"):
            fld_types = get_run_fld_char_types(node)
            was_in_field = field_depth > 0
            for ft in fld_types:
                if ft == "begin":
                    field_depth += 1
                elif ft == "end":
                    field_depth = max(0, field_depth - 1)
            if was_in_field or field_depth > 0 or fld_types or node.find(qn("w:instrText")) is not None:
                node_copy = copy.deepcopy(node)
                remove_underline_from_run(node_copy)
                current_nodes.append(node_copy)
                is_line_start = False
                after_strong_sep = False
                after_any_space = False
                i += 1
                continue

        # 1. Nếu là Tab run -> bật cờ after_strong_sep
        if node.tag == qn("w:r") and node.find(qn("w:tab")) is not None and not node.findall(qn("w:t")) and not run_has_non_text_media(node):
            after_strong_sep = True
            after_any_space = True
            i += 1
            continue

        # 2. Nếu là <w:r> chứa công thức MathType (<w:object>), hình vẽ (<w:drawing>, <w:pict>), <w:sym>
        if node.tag == qn("w:r") and run_has_non_text_media(node):
            node_copy = copy.deepcopy(node)
            remove_underline_from_run(node_copy)
            current_nodes.append(node_copy)
            is_line_start = False
            # Không bật after_strong_sep sau công thức để tránh cắt nhầm chữ cái ngay sau công thức
            after_strong_sep = False
            after_any_space = False
            i += 1
            continue

        # 3. Nếu là <w:r> chứa văn bản <w:t>
        if node.tag == qn("w:r"):
            lookahead_runs: List[Tuple[int, OxmlElement, str]] = []
            j = i
            while j < n and len(lookahead_runs) < 3:
                cand = children[j]
                if cand.tag != qn("w:r") or cand.find(qn("w:tab")) is not None or run_has_non_text_media(cand):
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

                leading_spaces = len(matched_prefix) - len(matched_prefix.lstrip(" \t"))
                is_genuine_option_marker = (
                    is_line_start
                    or after_strong_sep
                    or (leading_spaces >= 2)
                    or (
                        (after_any_space or leading_spaces >= 1)
                        and (label_is_bold or label_is_ul or has_ast or punct_char == ")")
                    )
                )

                if is_expected_label and is_genuine_option_marker:
                    if current_label is not None:
                        segments.append((current_label, current_correct, current_nodes))

                    consumed_chars = 0
                    advance_count = 0
                    leftover_runs: List[OxmlElement] = []

                    for _, r_cand, r_txt in lookahead_runs:
                        if consumed_chars >= prefix_len:
                            break
                        r_start = consumed_chars
                        r_end = consumed_chars + len(r_txt)
                        consumed_chars = r_end
                        advance_count += 1

                        if r_end > prefix_len:
                            offset = prefix_len - r_start
                            rem_str = r_txt[offset:].lstrip(" \t")
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

                    if leftover_runs:
                        current_nodes.extend(leftover_runs)
                        l_txt = "".join((t.text or "") for t in leftover_runs[-1].findall(qn("w:t")))
                        after_strong_sep = bool(re.search(r"(\t|\s{2,})$", l_txt))
                        after_any_space = bool(re.search(r"[\s\.\;\,]$", l_txt))
                    else:
                        after_strong_sep = False
                        after_any_space = True

                    i += max(1, advance_count)
                    continue

            # Không phải nhãn phương án mới -> đưa vào nội dung của phương án hiện tại
            txt_curr = "".join((t.text or "") for t in node.findall(qn("w:t")))
            if not txt_curr.strip():
                if len(txt_curr) >= 2 or "\t" in txt_curr:
                    after_strong_sep = True
                after_any_space = True
                if current_nodes:
                    node_copy = copy.deepcopy(node)
                    remove_underline_from_run(node_copy)
                    current_nodes.append(node_copy)
                i += 1
                continue

            node_copy = copy.deepcopy(node)
            if not current_nodes:
                if is_run_underlined(node_copy):
                    remove_underline_from_run(node_copy)
                t_els = node_copy.findall(qn("w:t"))
                if t_els and t_els[0].text:
                    t_els[0].text = t_els[0].text.lstrip(" \t")
                    t_els[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            current_nodes.append(node_copy)
            is_line_start = False
            after_strong_sep = bool(re.search(r"(\t|\s{2,})$", txt_curr))
            after_any_space = bool(re.search(r"[\s\.\;\,]$", txt_curr))
            i += 1
        else:
            # Các phần tử cấp đoạn văn khác như <m:oMath>, <m:oMathPara>, <mc:AlternateContent>, <w:fldSimple>...
            current_nodes.append(copy.deepcopy(node))
            is_line_start = False
            after_strong_sep = False
            after_any_space = False
            i += 1

    if current_label is not None:
        segments.append((current_label, current_correct, current_nodes))

    if not segments:
        return None

    # RẤT QUAN TRỌNG: Nếu cả đoạn <w:p> chỉ có đúng 1 phương án, trả về ngay bản cắt tại chỗ (in-place)
    # từ `strip_prefix_from_paragraph` (giống hệt cách cắt nhãn "Câu 1." ở thân câu hỏi)!
    if len(segments) == 1:
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
        result.append((lbl, is_corr, opt_p))

    return result


def prepend_bold_label_to_paragraph(p_elem: OxmlElement, label_text: str, color_hex: Optional[str] = None) -> OxmlElement:
    """
    Chèn một nhãn in đậm KHÔNG GẠCH CHÂN (ví dụ: 'Câu 1. ' hoặc 'A. ' hoặc 'a) ') vào đầu thẻ <w:p>
    mà giữ nguyên toàn bộ công thức, hình ảnh và định dạng phần thân của <w:p>.
    """
    new_p = copy.deepcopy(p_elem)
    remove_underline_from_paragraph_pPr(new_p)

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

    pPr = new_p.find(qn("w:pPr"))
    if pPr is not None:
        pPr.addnext(r_label)
    else:
        new_p.insert(0, r_label)

    return new_p


# ---------------------------------------------------------------------------
# Bộ phân tích cú pháp (Parser) file Word đề gốc (.docx)
# ---------------------------------------------------------------------------

PART1_HEADER_RE = re.compile(r"^\s*PHẦN\s*(?:I|1)\b", re.IGNORECASE)
PART2_HEADER_RE = re.compile(r"^\s*PHẦN\s*(?:II|2)\b", re.IGNORECASE)
PART3_HEADER_RE = re.compile(r"^\s*PHẦN\s*(?:III|3)\b", re.IGNORECASE)
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


def parse_docx_exam(docx_bytes: bytes) -> ParsedExam:
    """
    Đọc file Word (.docx) từ bytes và bóc tách thành 3 phần:
    - Phần I: Trắc nghiệm nhiều phương án lựa chọn (A, B, C, D)
    - Phần II: Trắc nghiệm Đúng/Sai (a, b, c, d)
    - Phần III: Trắc nghiệm trả lời ngắn
    """
    doc = docx.Document(io.BytesIO(docx_bytes))
    parsed = ParsedExam(source_doc=doc)

    body = doc.element.body
    current_part = 1
    current_q: Optional[QuestionBlock] = None

    def finalize_question(q: Optional[QuestionBlock]) -> None:
        if q is None:
            return
        if q.part == 1:
            if q.options and not any(opt.is_correct for opt in q.options):
                q.options[0].is_correct = True
                parsed.warnings.append(
                    f"Phần I - Câu {q.original_index}: Chưa gạch chân đáp án đúng, hệ thống tạm lấy phương án A."
                )
            for opt in q.options:
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

    for child in body:
        if child.tag == qn("w:sectPr"):
            continue

        if child.tag == qn("w:p"):
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
                break

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
                    current_q.stem_elements.append(copy.deepcopy(child))
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

        elif child.tag == qn("w:tbl"):
            if current_q is not None:
                if current_q.part in (1, 2) and current_q.options:
                    current_q.options[-1].xml_elements.append(copy.deepcopy(child))
                else:
                    current_q.stem_elements.append(copy.deepcopy(child))

    finalize_question(current_q)
    return parsed


# ---------------------------------------------------------------------------
# Bộ sinh mã đề đã trộn (.docx) và Bảng đáp án chấm thi (.xlsx)
# ---------------------------------------------------------------------------

@dataclass
class ExamMetadata:
    """Thông tin tiêu đề đề thi in ở đầu mỗi mã đề."""
    governing_body: str = "SỞ GD&ĐT BẮC NINH"
    school_name: str = "TRƯỜNG THPT HIỆP HÒA SỐ 3"
    exam_name: str = "ĐỀ KIỂM TRA ĐỊNH KỲ NĂM HỌC 2026 - 2027"
    subject_name: str = "Môn: TOÁN - Lớp 11"
    duration_minutes: str = "90 phút (không kể thời gian phát đề)"
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


def _set_cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for m, val in [("top", top), ("bottom", bottom), ("left", left), ("right", right)]:
        node = OxmlElement(f"w:{m}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _format_run(run, font_name="Times New Roman", size_pt=12, bold=False, italic=False, underline=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.underline = underline


def generate_single_variant_docx(
    source_docx_bytes: bytes,
    parsed: ParsedExam,
    exam_code: str,
    meta: ExamMetadata,
    rng: random.Random,
) -> ShuffledVariantResult:
    doc = docx.Document(io.BytesIO(source_docx_bytes))
    body = doc.element.body

    sectPr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not sectPr:
            body.remove(child)

    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.7)

    header_tbl = doc.add_table(rows=1, cols=2)
    header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    header_tbl.autofit = False

    cell_left = header_tbl.cell(0, 0)
    cell_right = header_tbl.cell(0, 1)
    cell_left.width = Inches(3.1)
    cell_right.width = Inches(3.8)
    _set_cell_margins(cell_left, 40, 40, 60, 60)
    _set_cell_margins(cell_right, 40, 40, 60, 60)

    p_left1 = cell_left.paragraphs[0]
    p_left1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left1.paragraph_format.space_after = Pt(2)
    r = p_left1.add_run(meta.governing_body.upper() + "\n")
    _format_run(r, size_pt=11, bold=False)
    r2 = p_left1.add_run(meta.school_name.upper())
    _format_run(r2, size_pt=11, bold=True, underline=True)

    p_left2 = cell_left.add_paragraph()
    p_left2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_left2.paragraph_format.space_before = Pt(4)
    p_left2.paragraph_format.space_after = Pt(2)
    r_official = p_left2.add_run("ĐỀ CHÍNH THỨC")
    _format_run(r_official, size_pt=11, bold=True)

    p_right1 = cell_right.paragraphs[0]
    p_right1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_right1.paragraph_format.space_after = Pt(2)
    r_exam = p_right1.add_run(meta.exam_name.upper() + "\n")
    _format_run(r_exam, size_pt=11, bold=True)
    r_subj = p_right1.add_run(meta.subject_name + "\n")
    _format_run(r_subj, size_pt=11, bold=True)
    r_time = p_right1.add_run(f"Thời gian làm bài: {meta.duration_minutes}")
    _format_run(r_time, size_pt=10.5, italic=True)

    p_code = cell_right.add_paragraph()
    p_code.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_code.paragraph_format.space_before = Pt(4)
    p_code.paragraph_format.space_after = Pt(2)
    r_code = p_code.add_run(f"Mã đề thi: {exam_code}")
    _format_run(r_code, size_pt=12, bold=True)

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(6)
    p_info.paragraph_format.space_after = Pt(6)
    r_info = p_info.add_run("Họ và tên thí sinh: ..................................................................... Số báo danh: ..........................")
    _format_run(r_info, size_pt=11, bold=False)

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
        p_h1.paragraph_format.space_before = Pt(6)
        p_h1.paragraph_format.space_after = Pt(4)
        r_h1 = p_h1.add_run("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. ")
        _format_run(r_h1, size_pt=12, bold=True)
        r_h1_desc = p_h1.add_run(
            f"Thí sinh trả lời từ câu 1 đến câu {len(parsed.part1_questions)}. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
        )
        _format_run(r_h1_desc, size_pt=12, italic=True)

        p1_list = list(parsed.part1_questions)
        if meta.shuffle_part1_questions:
            rng.shuffle(p1_list)

        letters_p1 = ["A", "B", "C", "D"]
        for idx, q in enumerate(p1_list, start=1):
            q_mapping["part1"].append(q.original_index)

            for s_idx, stem_el in enumerate(q.stem_elements):
                if s_idx == 0 and stem_el.tag == qn("w:p"):
                    new_stem = prepend_bold_label_to_paragraph(stem_el, f"Câu {idx}. ")
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
                for el_idx, el in enumerate(opt.xml_elements):
                    if el_idx == 0 and el.tag == qn("w:p"):
                        labeled_opt = prepend_bold_label_to_paragraph(el, f"{letter}. ")
                        append_element_to_body(labeled_opt)
                    else:
                        append_element_to_body(copy.deepcopy(el))

            part1_keys.append(correct_letter)

    # PHẦN II
    if parsed.part2_questions:
        p_h2 = doc.add_paragraph()
        p_h2.paragraph_format.space_before = Pt(8)
        p_h2.paragraph_format.space_after = Pt(4)
        r_h2 = p_h2.add_run("PHẦN II. Câu trắc nghiệm đúng sai. ")
        _format_run(r_h2, size_pt=12, bold=True)
        r_h2_desc = p_h2.add_run(
            f"Thí sinh trả lời từ câu 1 đến câu {len(parsed.part2_questions)}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
        )
        _format_run(r_h2_desc, size_pt=12, italic=True)

        p2_list = list(parsed.part2_questions)
        if meta.shuffle_part2_questions:
            rng.shuffle(p2_list)

        letters_p2 = ["a", "b", "c", "d"]
        for idx, q in enumerate(p2_list, start=1):
            q_mapping["part2"].append(q.original_index)

            for s_idx, stem_el in enumerate(q.stem_elements):
                if s_idx == 0 and stem_el.tag == qn("w:p"):
                    new_stem = prepend_bold_label_to_paragraph(stem_el, f"Câu {idx}. ")
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
                        labeled_opt = prepend_bold_label_to_paragraph(el, f"{letter}) ")
                        append_element_to_body(labeled_opt)
                    else:
                        append_element_to_body(copy.deepcopy(el))

            part2_keys.append("-".join(q_ds_list))
            part2_detail.append(q_ds_dict)

    # PHẦN III
    if parsed.part3_questions:
        p_h3 = doc.add_paragraph()
        p_h3.paragraph_format.space_before = Pt(8)
        p_h3.paragraph_format.space_after = Pt(4)
        r_h3 = p_h3.add_run("PHẦN III. Câu trắc nghiệm trả lời ngắn. ")
        _format_run(r_h3, size_pt=12, bold=True)
        r_h3_desc = p_h3.add_run(
            f"Thí sinh trả lời từ câu 1 đến câu {len(parsed.part3_questions)}."
        )
        _format_run(r_h3_desc, size_pt=12, italic=True)

        p3_list = list(parsed.part3_questions)
        if meta.shuffle_part3_questions:
            rng.shuffle(p3_list)

        for idx, q in enumerate(p3_list, start=1):
            q_mapping["part3"].append(q.original_index)
            for s_idx, stem_el in enumerate(q.stem_elements):
                if s_idx == 0 and stem_el.tag == qn("w:p"):
                    new_stem = prepend_bold_label_to_paragraph(stem_el, f"Câu {idx}. ")
                    append_element_to_body(new_stem)
                else:
                    append_element_to_body(copy.deepcopy(stem_el))
            part3_keys.append(q.short_answer)

    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(12)
    r_end = p_end.add_run("-------------- HẾT --------------\n")
    _format_run(r_end, size_pt=11, bold=True)
    r_note = p_end.add_run("(Thí sinh không được sử dụng tài liệu. Cán bộ coi thi không giải thích gì thêm)")
    _format_run(r_note, size_pt=10.5, italic=True)

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
