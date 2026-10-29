# -*- coding: utf-8 -*-
"""
Module lõi Trộn Đề Trắc Nghiệm Chuẩn GDPT 2018 (3 Phần: TNKQ 4 lựa chọn, Đúng/Sai, Trả lời ngắn).
Phiên bản v2.2 - Đã khắc phục triệt để lỗi mất nội dung phương án ở cả Phần I và Phần II:
1. Bảo toàn 100% thẻ <w:r> chứa công thức MathType OLE (<w:object>), hình ảnh (<w:drawing>, <w:pict>),
   ký tự đặc biệt (<w:sym>) ngay ở đầu phương án (không có thẻ <w:t> nên bản cũ tưởng nhầm là run rỗng và bỏ qua).
2. Chống cắt nhầm câu khi trong nội dung mệnh đề có tên điểm/biến kết thúc bằng dấu chấm như
   "điểm B.", "cạnh b.", "mặt phẳng (P).": chỉ nhận diện nhãn phương án nằm giữa dòng khi đứng sau phím Tab,
   >= 2 dấu cách, hoặc nhãn đó được định dạng In Đậm (<w:b>) / Gạch Chân (<w:u>) / có dấu *.
3. Hỗ trợ trường hợp phương án A. hoặc câu hỏi nằm nối tiếp ngay trên cùng dòng thân câu hỏi.
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
    hình vẽ (<w:drawing>, <w:pict>), ký hiệu (<w:sym>), hay công thức toán không.
    Rất quan trọng: Nếu có, tuyệt đối KHÔNG ĐƯỢC coi run này là rỗng!
    """
    for sub in r_elem:
        if sub.tag not in (qn("w:rPr"), qn("w:t"), qn("w:tab")):
            return True
    return False


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
    đồng thời bảo toàn mọi đối tượng công thức/hình ảnh.
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
    for r in direct_runs:
        if run_has_non_text_media(r):
            remove_underline_from_run(r)
            seen_meaningful_content = True
            continue
        txt = "".join((t.text or "") for t in r.findall(qn("w:t")))
        stripped = txt.strip()
        if not stripped:
            remove_underline_from_run(r)
            continue
        if not seen_meaningful_content:
            if re.fullmatch(r"[\.\:\)\-\s]+", stripped):
                remove_underline_from_run(r)
                continue
            if re.match(r"^[\.\:\)]\s*", stripped) and is_run_underlined(r):
                remove_underline_from_run(r)
            seen_meaningful_content = True


def strip_prefix_from_paragraph(p_elem: OxmlElement, regex_pattern: str) -> Tuple[OxmlElement, bool, str]:
    """
    Sao chép thẻ <w:p>, kiểm tra và cắt bỏ phần nhãn đầu câu/đáp án khớp với `regex_pattern`
    (ví dụ: 'Câu 12: ', 'Câu 1. ', 'A. ', '*A. ', 'a) '), đồng thời kiểm tra xem phần nhãn
    có được gạch chân hay đánh dấu * hay không.
    Bảo toàn 100% các node công thức/hình ảnh (<w:object>, <w:drawing>, <m:oMath>).
    """
    new_p = copy.deepcopy(p_elem)
    remove_underline_from_paragraph_pPr(new_p)

    t_nodes: List[Tuple[OxmlElement, OxmlElement, str]] = []
    for child in new_p:
        if child.tag == qn("w:r"):
            for sub in child:
                if sub.tag == qn("w:t") and sub.text:
                    t_nodes.append((child, sub, sub.text))
                elif sub.tag == qn("w:tab"):
                    t_nodes.append((child, sub, "\t"))

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
    Chuẩn hóa danh sách các phần tử con trong <w:p>:
    - Tách mỗi sub-element trong <w:r> (<w:t>, <w:tab>, <w:object>, <w:drawing>, <w:pict>, <w:sym>)
      thành một thẻ <w:r> độc lập có giữ nguyên <w:rPr> (để không bao giờ làm rơi <w:object> hay <w:drawing>).
    - Đối với <w:t> chứa nhiều phương án gõ liền nhau bằng khoảng trắng:
      + Nếu cách nhau bởi >= 2 khoảng trắng (hoặc Tab): luôn tách trước nhãn A., B., C., D. / a), b), c), d).
      + Nếu chỉ cách nhau bởi 1 khoảng trắng: chỉ tách nếu chính run đó đang được In Đậm (<w:b>) hoặc Gạch Chân (<w:u>)
        hoặc có dấu `*` hoặc dấu `)` (đối với Phần II: `a)`, `b)`, `c)`, `d)`), nhằm tránh cắt nhầm chữ
        "điểm B.", "cạnh b." trong văn bản bình thường!
    """
    expanded: List[OxmlElement] = []

    for child in p_elem:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag != qn("w:r"):
            expanded.append(copy.deepcopy(child))
            continue

        rPr = child.find(qn("w:rPr"))
        run_styled = is_rPr_bold(rPr) or is_rPr_underlined(rPr)

        for sub in child:
            if sub.tag == qn("w:rPr"):
                continue
            if sub.tag == qn("w:t") and sub.text:
                raw_txt = sub.text
                # Chọn biểu thức tách bên trong 1 <w:t> an toàn:
                if part == 1:
                    if run_styled:
                        inline_re = re.compile(r"(\s+)(\*?\s*[A-D]\s*[\.\:\)]\s*)")
                    else:
                        # Trong run chữ thường: phải có >= 2 dấu cách hoặc dấu * trước A., B., C., D.
                        inline_re = re.compile(r"(\s{2,}|\s+(?=\*))(\*?\s*[A-D]\s*[\.\:\)]\s*)")
                else:
                    if run_styled:
                        inline_re = re.compile(r"(\s+)(\*?\s*[a-d]\s*[\)\.]\s*)")
                    else:
                        # Trong run chữ thường Phần II: chấp nhận a), b), c), d) (dấu ngoặc đơn) hoặc >= 2 dấu cách trước a.
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
                    # Đánh dấu nếu khoảng trắng phân cách >= 2 ký tự
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
                # Bảo toàn nguyên vẹn <w:object> (MathType), <w:drawing>, <w:pict>, <w:sym>, <w:tab>
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
    - Bảo toàn 100% các thẻ `<w:r>` chứa công thức MathType (`<w:object>`), hình vẽ (`<w:drawing>`),
      ký hiệu (`<w:sym>`) và các thẻ `<m:oMath>` ở mọi vị trí (kể cả đứng ngay đầu phương án!).
    - Phân biệt chính xác nhãn phương án thật với tên điểm hình học / biến toán học ("điểm B.", "cạnh b.")
      nằm giữa câu văn.
    """
    if part == 1:
        start_check_re = re.compile(r"^\s*(\*?)\s*([A-D])\s*[\.\:\)]\s*", re.DOTALL)
        opt_prefix_re = re.compile(r"^(\s*\*?\s*([A-D])\s*([\.\:\)])\s*)")
        valid_seq = ["A", "B", "C", "D"]
    else:
        start_check_re = re.compile(r"^\s*(\*?)\s*([a-d])\s*[\)\.]\s*", re.DOTALL)
        opt_prefix_re = re.compile(r"^(\s*\*?\s*([a-d])\s*([\)\.])\s*)")
        valid_seq = ["a", "b", "c", "d"]

    raw_text = get_paragraph_text(p_elem)
    if not start_check_re.match(raw_text):
        return None

    pPr = p_elem.find(qn("w:pPr"))
    children = _split_runs_at_char_offsets(p_elem, part=part)

    segments: List[Tuple[str, bool, List[OxmlElement]]] = []
    current_label: Optional[str] = None
    current_correct: bool = False
    current_nodes: List[OxmlElement] = []

    # Cờ đánh dấu loại phân cách đứng ngay trước vị trí hiện tại:
    # - is_line_start: đầu dòng <w:p>
    # - after_strong_sep: sau <w:tab/>, sau >= 2 khoảng trắng, hoặc sau công thức <m:oMath>/<w:object>
    # - after_any_space: sau ít nhất 1 khoảng trắng
    is_line_start = True
    after_strong_sep = True
    after_any_space = True

    i = 0
    n = len(children)
    while i < n:
        node = children[i]

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
            after_strong_sep = True
            after_any_space = True
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
                lookahead_runs.append((j, cand, txt_j))
                j += 1

            combined_txt = "".join(item[2] for item in lookahead_runs)
            m_opt = opt_prefix_re.match(combined_txt) if lookahead_runs else None

            if m_opt:
                matched_prefix = m_opt.group(1)
                cand_raw_label = m_opt.group(2)
                punct_char = m_opt.group(3)  # '.', ':', hoặc ')'
                cand_label = cand_raw_label.upper() if part == 1 else cand_raw_label.lower()

                # Kiểm tra thứ tự nhãn A -> B -> C -> D (hoặc a -> b -> c -> d)
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

                # Kiểm tra xem các run tạo nên nhãn này có được in đậm (<w:b>), gạch chân (<w:u>) hay có dấu '*' không
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

                # Điều kiện chặt chẽ để công nhận đây là NHÃN PHƯƠNG ÁN THẬT (chứ không phải chữ "điểm B." hay "cạnh b."):
                # - Nếu ở đầu dòng (is_line_start): luôn là nhãn phương án.
                # - Nếu ở giữa dòng: phải đứng sau phân cách mạnh (Tab, >=2 dấu cách, công thức),
                #   HOẶC đứng sau dấu cách và bản thân nhãn được In Đậm / Gạch Chân / có dấu '*' / dùng dấu ngoặc ')'.
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
                    leftover_run: Optional[OxmlElement] = None

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
                            if rem_str:
                                leftover_run = copy.deepcopy(r_cand)
                                remove_underline_from_run(leftover_run)
                                for idx_t, t_el in enumerate(leftover_run.findall(qn("w:t"))):
                                    t_el.text = rem_str if idx_t == 0 else ""
                                    t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

                    current_label = cand_label
                    current_correct = has_ast or label_is_ul
                    current_nodes = []
                    is_line_start = False

                    if leftover_run is not None:
                        current_nodes.append(leftover_run)
                        l_txt = "".join((t.text or "") for t in leftover_run.findall(qn("w:t")))
                        after_strong_sep = bool(re.search(r"(\t|\s{2,})$", l_txt))
                        after_any_space = bool(re.search(r"[\s\.\;\,]$", l_txt))
                    else:
                        after_strong_sep = False
                        after_any_space = True

                    i += max(1, advance_count)
                    continue

            # Không phải nhãn phương án mới -> đưa vào nội dung của phương án hiện tại
            txt_curr = "".join((t.text or "") for t in node.findall(qn("w:t")))
            if not current_nodes and not txt_curr.strip():
                # Khoảng trắng thừa ngay sau nhãn phương án
                if len(txt_curr) >= 2 or "\t" in txt_curr:
                    after_strong_sep = True
                after_any_space = True
                i += 1
                continue

            # Nếu đây là một run chỉ toàn khoảng trắng nằm giữa các phần tử của phương án
            if not txt_curr.strip():
                if len(txt_curr) >= 2 or "\t" in txt_curr:
                    after_strong_sep = True
                after_any_space = True
                node_copy = copy.deepcopy(node)
                remove_underline_from_run(node_copy)
                current_nodes.append(node_copy)
                i += 1
                continue

            node_copy = copy.deepcopy(node)
            if not current_nodes and is_run_underlined(node_copy):
                remove_underline_from_run(node_copy)
            current_nodes.append(node_copy)
            is_line_start = False
            after_strong_sep = bool(re.search(r"(\t|\s{2,})$", txt_curr))
            after_any_space = bool(re.search(r"[\s\.\;\,]$", txt_curr))
            i += 1
        else:
            # Các phần tử cấp đoạn văn khác như <m:oMath>, <m:oMathPara>, <w:hyperlink>, <w:smartTag>...
            current_nodes.append(copy.deepcopy(node))
            is_line_start = False
            after_strong_sep = True
            after_any_space = True
            i += 1

    if current_label is not None:
        segments.append((current_label, current_correct, current_nodes))

    if not segments:
        return None

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

    num_p1 = len(variants[0].part1_keys) if variants else 0
    if num_p1 > 0:
        row_cursor = write_table_header(
            ws1,
            row_cursor,
            "PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN",
        )
        for q_i in range(num_p1):
            c_lbl = ws1.cell(row_cursor, 1, f"Câu {q_i + 1}")
            c_lbl.font = font_bold
            c_lbl.border = border_all
            c_lbl.alignment = align_center
            if q_i % 2 == 1:
                c_lbl.fill = fill_zebra

            for v_idx, v in enumerate(variants, start=2):
                val = v.part1_keys[q_i] if q_i < len(v.part1_keys) else ""
                c = ws1.cell(row_cursor, v_idx, val)
                c.font = font_bold
                c.border = border_all
                c.alignment = align_center
                if q_i % 2 == 1:
                    c.fill = fill_zebra
            ws1.row_dimensions[row_cursor].height = 20
            row_cursor += 1
        row_cursor += 1

    num_p2 = len(variants[0].part2_keys) if variants else 0
    if num_p2 > 0:
        row_cursor = write_table_header(
            ws1,
            row_cursor,
            "PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (Thứ tự ý a - b - c - d)",
        )
        for q_i in range(num_p2):
            for sub_letter in ["a", "b", "c", "d"]:
                c_lbl = ws1.cell(row_cursor, 1, f"Câu {q_i + 1}.{sub_letter})")
                c_lbl.font = font_cell
                c_lbl.border = border_all
                c_lbl.alignment = align_center
                if q_i % 2 == 1:
                    c_lbl.fill = fill_zebra

                for v_idx, v in enumerate(variants, start=2):
                    val = v.part2_detail[q_i].get(sub_letter, "") if q_i < len(v.part2_detail) else ""
                    c = ws1.cell(row_cursor, v_idx, val)
                    c.font = font_bold
                    c.border = border_all
                    c.alignment = align_center
                    if q_i % 2 == 1:
                        c.fill = fill_zebra
                ws1.row_dimensions[row_cursor].height = 19
                row_cursor += 1
        row_cursor += 1

    num_p3 = len(variants[0].part3_keys) if variants else 0
    if num_p3 > 0:
        row_cursor = write_table_header(
            ws1,
            row_cursor,
            "PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN",
        )
        for q_i in range(num_p3):
            c_lbl = ws1.cell(row_cursor, 1, f"Câu {q_i + 1}")
            c_lbl.font = font_bold
            c_lbl.border = border_all
            c_lbl.alignment = align_center
            if q_i % 2 == 1:
                c_lbl.fill = fill_zebra

            for v_idx, v in enumerate(variants, start=2):
                val = v.part3_keys[q_i] if q_i < len(v.part3_keys) else ""
                c = ws1.cell(row_cursor, v_idx, val)
                c.font = font_bold
                c.border = border_all
                c.alignment = align_center
                if q_i % 2 == 1:
                    c.fill = fill_zebra
            ws1.row_dimensions[row_cursor].height = 20
            row_cursor += 1

    ws1.column_dimensions["A"].width = 16
    for v_idx in range(2, 2 + len(variants)):
        ws1.column_dimensions[get_column_letter(v_idx)].width = 15

    ws2 = wb.create_sheet(title="Chi_Tiet_Cham_May")
    headers_s2 = ["Mã đề", "Phần", "Câu", "Đáp án"]
    for col_i, h_text in enumerate(headers_s2, start=1):
        c = ws2.cell(1, col_i, h_text)
        c.font = font_header
        c.fill = fill_header
        c.border = border_all
        c.alignment = align_center

    r_s2 = 2
    for v in variants:
        for idx, k in enumerate(v.part1_keys, start=1):
            ws2.cell(r_s2, 1, v.exam_code).alignment = align_center
            ws2.cell(r_s2, 2, "Phần I").alignment = align_center
            ws2.cell(r_s2, 3, f"Câu {idx}").alignment = align_center
            ws2.cell(r_s2, 4, k).alignment = align_center
            for col_i in range(1, 5):
                ws2.cell(r_s2, col_i).font = font_cell
                ws2.cell(r_s2, col_i).border = border_all
            r_s2 += 1
        for idx, k in enumerate(v.part2_keys, start=1):
            ws2.cell(r_s2, 1, v.exam_code).alignment = align_center
            ws2.cell(r_s2, 2, "Phần II").alignment = align_center
            ws2.cell(r_s2, 3, f"Câu {idx} (a-b-c-d)").alignment = align_center
            ws2.cell(r_s2, 4, k).alignment = align_center
            for col_i in range(1, 5):
                ws2.cell(r_s2, col_i).font = font_cell
                ws2.cell(r_s2, col_i).border = border_all
            r_s2 += 1
        for idx, k in enumerate(v.part3_keys, start=1):
            ws2.cell(r_s2, 1, v.exam_code).alignment = align_center
            ws2.cell(r_s2, 2, "Phần III").alignment = align_center
            ws2.cell(r_s2, 3, f"Câu {idx}").alignment = align_center
            ws2.cell(r_s2, 4, k).alignment = align_center
            for col_i in range(1, 5):
                ws2.cell(r_s2, col_i).font = font_cell
                ws2.cell(r_s2, col_i).border = border_all
            r_s2 += 1

    for col_i in range(1, 5):
        ws2.column_dimensions[get_column_letter(col_i)].width = 18

    ws3 = wb.create_sheet(title="Ma_Tran_Dao_Cau")
    ws3.cell(1, 1, "Phần thi").font = font_header
    ws3.cell(1, 2, "Câu ở đề trộn").font = font_header
    ws3.cell(1, 1).fill = fill_header
    ws3.cell(1, 2).fill = fill_header
    ws3.cell(1, 1).border = border_all
    ws3.cell(1, 2).border = border_all
    ws3.cell(1, 1).alignment = align_center
    ws3.cell(1, 2).alignment = align_center

    for v_idx, v in enumerate(variants, start=3):
        c = ws3.cell(1, v_idx, f"Mã đề {v.exam_code} (Câu gốc)")
        c.font = font_header
        c.fill = fill_header
        c.border = border_all
        c.alignment = align_center

    r_s3 = 2
    for part_key, part_label in [("part1", "Phần I"), ("part2", "Phần II"), ("part3", "Phần III")]:
        count_q = len(variants[0].question_mapping[part_key]) if variants else 0
        for q_i in range(count_q):
            ws3.cell(r_s3, 1, part_label).alignment = align_center
            ws3.cell(r_s3, 2, f"Câu {q_i + 1}").alignment = align_center
            ws3.cell(r_s3, 1).font = font_cell
            ws3.cell(r_s3, 2).font = font_bold
            ws3.cell(r_s3, 1).border = border_all
            ws3.cell(r_s3, 2).border = border_all
            for v_idx, v in enumerate(variants, start=3):
                orig_num = v.question_mapping[part_key][q_i]
                c = ws3.cell(r_s3, v_idx, f"Câu gốc {orig_num}")
                c.font = font_cell
                c.border = border_all
                c.alignment = align_center
            r_s3 += 1

    ws3.column_dimensions["A"].width = 14
    ws3.column_dimensions["B"].width = 16
    for v_idx in range(3, 3 + len(variants)):
        ws3.column_dimensions[get_column_letter(v_idx)].width = 20

    excel_buf = io.BytesIO()
    wb.save(excel_buf)
    return excel_buf.getvalue()


def generate_exam_codes(num_variants: int, starting_code: int = 101, custom_codes: Optional[List[str]] = None) -> List[str]:
    if custom_codes:
        cleaned = [c.strip() for c in custom_codes if c.strip()]
        if cleaned:
            return cleaned[:24]
    num_variants = max(1, min(num_variants, 48))
    return [str(starting_code + i) for i in range(num_variants)]


def shuffle_and_package_zip(
    source_docx_bytes: bytes,
    exam_codes: List[str],
    meta: ExamMetadata,
    seed: Optional[int] = None,
) -> Tuple[bytes, ParsedExam, List[ShuffledVariantResult]]:
    parsed = parse_docx_exam(source_docx_bytes)
    rng = random.Random(seed)

    variants: List[ShuffledVariantResult] = []
    for code in exam_codes:
        v_res = generate_single_variant_docx(
            source_docx_bytes=source_docx_bytes,
            parsed=parsed,
            exam_code=code,
            meta=meta,
            rng=rng,
        )
        variants.append(v_res)

    excel_bytes = generate_answer_key_excel(variants, meta)

    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("00_De_Goc_Tham_Chieu.docx", source_docx_bytes)
        for v in variants:
            zf.writestr(f"Ma_De_{v.exam_code}.docx", v.docx_bytes)
        zf.writestr("Bang_Dap_An_Cham_Thi.xlsx", excel_bytes)

    return zip_buf.getvalue(), parsed, variants
