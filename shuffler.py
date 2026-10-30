# -*- coding: utf-8 -*-
"""
Module lõi Trộn Đề Trắc Nghiệm Chuẩn GDPT 2018 (3 Phần: TNKQ 4 lựa chọn, Đúng/Sai, Trả lời ngắn).
Phiên bản v2.5:
1. Khắc phục triệt để lỗi mất phương án D. (hoặc B., C.) trong đề in sau khi trộn:
   - Loại bỏ toàn bộ các thẻ đánh dấu ẩn của Word (<w:proofErr>, <w:bookmarkStart>, <w:bookmarkEnd>,
     <w:lastRenderedPageBreak>) vốn nằm quanh dấu chấm "13 ." và làm đứt quãng việc nhận diện "D.".
   - Xử lý triệt để trường hợp chữ cái "D" nằm ở cuối thẻ <w:t> của run trước (ví dụ ".   D" hoặc ".D")
     còn dấu chấm "." nằm ở thẻ <w:r> kế tiếp (nhờ bước ghép nối liên run `_normalize_cross_run_option_markers`).
   - Không để trạng thái `field_depth` của công thức trước nuốt mất phương án D. ở cuối dòng.
   - Hỗ trợ trường hợp giáo viên xuống dòng bằng Shift+Enter (<w:br/>) hoặc gõ liền các phương án
     A. B. C. D. ngay trong cùng đoạn văn <w:p> với thân câu hỏi (`split_stem_and_trailing_options`).
   - Hỗ trợ cả trường hợp chữ cái nhãn phương án dùng định dạng viết hoa tự động (<w:caps/>) hoặc đứng
     ngay sau dấu chấm kết thúc phương án trước ("13 . D. 21 .").
2. Bảo toàn 100% công thức MathType OLE (<w:object>, <w:pict>, Word Field EMBED Equation.DSMT4),
   Word Equation (<m:oMath>, <m:oMathPara>), hình vẽ (<w:drawing>), ký hiệu (<w:sym>).
3. Xóa sạch 100% gạch chân ở nhãn phương án sau khi trộn.
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

# Các thẻ đánh dấu ẩn cấp đoạn văn (<w:p>) của Word không mang nội dung hiển thị
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

# Các phần tử con trong <w:r> thực sự là công thức / hình ảnh / trường dữ liệu
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
        elif elem.tag in (qn("w:tab"), qn("w:ptab")):
            texts.append("\t")
        elif elem.tag in (qn("w:br"), qn("w:cr")):
            texts.append("\n")
    return "".join(texts)


def run_has_non_text_media(r_elem: OxmlElement) -> bool:
    """
    Kiểm tra xem thẻ <w:r> có chứa đối tượng công thức MathType (<w:object>),
    hình vẽ (<w:drawing>, <w:pict>), ký hiệu (<w:sym>), trường công thức (<w:fldChar>, <w:instrText>)
    hay không. Bỏ qua các thẻ đánh dấu ẩn như <w:lastRenderedPageBreak/>.
    """
    for sub in r_elem:
        if sub.tag in MEDIA_OR_FIELD_RUN_CHILD_TAGS or sub.tag.endswith("AlternateContent"):
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
    """Kiểm tra xem một thẻ <w:rPr> có thuộc tính in đậm <w:b> hoặc <w:bCs> hay không."""
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
    """Kiểm tra xem <w:rPr> có bật thuộc tính viết hoa tự động <w:caps> hay <w:smallCaps> không."""
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
    """Xóa thuộc tính gạch chân mặc định và các thuộc tính khung cố định ở cấp đoạn văn <w:pPr>."""
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is not None:
        remove_underline_from_rPr(pPr.find(qn("w:rPr")))
        for bad_tag in ("w:framePr", "w:pageBreakBefore"):
            for el in list(pPr.findall(qn(bad_tag))):
                pPr.remove(el)


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
    """
    Trải phẳng các thẻ bao trung gian như <w:ins>, <w:smartTag>, <w:sdt> và loại bỏ các thẻ đánh dấu ẩn
    (<w:proofErr>, <w:bookmarkStart>, <w:bookmarkEnd>...) để không bao giờ làm đứt quãng việc
    nhận diện các phương án A, B, C, D.
    """
    out: List[OxmlElement] = []
    for child in p_elem:
        if child.tag == qn("w:pPr") or child.tag in IGNORED_PARAGRAPH_MARKER_TAGS:
            continue
        if child.tag in (qn("w:ins"), qn("w:smartTag"), qn("w:customXml")):
            out.extend(_iter_paragraph_effective_children(child))
        elif child.tag == qn("w:sdt"):
            sdt_content = child.find(qn("w:sdtContent"))
            if sdt_content is not None:
                out.extend(_iter_paragraph_effective_children(sdt_content))
        else:
            out.append(child)
    return out


def strip_prefix_from_paragraph(p_elem: OxmlElement, regex_pattern: str) -> Tuple[OxmlElement, bool, str]:
    """
    Sao chép thẻ <w:p>, kiểm tra và cắt bỏ phần nhãn đầu câu/đáp án khớp với `regex_pattern`
    (ví dụ: 'Câu 12: ', 'Câu 1. ', 'A. ', '*A. ', 'a) ') ngay tại chỗ (in-place),
    đồng thời kiểm tra xem phần nhãn có được gạch chân hay đánh dấu * hay không.
    Bỏ qua mọi thẻ đánh dấu ẩn (<w:proofErr>, <w:bookmarkStart>) ở đầu dòng và dừng thu thập
    ngay khi gặp công thức (<w:object>, <w:pict>, <m:oMath>) hoặc hình vẽ (<w:drawing>).
    """
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
    return new_p, (is_underlined or has_asterisk), label_char


def _split_text_into_option_pieces(raw_txt: str, part: int, run_styled: bool) -> List[Tuple[str, bool]]:
    """
    Tách một chuỗi `raw_txt` trong thẻ `<w:t>` nếu bên trong chuỗi có chứa nhãn phương án nối tiếp:
    - Hỗ trợ cả trường hợp có đầy đủ dấu chấm trong cùng `<w:t>` (ví dụ `".   D. "`, `". D. "`, `".D. "`)
    - Hỗ trợ trường hợp chữ cái `D` nằm ở cuối `<w:t>` (ví dụ `".   D"`, `".D"`) còn dấu `"."` nằm ở `<w:r>` kế tiếp!
    """
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
    """
    Chuẩn hóa danh sách các phần tử con trong <w:p> để phục vụ tách nhiều phương án trên cùng 1 dòng:
    - Bỏ qua các thẻ đánh dấu ẩn (<w:proofErr>, <w:bookmarkStart>, <w:bookmarkEnd>, <w:lastRenderedPageBreak>).
    - Nếu <w:r> chỉ chứa công thức MathType (<w:object>, <w:pict>), hình ảnh (<w:drawing>) mà không có <w:t>/<w:tab>/<w:br>,
      giữ nguyên bản sao <w:r> gốc (copy.deepcopy(child)).
    - Tách các nhãn phương án nằm chung trong cùng 1 <w:t> hoặc chữ cái nằm cuối <w:t> còn dấu '.' nằm ở run sau.
    """
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
    """
    Tách các phương án A., B., C., D. (Phần I) hoặc a), b), c), d) (Phần II) trong một đoạn văn `<w:p>`:
    - Nếu đoạn văn chỉ chứa 1 phương án duy nhất: dùng ngay `strip_prefix_from_paragraph` (giống hệt cách
      xử lý thân câu hỏi) để giữ nguyên 100% cấu trúc `<w:p>` và mọi công thức MathType/Equation.
    - Nếu đoạn văn chứa nhiều phương án trên cùng 1 dòng: tách chính xác cả 4 phương án A, B, C, D.
    """
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

        # 1. Nếu là Tab run hoặc Line-break run (<w:tab/>, <w:ptab/>, <w:br/>, <w:cr/>)
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

        # 2. Nếu là <w:r> chứa công thức MathType (<w:object>), hình vẽ (<w:drawing>, <w:pict>), <w:sym>, <w:fldChar>
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

        # 3. Nếu là <w:r> chứa văn bản <w:t>
        if node.tag == qn("w:r"):
            txt_curr = "".join((t.text or "") for t in node.findall(qn("w:t")))

            # Xử lý ngay các run chỉ chứa khoảng trắng trước khi lookahead
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

    # Nếu cả đoạn <w:p> chỉ có đúng 1 phương án, trả về ngay bản cắt tại chỗ (in-place)
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
        result.append((lbl, is_corr, opt_p))

    return result


def split_paragraph_on_line_breaks_before_options(p_elem: OxmlElement, part: int = 1) -> List[OxmlElement]:
    """
    Trường hợp giáo viên dùng Shift+Enter (<w:br/>) để xuống dòng gõ các phương án A. B. C. D.
    ngay trong cùng một đoạn văn <w:p> với thân câu hỏi (hoặc dùng Shift+Enter giữa C. và D.):
    Hàm này tách <w:p> tại các vị trí <w:br/> mà dòng phía sau bắt đầu bằng nhãn phương án.
    """
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
    """
    Trường hợp giáo viên không nhấn Enter mà gõ nối tiếp các phương án A. ... B. ... C. ... D. ...
    ngay ở cuối đoạn văn thân câu hỏi (ví dụ dùng phím Tab hoặc dấu cách để Word tự đẩy A. B. C. D.
    xuống dòng tiếp theo cùng thuộc 1 thẻ <w:p>):
    Hàm này tìm vị trí bắt đầu nhãn A. (hoặc a)) ở phần cuối của <w:p>, tách riêng phần thân câu hỏi
    và danh sách các phương án.
    """
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
        # Ghép thử tối đa 3 run liên tiếp từ idx
        comb = ""
        for k in range(idx, min(len(children), idx + 3)):
            ck = children[k]
            if ck.tag != qn("w:r") or run_has_non_text_media(ck):
                break
            comb += "".join((t.text or "") for t in ck.findall(qn("w:t")))
        if not first_opt_re.match(comb):
            continue

        # Kiểm tra điều kiện: nhãn A. phải được In Đậm / Gạch Chân hoặc đứng sau Tab / >= 2 dấu cách
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
            stem_Slice = children[:idx]
            while stem_Slice and stem_Slice[-1].tag == qn("w:r") and not run_has_non_text_media(stem_Slice[-1]):
                last_r = stem_Slice[-1]
                has_tb = last_r.find(qn("w:tab")) is not None or last_r.find(qn("w:ptab")) is not None
                l_txt = "".join((t.text or "") for t in last_r.findall(qn("w:t")))
                if has_tb and not l_txt.strip():
                    stem_Slice.pop()
                elif not l_txt.strip() and not has_tb:
                    stem_Slice.pop()
                else:
                    break
            for s_nd in stem_Slice:
                new_stem_p.append(copy.deepcopy(s_nd))
            return new_stem_p, opts

    return stem_p, None


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
