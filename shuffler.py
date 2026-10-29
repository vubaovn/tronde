# -*- coding: utf-8 -*-
"""
Module lõi Trộn Đề Trắc Nghiệm Chuẩn GDPT 2018 (3 Phần: TNKQ 4 lựa chọn, Đúng/Sai, Trả lời ngắn).
Bảo toàn 100% cấu trúc XML của Microsoft Word (.docx):
- Công thức Toán/Lý/Hóa: Office MathML (<m:oMath>, <m:oMathPara>), MathType OLE (<w:object>), Equation.
- Hình ảnh & Đồ thị: DrawingML (<w:drawing>), VML (<w:pict>), mối quan hệ rId (Relationships).
- Định dạng văn bản: In đậm, in nghiêng, chỉ số trên/dưới (superscript/subscript), bảng biểu.

Quy ước nhận diện đáp án từ đề gốc:
1. Phần I (Trắc nghiệm 4 lựa chọn A, B, C, D):
   - Đáp án đúng được gạch chân ở chữ cái đầu (ví dụ: <u>A</u>. hoặc <u>A.</u>) hoặc có dấu * trước (ví dụ: *A.).
   - Hỗ trợ cả trường hợp 4 đáp án nằm trên 4 dòng riêng biệt hoặc nằm chung trên 1-2 dòng (ngăn cách bởi Tab/khoảng trắng).
2. Phần II (Trắc nghiệm Đúng / Sai - a, b, c, d):
   - Ý đúng (Đ) được gạch chân chữ cái đầu (ví dụ: <u>a</u>) hoặc <u>a)</u>) hoặc đánh dấu *a).
   - Ý không gạch chân mặc định là Sai (S). (Cũng hỗ trợ ghi rõ [Đ]/[S] ở cuối ý nếu giáo viên muốn tường minh).
3. Phần III (Trả lời ngắn):
   - Đáp án được đặt ở dòng bắt đầu bằng "Đáp án:" / "ĐS:" / "Key:" (ví dụ: "Đáp án: 12,5" hoặc đoạn văn có gạch chân giá trị đáp án).
   - Khi xuất đề thi cho học sinh, dòng "Đáp án: ..." sẽ tự động được ẩn đi và đưa vào file Excel chấm thi.
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
from docx.shared import Inches, Pt, RGBColor
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


def is_run_underlined(r_elem: OxmlElement) -> bool:
    """Kiểm tra xem một thẻ <w:r> có định dạng gạch chân (<w:u>) hay không."""
    rPr = r_elem.find(qn("w:rPr"))
    if rPr is None:
        return False
    u_elem = rPr.find(qn("w:u"))
    if u_elem is None:
        return False
    val = u_elem.get(qn("w:val"))
    # Trong OpenXML, <w:u/> (không có w:val) hoặc w:val khác 'none' đều là có gạch chân
    return val is None or val.lower() != "none"


def remove_underline_from_run(r_elem: OxmlElement) -> None:
    """Xóa thuộc tính gạch chân khỏi thẻ <w:r> (để học sinh không nhìn thấy dấu hiệu đáp án đúng)."""
    rPr = r_elem.find(qn("w:rPr"))
    if rPr is not None:
        u_elem = rPr.find(qn("w:u"))
        if u_elem is not None:
            rPr.remove(u_elem)


def strip_prefix_from_paragraph(p_elem: OxmlElement, regex_pattern: str) -> Tuple[OxmlElement, bool, str]:
    """
    Sao chép thẻ <w:p>, kiểm tra và cắt bỏ phần nhãn đầu câu/đáp án khớp với `regex_pattern`
    (ví dụ: 'Câu 12: ', 'A. ', '*A. ', 'a) '), đồng thời kiểm tra xem phần nhãn chữ cái đó
    có được gạch chân hay đánh dấu * hay không.

    Trả về: (new_p_elem, is_marked_correct, matched_label)
    """
    new_p = copy.deepcopy(p_elem)

    # Thu thập danh sách các node văn bản <w:t> kèm thẻ <w:r> chứa nó
    t_nodes: List[Tuple[OxmlElement, OxmlElement, str]] = []
    for child in new_p.iter():
        if child.tag == qn("w:r"):
            for t in child.findall(qn("w:t")):
                txt = t.text or ""
                if txt:
                    t_nodes.append((child, t, txt))

    full_text = "".join(item[2] for item in t_nodes)
    match = re.match(regex_pattern, full_text, flags=re.IGNORECASE)
    if not match:
        return new_p, False, ""

    matched_str = match.group(0)
    label_char = match.group("label") if "label" in match.groupdict() else ""
    has_asterisk = "*" in matched_str

    # Xác định vị trí ký tự chữ cái nhãn (A, B, C, D hoặc a, b, c, d) trong chuỗi đầu
    label_start_idx = match.start("label") if "label" in match.groupdict() else 0
    label_end_idx = match.end("label") if "label" in match.groupdict() else len(matched_str)
    cut_len = len(matched_str)

    is_underlined = False
    curr_pos = 0

    for r_elem, t_elem, txt in t_nodes:
        node_start = curr_pos
        node_end = curr_pos + len(txt)
        curr_pos = node_end

        # Kiểm tra xem ký tự label có nằm trong run này và được gạch chân không
        if node_start < label_end_idx and node_end > label_start_idx:
            if is_run_underlined(r_elem):
                is_underlined = True

        # Cắt bỏ các ký tự thuộc phần prefix (từ 0 đến cut_len)
        if node_end <= cut_len:
            t_elem.text = ""
            remove_underline_from_run(r_elem)
        elif node_start < cut_len < node_end:
            offset = cut_len - node_start
            t_elem.text = txt[offset:].lstrip(" \t")
            t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            # Nếu run này chỉ gạch chân ở ký tự đầu và kéo dài sang phần sau do gõ dính,
            # chỉ xóa gạch chân nếu toàn bộ run là nhãn
            if offset >= len(txt.strip()):
                remove_underline_from_run(r_elem)

    return new_p, (is_underlined or has_asterisk), label_char


def split_inline_options_in_paragraph(p_elem: OxmlElement, part: int = 1) -> Optional[List[Tuple[str, bool, OxmlElement]]]:
    """
    Hỗ trợ trường hợp nhiều phương án A. B. C. D. (hoặc a) b) c) d)) được gõ trên CÙNG MỘT DÒNG
    ngăn cách bởi phím Tab hoặc >= 2 khoảng trắng, trong đó mỗi phương án có thể chứa công thức <m:oMath>.

    Trả về danh sách: [(label_char, is_correct, option_p_elem), ...] hoặc None nếu dòng không bắt đầu bằng nhãn phương án.
    """
    if part == 1:
        label_regex = re.compile(r"^\s*(\*?)\s*([A-D])\s*[\.\:\)]\s*", re.DOTALL)
        split_token_regex = re.compile(r"(?:^|\t|\s{2,})(\*?)\s*([A-D])\s*[\.\:\)]\s*")
    else:
        label_regex = re.compile(r"^\s*(\*?)\s*([a-d])\s*[\)\.]\s*", re.DOTALL)
        split_token_regex = re.compile(r"(?:^|\t|\s{2,})(\*?)\s*([a-d])\s*[\)\.]\s*")

    raw_text = get_paragraph_text(p_elem)
    if not label_regex.match(raw_text):
        return None

    # Để bảo toàn tuyệt đối các node <m:oMath>, <w:drawing>, <w:object> xen kẽ giữa các <w:r>,
    # ta duyệt các phần tử con trực tiếp của <w:p> và tách thành các nhóm khi gặp ranh giới phương án mới.
    pPr = p_elem.find(qn("w:pPr"))

    # Bước 1: Chuẩn hóa các run chứa tab hoặc nhiều phương án gộp chung trong 1 <w:r>
    expanded_children: List[OxmlElement] = []
    for child in p_elem:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag == qn("w:r"):
            # Kiểm tra nếu trong 1 run vừa có text vừa có <w:tab/> hoặc nhiều nhãn A. B. C. D.
            rPr = child.find(qn("w:rPr"))
            sub_nodes = list(child)
            has_tab = any(sn.tag == qn("w:tab") for sn in sub_nodes)
            if has_tab:
                for sn in sub_nodes:
                    if sn.tag == qn("w:rPr"):
                        continue
                    new_r = OxmlElement("w:r")
                    if rPr is not None:
                        new_r.append(copy.deepcopy(rPr))
                    new_r.append(copy.deepcopy(sn))
                    expanded_children.append(new_r)
            else:
                expanded_children.append(copy.deepcopy(child))
        else:
            expanded_children.append(copy.deepcopy(child))

    # Bước 2: Duyệt qua expanded_children để nhận diện điểm bắt đầu của từng phương án A, B, C, D
    segments: List[Tuple[str, bool, List[OxmlElement]]] = []
    current_label: Optional[str] = None
    current_correct: bool = False
    current_nodes: List[OxmlElement] = []
    after_separator = True  # Đầu dòng luôn được coi là sau dấu phân cách

    i = 0
    n = len(expanded_children)
    while i < n:
        node = expanded_children[i]
        if node.tag == qn("w:r"):
            # Nếu là tab run -> bật cờ after_separator
            if node.find(qn("w:tab")) is not None and not node.findall(qn("w:t")):
                after_separator = True
                i += 1
                continue

            # Lấy text của run hiện tại (và ghép thử với run kế tiếp nếu nhãn bị tách làm 2 run kiểu 'A' + '.')
            t_list = node.findall(qn("w:t"))
            txt = "".join((t.text or "") for t in t_list)

            # Kiểm tra nếu trong chính `txt` có nhiều phương án cách nhau bởi >=2 dấu cách
            # (ví dụ: "A. 10    B. 20")
            if current_label is not None and not after_separator:
                m_inner = re.search(
                    r"(\s{2,})(\*?\s*[A-Da-d]\s*[\.\:\)]\s*)",
                    txt,
                )
                if m_inner:
                    before_txt = txt[: m_inner.start(1)]
                    rest_txt = txt[m_inner.start(2) :]
                    if before_txt:
                        r_before = copy.deepcopy(node)
                        for idx_t, t_el in enumerate(r_before.findall(qn("w:t"))):
                            t_el.text = before_txt if idx_t == 0 else ""
                        current_nodes.append(r_before)
                    # Cập nhật node hiện tại thành phần còn lại và bật cờ after_separator
                    for idx_t, t_el in enumerate(node.findall(qn("w:t"))):
                        t_el.text = rest_txt if idx_t == 0 else ""
                    after_separator = True
                    continue

            # Kiểm tra xem tại vị trí này có bắt đầu một phương án mới không
            combined_txt = txt
            next_node = expanded_children[i + 1] if (i + 1 < n and expanded_children[i + 1].tag == qn("w:r")) else None
            next_txt = ""
            if next_node is not None:
                next_txt = "".join((t.text or "") for t in next_node.findall(qn("w:t")))
                combined_txt = txt + next_txt

            m_opt = label_regex.match(combined_txt) if (after_separator or txt.startswith(("  ", "\t"))) else None
            if m_opt:
                # Lưu phương án trước đó (nếu có)
                if current_label is not None:
                    segments.append((current_label, current_correct, current_nodes))

                asterisk = m_opt.group(1)
                lbl = m_opt.group(2)
                matched_len = len(m_opt.group(0))

                # Kiểm tra gạch chân trên node chứa chữ cái nhãn
                is_ul = is_run_underlined(node)
                current_label = lbl.upper() if part == 1 else lbl.lower()
                current_correct = bool(asterisk) or is_ul
                current_nodes = []
                after_separator = False

                # Cắt bỏ phần nhãn khỏi node (và next_node nếu nhãn trải qua 2 run)
                if len(txt) >= matched_len:
                    rem = txt[matched_len:]
                    if rem:
                        r_rem = copy.deepcopy(node)
                        remove_underline_from_run(r_rem)
                        for idx_t, t_el in enumerate(r_rem.findall(qn("w:t"))):
                            t_el.text = rem if idx_t == 0 else ""
                            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        current_nodes.append(r_rem)
                    i += 1
                else:
                    # Nhãn nằm vắt sang next_node
                    take_from_next = matched_len - len(txt)
                    rem_next = next_txt[take_from_next:]
                    if rem_next and next_node is not None:
                        r_rem = copy.deepcopy(next_node)
                        remove_underline_from_run(r_rem)
                        for idx_t, t_el in enumerate(r_rem.findall(qn("w:t"))):
                            t_el.text = rem_next if idx_t == 0 else ""
                            t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                        current_nodes.append(r_rem)
                    i += 2
                continue

            after_separator = txt.endswith(("  ", "\t"))
            current_nodes.append(node)
            i += 1
        else:
            # Công thức toán <m:oMath>, hình ảnh <w:drawing>, v.v.
            after_separator = False
            current_nodes.append(node)
            i += 1

    if current_label is not None:
        segments.append((current_label, current_correct, current_nodes))

    if not segments:
        return None

    result: List[Tuple[str, bool, OxmlElement]] = []
    for lbl, is_corr, nodes in segments:
        opt_p = OxmlElement("w:p")
        if pPr is not None:
            opt_p.append(copy.deepcopy(pPr))
        for nd in nodes:
            opt_p.append(nd)
        result.append((lbl, is_corr, opt_p))

    return result


def prepend_bold_label_to_paragraph(p_elem: OxmlElement, label_text: str, color_hex: Optional[str] = None) -> OxmlElement:
    """
    Chèn một nhãn in đậm (ví dụ: 'Câu 1: ' hoặc 'A. ' hoặc 'a) ') vào đầu thẻ <w:p>
    mà giữ nguyên toàn bộ công thức, hình ảnh và định dạng phần thân của <w:p>.
    """
    new_p = copy.deepcopy(p_elem)
    r_label = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    b_el = OxmlElement("w:b")
    rPr.append(b_el)

    # Giữ font chữ đồng bộ với đoạn văn hoặc dùng Times New Roman 12pt chuẩn đề thi Việt Nam
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

    # Chèn ngay sau <w:pPr> (nếu có), hoặc ở vị trí đầu tiên của <w:p>
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
    current_part = 1  # Mặc định Phần I nếu đề không ghi tiêu đề PHẦN I
    current_q: Optional[QuestionBlock] = None

    def finalize_question(q: Optional[QuestionBlock]) -> None:
        if q is None:
            return
        if q.part == 1:
            # Nếu chưa có đáp án nào được đánh dấu đúng (do người soạn quên gạch chân),
            # mặc định phương án đầu tiên và ghi cảnh báo
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
            # Với Phần II: hỗ trợ thêm trường hợp giáo viên ghi (Đ)/(S) hoặc [Đ]/[S] ở cuối mệnh đề
            for opt in q.options:
                if opt.xml_elements:
                    last_p = opt.xml_elements[-1]
                    txt = get_paragraph_text(last_p).strip()
                    m_ds = re.search(r"[\(\[](Đ|Đúng|S|Sai)[\)\]]\s*\.?\s*$", txt, flags=re.IGNORECASE)
                    if m_ds:
                        val = m_ds.group(1).upper()
                        opt.is_correct = val.startswith("Đ")
                        # Xóa ký hiệu (Đ)/(S) ở cuối để học sinh không thấy
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
            # Nếu chưa tìm thấy dòng "Đáp án: ...", thử tìm text được gạch chân trong thân câu hỏi
            if not q.short_answer:
                underlined_chunks = []
                for el in q.stem_elements:
                    if el.tag == qn("w:p"):
                        for r_el in el.findall(qn("w:r")):
                            if is_run_underlined(r_el):
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
            if not text and len(child.findall(".//" + qn("w:drawing"))) == 0 and len(child.findall(".//" + qn("m:oMath"))) == 0 and len(child.findall(".//" + qn("w:object"))) == 0:
                continue

            if END_MARKER_RE.match(text):
                break

            # Kiểm tra chuyển Phần thi (PHẦN I, PHẦN II, PHẦN III)
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

            # Kiểm tra bắt đầu câu hỏi mới ("Câu 1:", "Câu 2.", ...)
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

            # Xử lý bên trong một câu hỏi đang mở
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

                # Nếu chưa có option nào -> đoạn văn/hình ảnh này thuộc phần thân câu hỏi (stem)
                if not current_q.options:
                    current_q.stem_elements.append(copy.deepcopy(child))
                else:
                    # Nếu đã có option -> đoạn văn/hình ảnh này là dòng tiếp theo của option vừa mở
                    current_q.options[-1].xml_elements.append(copy.deepcopy(child))

            elif current_part == 3:
                m_ans = SHORT_ANS_LINE_RE.match(text)
                if m_ans:
                    current_q.short_answer = m_ans.group("ans").strip().rstrip(".")
                else:
                    current_q.stem_elements.append(copy.deepcopy(child))

        elif child.tag == qn("w:tbl"):
            # Nếu trong câu hỏi hoặc phương án có bảng biểu (<w:tbl>), giữ nguyên toàn bộ bảng
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
    school_name: str = "TRƯỜNG THPT CHUYÊN"
    exam_name: str = "ĐỀ KIỂM TRA ĐỊNH KỲ NĂM HỌC 2026 - 2027"
    subject_name: str = "Môn: TOÁN - Lớp 12"
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
    part1_keys: List[str]          # Danh sách đáp án đúng ['A', 'C', 'B', ...] theo thứ tự câu mới
    part2_keys: List[str]          # Danh sách chuỗi Đ/S kiểu 'Đ-S-Đ-S' cho từng câu ở Phần II
    part2_detail: List[Dict[str, str]]  # [{'a': 'Đ', 'b': 'S', 'c': 'Đ', 'd': 'S'}, ...]
    part3_keys: List[str]          # Danh sách đáp án trả lời ngắn cho từng câu ở Phần III
    question_mapping: Dict[str, List[int]]  # Lưu vết câu mới ứng với câu gốc số mấy


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
    """
    Tạo một file .docx cho mã đề `exam_code` bằng cách:
    - Mở một bản sao của chính file `source_docx_bytes` (nhờ vậy giữ nguyên 100% media/hình ảnh,
      OLE MathType, style, numbering, font table và relationships).
    - Xóa các đoạn văn cũ trong `<w:body>` (trừ `<w:sectPr>`).
    - Đóng khung tiêu đề đề thi chuẩn Bộ GD&ĐT và gắn các khối câu hỏi/phương án đã đảo.
    """
    doc = docx.Document(io.BytesIO(source_docx_bytes))
    body = doc.element.body

    # Giữ lại cấu hình trang <w:sectPr> của file gốc, xóa các phần tử nội dung cũ
    sectPr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child is not sectPr:
            body.remove(child)

    # Thiết lập lề trang chuẩn A4 nếu cần
    for section in doc.sections:
        section.top_margin = Inches(0.65)
        section.bottom_margin = Inches(0.65)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.7)

    # 1. Tạo bảng Header chuẩn đề thi THPT Quốc Gia / Sở GD&ĐT
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

    # Dòng Họ và tên thí sinh / Số báo danh
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

    # -----------------------------------------------------------------------
    # PHẦN I: TRẮC NGHIỆM 4 LỰA CHỌN (A, B, C, D)
    # -----------------------------------------------------------------------
    if parsed.part1_questions:
        p_h1 = doc.add_paragraph()
        p_h1.paragraph_format.space_before = Pt(6)
        p_h1.paragraph_format.space_after = Pt(4)
        r_h1 = p_h1.add_run(
            f"PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. "
        )
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

            # Gắn thân câu hỏi (Stem)
            for s_idx, stem_el in enumerate(q.stem_elements):
                if s_idx == 0 and stem_el.tag == qn("w:p"):
                    new_stem = prepend_bold_label_to_paragraph(stem_el, f"Câu {idx}: ")
                    append_element_to_body(new_stem)
                else:
                    append_element_to_body(copy.deepcopy(stem_el))

            # Trộn các phương án A, B, C, D trong câu
            opts = list(q.options)
            if meta.shuffle_part1_options and len(opts) > 1:
                # Giữ các phương án dạng "Cả A, B, C đều đúng" ở vị trí cố định nếu có
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

    # -----------------------------------------------------------------------
    # PHẦN II: TRẮC NGHIỆM ĐÚNG / SAI (a, b, c, d)
    # -----------------------------------------------------------------------
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
                    new_stem = prepend_bold_label_to_paragraph(stem_el, f"Câu {idx}: ")
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

    # -----------------------------------------------------------------------
    # PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN
    # -----------------------------------------------------------------------
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
                    new_stem = prepend_bold_label_to_paragraph(stem_el, f"Câu {idx}: ")
                    append_element_to_body(new_stem)
                else:
                    append_element_to_body(copy.deepcopy(stem_el))
            part3_keys.append(q.short_answer)

    # Dòng chữ ------ HẾT ------ ở cuối đề
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
    """
    Tạo file Excel (.xlsx) bảng đáp án tổng hợp tất cả mã đề, gồm 3 Sheet:
    - Sheet 1: 'Bang_Dap_An_Tong_Hop' (Dạng cột theo từng Mã đề cho cả 3 Phần, trình bày chuyên nghiệp để in dán bảng hoặc tra cứu).
    - Sheet 2: 'Chi_Tiet_Cham_May' (Dạng hàng ngang phẳng chuẩn phần mềm chấm trắc nghiệm bằng máy / điện thoại).
    - Sheet 3: 'Ma_Tran_Dao_Cau' (Bảng đối chiếu câu hỏi của từng mã đề ứng với câu gốc số mấy).
    """
    wb = openpyxl.Workbook()

    # Styles
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

    # -----------------------------------------------------------------------
    # SHEET 1: BẢNG ĐÁP ÁN TỔNG HỢP
    # -----------------------------------------------------------------------
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

    # Phần I
    num_p1 = len(variants[0].part1_keys) if variants else 0
    if num_p1 > 0:
        row_cursor = write_table_header(
            ws1,
            row_cursor,
            "PHẦN I: TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN (Mỗi câu 0,25 điểm)",
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

    # Phần II
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

    # Phần III
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

    # -----------------------------------------------------------------------
    # SHEET 2: CHI TIẾT CHẤM MÁY (Dạng bảng ngang chuẩn import phần mềm chấm)
    # -----------------------------------------------------------------------
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

    # -----------------------------------------------------------------------
    # SHEET 3: MA TRẬN ĐỐI CHIẾU CÂU GỐC
    # -----------------------------------------------------------------------
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
    """Sinh danh sách mã đề (ví dụ: 101, 102, 103, 104...)."""
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
    """
    Hàm tổng hợp:
    1. Phân tích đề gốc (.docx)
    2. Trộn tạo danh sách các mã đề `.docx`
    3. Sinh bảng đáp án chấm thi `.xlsx`
    4. Đóng gói tất cả vào bộ nhớ đệm `.zip` (in-memory) để trả thẳng về trình duyệt.
    """
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
        # Lưu đề gốc đã chuẩn hóa vào thư mục De_Goc
        zf.writestr("00_De_Goc_Tham_Chieu.docx", source_docx_bytes)
        # Lưu từng mã đề Word
        for v in variants:
            zf.writestr(f"Ma_De_{v.exam_code}.docx", v.docx_bytes)
        # Lưu file Excel đáp án
        zf.writestr("Bang_Dap_An_Cham_Thi.xlsx", excel_bytes)

    return zip_buf.getvalue(), parsed, variants
