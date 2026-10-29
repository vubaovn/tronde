# -*- coding: utf-8 -*-
"""
Backend Web Server Trộn Đề Trắc Nghiệm Trực Tuyến (Chuẩn GDPT 2018 - 3 Phần).
- Không cần đăng nhập, hoàn toàn miễn phí.
- Xử lý trực tiếp trên bộ nhớ RAM (in-memory), không lưu trữ đề thi của giáo viên trên ổ cứng server để bảo mật và tiết kiệm dung lượng hosting.
- Tương thích triển khai trên: Render.com, Railway, PythonAnywhere (WSGI), VPS Ubuntu/Nginx, hoặc chạy trực tiếp trên máy tính cá nhân.
"""

import base64
import io
import os
from flask import Flask, jsonify, render_template, request, send_file

from shuffler import (
    ExamMetadata,
    generate_exam_codes,
    parse_docx_exam,
    shuffle_and_package_zip,
    get_paragraph_text,
)

app = Flask(__name__, static_folder="static", template_folder="templates")
# Giới hạn dung lượng file Word tải lên tối đa 25MB (đủ cho đề có nhiều hình ảnh/công thức)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024


@app.route("/", methods=["GET"])
def index():
    """Trang chủ giao diện trộn đề trực tuyến."""
    return render_template("index.html")


@app.route("/api/sample-docx", methods=["GET"])
def download_sample_docx():
    """Tải về file đề gốc mẫu chuẩn 3 phần (.docx) có sẵn công thức Toán và hình ảnh."""
    sample_path = os.path.join(app.static_folder, "de_mau_chuan_2025.docx")
    if not os.path.exists(sample_path):
        from create_sample_exam import build_sample_exam_docx
        build_sample_exam_docx(sample_path)
    return send_file(
        sample_path,
        as_attachment=True,
        download_name="De_Goc_Mau_Chuan_3_Phan.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/api/analyze", methods=["POST"])
def analyze_exam():
    """
    API phân tích nhanh file Word (.docx) vừa tải lên để hiển thị:
    - Số lượng câu hỏi từng phần (Phần I, Phần II, Phần III).
    - Đáp án gốc nhận diện được của từng câu.
    - Cảnh báo nếu có câu nào quên gạch chân đáp án đúng.
    """
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Vui lòng chọn file Word (.docx) đề gốc."}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".docx"):
        return jsonify({"ok": False, "error": "Hệ thống chỉ nhận định dạng file Microsoft Word (.docx)."}), 400

    try:
        docx_bytes = f.read()
        parsed = parse_docx_exam(docx_bytes)

        def summarize_q(q):
            stem_preview = " ".join(get_paragraph_text(el).strip() for el in q.stem_elements).strip()
            if len(stem_preview) > 140:
                stem_preview = stem_preview[:140] + "..."
            if q.part == 1:
                ans = next((o.original_label for o in q.options if o.is_correct), "?")
            elif q.part == 2:
                ans = "-".join(("Đ" if o.is_correct else "S") for o in q.options)
            else:
                ans = q.short_answer or "(Chưa có)"
            return {
                "index": q.original_index,
                "stem": stem_preview,
                "option_count": len(q.options),
                "detected_key": ans,
            }

        return jsonify({
            "ok": True,
            "counts": {
                "part1": len(parsed.part1_questions),
                "part2": len(parsed.part2_questions),
                "part3": len(parsed.part3_questions),
                "total": len(parsed.part1_questions) + len(parsed.part2_questions) + len(parsed.part3_questions),
            },
            "part1": [summarize_q(q) for q in parsed.part1_questions],
            "part2": [summarize_q(q) for q in parsed.part2_questions],
            "part3": [summarize_q(q) for q in parsed.part3_questions],
            "warnings": parsed.warnings,
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Không thể đọc cấu trúc file Word: {exc}"}), 500


@app.route("/api/shuffle", methods=["POST"])
def shuffle_exam_api():
    """
    API nhận file .docx + cấu hình trộn đề -> Trả về JSON gồm:
    - File nén .zip (base64) chứa các file Word mã đề + file Excel đáp án.
    - Bảng đáp án của tất cả mã đề để hiển thị trực tiếp trên giao diện web.
    """
    if "file" not in request.files:
        return jsonify({"ok": False, "error": "Vui lòng chọn file Word (.docx) đề gốc."}), 400

    f = request.files["file"]
    if not f.filename or not f.filename.lower().endswith(".docx"):
        return jsonify({"ok": False, "error": "Vui lòng tải lên file định dạng .docx."}), 400

    try:
        docx_bytes = f.read()

        num_variants = int(request.form.get("num_variants", 4))
        starting_code = int(request.form.get("starting_code", 101))
        custom_codes_raw = request.form.get("custom_codes", "").strip()
        custom_codes = [c.strip() for c in custom_codes_raw.split(",") if c.strip()] if custom_codes_raw else None

        exam_codes = generate_exam_codes(
            num_variants=num_variants,
            starting_code=starting_code,
            custom_codes=custom_codes,
        )

        meta = ExamMetadata(
            governing_body=request.form.get("governing_body", "SỞ GD&ĐT BẮC NINH").strip(),
            school_name=request.form.get("school_name", "TRƯỜNG THPT CHUYÊN").strip(),
            exam_name=request.form.get("exam_name", "ĐỀ KIỂM TRA ĐỊNH KỲ NĂM HỌC 2026 - 2027").strip(),
            subject_name=request.form.get("subject_name", "Môn: TOÁN - Lớp 12").strip(),
            duration_minutes=request.form.get("duration_minutes", "90 phút (không kể thời gian phát đề)").strip(),
            shuffle_part1_questions=request.form.get("shuffle_part1_questions", "true").lower() == "true",
            shuffle_part1_options=request.form.get("shuffle_part1_options", "true").lower() == "true",
            shuffle_part2_questions=request.form.get("shuffle_part2_questions", "true").lower() == "true",
            shuffle_part2_options=request.form.get("shuffle_part2_options", "true").lower() == "true",
            shuffle_part3_questions=request.form.get("shuffle_part3_questions", "true").lower() == "true",
        )

        zip_bytes, parsed, variants = shuffle_and_package_zip(
            source_docx_bytes=docx_bytes,
            exam_codes=exam_codes,
            meta=meta,
        )

        total_q = len(parsed.part1_questions) + len(parsed.part2_questions) + len(parsed.part3_questions)
        if total_q == 0:
            return jsonify({
                "ok": False,
                "error": "Không tìm thấy câu hỏi nào trong file. Hãy đảm bảo mỗi câu bắt đầu bằng 'Câu 1:', 'Câu 2:'..."
            }), 400

        variants_summary = []
        for v in variants:
            variants_summary.append({
                "exam_code": v.exam_code,
                "part1_keys": v.part1_keys,
                "part2_keys": v.part2_keys,
                "part3_keys": v.part3_keys,
            })

        return jsonify({
            "ok": True,
            "zip_base64": base64.b64encode(zip_bytes).decode("ascii"),
            "zip_filename": f"Bo_De_Tron_{len(variants)}_Ma_De.zip",
            "counts": {
                "part1": len(parsed.part1_questions),
                "part2": len(parsed.part2_questions),
                "part3": len(parsed.part3_questions),
                "total": total_q,
            },
            "warnings": parsed.warnings,
            "variants": variants_summary,
        })
    except Exception as exc:
        return jsonify({"ok": False, "error": f"Lỗi trong quá trình trộn đề: {exc}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
