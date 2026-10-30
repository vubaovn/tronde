# HƯỚNG DẪN ĐƯA WEBSITE TRỘN ĐỀ TRẮC NGHIỆM PYTHON LÊN HOSTING & TRỎ TÊN MIỀN

Bộ mã nguồn này được thiết kế hoàn chỉnh để bạn có thể chạy theo **2 chế độ**:
1. **Chế độ Full-Stack Python (`app.py` + `shuffler.py`)**: Chạy trên Python Hosting (Render, Railway, PythonAnywhere, cPanel hỗ trợ Setup Python App, hoặc VPS Ubuntu). Xử lý ở cấp độ hạt nhân OpenXML (`lxml` + `python-docx`), bảo toàn **100% công thức Toán/Lý/Hóa (MathType OLE, Word Equation OMML)**, hình ảnh đồ thị, chỉ số trên/dưới và bảng biểu.
2. **Chế độ Trình duyệt Độc lập (`templates/index.html`)**: Giao diện của ứng dụng còn tích hợp sẵn bộ giải nén & trộn cấu trúc XML của file `.docx` ngay trên trình duyệt (Client-side). Nhờ vậy, ngay cả khi bạn dùng thử trực tiếp hoặc mở file HTML độc lập, mọi nút bấm và tính năng kéo thả `.docx` -> xuất `.zip` (`.docx` + `.xlsx`) đều hoạt động tức thì.

---

## PHẦN 1: CẤU TRÚC THƯ MỤC DỰ ÁN

```text
tron_de_online/
├── app.py                      # Server Python (Flask) xử lý API phân tích & trộn đề
├── shuffler.py                 # Thuật toán lõi trộn XML Word (.docx) & xuất Excel (.xlsx)
├── create_sample_exam.py       # Script sinh file đề mẫu chuẩn 3 phần có công thức & đồ thị
├── requirements.txt            # Danh sách thư viện Python cần cài đặt
├── Procfile                    # Cấu hình khởi chạy Gunicorn cho Render / Railway / Heroku
├── Dockerfile                  # Cấu hình đóng gói Docker container cho VPS / Cloud Run
├── render.yaml                 # Cấu hình tự động triển khai 1-click trên Render.com
├── static/
│   └── de_mau_chuan_2025.docx  # File đề gốc mẫu chuẩn GDPT 2018 để người dùng tải thử
└── templates/
    └── index.html              # Giao diện Web App (hỗ trợ cả Python API lẫn Client-side)
```

---

## PHẦN 2: QUY ƯỚC SOẠN ĐỀ GỐC WORD (`.DOCX`) CHUẨN 3 PHẦN

Phần mềm nhận diện tự động 3 phần theo đúng cấu trúc đề thi tốt nghiệp THPT mới (Chương trình GDPT 2018):

1. **PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn (A, B, C, D)**
   - Dòng tiêu đề bắt đầu bằng: `PHẦN I` (hoặc `PHẦN 1`).
   - Mỗi câu hỏi bắt đầu bằng: `Câu 1:`, `Câu 2.`, ...
   - Các phương án bắt đầu bằng `A.`, `B.`, `C.`, `D.` (có thể để trên 4 dòng riêng biệt hoặc gõ chung trên 1–2 dòng cách nhau bởi phím `Tab` hoặc từ 2 dấu cách trở lên).
   - **Đánh dấu đáp án đúng**: **Gạch chân** chữ cái đầu của phương án đúng (ví dụ: <u>**A**</u>. hoặc <u>**A.**</u>). Hệ thống sẽ tự động xóa gạch chân này khi xuất ra đề thi cho học sinh và lưu vị trí mới vào file Excel.
   - *Mẹo*: Nếu câu có phương án *"Cả A, B, C đều đúng"* hoặc *"Tất cả đều đúng"*, hệ thống sẽ tự động giữ cố định phương án đó ở vị trí D cuối cùng.

2. **PHẦN II. Câu trắc nghiệm đúng sai (a, b, c, d)**
   - Dòng tiêu đề bắt đầu bằng: `PHẦN II` (hoặc `PHẦN 2`).
   - Mỗi câu hỏi bắt đầu bằng: `Câu 1:`, `Câu 2:`, ...
   - 4 mệnh đề con bắt đầu bằng: `a)`, `b)`, `c)`, `d)` (hoặc `a.`, `b.`, `c.`, `d.`).
   - **Đánh dấu mệnh đề Đúng/Sai**:
     - Mệnh đề nào **ĐÚNG (Đ)** thì **gạch chân** chữ cái đầu (ví dụ: <u>**a**</u>) hoặc <u>**a)**</u>).
     - Mệnh đề nào **SAI (S)** thì để bình thường **không gạch chân**.
     - Khi trộn đề, cả thứ tự các câu trong Phần II và thứ tự 4 ý `a), b), c), d)` bên trong mỗi câu đều được đảo ngẫu nhiên và tự động gắn lại nhãn `a), b), c), d)` theo thứ tự mới, đồng thời cập nhật chuỗi đáp án `Đ - S - Đ - S` vào bảng Excel.

3. **PHẦN III. Câu trắc nghiệm trả lời ngắn**
   - Dòng tiêu đề bắt đầu bằng: `PHẦN III` (hoặc `PHẦN 3`).
   - Mỗi câu hỏi bắt đầu bằng: `Câu 1:`, `Câu 2:`, ...
   - **Đánh dấu đáp án trả lời ngắn**: Thêm một dòng ở cuối câu hỏi theo cú pháp `Đáp án: 40` (hoặc `ĐS: -2,5`). Khi xuất đề cho học sinh, phần mềm sẽ **tự động ẩn dòng Đáp án này đi** và chỉ đưa vào bảng đáp án Excel chấm thi.

---

## PHẦN 3: HƯỚNG DẪN CHỌN HOSTING & ĐƯA WEB LÊN MẠNG (ĐÃ CÓ TÊN MIỀN)

Vì bạn đã có sẵn **tên miền (domain)** và muốn chạy bằng **Python**, dưới đây là 3 phương án tốt nhất từ **Miễn phí** đến **Hosting có cPanel / VPS**:

### Phương án A: Triển khai MIỄN PHÍ 100% lên Render.com (Khuyên dùng)
Render.com hỗ trợ chạy code Python (Flask/Gunicorn) và cho phép gắn tên miền riêng (**Custom Domain**) kèm chứng chỉ bảo mật **SSL (HTTPS) miễn phí**.

1. **Bước 1**: Tạo một tài khoản miễn phí tại [github.com](https://github.com) và tải toàn bộ các file trong thư mục dự án này lên một Repository mới (ví dụ: `tron-de-online`).
2. **Bước 2**: Truy cập [render.com](https://render.com), đăng nhập bằng tài khoản GitHub vừa tạo.
3. **Bước 3**: Bấm nút **New +** -> chọn **Web Service** -> chọn Repository `tron-de-online` của bạn.
4. **Bước 4**: Điền thông số cấu hình:
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 120`
   - **Instance Type**: Chọn gói **Free** (hoặc gói Starter $7/tháng nếu muốn server luôn thức 24/24 không nghỉ).
5. **Bước 5 (Trỏ tên miền của bạn về Render)**:
   - Trong trang quản lý dịch vụ trên Render, vào mục **Settings** -> cuộn xuống phần **Custom Domains** -> bấm **Add Custom Domain** và nhập tên miền của bạn (ví dụ: `tronde.tenmiencuaban.com` hoặc `tenmiencuaban.com`).
   - Mở trang quản trị DNS nơi bạn mua tên miền (Tenten, Mắt Bão, Nhân Hòa, P.A Việt Nam, Cloudflare...), thêm bản ghi theo hướng dẫn của Render:
     - Loại `CNAME`: Tên `@` hoặc `www` -> Giá trị trỏ về địa chỉ `<ten-app-cua-ban>.onrender.com`.
     - Loại `A`: Tên `@` -> Giá trị `216.24.57.1`.
   - Chờ 5–15 phút để hệ thống kích hoạt HTTPS tự động là xong!

---

### Phương án B: Mua Shared Hosting có hỗ trợ cPanel ("Setup Python App") tại Việt Nam
Nếu bạn muốn mua gói Hosting tại các nhà cung cấp Việt Nam (như AZDIGI, Hostvn, TinoHost, Nhân Hòa, Mắt Bão...), khi mua hãy lưu ý: **Chọn gói Hosting có cPanel hỗ trợ "Setup Python App" (CloudLinux Passenger)**.

1. **Bước 1**: Trỏ tên miền của bạn về IP của Hosting (bản ghi `A`).
2. **Bước 2**: Đăng nhập **cPanel** -> vào **File Manager** -> tạo một thư mục `tron_de_online` ở thư mục gốc và tải file nén mã nguồn lên, giải nén ra.
3. **Bước 3**: Quay lại màn hình chính cPanel -> tìm biểu tượng **Setup Python App** -> bấm **Create Application**:
   - **Python version**: Chọn `3.10` hoặc `3.11`.
   - **Application root**: `tron_de_online`
   - **Application URL**: Chọn tên miền của bạn.
   - **Application startup file**: `app.py`
   - **Application Entry point**: `app`
   - Bấm **Create**.
4. **Bước 4**: Ngay trong màn hình cấu hình Python App, mục **Configuration files**, gõ `requirements.txt` -> bấm **Add** -> bấm nút **Run Pip Install** để hệ thống tự cài `flask`, `python-docx`, `openpyxl`, `lxml`.
5. **Bước 5**: Bấm **Restart** ứng dụng là website chính thức hoạt động trên tên miền của bạn!

---

### Phương án C: Chạy trực tiếp trên máy tính cá nhân (Localhost) để kiểm tra trước
Nếu bạn muốn chạy thử ngay trên máy tính Windows/Mac của mình:
```bash
# 1. Cài đặt các thư viện cần thiết
pip install -r requirements.txt

# 2. Khởi chạy server
python app.py
```
Sau đó mở trình duyệt truy cập `http://localhost:5000`.
