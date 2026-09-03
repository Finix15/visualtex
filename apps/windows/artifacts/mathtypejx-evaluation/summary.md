# Đánh giá mathtypejx — tài liệu thực tế 147 OLE

## Kết luận

Tài liệu hiện có chứa **147** OLE `Equation.DSMT4`, 147 stream đều được `mathtypejx` xác định là **MTEF v5**; không có stream MTEF v3. Word vẫn có thể hiện hộp thoại “Equation Editor 3.0” cho nhóm OLE toán cũ này, nên nhãn giao diện Word không đồng nghĩa với phiên bản MTEF trong stream.

Tài liệu nguồn không bị ghi đè. SHA-256 trước và sau lượt tuần tự đều là `daa4fa36ed4321d86dd6c51b63a27cc2d110f543725204e76b178c3ad188db07`.

## Môi trường

- mathtypejx commit: `7d90e7274c85cf56ac28d4d15e593044693d7e70`
- Package version: `0.1.0`
- License: MIT, Copyright (c) 2026 a917470154
- Python: 3.12.13 x64, virtual environment riêng
- lxml: 6.1.3
- olefile: 0.47
- python-docx: 1.2.0
- pytest: 9.1.1
- pytest-cov: 7.1.0
- Office XSLT: `C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL`
- `mathtypejx health`: READY
- Upstream tests: 69 passed, 12 skipped

## Kết quả chuyển đổi

- Đầu vào phát hiện: 147 MathType OLE
- Chuyển/thay: 116
- Thất bại và giữ OLE: 31
- OLE còn lại trong đầu ra: 31
- OMML trong đầu ra: 116
- Thời gian: 20.685 giây
- `remove_edit_info=False`
- `parallel=False`

Lượt 4 luồng (`parallel=True`, `max_workers=4`) cũng chuyển 116 và giữ lại 31, hoàn thành trong 8.496 giây. Đối chiếu từng formula ID cho thấy trạng thái, risk, cảnh báo, lỗi, MathML chuẩn hóa và OMML chuẩn hóa giống hoàn toàn với lượt tuần tự (0 khác biệt).

31 công thức được giữ nguyên đều do cổng chất lượng báo nguy cơ mất delimiter `()`, `|` hoặc `{`; không ép thay các công thức này.

Hai DOCX đầu ra đều vượt qua kiểm tra ZIP/OOXML, giữ nguyên hash văn bản, media và embedding giữa hai lượt. Tài liệu không được render tự động vì môi trường không có LibreOffice; cần mở `converted-serial.docx` trong Word để duyệt hình ảnh cuối cùng.

## Phạm vi còn lại

Validator delimiter đã được sửa trong bản clone thử nghiệm để đếm dấu co giãn lưu bằng `m:d`, `m:begChr` và `m:endChr`, bao gồm cặp ngoặc mặc định của OMML. Ba regression test mới bao phủ ngoặc mặc định, dấu ngoặc nhọn một phía và cặp dấu `|`.

- Toàn bộ test sau vá: 72 passed, 12 skipped.
- Kết quả mới: 147/147 OLE được chuyển, 0 thất bại, 0 OLE còn lại, 147 OMML.
- ZIP/OOXML hợp lệ và hash nội dung văn bản được giữ nguyên.
- File nguồn vẫn giữ nguyên SHA-256.

Đầu ra mới là `converted-delimiter-fixed.docx`. Cần duyệt trực quan trong Word, đặc biệt 31 công thức trước đây bị chặn, trước khi tích hợp bản vá vào VisualTeX.
