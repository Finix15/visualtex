<div align="center">
  <img src="apps/macos/src-tauri/app-icon.svg" width="200" alt="Biểu tượng VisualTeX" />
  <h1>VisualTeX</h1>
  <p><strong>Trình soạn thảo công thức trực quan tích hợp nguyên bản với Microsoft Office</strong></p>
  <p>
    <a href="https://visualtex.pauljianliao.com/">Trang chủ</a> ·
    <a href="https://github.com/Finix15/visualtex/releases">Tải xuống</a> ·
    <a href="huongdan.txt">Hướng dẫn cài đặt</a>
  </p>
  <p>
    <a href="https://github.com/Finix15/visualtex/stargazers"><img src="https://img.shields.io/github/stars/Finix15/visualtex?style=for-the-badge&logo=github&label=SAO" alt="Số sao GitHub" /></a>
    <a href="https://github.com/Finix15/visualtex/releases/latest"><img src="https://img.shields.io/github/v/release/Finix15/visualtex?style=for-the-badge&logo=github&label=PHIÊN%20BẢN" alt="Phiên bản mới nhất" /></a>
    <a href="https://github.com/Finix15/visualtex/releases"><img src="https://img.shields.io/github/downloads/Finix15/visualtex/total?style=for-the-badge&logo=github&label=LƯỢT%20TẢI" alt="Tổng lượt tải" /></a>
  </p>
  <p>
    <img src="https://img.shields.io/badge/WORD-HỖ_TRỢ-0099E5?style=for-the-badge&labelColor=555555" alt="Hỗ trợ Word" />
    <img src="https://img.shields.io/badge/POWERPOINT-HỖ_TRỢ-0099E5?style=for-the-badge&labelColor=555555" alt="Hỗ trợ PowerPoint" />
    <img src="https://img.shields.io/badge/LaTeX-NHẬP_LIỆU-0099E5?style=for-the-badge&labelColor=555555" alt="Nhập liệu LaTeX" />
    <img src="https://img.shields.io/badge/GIẤY_PHÉP-MIT-55B800?style=for-the-badge&labelColor=555555" alt="Giấy phép MIT" />
  </p>
</div>

---

VisualTeX là ứng dụng soạn thảo công thức dành cho toán học, vật lý, kỹ thuật và viết tài liệu nghiên cứu. Ứng dụng kết hợp nhập liệu trực quan có cấu trúc, đồng bộ hai chiều với mã nguồn LaTeX, nhận dạng công thức từ ảnh ngay trên máy và quy trình làm việc nguyên bản trong Word, PowerPoint.

Tiện ích Word hỗ trợ vẽ lại hàng loạt mã LaTeX trong vùng chọn hoặc toàn bộ tài liệu thành công thức nguyên bản hay công thức VisualTeX có thể chỉnh sửa lại. Ứng dụng cũng cung cấp đánh số, tham chiếu chéo và nhập tài liệu.

<div align="center">
  <p><strong>Nhóm trao đổi VisualTeX trên QQ: <code>1045801770</code></strong></p>
  <p>Chào đón mọi trao đổi về cách sử dụng, đề xuất tính năng, tiện ích Office và phát triển dự án.</p>
  <img src="apps/macos/public/qq-group-card.svg" width="300" alt="Nhóm QQ VisualTeX 1045801770" />
</div>

## Tải và cài đặt

Tải ứng dụng cùng các mô hình OCR PP-FormulaNet tại **[trang chủ VisualTeX](https://visualtex.pauljianliao.com/)** hoặc [GitHub Releases](https://github.com/Finix15/visualtex/releases).

| Hệ điều hành | Bộ cài | Yêu cầu chính |
| --- | --- | --- |
| Windows | `.exe` | Windows 10/11, 64-bit |
| macOS | `.dmg` | macOS 11 trở lên, Apple Silicon |

Xem [hướng dẫn cài đặt chi tiết](huongdan.txt) để biết cách chọn đúng tệp, xử lý cảnh báo bảo mật và cài tích hợp Office.

## Điểm mới trong VisualTeX 1.2.5

- **Chế độ bàn phím số:** bố cục nhập công thức nhỏ gọn; chế độ thường và bàn phím số ghi nhớ kích thước cửa sổ riêng.
- **Căn chỉnh và chỉnh sửa hàm từng phần:** dùng trực tiếp `&` trong `align` / `aligned` để chọn điểm căn chỉnh; `cases` được cải thiện về nhiều dòng, dấu ngoặc, căn chỉnh và phím Enter.
- **Phím tắt và nhập nhanh tùy chỉnh:** gán phím tắt lâu dài cho công cụ công thức, công cụ thường dùng và ô tùy chỉnh; hỗ trợ nhập nhanh chữ cái Hy Lạp.
- **Trình thiết kế ký hiệu tùy chỉnh:** vẽ, kết hợp và đăng ký ký hiệu toán học để dùng trong thanh công cụ, tự động hoàn thành, MathLive và xuất SVG/PNG.
- **Tùy chỉnh giao diện đầy đủ:** quản lý thống nhất màu nền, trang giấy, lớp bề mặt, trạng thái chọn, màu nhấn và con trỏ công thức.
- **OCR nhanh và OCR nền:** nhận dạng từ ảnh chụp màn hình hoặc chạy nền rồi ghi LaTeX vào bảng nhớ tạm, kèm thông báo trạng thái.
- **Tối ưu trình soạn thảo Office:** duy trì hoặc khởi động sẵn cửa sổ Word/PowerPoint để giảm thời gian chờ.
- **Cài đặt, cập nhật và sửa chữa tiện ích Office:** phát hiện DOTM/PPAM bị thiếu, lỗi thời, không đầy đủ hoặc chưa đăng ký.
- **Làm lại hệ thống đánh số và tham chiếu chéo trong Word:** hỗ trợ số thứ tự, chương, mục cùng nhiều dấu phân cách; cải thiện hiệu năng với tài liệu lớn.
- **Cải thiện vẽ lại và nhập tài liệu Word:** quản lý định danh, siêu dữ liệu, bố cục đoạn và đích chỉnh sửa khi nhấp đúp ổn định hơn.
- **Tài liệu trợ giúp tích hợp:** hướng dẫn trình soạn thảo, phím tắt, `align`, `cases`, ma trận, OCR, ký hiệu tùy chỉnh, Word và PowerPoint.
- **Chuyển đổi định dạng trong PowerPoint trên Windows:** chuyển đổi giữa công thức dạng ảnh, OLE và OMML.

Xem đầy đủ tại [ghi chú phát hành VisualTeX 1.2.5](https://github.com/Finix15/visualtex/releases/tag/v1.2.5).

## Giao diện thực tế

Các hình dưới đây được chụp trực tiếp từ VisualTeX, không phải bản thiết kế hoặc mô phỏng.

### Không gian làm việc và bảng điều khiển

<p align="center">
  <img src="docs/images/visualtex-readme-1.webp" width="45%" alt="Không gian làm việc VisualTeX với giao diện xanh trắng" />
  <img src="docs/images/visualtex-readme-2.webp" width="45%" alt="Thiết lập bố cục và tính năng VisualTeX" />
</p>
<p align="center"><sub>Bố cục cổ điển xanh trắng · Tùy chọn bố cục, nhập liệu và xuất tệp</sub></p>

### Các chủ đề màu sắc

<p align="center">
  <img src="docs/images/visualtex-readme-3.webp" width="30%" alt="Chủ đề màu be ấm" />
  <img src="docs/images/visualtex-readme-4.webp" width="30%" alt="Chủ đề tím đậm" />
  <img src="docs/images/visualtex-readme-5.webp" width="30%" alt="Chủ đề xanh lục đậm" />
</p>
<p align="center"><sub>Màu be ấm · Tím đậm · Xanh lục đậm</sub></p>

## Tính năng chính

### Soạn thảo công thức trực quan

- Nhập phân số, căn thức, tích phân, tổng, giới hạn, chỉ số, chữ cái Hy Lạp, tập hợp và quan hệ theo cấu trúc.
- Hỗ trợ nhiều dòng công thức, ma trận từ 1×1 đến 10×10, dấu phân cách và chèn cấu trúc dựa trên vùng chọn.
- Có chế độ bàn phím số với kích thước cửa sổ được ghi nhớ độc lập với chế độ thường.
- Tự chọn điểm căn chỉnh bằng `&` trong `align` / `aligned`; chỉnh sửa `cases` nhiều dòng ổn định hơn.
- Cung cấp bố cục tiêu chuẩn và cổ điển; khu vực công cụ và mã LaTeX có thể thu gọn.
- Cho phép chuyển đổi văn bản thường: `alpha` thành chữ cái Hy Lạp, `>=` thành dấu lớn hơn hoặc bằng; `pp`, `ss`, `mm`, `dd`, `eq` thành cộng, trừ, nhân, chia và bằng.
- Điều khiển độc lập việc tự thoát khỏi chỉ số, dấu trọng âm, lệnh phông chữ và từng bảng gợi ý lệnh.
- Hỗ trợ `mathbb`, `mathbf`, `mathcal`; nút chữ đậm dùng chuẩn `\mathbf{...}` và có thể bật/tắt kiểu chữ.
- Tự chuẩn hóa vi phân `d`, hằng số `e`, đơn vị ảo `i/j` và các toán tử thông dụng thành kiểu chữ đứng.
- Hoàn tác và làm lại ở cấp tài liệu, khôi phục nội dung, dòng hiện hành, con trỏ và vùng chọn.
- Gán phím tắt cho công cụ công thức, công cụ thường dùng và ô tùy chỉnh, có quản lý xung đột.
- Tích hợp trình thiết kế ký hiệu để vẽ, kết hợp, cắt và đăng ký ký tự toán học riêng.
- Hỗ trợ thu phóng, nhiều chủ đề sáng/tối, tùy chỉnh màu giao diện, lịch sử cục bộ và tài liệu JSON.

### Mã nguồn LaTeX

- Đồng bộ hai chiều giữa trình soạn thảo trực quan và vùng mã nguồn CodeMirror.
- Hỗ trợ LaTeX thuần, `$...$`, `\(...\)`, `\[...\]`, `equation`, `align`, `gather`, `multline`, `split` và các định dạng liên quan.
- Tự xử lý dấu căn chỉnh cấp cao nhất nhưng vẫn bảo toàn cấu trúc bên trong ma trận.
- Sao chép công thức mà không cần cài TeX Live.

### Xuất và cập nhật tài liệu

- Xuất tài liệu công thức thành Markdown, SVG hoặc PNG và ghi nhớ thư mục đã chọn.
- Markdown dùng khối công thức phổ biến; SVG và PNG phù hợp cho trang web, ghi chú và bài trình chiếu.
- Tự động kiểm tra cập nhật khi khởi động, hiển thị ghi chú phát hành và cho phép kiểm tra thủ công.

### Vẽ lại LaTeX trong Word

- Quét `$...$`, `\(...\)`, `$$...$$` và `\[...\]` trong vùng chọn hoặc toàn bộ tài liệu.
- Chuyển đổi hàng loạt thành OMML nguyên bản hoặc VisualTeX OLE có thể chỉnh sửa bằng cách nhấp đúp.
- Kế thừa cỡ chữ theo nội dung xung quanh và tái sử dụng đoạn gốc để không sinh dòng trống.
- Dùng bộ chuyển đổi khởi động sẵn, bộ nhớ đệm và thay thế theo thứ tự ngược để tăng độ an toàn, hiệu năng.

### Nhận dạng công thức từ ảnh

- Chọn, kéo thả hoặc dán ảnh rồi chèn LaTeX vào vị trí con trỏ đã lưu.
- Hỗ trợ PaddleOCR PP-FormulaNet plus-S, plus-M và plus-L.
- Xử lý ảnh nền tối hoặc trong suốt, hiển thị tiến độ và cho phép hủy.
- Ảnh được xử lý ngay trên thiết bị, không tải lên dịch vụ bên thứ ba.
- OCR nhanh hỗ trợ ảnh chụp màn hình; OCR nền ghi LaTeX thẳng vào bảng nhớ tạm.

## Phiên bản macOS

Ứng dụng macOS nằm tại [`apps/macos`](apps/macos) và tích hợp Office nguyên bản, hoàn toàn ngoại tuyến:

![Tiện ích Word nguyên bản của VisualTeX trên macOS](docs/images/visualtex-macos-word-ribbon.png)

- Word dùng mẫu DOTM toàn cục; PowerPoint dùng tiện ích PPAM tại đường dẫn cố định.
- VBA, AppleScriptTask, Office Group Container và cửa sổ Tauri cục bộ phối hợp thành quy trình phiên ngoại tuyến.
- Word hỗ trợ công thức ảnh, OMML cùng dòng/riêng dòng, đánh số, tham chiếu chéo, nhập tài liệu và vẽ lại LaTeX.
- Cài đặt, cập nhật, sửa chữa và phát hiện phiên bản DOTM/PPAM; có thể thay thế trực tiếp bản cũ.
- PowerPoint hỗ trợ tạo, thay thế, xóa, đổi cỡ và nhấp đúp để chỉnh sửa công thức VisualTeX.
- Tiện ích không phụ thuộc Office.js, XML Manifest, chứng chỉ tin cậy của hệ thống hoặc mạng ngoài.

## Phiên bản Windows

Ứng dụng Windows nằm tại [`apps/windows`](apps/windows), sử dụng VSTO và đối tượng COM/OLE thực:

![Tiện ích Word nguyên bản của VisualTeX trên Windows](docs/images/visualtex-windows-word-ribbon.png)

- Word và PowerPoint sử dụng Ribbon VSTO cùng sự kiện Office nguyên bản.
- Chế độ chuyên nghiệp chèn đối tượng OLE `VisualTeX.Formula.1` thực với siêu dữ liệu và bản xem trước EMF/PNG.
- Word hỗ trợ OLE/OMML cùng dòng hoặc riêng dòng, chuyển đổi, đánh số, tham chiếu, nhập và vẽ lại LaTeX hàng loạt.
- PowerPoint hỗ trợ tạo, chỉnh sửa, xóa, xuất ảnh và chuyển đổi giữa công thức dạng ảnh, OLE, OMML.
- Nhấp đúp theo cơ chế Office nguyên bản để mở lại trình soạn thảo VisualTeX.
- Chế độ ảnh phục vụ tài liệu đa nền tảng và di chuyển công thức cũ.

## Cấu trúc kho mã nguồn

```text
visualtex/
├── apps/
│   ├── macos/       # Ứng dụng macOS độc lập, Tauri, DOTM/PPAM, OCR và kiểm thử
│   └── windows/     # Ứng dụng Windows độc lập, Tauri, VSTO/OLE, OCR và kiểm thử
├── docs/            # Kiến trúc và ảnh chụp giao diện thực tế
├── tools/           # Công cụ kiểm tra cấu trúc cấp cao nhất
├── .github/         # Quy trình CI riêng cho từng nền tảng
├── package.json     # Chỉ điều phối hai dự án con
└── README.md
```

Hai ứng dụng không dùng chung `src`, `src-tauri`, thành phần trình soạn thảo, cách triển khai phiên Office, `package-lock.json` hoặc `Cargo.lock`.

## Phát triển cục bộ

```bash
# Cài dependency cho cả hai nền tảng
npm run bootstrap

# Build riêng từng frontend
npm run build:macos
npm run build:windows

# Kiểm tra cấu trúc kho mã nguồn và build cả hai frontend
npm run check
```

Đóng gói ứng dụng nguyên bản và kiểm thử Office trong thư mục tương ứng:

```bash
cd apps/macos
npm run tauri:build

cd ../windows
npm run tauri:build
```

Xem [tài liệu kiến trúc](docs/ARCHITECTURE.md) để biết thiết kế chi tiết và [hướng dẫn build, đóng gói](BUILD_PACKAGING_VI.md) để tạo bộ cài từ mã nguồn.

## Giấy phép

VisualTeX được phát hành theo [Giấy phép MIT](LICENSE).
