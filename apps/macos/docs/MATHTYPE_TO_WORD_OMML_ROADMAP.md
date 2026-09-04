# Định hướng chuyển MathType sang Word Equation trên macOS

Trạng thái: tài liệu bàn giao triển khai sau bản Windows `VisualTeX 1.2.7` / Office Integration `1.0.46.0`.

## 1. Mục tiêu

Đưa hai lệnh sau lên Ribbon VisualTeX của Word for Mac:

- **Chuyển MathType đã chọn**.
- **Chuyển toàn bộ MathType**.

Kết quả phải là Word Equation native (OMML), sửa được bằng Word và VisualTeX, chạy hoàn toàn offline, không yêu cầu cài MathType và không dùng OCR làm fallback.

Ưu tiên tuyệt đối là đúng ngữ nghĩa và giữ nguyên tài liệu. Công thức chưa được xác minh phải còn nguyên OLE, có mã lý do rõ ràng và không được thay bằng kết quả phỏng đoán.

## 2. Phạm vi phiên bản đầu

Hỗ trợ:

- Word trên macOS, Apple Silicon trước.
- DOCX và DOCM đã được lưu trên ổ đĩa cục bộ.
- Main document story.
- Embedded inline OLE có `ProgID=Equation.DSMT4` và MTEF v5.
- Công thức inline và công thức đứng riêng một đoạn.

Chưa hỗ trợ trong phiên bản đầu:

- PowerPoint.
- Floating OLE, linked OLE, header/footer, footnote/endnote và text box.
- Equation Editor 3.0/MTEF v3 nếu chưa có fixture riêng vượt qua acceptance.
- OCR hoặc Word Equation Editor converter làm fallback.
- Intel Mac hoặc Universal Binary trước khi bản arm64 đạt nghiệm thu.

## 3. Nguồn tham chiếu bắt buộc từ Windows

Không port VSTO, COM Interop hoặc C# orchestration sang macOS. Chỉ tái sử dụng giao thức, engine toán học, quy tắc an toàn và bộ golden.

Các nguồn cần đọc trước khi triển khai:

- `apps/windows/src-windows/VisualTeX.MathTypeSidecar/`
- `apps/windows/src-windows/VisualTeX.MathTypeConversion/`
- `apps/windows/src-windows/VisualTeX.WordVsto/WordFormulaService.cs`
- `apps/windows/src-windows/VisualTeX.WindowsOffice.VstoShared/WordMathTypeLayoutValidation.cs`
- `apps/windows/src-windows/VisualTeX.WindowsOffice.Tests/MathTypeConversionTests.cs`
- `apps/windows/src-windows/VisualTeX.WindowsOffice.Tests/WordMathTypeLayoutValidationTests.cs`

Engine được khóa như bản Windows:

- Upstream: `https://github.com/a917470154/mathtypejx`
- Commit: `7d90e7274c85cf56ac28d4d15e593044693d7e70`
- Phiên bản upstream: `0.1.0`
- Bản vá VisualTeX: validator nhận đúng delimiter mặc định và delimiter có `m:begChr`/`m:endChr`.
- Giao thức sidecar: version `1`.
- Giới hạn payload: 16 MiB cho mỗi công thức.
- Tối đa 4 worker, không worker nào được gọi Word/VBA.

Không lấy file trong `apps/windows/artifacts/`, virtual environment, EXE/MSI đã build hoặc tài liệu người dùng làm mã nguồn macOS.

## 4. Kiến trúc macOS được chọn

Tích hợp phải dựa trên kiến trúc Office macOS hiện có:

```text
Ribbon Word (VisualTeX.dotm / VBA)
  -> AppleScriptTask (chỉ truyền UUID)
  -> Office Group Container / OfficeSessions/<session-id>
  -> VisualTeX Tauri app
  -> MathType sidecar arm64 (không Office automation)
  -> kết quả đã xác minh
  -> VBA materialize OMML vào Word
```

Các ràng buộc trong `MACOS_OFFLINE_OFFICE_ARCHITECTURE.md` tiếp tục áp dụng:

- VBA không gửi shell command, đường dẫn tùy ý hoặc nội dung công thức qua `AppleScriptTask`.
- Mọi đường dẫn được tạo từ UUID hợp lệ trong Office Group Container.
- App đọc/ghi manifest bằng thao tác atomic.
- Chỉ VBA chạy trên luồng Word mới được sửa tài liệu.
- Sidecar không được điều khiển Word, AppleScript UI, chuột hoặc bàn phím.

### 4.1. Trách nhiệm của VBA

VBA thực hiện:

- Kiểm tra tài liệu là DOCX/DOCM, có đường dẫn và `Saved=True`.
- Xác định lệnh vùng chọn hay toàn tài liệu.
- Tạo request với document identity, scope và UUID; không nhúng OLE vào request.
- Gọi app qua `AppleScriptTask` theo cơ chế hiện có.
- Sau khi app chuẩn bị xong, xác minh document identity và trạng thái đã lưu.
- Tạo bookmark động quanh các OLE mục tiêu, thay từ cuối lên đầu và dùng một Undo Record.
- Ép OMath về inline/display theo phân loại nguồn.
- Kiểm tra prefix/suffix, paragraph, tab/manual break và metadata sau mỗi lần thay.
- Xóa toàn bộ bookmark tạm trong `finally`.

### 4.2. Trách nhiệm của app Tauri

App thực hiện:

- Đọc DOCX/DOCM đã lưu một lần.
- Tạo snapshot gồm hash tài liệu, thứ tự OLE, relationship, part, ProgID, SHA-256 và bố cục nguồn.
- Tạo backup byte-for-byte cạnh file nguồn trước khi hiện xác nhận chuyển toàn bộ.
- Đối chiếu SHA-256 nguồn và backup, kiểm tra ZIP/OOXML và số OLE dự kiến.
- Gửi batch OLE cho sidecar, kiểm tra response version, operation ID, formula ID và fingerprint.
- Xác minh MathML/OMML, phân loại risk và chuẩn bị payload dành cho Word.
- Báo thống kê, cảnh báo và thời gian từng giai đoạn mà không ghi nội dung công thức vào log thường.

### 4.3. Trách nhiệm của sidecar

Sidecar thực hiện thuần dữ liệu:

```text
Compound OLE -> Equation Native -> MTEF -> MathML -> OMML/validation result
```

Mỗi kết quả phải có:

- `formulaId`
- `status`: `convertible`, `unsupported` hoặc `corrupt`
- `risk`: `auto_replace`, `spot_check`, `manual_review` hoặc `blocked`
- `mtefVersion`
- `fingerprint`
- `mathMl`
- `omml`
- `warnings`, `errors`, `reasonCode`

Sidecar phải fail closed khi thiếu file, sai hash, vượt giới hạn, timeout, crash, sai protocol hoặc trả XML không hợp lệ. Không gọi parser C# hay OCR làm fallback.

## 5. Cổng kỹ thuật phải hoàn thành trước khi viết tính năng

### Gate A — MathML sang OMML trên Word for Mac

Bản Windows dùng `MML2OMML.XSL` từ Microsoft Office. Mã macOS hiện đã dùng stylesheet chiều ngược tại:

```text
/Applications/Microsoft Word.app/Contents/Resources/omml2mathml.xsl
```

Trên máy Mac thật phải kiểm tra:

1. Word for Mac có cung cấp `MML2OMML.XSL` tương đương hay không.
2. Kết quả XSLT có chuẩn hóa giống Windows trên bộ golden hay không.
3. Điều khoản phân phối có cho phép nhúng stylesheet vào app hay chỉ được dùng bản nằm trong Office.

Quyết định:

- Nếu stylesheet có sẵn trong Word và kết quả golden khớp: dùng bản của Word theo đường dẫn được dò và xác minh hash/version lúc chạy.
- Nếu không có hoặc không được phép phân phối: dùng bộ sinh OMML hiện có của VisualTeX macOS, nhưng chỉ sau khi test MathML -> OMML riêng đạt cùng semantic signature với golden Windows.
- Nếu cả hai đường không đạt: dừng tại Gate A và báo người dùng; không chuyển qua MathML -> LaTeX -> OMML nếu chưa chứng minh không mất ngữ nghĩa.

### Gate B — Word for Mac có giữ legacy OLE trong live document

Tạo probe chỉ đọc để xác nhận:

- `ActiveDocument.Content.WordOpenXML` còn chứa `o:OLEObject` và relationship tương ứng.
- Word VBA liệt kê được đối tượng inline và giữ đúng thứ tự so với package.
- Có thể tạo bookmark quanh đúng một OLE mà không kích hoạt OLE server.
- Có thể sao chép `FormattedText` sang scratch document và phục hồi đúng đối tượng.

Nếu scratch không bảo toàn đúng OLE part SHA-256 trên nhiều mẫu, không triển khai thay tại chỗ. Phiên bản đầu sẽ chỉ sinh **converted copy** và mở bản đó trong Word, giữ tài liệu nguồn bất biến.

## 6. Quy trình chuyển toàn tài liệu

1. VBA yêu cầu tài liệu đã lưu; nếu chưa lưu thì hướng dẫn `Command+S` và dừng.
2. App đọc package một lượt, ghi SHA-256 và lập snapshot OLE.
3. App ánh xạ snapshot với danh sách live Word theo thứ tự; sai số lượng, ProgID hoặc loại inline thì dừng trước khi sửa.
4. App tạo `<tên>.visualtex-mathtype-backup-YYYYMMDD-HHmmss.<docx|docm>` bằng copy byte trực tiếp.
5. App xác minh hash, OOXML và số OLE của backup.
6. Sidecar giải mã tối đa 4 worker và trả kết quả theo formula ID.
7. App xác minh toàn bộ output, tính các nhóm risk và hiện hộp xác nhận.
8. Trước khi thay, app kiểm tra hash file đã lưu vẫn trùng snapshot.
9. VBA quét lại live objects và tạo bookmark tạm cho từng công thức convertible.
10. VBA thay từ cuối lên đầu; mỗi lần đều lấy range hiện tại từ bookmark.
11. VBA kiểm tra đúng một OMath, đúng inline/display, cùng paragraph, prefix/suffix và không sinh control character ngoài dự kiến.
12. Nếu một công thức lỗi, phục hồi OLE từ scratch và xác minh fingerprint. Nếu mất boundary hoặc rollback không khớp, hoàn tác toàn bộ Undo Record và dừng batch.
13. Xóa bookmark tạm, đóng scratch document và dọn session trong `finally`.

Nếu Gate B buộc dùng converted-copy, các bước 9–12 được thay bằng một bộ sửa package thuần dữ liệu. App tạo file kết quả mới, không ghi đè nguồn, rồi Word mở file kết quả sau khi package validation hoàn tất.

## 7. Quy tắc bố cục

Không dùng `display="block"` từ MathML để quyết định bố cục Word.

- Có chữ, dấu câu, khoảng trắng bố cục, tab hoặc manual break trước/sau OLE trong cùng paragraph: `Inline`.
- Paragraph thực sự chỉ có đúng một OLE và paragraph/cell terminator: `Display`.
- Công thức inline được thay đúng tại một ký tự OLE; không thêm paragraph mark hoặc line break.
- Giữ paragraph alignment, indent, spacing, line spacing và tab stops.
- Cho phép Word tăng nhẹ chiều cao dòng đối với phân số/căn cao; không ép line-height làm cắt công thức.

Các regression Windows `F0548`, `F0568`, `F0630`, `F0645`, `F0659` là test bắt buộc trên macOS vì chúng dùng tab để căn vị trí.

## 8. Backup, rollback và bảo mật

- Không tự động lưu tài liệu đang mở.
- Không dùng `SaveAs`, UI automation hay quyền Accessibility.
- Backup phải khớp SHA-256 trước khi người dùng được phép xác nhận chuyển toàn bộ.
- Công thức thất bại phải giữ OLE byte-identical.
- Mọi file tạm nằm trong session UUID và được dọn trong `finally`.
- Log thường chỉ chứa formula ID, reason code, fingerprint rút gọn và thời gian; không chứa MathML, OMML, MTEF hay nội dung công thức.
- Manifest và output phải có giới hạn kích thước, field allowlist và ghi atomic.
- App từ chối symlink/path traversal và không nhận đường dẫn tùy ý từ VBA.

## 9. Đóng gói sidecar arm64

Tạo runtime Python riêng cho VisualTeX, không phụ thuộc Python hệ thống:

- CPython 3.12 arm64 tối thiểu.
- `mathtypejx` đã vendor tại commit khóa.
- `lxml==6.1.3` arm64.
- `olefile==0.47`.
- XSLT/font maps, NOTICE và LICENSE cần thiết.

Không đưa OCR runtime vào sidecar MathType. Build phải sinh manifest SHA-256 cho từng file và app phải xác minh trước khi chạy. Sidecar được ký cùng app, nằm trong app bundle và chạy không hiện cửa sổ.

Sau khi arm64 đạt acceptance mới mở thêm x86_64/universal. Hai kiến trúc phải cho kết quả MathML/OMML chuẩn hóa giống nhau trên cùng fixture.

## 10. Giao diện người dùng

Hộp xác nhận chuyển toàn bộ hiển thị:

- Tổng OLE tìm thấy.
- Số convertible đã xác minh.
- `auto_replace`, `spot_check`, `manual_review`.
- `unsupported`, `corrupt`, skipped.
- Thời gian đọc package, ánh xạ Word, sidecar và chuẩn bị OMML.
- Đường dẫn backup.
- Một số reason code đầu tiên kèm formula ID.

Hộp kết quả hiển thị số đã chuyển, giữ nguyên, thất bại, thời gian chèn Word và lỗi theo `formulaId | stage | reasonCode`.

## 11. Kế hoạch triển khai theo pha

### Pha 0 — Probe trên máy Mac thật

- Hoàn thành Gate A và Gate B.
- Lưu báo cáo phiên bản macOS, Word, CPU, đường dẫn stylesheet và hash.
- Chưa sửa Ribbon hoặc tài liệu người dùng.

### Pha 1 — Engine và package scanner

- Tách phần Python dùng chung khỏi thư mục `apps/windows` hoặc tạo package vendored dùng chung có provenance duy nhất.
- Build sidecar arm64 và test protocol.
- Viết scanner DOCX/DOCM một lượt và fixture tests.
- Chạy golden Windows trên macOS, chưa tích hợp Word.

### Pha 2 — Orchestration Word for Mac

- Thêm hai nút Ribbon và callback VBA.
- Bổ sung request/result cho OfficeSessions.
- Thực hiện backup, confirm, dynamic bookmark, Undo và rollback.
- Nếu Gate B không đạt, triển khai converted-copy thay cho in-place.

### Pha 3 — Nghiệm thu thực tế

- Test fixture nhỏ theo từng cấu trúc.
- Test tài liệu 680 công thức.
- Test hủy, crash, timeout, thiếu sidecar, tài liệu đổi sau scan và rollback lỗi.
- Duyệt trực quan toàn bộ nhóm `manual_review`, toàn bộ `spot_check` có cảnh báo và lấy mẫu `auto_replace`.

### Pha 4 — Phát hành

- Cập nhật `VisualTeX.dotm` từ source VBA đã review.
- Package sidecar arm64 và licenses vào app.
- Ký, notarize và đóng gói DMG.
- Chạy full macOS Office acceptance hai lần đóng/mở Word.

## 12. Tiêu chí nghiệm thu

Fixture 680 công thức phải đạt:

- Phát hiện đúng 680 OLE `Equation.DSMT4`.
- 678 công thức đạt validator trở thành OMML; đúng 2 công thức bị chặn vẫn là OLE byte-identical.
- Không có lỗi chèn hàng loạt hoặc sai vị trí.
- Không tăng số paragraph và không sinh manual break mới.
- Không còn bookmark tạm `VTMT_*`.
- Năm công thức căn bằng tab chuyển thành inline, giữ nguyên tab.
- Một lần Undo phục hồi toàn bộ batch khi dùng chế độ in-place.
- Lưu, đóng, mở lại và sửa được bằng Word/VisualTeX.
- Quét và chuẩn bị không quá 60 giây; toàn bộ không quá 3 phút trên máy nghiệm thu.
- Kết quả 1 worker và 4 worker giống nhau sau chuẩn hóa.

Chỉ tuyên bố “đúng 100%” theo nghĩa: 100% công thức **đã được chuyển** không có lỗi ngữ nghĩa phát hiện được. Không cam kết chuyển tất cả công thức; trường hợp không chắc chắn phải giữ nguyên OLE.

## 13. Đồng bộ repository sang macOS

Đồng bộ toàn bộ repository để giữ chung fixture, protocol và lịch sử; không chỉ chép thư mục `apps/macos`.

```bash
git clone https://github.com/Finix15/visualtex.git
cd visualtex
git checkout working
git pull --ff-only origin working
```

Sau khi checkout, xác nhận commit bàn giao:

```bash
git rev-parse HEAD
git status --short
```

Không chạy build trước khi trạng thái Git sạch. Trên Mac cần cài toolchain Node/npm, Rust/Tauri, Xcode Command Line Tools và Python dùng để tạo sidecar; Python runtime phát hành vẫn phải được đóng gói riêng, không dựa vào Python cài trên máy người dùng.

## 14. Quy tắc dừng

Dừng ở trạng thái an toàn và báo người dùng trước khi thay đổi phạm vi nếu gặp một trong các trường hợp:

- Cần thêm dependency hoặc cần phân phối tài nguyên chưa rõ giấy phép.
- Word for Mac không cung cấp primitive cần thiết cho bookmark/rollback.
- Golden macOS khác Windows về ngữ nghĩa.
- Cần bật Accessibility, UI automation hoặc gọi MathType.
- Muốn mở rộng sang MTEF v3, floating OLE, story khác main document hoặc PowerPoint.
- Muốn đổi từ in-place sang converted-copy hoặc ngược lại sau khi Gate B đã kết luận.

Không tự động chuyển sang OCR khi bất kỳ gate nào thất bại.
