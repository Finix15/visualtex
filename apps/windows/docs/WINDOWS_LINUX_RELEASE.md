# Phát hành VisualTeX 1.2.6 cho Windows

VisualTeX hiện có pipeline Windows tại `.github/workflows/windows.yml`. Workflow này chạy kiểm tra cấu trúc repository, frontend, metadata Office và kiến trúc Office native. Bộ cài release đầy đủ cần được dựng trên máy Windows đã có toolchain native như mô tả bên dưới.

## Yêu cầu máy build

- Node.js và `npm` theo phiên bản dự án hỗ trợ;
- Rust stable với target `x86_64-pc-windows-msvc`;
- .NET SDK 8 (dự án dùng `global.json` để tránh chọn nhầm SDK mới hơn);
- Visual Studio Build Tools 2022/2026 với MSBuild, VC++ x86/x64 và ATL;
- Windows 11 SDK;
- .NET Framework 4.8/4.8.1 SDK và targeting pack (bao gồm `AxImp.exe`).

## Dựng và kiểm tra

Từ `apps/windows`:

```powershell
npm ci
npm run test:windows-office-architecture
npm run tauri:build
```

`npm run tauri:build` chuẩn bị Python OCR 3.12.10 riêng và wheelhouse offline, VSTO Runtime, Office native x86/x64, frontend và Rust release; sau đó đóng gói NSIS, xác minh artifact và cài thử vào thư mục acceptance sạch.

Artifact chính:

- `src-tauri/target/release/bundle/nsis/VisualTeX_1.2.6_x64-setup.exe`;
- `src-tauri/resources/windows-office/VisualTeX-WindowsOffice-VSTO-x64.msi`;
- `src-tauri/resources/windows-office/VisualTeX-WindowsOffice-VSTO-x86.msi`.

Các MSI Office có `ProductVersion` nội bộ `1.0.41.0`; phiên bản ứng dụng/NSIS là `1.2.6`.

## OCR

Bộ cài có tùy chọn cài runtime OCR offline gồm Python riêng, PaddlePaddle, PaddleOCR và dependency wheel đã khóa hash. Các model S/M/L không nhúng trong installer và được tải riêng khi người dùng cần, giúp tránh làm bộ cài lớn hơn nữa.

## Trước khi phát hành công khai

- Ký Authenticode cho EXE/MSI bằng chứng thư phát hành và timestamp server tin cậy;
- chạy lại gate xác minh sau khi ký;
- kiểm tra trên máy Windows sạch với Office x86 và Office x64 thực tế (Word và PowerPoint), bao gồm cài mới, nâng cấp, gỡ cài đặt và luồng công thức;
- kiểm tra SmartScreen/Defender và quyền UAC trên máy không có toolchain phát triển.

Không phát hành artifact unsigned như một bản production chính thức. Repository hiện chưa chứa workflow tự động tạo GitHub Release hoặc cấu hình chứng thư ký; hai phần này cần secret/hạ tầng phát hành của chủ dự án.
