# Hướng dẫn build và đóng gói VisualTeX

Tài liệu này hướng dẫn tạo bộ cài VisualTeX từ mã nguồn. Hai ứng dụng macOS và Windows độc lập, vì vậy phải đóng gói trên đúng hệ điều hành đích:

- macOS: build trên máy Mac;
- Windows: build trên máy Windows;
- không dùng lệnh build ở thư mục gốc để tạo bộ cài native cho cả hai hệ điều hành.

## Phân biệt các loại build

| Mục đích | Lệnh | Kết quả |
| --- | --- | --- |
| Build frontend macOS | `npm run build:macos` tại thư mục gốc | `apps/macos/dist/` |
| Build frontend Windows | `npm run build:windows` tại thư mục gốc | `apps/windows/dist/` |
| Đóng gói macOS | `npm run tauri:build` trong `apps/macos` | `.app` và `.dmg` |
| Đóng gói Windows | `npm run tauri:build` trong `apps/windows` | `.exe`, cùng các MSI Office |

Thư mục `dist/` chỉ chứa frontend đã biên dịch, không phải bộ cài dành cho người dùng.

## Đóng gói trên macOS

### 1. Yêu cầu môi trường

- macOS 11 trở lên;
- máy Apple Silicon (`arm64`) cho pipeline OCR hiện tại;
- Xcode Command Line Tools;
- Node.js 22 và npm;
- Rust stable;
- Python 3, gọi được bằng lệnh `python3`;
- kết nối Internet ở lần build đầu để tải dependency và runtime OCR.

Cài Xcode Command Line Tools:

```bash
xcode-select --install
```

Cài Rust nếu máy chưa có:

```bash
curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh
source "$HOME/.cargo/env"
```

Kiểm tra công cụ:

```bash
node --version
npm --version
rustc --version
cargo --version
python3 --version
```

### 2. Cài dependency

Từ thư mục repository:

```bash
cd apps/macos
npm ci
```

Nên dùng `npm ci` thay cho `npm install` để cài đúng phiên bản trong `package-lock.json`.

### 3. Kiểm tra frontend

```bash
npm run build:desktop
```

Frontend được tạo tại:

```text
apps/macos/dist/
```

### 4. Tạo ứng dụng và DMG

```bash
npm run tauri:build
```

Lệnh này tự động:

1. kiểm tra tài nguyên Office `VisualTeX.dotm` và `VisualTeX.ppam`;
2. build frontend;
3. chuẩn bị runtime OCR offline cho macOS ARM64;
4. build Rust/Tauri ở chế độ release;
5. đóng gói ứng dụng và DMG.

Quá trình chuẩn bị OCR có thể khá lâu ở lần đầu và cần tải nhiều dependency.

### 5. Vị trí file sau build

Ứng dụng macOS:

```text
apps/macos/src-tauri/target/release/bundle/macos/VisualTeX.app
```

Bộ cài DMG:

```text
apps/macos/src-tauri/target/release/bundle/dmg/VisualTeX_<version>_aarch64.dmg
```

Ví dụ với phiên bản 1.2.6:

```text
apps/macos/src-tauri/target/release/bundle/dmg/VisualTeX_1.2.6_aarch64.dmg
```

Tìm artifact nhanh:

```bash
find apps/macos/src-tauri/target/release/bundle \
  \( -name "*.dmg" -o -name "*.app" \)
```

Kiểm tra DMG:

```bash
npm run verify:mac-dmg
```

### 6. Ký ứng dụng macOS

Cấu hình hiện tại dùng chữ ký ad-hoc (`signingIdentity: "-"`). File tạo ra phù hợp để kiểm thử nội bộ nhưng chưa phải artifact production đã ký Developer ID và notarize.

Trước khi phát hành công khai cần:

- dùng chứng thư Apple Developer ID Application;
- ký toàn bộ app bundle;
- notarize với Apple;
- staple kết quả notarization vào `.app` hoặc `.dmg`;
- kiểm tra Gatekeeper trên một máy Mac sạch.

## Đóng gói trên Windows

### 1. Yêu cầu môi trường

- Windows 10 hoặc Windows 11 x64;
- Node.js 22 và npm;
- Rust stable với target MSVC x64;
- .NET SDK 8;
- Visual Studio 2022 Build Tools;
- Windows 11 SDK;
- .NET Framework 4.8 hoặc 4.8.1 SDK và Targeting Pack.

Trong Visual Studio Installer, cài các thành phần sau:

- MSBuild;
- Desktop development with C++;
- MSVC x86/x64 build tools;
- C++ ATL cho x86 và x64;
- Windows 11 SDK;
- .NET Framework 4.8/4.8.1 SDK;
- .NET Framework 4.8/4.8.1 Targeting Pack.

Cài và cấu hình Rust:

```powershell
winget install Rustlang.Rustup
rustup default stable-msvc
rustup target add x86_64-pc-windows-msvc
```

Kiểm tra công cụ:

```powershell
node --version
npm --version
rustc --version
cargo --version
dotnet --version
```

`dotnet --version` phải là dòng 8.x. Dự án có `apps/windows/global.json` để chọn SDK phù hợp.

### 2. Cài dependency

Mở PowerShell hoặc Developer PowerShell:

```powershell
cd E:\Code\visualtex\apps\windows
npm ci
```

Nếu repository nằm ở vị trí khác, thay đường dẫn trên bằng đường dẫn thực tế.

### 3. Kiểm tra trước khi đóng gói

```powershell
npm run build:desktop
npm run test:windows-office-architecture
```

Frontend được tạo tại:

```text
apps/windows/dist/
```

### 4. Tạo bộ cài Windows

```powershell
npm run tauri:build
```

Pipeline này tự động:

1. dọn các release output cũ có thể làm Tauri dùng nhầm asset;
2. chuẩn bị Python OCR 3.12.10 embedded và wheelhouse offline;
3. tải và xác minh Microsoft VSTO Runtime;
4. build Office UI;
5. build VSTO, COM và OLE cho Office x86/x64;
6. build frontend và ứng dụng Rust/Tauri release;
7. tạo NSIS installer;
8. xác minh artifact;
9. cài thử vào thư mục acceptance sạch và chạy smoke test.

Lần build đầu cần Internet để tải dependency, Python OCR và VSTO Runtime. Bước smoke test có thể cần quyền phù hợp để chạy bộ cài.

### 5. Vị trí file sau build

Bộ cài chính:

```text
apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_<version>_x64-setup.exe
```

Ví dụ với phiên bản 1.2.6:

```text
E:\Code\visualtex\apps\windows\src-tauri\target\release\bundle\nsis\VisualTeX_1.2.6_x64-setup.exe
```

File chương trình chưa đóng gói:

```text
apps/windows/src-tauri/target/release/visualtex.exe
```

Các gói Office được tạo tại:

```text
apps/windows/src-tauri/resources/windows-office/VisualTeX-WindowsOffice-VSTO-x64.msi
apps/windows/src-tauri/resources/windows-office/VisualTeX-WindowsOffice-VSTO-x86.msi
```

Tìm artifact nhanh:

```powershell
Get-ChildItem .\src-tauri\target\release -Recurse -Include *.exe,*.msi
Get-ChildItem .\src-tauri\resources\windows-office -Include *.msi
```

### 6. Bỏ qua bước cài thử khi chẩn đoán

Chỉ dùng khi cần chẩn đoán lỗi đóng gói:

```powershell
$env:VISUALTEX_SKIP_INSTALL_SMOKE = "1"
npm run tauri:build
Remove-Item Env:VISUALTEX_SKIP_INSTALL_SMOKE
```

Artifact phát hành chính thức nên được build lại mà không bỏ qua smoke test.

### 7. Ký bộ cài Windows

Repository chưa chứa chứng thư và hạ tầng ký release. Trước khi phát hành công khai cần:

- ký Authenticode cho EXE và MSI;
- sử dụng timestamp server tin cậy;
- chạy lại bước xác minh sau khi ký;
- thử cài mới, nâng cấp và gỡ cài đặt trên máy sạch;
- thử với Office x86 và Office x64;
- kiểm tra SmartScreen, Defender và UAC.

Không nên phát hành artifact unsigned như bản production chính thức.

## Các lệnh tiện ích ở thư mục gốc

Cài dependency riêng cho từng nền tảng:

```bash
npm run bootstrap:macos
npm run bootstrap:windows
```

Build frontend riêng:

```bash
npm run build:macos
npm run build:windows
```

Kiểm tra cấu trúc repository:

```bash
npm run test:repository
```

Lệnh `npm run build` ở thư mục gốc chỉ build frontend của cả hai thư mục. Lệnh này không thay thế `npm run tauri:build` và không tạo bộ cài native hoàn chỉnh.

## Xử lý lỗi thường gặp

### `npm ci` báo lockfile không khớp

Đảm bảo đang chạy trong đúng thư mục `apps/macos` hoặc `apps/windows`. Mỗi ứng dụng có `package-lock.json` độc lập.

### Không tìm thấy `cargo`, `rustc` hoặc `tauri`

Khởi động lại terminal sau khi cài Rust và kiểm tra thư mục Cargo đã có trong `PATH`. Tauri CLI của dự án được cài qua `npm ci`, không cần cài global.

### Windows không tìm thấy MSBuild hoặc ATL

Mở Visual Studio Installer, chọn Modify cho Build Tools và bổ sung MSBuild, MSVC x86/x64, ATL x86/x64, Windows SDK cùng .NET Framework SDK/Targeting Pack.

### Windows báo sai phiên bản .NET SDK

Chạy lệnh trong `apps/windows` để `global.json` có hiệu lực:

```powershell
cd E:\Code\visualtex\apps\windows
dotnet --version
```

Nếu vẫn không phải 8.x, cài .NET SDK 8 và kiểm tra lại biến `PATH`.

### Build Windows lỗi khi tải OCR hoặc VSTO

Kiểm tra kết nối Internet, proxy và quyền truy cập các nguồn chính thức. Script xác minh hash và chữ ký nên file tải thiếu hoặc bị proxy thay đổi sẽ bị từ chối.

### macOS lỗi ở bước chuẩn bị OCR

Pipeline OCR hiện tạo bundle `macos-arm64`. Hãy kiểm tra máy là Apple Silicon, có `python3`, còn đủ dung lượng đĩa và có thể truy cập nguồn dependency.

### Không thấy file bộ cài

Kiểm tra lệnh vừa chạy:

- `npm run build:desktop`: chỉ tạo `dist/`;
- `npm run tauri:build`: mới tạo `.dmg` hoặc NSIS `.exe`.

Nếu `tauri:build` kết thúc với lỗi thì artifact cuối có thể chưa được tạo hoặc đã bị bước xác minh từ chối.

## Checklist trước khi bàn giao

- [ ] `npm ci` hoàn tất không lỗi;
- [ ] frontend build thành công;
- [ ] các test kiến trúc nền tảng hoàn tất;
- [ ] `npm run tauri:build` kết thúc với exit code 0;
- [ ] artifact tồn tại đúng thư mục;
- [ ] cài và mở được ứng dụng trên máy sạch;
- [ ] kiểm tra Word và PowerPoint nếu phát hành Office integration;
- [ ] kiểm tra OCR offline;
- [ ] ký và xác minh chữ ký production;
- [ ] ghi SHA-256 của artifact phát hành.

Tạo SHA-256 trên macOS:

```bash
shasum -a 256 apps/macos/src-tauri/target/release/bundle/dmg/*.dmg
```

Tạo SHA-256 trên Windows:

```powershell
Get-FileHash .\src-tauri\target\release\bundle\nsis\*.exe -Algorithm SHA256
Get-FileHash .\src-tauri\resources\windows-office\*.msi -Algorithm SHA256
```

