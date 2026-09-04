# Kế hoạch triển khai bộ chuyển đổi MathType và Equation Editor 3.0 trên macOS bằng Codex

Trạng thái: kế hoạch triển khai, chưa phải xác nhận tính năng đã hoàn thành.

Tài liệu này là nguồn yêu cầu kỹ thuật để Codex chạy trực tiếp trên máy Mac triển khai bộ chuyển đổi công thức MathType/Equation Editor 3.0 cho VisualTeX. Codex phải làm theo từng phase, chạy test gate của phase hiện tại và không tuyên bố hoàn thành khi chưa vượt qua kiểm thử Microsoft Word thực trên macOS.

## 1. Chỉ thị thực thi cho Codex

Khi nhận nhiệm vụ triển khai tài liệu này, Codex phải:

1. Đọc toàn bộ tài liệu này trước khi sửa code.
2. Đọc các tài liệu liên quan:
   - `docs/ARCHITECTURE.md`;
   - `apps/macos/docs/MACOS_OFFLINE_OFFICE_ARCHITECTURE.md`;
   - `apps/macos/docs/MACOS_OFFLINE_OFFICE_ACCEPTANCE.md`;
   - `apps/macos/docs/OFFICE_PERFORMANCE_BUDGET.md`;
   - `tools/mathtypejx/README.md`.
3. Chạy `git status --short` và giữ nguyên mọi thay đổi không liên quan của người dùng.
4. Không chạy lệnh phá hủy như `git reset --hard`, `git clean -fd`, hoặc xóa recursive các thư mục không được xác thực.
5. Không sửa hành vi Office Windows ngoài phần cần thiết để đưa parser thành core dùng chung.
6. Không phụ thuộc Python hệ thống ở runtime phát hành.
7. Không phụ thuộc runtime OCR tùy chọn để sử dụng bộ chuyển đổi.
8. Không ghi đè tài liệu nguồn. Đầu ra luôn là file mới cho đến khi có yêu cầu sản phẩm khác được duyệt.
9. Không dùng COM, VSTO, registry hoặc Windows OLE activation trên macOS.
10. Không chuyển qua LaTeX làm định dạng trung gian giữa MathML và OMML.
11. Không tải hay đóng gói `MML2OMML.XSL` của Microsoft nếu chưa xác minh quyền phân phối.
12. Mọi lỗi chuyển đổi từng công thức phải fail closed: giữ nguyên OLE gốc trong tài liệu đầu ra và ghi lỗi vào report.
13. Sau mỗi phase, cập nhật implementation log, liệt kê file đã đổi, test đã chạy và vấn đề còn lại.
14. Không commit, push, tạo release hoặc notarize bản phân phối thật nếu người dùng chưa yêu cầu rõ.

Nhánh khuyến nghị nếu người dùng yêu cầu tạo nhánh:

```text
codex/macos-mathtype-equation3-converter
```

## 2. Mục tiêu sản phẩm

Xây dựng một bộ chuyển đổi offline trên macOS có thể:

- mở tài liệu Word Open XML;
- phát hiện OLE công thức có ProgID `Equation.DSMT4`, `Equation.3` và, nếu fixture xác nhận, `Equation.2`;
- trích xuất stream MTEF từ OLE Compound File;
- phân tích MTEF v5 của MathType và MTEF v3 của Equation Editor 3.x;
- chuyển MTEF thành Presentation MathML chuẩn hóa;
- chuyển MathML thành OMML Word native bằng converter thuộc VisualTeX;
- thay từng OLE hợp lệ bằng OMML theo transaction;
- giữ nguyên OLE bị lỗi hoặc không được hỗ trợ;
- tạo file đầu ra mới và báo cáo đầy đủ;
- chạy độc lập từ VisualTeX và có thể khởi động từ Word Ribbon trên macOS.

## 3. Phạm vi phát hành

### 3.1 MVP bắt buộc

- Đầu vào `.docx`.
- Chuyển toàn tài liệu.
- Đầu ra mặc định `<stem>_VisualTeX_OMML.docx`.
- Công thức inline và display.
- Công thức trong document body, headers, footers, footnotes, endnotes và textboxes có relationship hợp lệ.
- Progress, cancel, report và retry.
- Chạy hoàn toàn offline.
- Apple Silicon bắt buộc.
- Intel macOS chỉ bắt buộc nếu bản VisualTeX hiện tại vẫn phát hành cho Intel.

### 3.2 Không thuộc MVP

- Parser trực tiếp định dạng Word binary `.doc`.
- Parser trực tiếp PowerPoint binary `.ppt`.
- Chuyển `.pptx`.
- Nhận dạng công thức chỉ còn WMF/PICT/ảnh và không có MTEF.
- Ghi đè file nguồn.
- Chỉnh trực tiếp package của tài liệu đang được Word mở và giữ file lock.
- Tạo hoặc kích hoạt OLE object trên macOS.

### 3.3 Phase sau MVP

- `.doc` được Word for Mac lưu thành bản sao `.docx` qua `SaveAs2`, sau đó chạy pipeline MVP.
- `.pptx` dùng pipeline riêng: MTEF thành SVG-backed VisualTeX Shape, giữ center, rotation, z-order và metadata.
- Chuyển selection hoặc một công thức đang chọn trong Word.

## 4. Gate 0: xác nhận source đầu vào

Source chính thức được Git quản lý phải nằm tại:

```text
tools/mathtypejx/
```

Source này đã được promote từ artifact đánh giá trên Windows. Thư mục cũ
`apps/windows/artifacts/mathtypejx-evaluation/source/` chỉ là nguồn migration cục
bộ và không phải dependency production. Trước khi triển khai trên Mac, Codex phải
kiểm tra ít nhất các file sau tồn tại dưới `tools/mathtypejx/`:

```text
README.md
NOTICE
LICENSE
pyproject.toml
src/mathtypejx/scanner.py
src/mathtypejx/extractor.py
src/mathtypejx/converter.py
src/mathtypejx/replacer.py
src/mathtypejx/validator.py
src/mathtypejx/mtef/records3.py
src/mathtypejx/mtef/records5.py
tests/fixtures/oleObject1.bin
tests/fixtures/oleObject2.bin
tests/fixtures/oleObject3.bin
```

Nếu source hoặc fixtures không có trên Mac:

- dừng Phase 1;
- báo chính xác các file bị thiếu;
- không tự viết lại parser từ mô tả trong README;
- không dùng package trên Internet thay thế mà chưa được người dùng duyệt.

Gate 0 hoàn thành khi:

- source và third-party notices đã có trong working tree trên Mac;
- có thể chạy unit test hiện hữu;
- đã lưu baseline số test pass/fail;
- đã xác nhận license của từng nhóm XSLT/fontmap được giữ nguyên.

## 5. Kiến trúc mục tiêu

```text
VisualTeX UI / Word Ribbon
          |
          v
Tauri legacy-equation job manager
          |
          v
Signed mathtype worker
  DOCX -> OOXML relationships -> OLE/CFB -> MTEF v3/v5
          |
          v
Canonical Presentation MathML
          |
          v
VisualTeX MathML -> OMML converter
          |
          v
Transactional OOXML writer
          |
          +-- output DOCX
          +-- conversion-report.json
          +-- conversion-report.txt
```

### 5.1 Bố cục source mục tiêu

Promote parser ra khỏi `artifacts`:

```text
tools/mathtypejx/
  pyproject.toml
  README.md
  LICENSE
  NOTICE
  src/mathtypejx/
  tests/
```

Backend macOS:

```text
apps/macos/src-tauri/src/legacy_equations/
  mod.rs
  commands.rs
  jobs.rs
  worker.rs
  package.rs
  report.rs
```

Frontend macOS:

```text
apps/macos/src/office/legacyEquations/
  LegacyEquationConverterApp.tsx
  legacyEquationClient.ts
  legacyEquationTypes.ts
  legacyEquationErrors.ts
```

Build/test:

```text
apps/macos/scripts/build_mathtype_worker.py
apps/macos/scripts/verify_mathtype_worker.mjs
apps/macos/scripts/macos_mathtype_converter_smoke.mjs
apps/macos/scripts/run_macos_mathtype_acceptance.mjs
```

Office integration:

```text
apps/macos/office/macos-offline/word/VTWordAdapter.bas
apps/macos/office/macos-offline/word/VTRibbonCallbacks.bas
apps/macos/office/macos-offline/word/customUI14.xml
```

### 5.2 Ranh giới trách nhiệm

Worker Python chịu trách nhiệm:

- mở và xác thực package;
- duyệt relationships;
- đọc OLE CFB;
- trích xuất MTEF;
- parse MTEF v3/v5;
- tạo MathML chuẩn hóa;
- nhận OMML đã xác thực để finalize package;
- tạo report.

Frontend/TypeScript chịu trách nhiệm:

- MathML thành OMML;
- preview;
- lựa chọn công thức được phép chuyển;
- progress và kết quả giao diện.

Tauri/Rust chịu trách nhiệm:

- đường dẫn và file dialog;
- job lifecycle;
- khởi chạy worker không qua shell;
- giới hạn tài nguyên;
- staging và atomic publish;
- cancellation;
- dọn file theo allowlist;
- cung cấp command cho frontend.

VBA chịu trách nhiệm:

- tạo bản sao tài liệu đang mở;
- tạo UUID session;
- ghi request trong Office Group Container;
- gọi AppleScriptTask với UUID;
- không truyền shell command, MathML, OMML hay đường dẫn tùy ý qua AppleScriptTask.

## 6. Protocol và data model

Mọi job dùng UUID v4 canonical. Job data nằm dưới thư mục do Tauri tạo:

```text
~/Library/Application Support/com.visualtex.studio/legacy-equation-jobs/<job-id>/
```

Chỉ các file sau được phép xuất hiện:

```text
request.json
scan-report.json
formula-batch-<n>.json
omml-batch-<n>.json
conversion-report.json
conversion-report.txt
candidate.docx
worker.stderr.log
```

Không dùng stdout để vận chuyển toàn bộ tài liệu hoặc hàng nghìn công thức. Worker chỉ in một JSON status nhỏ; payload lớn dùng file trong job directory.

### 6.1 LegacyFormula

```ts
interface LegacyFormula {
  formulaId: string;
  partName: string;
  relationshipId: string;
  olePartName: string;
  progId: "Equation.DSMT4" | "Equation.3" | "Equation.2";
  mtefVersion: 3 | 5 | null;
  displayMode: "inline" | "block";
  status:
    | "detected"
    | "extracted"
    | "converted"
    | "replaced"
    | "preserved"
    | "failed";
  riskLevel: "auto-replace" | "spot-check" | "manual-review" | "blocked";
  mathMlPath?: string;
  warnings: string[];
  errors: string[];
}
```

### 6.2 ConversionReport

Report phải chứa:

- protocol version;
- VisualTeX version;
- worker version;
- input/output SHA-256;
- input/output file size;
- start/end/duration;
- counts detected/extracted/replaced/preserved/failed/skipped;
- counts theo ProgID và MTEF version;
- từng formula status, warnings, errors và part location;
- kết quả package validation;
- cờ `countConserved`;
- cờ `sourceUnmodified`.

Không ghi nội dung công thức đầy đủ vào log mặc định. MathML/OMML chỉ nằm trong job folder và bị dọn khi job hết hạn hoặc người dùng xác nhận đóng report.

## 7. Thiết kế an toàn file và package

### 7.1 Input validation

- Chỉ nhận regular file.
- MVP chỉ nhận extension `.docx` không phân biệt hoa/thường.
- Kiểm tra ZIP magic và central directory.
- Reject encrypted ZIP.
- Reject absolute paths, `..`, symlink và path traversal trong ZIP.
- Đặt giới hạn cấu hình cho:
  - kích thước file nén;
  - tổng kích thước giải nén;
  - số ZIP entries;
  - kích thước XML part;
  - kích thước OLE part;
  - số công thức;
  - kích thước mỗi MTEF stream.
- XML parser phải tắt DTD, external entity và network resolution.

Giá trị khởi đầu để test, sau đó điều chỉnh bằng corpus thực:

```text
compressed DOCX max:       500 MB
uncompressed package max:  2 GB
ZIP entries max:           50,000
XML part max:              64 MB
OLE part max:              32 MB
formula count max:         20,000
MTEF stream max:           8 MB
```

### 7.2 Relationship traversal

Không hard-code `header1.xml`, `header2.xml`, `header3.xml`.

Phải:

1. Bắt đầu từ package relationships.
2. Duyệt office document part.
3. Duyệt các relationship nội bộ có thể chứa WordprocessingML.
4. Tìm `o:OLEObject` có ProgID được hỗ trợ.
5. Resolve target theo URI rules của OPC, không dùng `Path.resolve()` phụ thuộc current working directory.
6. Xác thực target vẫn nằm trong package.

### 7.3 Transaction

1. Không sửa input.
2. Tạo staging cùng filesystem với output nếu có thể.
3. Copy package metadata cần bảo toàn.
4. Chỉ thay công thức có OMML pass validator.
5. Chỉ xóa relationship và embedding khi không còn tham chiếu.
6. Repack candidate.
7. Mở lại candidate bằng ZIP/XML validator.
8. Xác nhận count conservation.
9. `fsync` file và thư mục cha nếu platform API hỗ trợ.
10. Atomic rename candidate sang output.
11. Nếu bất kỳ bước nào thất bại, không publish output nửa chừng.

Tài liệu có digital signature phải được phát hiện. MVP phải block hoặc yêu cầu người dùng chấp nhận chữ ký sẽ không còn hợp lệ; không âm thầm sửa tài liệu đã ký.

## 8. Pipeline MathML và OMML

### 8.1 Không dùng LaTeX trung gian

Pipeline bắt buộc:

```text
MTEF -> parsed record tree -> canonical MathML -> OMML
```

Pipeline bị cấm:

```text
MTEF -> MathML -> LaTeX -> OMML
```

Lý do: LaTeX trung gian có thể làm mất font style, spacing, prescript, matrix alignment, embellishment và một số MathType variations.

### 8.2 Refactor converter macOS

Từ `apps/macos/src/office/omml/latexToOmml.ts`, tách API mới:

```ts
export function mathMlToOmmlArtifacts(
  mathMl: string,
  displayMode: OmmlDisplayMode,
  preferences?: OmmlFontPreferences,
): OmmlArtifacts
```

`latexLinesToOmmlArtifacts` tiếp tục dùng API chung sau khi MathJax tạo MathML. Không được tạo hai implementation MathML-to-OMML khác nhau.

Converter phải hỗ trợ và có test cho:

- fraction;
- square root và indexed root;
- subscript, superscript, sub-sup;
- prescript/isotope;
- n-ary và limits;
- matrix, pile, cases và aligned rows;
- delimiter hai phía, một phía và rỗng;
- overbar, underbar, hat, tilde, vector, arc;
- boxed/crossed terms;
- long division nếu MTEF parser xuất được cấu trúc;
- text-mode units;
- normal, italic, bold và color nếu OMML biểu diễn được;
- Unicode Greek, operators và supplementary characters;
- private-use mappings đã có provenance rõ ràng.

### 8.3 Quality gate từng công thức

Không thay OLE nếu có một trong các lỗi:

- MathML hoặc OMML malformed;
- critical empty slot;
- token loss;
- matrix row/cell loss;
- root degree loss;
- delimiter loss;
- accent/bar loss;
- n-ary limit loss;
- prescript loss;
- formula output vượt giới hạn kích thước.

Warning về style hoặc spacing được phép chuyển chỉ khi risk level là `spot-check` hoặc `manual-review`, và UI cho người dùng bỏ chọn.

## 9. Phase triển khai

### Phase 0: baseline và provenance

Công việc:

- hoàn thành Gate 0;
- xác nhận source và notices đã nằm trong `tools/mathtypejx`;
- chạy test hiện hữu trên Mac;
- ghi baseline corpus và performance;
- xác nhận Python version hỗ trợ;
- lập danh sách third-party assets và license.

Test gate:

```bash
cd <repo-root>
python3 -m venv /tmp/visualtex-mathtype-venv
source /tmp/visualtex-mathtype-venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e "tools/mathtypejx[dev]"
python -m pytest -q tools/mathtypejx/tests
```

Không chạy test production từ source artifact cũ trên Windows.

Definition of done:

- tất cả test baseline pass hoặc mọi failure đã được ghi rõ;
- `tools/mathtypejx` không phụ thuộc đường dẫn Windows;
- NOTICE/LICENSE đi cùng source;
- không còn test production đọc từ `apps/windows/artifacts`.

### Phase 1: harden scanner, extractor và package writer

Công việc:

- safe ZIP reader;
- OPC relationship traversal;
- CFB/MTEF bounds checking;
- dynamic Word part discovery;
- package replacement theo transaction;
- bảo toàn unknown parts;
- relationship/reference cleanup đúng điều kiện;
- tách tính năng xóa dòng biên tập tiếng Trung khỏi conversion mặc định.

Test gate:

- ZIP traversal fixture bị reject;
- ZIP bomb metadata bị reject trước giải nén lớn;
- malformed XML không crash;
- broken relationship giữ nguyên OLE;
- một embedding được nhiều object tham chiếu không bị xóa sớm;
- multiple formulas trong cùng run được xử lý đúng;
- header/footer có chỉ số lớn hơn 3 được phát hiện;
- source hash trước và sau giống nhau.

### Phase 2: MathML-to-OMML cross-platform

Công việc:

- export `mathMlToOmmlArtifacts`;
- thêm test corpus MathML lấy từ MTEF fixtures;
- so sánh semantics với pipeline Windows hiện tại;
- nối validator;
- bảo toàn inline/display.

Test gate:

```bash
cd <repo-root>/apps/macos
npm ci
npm run build:desktop
npm run test:word-omml
npm run test:document-import:parser
```

Thêm script riêng:

```bash
npm run test:mathtype-omml
```

Definition of done:

- mọi fixture MTEF có MathML và OMML xác định;
- validator không báo mất token/cấu trúc;
- không tìm hoặc tải `MML2OMML.XSL` ở runtime;
- existing OMML tests không regression.

### Phase 3: Tauri job manager và worker

Công việc:

- thêm module `legacy_equations`;
- command API;
- worker process không qua shell;
- stdout status JSON giới hạn kích thước;
- stderr log giới hạn và redaction;
- cancellation với process ownership;
- TTL job cleanup theo allowlist;
- app restart có thể đọc lại job chưa hết hạn;
- không cho job ID tạo path traversal.

Command API đề xuất:

```text
create_legacy_equation_job
get_legacy_equation_job
read_legacy_equation_batch
submit_legacy_omml_batch
finalize_legacy_equation_job
cancel_legacy_equation_job
delete_legacy_equation_job
```

Test gate:

```bash
cargo test --manifest-path apps/macos/src-tauri/Cargo.toml --lib
```

Phải có test:

- invalid UUID;
- input bằng output;
- output đã tồn tại;
- cancel giữa convert;
- worker crash;
- output write failure;
- job folder chứa unknown file thì không recursive-delete;
- concurrent jobs không dùng nhầm payload;
- symlink swap không thoát khỏi job root.

### Phase 4: UI standalone

UI tối thiểu:

- nút chọn DOCX;
- scan summary;
- số MathType/Equation 3/unknown OLE;
- filter theo risk/status;
- preview công thức nếu render được;
- checkbox chọn/bỏ từng công thức;
- chọn output;
- progress theo `scan`, `extract`, `convert`, `validate`, `write`;
- cancel;
- kết quả và đường dẫn report;
- nút mở output bằng Word/Finder.

Tên hiển thị tiếng Việt:

```text
Chuyển công thức MathType cũ
MathType 4 trở lên
Microsoft Equation Editor 3.0
An toàn để chuyển
Cần kiểm tra
Được giữ nguyên do lỗi
```

Không tự động ghi khi scan xong. Người dùng phải bấm nút xác nhận chuyển đổi.

Test gate:

- TypeScript build;
- browser regression;
- file có 0 formula;
- file có mixed success/failure;
- cancel;
- output collision;
- Vietnamese/English strings không mojibake;
- keyboard navigation và error focus cơ bản.

### Phase 5: Word Ribbon trên macOS

Thêm nút Ribbon:

```text
Chuyển MathType / Equation 3.0
```

Luồng bắt buộc:

1. Xác nhận có ActiveDocument.
2. Nếu tài liệu chưa lưu, yêu cầu người dùng lưu trước.
3. Nếu đang ở `.doc`, chỉ bật flow SaveAs2-to-DOCX khi Phase `.doc` đã được duyệt; MVP báo chưa hỗ trợ.
4. Tạo session UUID.
5. Tạo bản sao DOCX trong Office Group Container.
6. Ghi request atomically.
7. Gọi AppleScriptTask với UUID duy nhất.
8. VisualTeX mở converter UI cho bản sao.
9. Output là file mới; không thay package đang mở.

Không đưa file path trực tiếp vào AppleScriptTask. Giữ contract UUID-only của kiến trúc macOS hiện tại.

Test gate:

```bash
cd <repo-root>/apps/macos
npm run test:macos-offline-office
npm run package:macos-offline-office
```

Real Word gate:

- Ribbon hiện sau 10 lần restart Word;
- scan tài liệu tên Unicode;
- chuyển khi offline;
- đổi selection trong lúc job chạy không làm đổi tài liệu đích;
- source đang mở không thay đổi;
- output mở, save, close, reopen thành công;
- uninstall không làm mất công thức trong tài liệu đã tạo.

### Phase 6: corpus và acceptance

Corpus bắt buộc:

- bộ 344 tài liệu/7.888 công thức hiện có nếu được phép dùng nội bộ;
- ít nhất 500 công thức Equation Editor 3.0 thực;
- inline/display;
- body/header/footer/footnote/endnote/textbox;
- tracked changes;
- comments/bookmarks/content controls;
- Chinese, Vietnamese, Greek và supplementary Unicode;
- malformed/truncated OLE fixtures;
- formula rất lớn và document rất nhiều formula.

Release gate:

- 100% công thức được hạch toán:

```text
detected = replaced + preserved + skipped + failed
```

- 0 công thức biến mất âm thầm;
- 0 DOCX corrupt trong corpus;
- source hash không đổi;
- auto-convert corpus đã biết đạt ít nhất 99,9%;
- mọi formula không đạt validator vẫn còn OLE gốc;
- Word macOS mở/save/reopen được toàn bộ sample gate;
- Word Windows mở được sample cross-platform;
- không có network request;
- không có dependency vào Office-installed XSL.

Performance budget ban đầu trên máy Apple Silicon tham chiếu:

- UI phản hồi bắt đầu job dưới 300 ms;
- scan 100 formula dưới 3 giây;
- sau warm-up, p95 convert mỗi formula dưới 100 ms với formula thông thường;
- 1.000 formula hoàn thành dưới 2 phút;
- peak memory dưới 750 MB;
- cancel phản hồi dưới 2 giây.

Nếu không đạt, Codex phải ghi benchmark, profile bottleneck và xin duyệt thay đổi budget; không âm thầm xóa performance gate.

### Phase 7: đóng gói macOS

Worker phát hành phải là executable tự chứa cho từng architecture. Không gọi `/usr/bin/python3`, Homebrew Python hoặc virtualenv của developer.

Khuyến nghị:

- build worker bằng PyInstaller hoặc công cụ tương đương đã pin version;
- build riêng `arm64` và `x86_64` trên runner/máy phù hợp;
- không dùng cross-compiled Python binary chưa được test;
- tạo manifest SHA-256 cho worker và assets;
- thêm third-party notices vào bundle;
- verify worker trước khi Tauri build;
- ký nested executable trước khi ký `.app`;
- chạy notarization và staple ở release workflow được người dùng cho phép.

Build scripts phải fail closed khi worker hoặc checksum bị thiếu.

Smoke commands dự kiến:

```bash
cd <repo-root>/apps/macos
python3 scripts/build_mathtype_worker.py --target macos-arm64
node scripts/verify_mathtype_worker.mjs
npm run build:desktop
cargo test --manifest-path src-tauri/Cargo.toml --lib
npm run test:mathtype-converter
npm run tauri:build
```

Kiểm tra release artifact:

```bash
codesign --verify --deep --strict --verbose=2 "/path/to/VisualTeX.app"
spctl --assess --type execute --verbose=4 "/path/to/VisualTeX.app"
```

Chỉ chạy `notarytool` khi có authorization rõ và credential release phù hợp.

## 10. Kế hoạch `.doc` sau MVP

Không viết parser Word binary trong scope này.

Flow:

1. Word mở `.doc`.
2. VBA tạo bản sao bằng `Document.SaveAs2` với định dạng DOCX.
3. Không overwrite `.doc`.
4. Pipeline DOCX chạy trên bản sao.
5. Report ghi rõ `sourceFormat=doc` và `intermediateFormat=docx`.

Phải block hoặc cảnh báo rõ:

- password protection;
- IRM/restricted document;
- digital signature;
- conversion của Word báo compatibility loss;
- tài liệu chưa lưu hoặc read-only location.

## 11. Kế hoạch PowerPoint sau MVP

PowerPoint là project con riêng, không gộp vào Word MVP.

Pipeline dự kiến:

```text
PPTX slide relationships
  -> OLE/CFB
  -> MTEF
  -> MathML
  -> VisualTeX render SVG
  -> replace OLE with SVG-backed Shape
  -> attach VisualTeX metadata
```

Phải giữ:

- slide ID/index;
- shape center;
- width/height theo natural aspect ratio;
- rotation;
- z-order;
- group membership nếu hỗ trợ an toàn;
- alt text và metadata liên quan.

Failed formula giữ nguyên OLE. Không gọi đối tượng thay thế là OLE và không đăng ký `VisualTeX.Formula.1` trên macOS.

## 12. Ma trận test tối thiểu

| Nhóm | Trường hợp |
|---|---|
| ProgID | `Equation.DSMT4`, `Equation.3`, `Equation.2`, unknown OLE |
| MTEF | v3, v5, truncated, future record, unknown selector |
| Cấu trúc | fraction, root, scripts, prescripts, matrix, pile, cases, n-ary, limits |
| Embellishment | bar, vector, hat, tilde, arc, underbar, strike |
| Văn bản | units, Unicode, Greek, CJK, Vietnamese |
| Word parts | body, headers, footers, footnotes, endnotes, textbox |
| OOXML | multiple formulas/run, shared embedding, broken rel, unknown parts |
| Tài liệu | tracked changes, bookmarks, comments, content controls |
| Failure | malformed ZIP/XML/CFB/MTEF, write failure, worker crash, cancel |
| Security | traversal, symlink swap, ZIP bomb, XXE, oversized payload |
| Runtime | offline, app restart, concurrent jobs, Unicode path |
| Office | Word Mac open/save/reopen, Word Windows cross-open |

## 13. Definition of done toàn dự án

Dự án chỉ được xem là hoàn thành khi:

- shared parser nằm ngoài `artifacts` và được Git quản lý;
- unit/integration/security tests pass trên Mac;
- frontend và Tauri build pass;
- bộ chuyển đổi standalone hoạt động trên Apple Silicon không cần Python hệ thống;
- source document không bị sửa;
- failed formulas được bảo toàn;
- report conservation pass;
- corpus gate pass;
- real Word Mac acceptance pass;
- sample output mở được trên Word Windows;
- worker và app bundle ký hợp lệ trong release candidate;
- notarized DMG pass Gatekeeper khi release được ủy quyền;
- tài liệu kiến trúc, acceptance, help và third-party notices đã cập nhật.

## 14. Báo cáo Codex phải trả sau mỗi phase

Codex phải báo theo mẫu:

```text
Phase:
Outcome:
Files changed:
Tests passed:
Tests not run and reason:
Corpus result:
Security checks:
Known limitations:
Next phase:
User action required:
```

Nếu bị block, báo blocker cụ thể cùng bằng chứng, không đánh dấu phase complete.

## 15. Lệnh khởi đầu trên Mac

Chạy từ Terminal trên Mac sau khi repository và source parser đã có đầy đủ:

```bash
cd /path/to/visualtex
git status --short

test -f apps/macos/docs/MACOS_MATHTYPE_EQUATION3_CODEX_IMPLEMENTATION_VI.md
test -f tools/mathtypejx/pyproject.toml
test -f tools/mathtypejx/src/mathtypejx/mtef/records3.py
test -f tools/mathtypejx/src/mathtypejx/mtef/records5.py
test -f tools/mathtypejx/tests/fixtures/oleObject1.bin

cd apps/macos
npm ci
npm run build:desktop
cargo test --manifest-path src-tauri/Cargo.toml --lib
npm run test:macos-offline-office
```

Prompt ngắn có thể giao cho Codex trên Mac:

```text
Đọc toàn bộ apps/macos/docs/MACOS_MATHTYPE_EQUATION3_CODEX_IMPLEMENTATION_VI.md và các tài liệu kiến trúc được liên kết. Triển khai Phase 0 trước, kiểm tra source/fixtures và baseline trên macOS. Không chuyển sang Phase 1 nếu Gate 0 chưa đạt. Giữ nguyên thay đổi không liên quan, không sửa tài liệu nguồn, không dùng COM/VSTO, không phụ thuộc Python hệ thống ở runtime và không commit/push nếu tôi chưa yêu cầu. Sau khi hoàn tất Phase 0, báo theo mẫu trong mục 14 cùng toàn bộ test thực tế đã chạy.
```
