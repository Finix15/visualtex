# Nhật ký triển khai MathType và Equation Editor 3.0 trên macOS

## Phase 0 — Baseline và provenance

Ngày kiểm tra: 2026-09-04
Git revision: `69ebfaba611f7291d5d67f3ee57c759979ee2351`  
Máy kiểm tra: Apple Silicon `arm64`, macOS 15.7.7 (24G720)

### Outcome

Gate 0 đạt. Source chính thức, fixtures và third-party notices đều được Git quản lý dưới
`tools/mathtypejx`. Baseline Python chạy thành công trên macOS và không đọc source từ
`apps/windows/artifacts`.

### Source và fixtures

- Đã xác nhận đủ toàn bộ file bắt buộc trong mục Gate 0, gồm parser MTEF v3/v5,
  scanner, extractor, converter, replacer, validator, `README.md`, `NOTICE`, `LICENSE`
  và `pyproject.toml`.
- Ba fixture `oleObject1.bin`, `oleObject2.bin`, `oleObject3.bin` đều là OLE Compound
  File và đều chứa MTEF v5. Cả ba sinh được Presentation MathML trong smoke probe.
- SHA-256 fixture:
  - `oleObject1.bin`: `3cfcc0a6da2c4185bd954caa042dfc61c2aecc65fb143229fe8c3a9c9c68b34d`
  - `oleObject2.bin`: `e44dfe5a274efb72a4695060036ce951986d68ba3679ac0d642dd15c665014e6`
  - `oleObject3.bin`: `687f2f713a0ae988419f86b2fdb622e6d70f152dc86af9646960c6a4df5d7838`

### License và provenance

- Mã package: MIT, được giữ trong `tools/mathtypejx/LICENSE`.
- `NOTICE` ghi nhận các nguồn tham chiếu `jure/mathtype`, `sbulka/mathtype`,
  `jure/mathtype_to_mathml`, `transpect/mathtype-extension` và
  `mathtype_to_mathml_plus`.
- 35 file trong nhóm fontmap giữ license BSD của transpect.io tại
  `src/mathtypejx/mtef/xslt/xsl/fontmaps/LICENSE`.
- Tổng cộng 58 asset dưới cây XSLT/fontmap được Git quản lý. Digest SHA-256 tổng hợp
  tại baseline: `3753c376465aa329a0c33e89fc7f61d5c9ce6024500466cf23258bc6b20ddb69`.

### Môi trường và dependency

- Python hệ thống Apple 3.9.6 không đạt yêu cầu `requires-python >=3.10` và không được
  dùng cho baseline.
- Baseline dùng Python 3.12.14 trong virtualenv tạm
  `/tmp/visualtex-mathtype-venv.3ghFiE`.
- Dependency đã cài: `lxml 6.1.3`, `olefile 0.47`, `python-docx 1.2.0`,
  `pytest 9.1.1`, `pytest-cov 7.1.0` cùng các dependency gián tiếp.

### Tests

Lệnh:

```bash
python -m pytest -q tools/mathtypejx/tests
```

Kết quả: `65 passed, 19 skipped`, không có failure; thời gian pytest 0,92 giây,
wall-clock 1,18 giây ở lần baseline có đo thời gian.

Các skip đã biết:

- 12 lượt do không có fixture DOCX MathType thực/private corpus trong repository.
- 7 lượt do Microsoft `MML2OMML.XSL` không có trên macOS.

CLI `mathtypejx health` xác nhận `python-lxml` và `python-olefile` hoạt động; trạng thái
toàn pipeline kiểu Windows là `NOT READY` chỉ vì thiếu `MML2OMML.XSL`. Phase 2 phải dùng
converter MathML-to-OMML thuộc VisualTeX theo kế hoạch, không tải stylesheet này.

### Giới hạn chuyển sang phase tiếp theo

- Public fixtures hiện chỉ bao phủ MTEF v5; chưa có fixture MTEF v3/Equation Editor 3.0.
- Private corpus 344 tài liệu/7.888 công thức không có trong repository nên chưa chạy.
- Chưa chạy benchmark corpus hoặc Microsoft Word thực; các kiểm tra này thuộc Phase 6.
- Phase 1 chưa được bắt đầu trong lần thực thi này.

## Phase 1 — Harden scanner, extractor và package writer

Ngày kiểm tra: 2026-09-04
Base revision: `971e4f1` (`Document macOS MathType Phase 0 baseline`)

### Outcome

Phase 1 đạt test gate tự động. Scanner và writer hiện fail closed với package không an
toàn, duyệt WordprocessingML parts qua OPC relationships thay vì danh sách filename cố
định, và chỉ publish output sau khi candidate đã được mở lại và xác thực. Tầng parser
MTEF, builder MathML và converter semantic không thay đổi.

### Thay đổi triển khai

- Thêm giới hạn package tập trung cho DOCX nén/giải nén, ZIP entries, XML, OLE, số
  công thức và MTEF stream.
- Thay `ZipFile.extractall()` bằng extraction thủ công: reject traversal, absolute path,
  backslash path, URI lạ, duplicate entry, encrypted entry, ZIP symlink, ZIP bomb theo
  metadata và digital signature.
- XML parser tắt DTD, entity resolution, network và huge-tree; malformed XML dừng an toàn.
- Bắt đầu từ package-level `officeDocument` relationship và duyệt động header, footer,
  footnote, endnote; textbox được phát hiện trong story part chứa nó.
- Resolve relationship bằng POSIX OPC URI, không phụ thuộc current working directory;
  lưu canonical `ole_part_name` cho từng công thức.
- Extractor giới hạn kích thước OLE/MTEF, kiểm tra `cbObject` không vượt stream thực và
  đóng CFB handle trong mọi nhánh.
- Writer làm việc trên staging copy, giữ nguyên source tree, bảo toàn unknown parts,
  xử lý nhiều OLE trong cùng run mà không làm mất text, và chỉ xóa relationship/embedding
  khi không còn tham chiếu.
- Candidate được repack, mở lại bằng safe reader, parse lại `word/document.xml`, `fsync`
  và publish bằng hard-link không ghi đè. Collision hoặc lỗi trước khi publish không để
  output nửa chừng.
- Xóa thông tin biên tập tiếng Trung đã được tách khỏi hành vi mặc định
  (`remove_edit_info=False`). Output mặc định đổi thành `<stem>_VisualTeX_OMML.docx`.

### Tests

Lệnh gate:

```bash
python -m compileall -q tools/mathtypejx/src tools/mathtypejx/tests
python -m pytest -q -rs tools/mathtypejx/tests
```

Kết quả: `78 passed, 19 skipped`, không có failure; pytest 0,26 giây, wall-clock
0,63 giây. Có 13 test Phase 1 mới cho:

- ZIP traversal và ZIP bomb metadata;
- malformed XML và XXE/DTD;
- digital signature;
- relationship target thoát package;
- broken OLE relationship;
- header có chỉ số 27;
- nhiều công thức trong cùng run;
- shared embedding và cleanup embedding không còn tham chiếu;
- bảo toàn unknown part và source SHA-256;
- output collision;
- giới hạn OLE và MTEF header bị truncate.

Ba fixture MTEF v5 tiếp tục sinh MathML. Không có diff dưới
`src/mathtypejx/mtef/` hoặc `converter.py`, nên Phase 1 không thay đổi kết quả semantic
MTEF hiện tại.

### Release blockers và giới hạn

- **MTEF v3/Equation Editor 3.0:** repository chưa có fixture hoặc corpus thực. Đây là
  release blocker riêng; Phase 1 không tạo fixture giả và không xác nhận hỗ trợ Equation
  Editor 3.0.
- 12 test vẫn skip vì thiếu DOCX MathType thực/private corpus.
- 7 test vẫn skip vì không có Microsoft `MML2OMML.XSL`; Phase 2 phải nối MathML vào
  converter OMML của VisualTeX, không tải stylesheet này.
- Chưa chạy corpus 344 tài liệu/7.888 công thức hoặc Microsoft Word thực.
- Phase 2 chưa được bắt đầu.

## Phase 2 — MathML-to-OMML cross-platform

Ngày kiểm tra: 2026-09-04
Base revision: `0847b33` (`Harden MathType DOCX package processing`)

### Outcome

Phase 2 đạt test gate tự động. macOS công khai `mathMlToOmmlArtifacts()` và dùng một
implementation MathML-to-OMML chung cho cả đầu vào MathML từ MTEF lẫn đầu vào LaTeX sau
khi MathJax tạo Presentation MathML. Production path không dùng LaTeX làm trung gian từ
MTEF và không tìm, tải hoặc đóng gói `MML2OMML.XSL`.

### Thay đổi triển khai

- Thêm API `mathMlToOmmlArtifacts(mathMl, displayMode, preferences?)` trả OMML,
  Base64URL và DOCX tối thiểu giống contract hiện hữu.
- `latexLinesToOmmlArtifacts()` tái sử dụng cùng parser, tree converter, font mapping,
  validator và artifact builder; không có converter MathML-to-OMML thứ hai.
- Thêm chuyển đổi `mmultiscripts`/`mprescripts` sang `m:sPre`, giữ postscript và các
  slot presubscript/presuperscript.
- Thêm validator TypeScript fail-closed cho XML/DTD/entity, giới hạn MathML 8 MB,
  critical empty slot, token multiplicity, structure, root degree, matrix row/cell,
  delimiter, accent/bar/group character, n-ary limits và prescript.
- Thêm contract cố định lấy trực tiếp từ `oleObject1.bin`–`oleObject3.bin`, khóa bằng
  SHA-256 và xác nhận cả ba là MTEF v5. Test riêng phủ fraction, square/indexed root,
  scripts, matrix, delimiter một phía/rỗng, accents, n-ary/limits, prescripts,
  text-mode, Unicode Greek/CJK/Vietnamese và supplementary character.
- Browser regression tự chọn Google Chrome hoặc Microsoft Edge có sẵn trên macOS.
- Khôi phục contract parser document-import bị regression từ trước: nhãn theorem family
  tiếng Trung và logic strip LaTeX comments không còn tự bảo vệ comment khỏi chính bước
  loại comment. Không sửa expectation để làm test pass.

### Tests

Gate chạy sau `npm ci`:

```bash
npm run build:desktop
npm run test:word-omml
npm run test:document-import:parser
npm run test:mathtype-omml
python -m pytest -q -rs tools/mathtypejx/tests
```

Kết quả:

- desktop TypeScript/Vite build PASS, 2.314 modules transformed;
- Word structural OMML regression PASS;
- document-import parser PASS: 88 syntax fixtures, 122 line-ending cases,
  144 rendered formulas;
- MathType OMML contract PASS: 12 cases, gồm 3 OLE MTEF v5 thật;
- Python core regression: `78 passed, 19 skipped`, không failure.

19 skip Python được giữ nguyên: 12 lượt thiếu fixture DOCX/private corpus được cấu hình
trong test hiện hữu và 7 lượt thiếu Microsoft `MML2OMML.XSL`. Không test nào bị xóa hoặc
sửa để giảm skip. `npm ci` báo 7 advisory dependency hiện hữu (4 moderate, 3 high); không
chạy `npm audit fix` vì có thể thay dependency ngoài scope Phase 2.

### Corpus bổ sung

File người dùng cung cấp
`BÀI 3 VẬN TỐC - GIA TỐC TRONG DAO ĐỘNG ĐIỀU HOÀ.docx` được scan read-only:

- phát hiện 680 OLE, toàn bộ có ProgID `Equation.DSMT4`;
- extract thành công 680/680, toàn bộ MTEF v5;
- MTEF-to-MathML thành công 680/680;
- cấu trúc quan sát được gồm 474 fraction, 131 square root, 279 subscript,
  340 superscript, 92 sub-sup, 13 matrix, 7 mover và 1 multiscript;
- file nguồn không bị sửa; chưa chạy package replacement hay Word open/save/reopen trên
  file này trong Phase 2.

### Security và giới hạn

- MathML có DTD/entity, XML malformed, quá 8 MB hoặc conversion mất token/cấu trúc bị
  chặn trước khi artifact được trả cho writer.
- Không có chuỗi `MML2OMML.XSL` trong production source/scripts của ứng dụng macOS.
- **MTEF v3/Equation Editor 3.0 vẫn là release blocker:** cả ba fixture repository và
  680 công thức bổ sung đều là MTEF v5. Phase 2 không tạo fixture giả và không xác nhận
  Equation Editor 3.0.
- Chưa chạy private corpus 344 tài liệu/7.888 công thức, Word macOS thực hoặc Word Windows
  cross-open; các gate này thuộc Phase 6.
- Phase 3 chưa được bắt đầu.

## Phase 3 — Tauri job manager và worker protocol

Ngày kiểm tra: 2026-09-04
Base revision: `b9bc6651479bb0a19047d06147d21948576962aa`

### Outcome

Phase 3 đạt test gate tự động. Backend `legacy_equations` quản lý job UUID v4 trong
VisualTeX app data, trao đổi file payload có giới hạn với worker bằng hai operation cố
định `scan`/`finalize`, và chỉ publish candidate sau khi report, SHA-256, package
validation và count conservation đều hợp lệ. Phase 4 chưa được bắt đầu.

### Thay đổi triển khai

- Thêm sáu module `mod.rs`, `commands.rs`, `jobs.rs`, `worker.rs`, `package.rs` và
  `report.rs`; đăng ký đủ bảy Tauri command theo kế hoạch.
- Frontend chỉ truyền job ID canonical UUID v4. Job root luôn được suy ra từ app data;
  mọi lần truy cập đều kiểm tra thư mục thật, không symlink và không thoát root.
- Worker được gọi trực tiếp bằng executable cùng argument array cố định, không qua
  shell. Production thiếu bundled worker sẽ fail closed. Fallback
  `VISUALTEX_MATHTYPE_WORKER_DEV` chỉ được biên dịch trong debug và yêu cầu executable
  tuyệt đối, regular, không symlink; không có contract phụ thuộc Python hệ thống/OCR.
- Giới hạn stdout 64 KiB, stderr 1 MiB, manifest/report 4 MiB và batch 16 MiB. Stderr
  được redact job path trước khi ghi log và mọi JSON protocol đều fail closed khi sai
  version, job ID, status hoặc schema.
- Worker process được sở hữu riêng theo job; cancel chỉ kill đúng child của job đó.
  Job `running`/`finalizing` còn lại sau restart được phục hồi thành `failed` với lý do
  rõ ràng; startup đồng thời chạy TTL cleanup 24 giờ theo file allowlist.
- Source DOCX được mở read-only để hash, từ chối symlink/traversal và được kiểm SHA-256
  trước lẫn sau publish. Output collision bị từ chối. Candidate chỉ publish bằng
  hard-link transaction không overwrite; link/fsync failure không để output tồn tại.
- Formula có validation error không được đánh dấu `replaced`; các nhánh non-replaced
  vẫn thuộc `preserved`/`skipped`/`failed`, và report bắt buộc
  `detected = replaced + preserved + skipped + failed`, `sourceUnmodified=true` cùng
  package validation và input/output SHA-256.
- `cargo fmt` hiện hành cũng chuẩn hóa định dạng các Rust module hiện hữu mà crate
  đưa vào format gate; không thay đổi logic của các module đó. Hai assertion test cũ
  được sửa đúng contract thực: decode Unicode trả chuỗi Trung văn và test process dùng
  `/bin/sleep` tuyệt đối.

### Tests

Các gate đã chạy:

```bash
cargo fmt --manifest-path apps/macos/src-tauri/Cargo.toml --check
cargo test --manifest-path apps/macos/src-tauri/Cargo.toml --lib
npm run build:desktop
npm run test:mathtype-omml
python -m pytest -q -rs tools/mathtypejx/tests
git diff --check
```

Kết quả:

- Rust full suite: `125 passed, 2 ignored`, không failure; riêng
  `legacy_equations`: `23 passed`.
- Desktop TypeScript/Vite build PASS, 2.314 modules transformed.
- MathType OMML contract PASS: 12 cases, gồm 3 fixture MTEF v5 thật.
- Python core regression: `78 passed, 19 skipped`, không failure.
- `cargo fmt --check` và `git diff --check` PASS.

Security/lifecycle tests phủ invalid/non-v4 UUID, input=output, output collision,
traversal, input/output/job symlink swap, worker crash/non-zero, malformed/oversized
stdout, oversized stderr/manifest/batch, stderr redaction, cancel cô lập process, hai job
concurrent, write/atomic-link/fsync failure, count conservation, source hash mutation,
unknown-file cleanup guard, TTL expiry, restart recovery và production worker missing.

### Tests chưa chạy và giới hạn

- Production worker chưa được đóng gói theo đúng phân chia Phase 7; production missing
  worker hiện fail closed và đã có unit test.
- Chưa chạy UI standalone, Word Ribbon, ký/notarize, Word macOS open/save/reopen hoặc
  Word Windows cross-open; các gate này thuộc các phase sau.
- Hai Rust test được ignore có chủ đích vì cần Office session/probe đích thực.
- 19 skip Python được giữ nguyên: 12 lượt thiếu DOCX/private corpus và 7 lượt thiếu
  Microsoft `MML2OMML.XSL`; không test nào bị sửa để giảm skip.
- **MTEF v3/Equation Editor 3.0 vẫn là release blocker:** fixture repository và corpus
  người dùng cung cấp đều chỉ xác nhận MTEF v5. Phase 3 không tạo fixture giả và không
  tuyên bố hỗ trợ Equation Editor 3.0 đã được xác nhận.
