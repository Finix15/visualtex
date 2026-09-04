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
