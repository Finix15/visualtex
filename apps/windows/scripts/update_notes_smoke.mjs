import assert from "node:assert/strict";
import { localizeReleaseNotes } from "../src/update/releaseNotes.ts";

const bilingualNotes = `
## Tiếng Việt

### Tính năng mới
- Thêm thông báo cập nhật tự động.
- Hiển thị ghi chú phát hành theo ngôn ngữ hiện tại.

### Sửa lỗi
- Sửa lỗi con trỏ bị kẹt sau khi xóa ký hiệu độ.

## English

### New features
- Added automatic update notifications.
- Added localized release notes.

### Bug fixes
- Fixed the degree-symbol caret getting stuck after deletion.
`;

assert.deepEqual(localizeReleaseNotes(bilingualNotes, "vi"), {
  features: ["Thêm thông báo cập nhật tự động.", "Hiển thị ghi chú phát hành theo ngôn ngữ hiện tại."],
  fixes: ["Sửa lỗi con trỏ bị kẹt sau khi xóa ký hiệu độ."],
  other: [],
});

assert.deepEqual(localizeReleaseNotes(bilingualNotes, "en"), {
  features: [
    "Added automatic update notifications.",
    "Added localized release notes.",
  ],
  fixes: ["Fixed the degree-symbol caret getting stuck after deletion."],
  other: [],
});

const legacyNotes = `VisualTeX improves editing stability.\n\n- Existing release note.`;
assert.deepEqual(localizeReleaseNotes(legacyNotes, "vi"), {
  features: [],
  fixes: [],
  other: ["VisualTeX improves editing stability.", "Existing release note."],
});

console.log("Localized update notes smoke test passed");
