import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  localizeReleaseNotes,
  stripPersonalAuthorNames,
} from "../src/update/releaseNotes.ts";

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

const privateChineseName = String.fromCodePoint(24278, 29632, 20581);
const privateEnglishName = ["Liao", "Pojian"].join(" ");
const privateIdentityNotes =
  `## Tiếng Việt\n- Tác giả: ${privateChineseName} (paulhe666)\n` +
  `## English\n- Author: ${privateEnglishName} (paulhe666)`;
for (const language of ["vi", "en"]) {
  const localized = localizeReleaseNotes(privateIdentityNotes, language);
  const visible = [
    ...localized.features,
    ...localized.fixes,
    ...localized.other,
  ].join(" ");
  assert(!visible.includes(privateChineseName));
  assert(!visible.toLocaleLowerCase().includes(privateEnglishName.toLocaleLowerCase()));
}
assert.equal(
  stripPersonalAuthorNames(`Release by ${privateChineseName}`),
  "Release by VisualTeX",
);

const updateDialogSource = await readFile("src/components/UpdateDialog.tsx", "utf8");
const qqGroupCard = await readFile("public/qq-group-card.svg", "utf8");
assert(updateDialogSource.includes('const QQ_GROUP_NUMBER = "1045801770"'));
assert(updateDialogSource.includes('const QQ_GROUP_IMAGE_URL = "/qq-group-card.svg"'));
assert(updateDialogSource.includes('className="update-community-card"'));
assert(updateDialogSource.includes("Join the VisualTeX QQ communication group"));
assert(!updateDialogSource.includes(privateChineseName));
assert(
  !updateDialogSource
    .toLocaleLowerCase()
    .includes(privateEnglishName.toLocaleLowerCase()),
);
assert(qqGroupCard.includes("https://qm.qq.com/q/TppXdoOO8Q") === false);
assert(qqGroupCard.includes("1045801770"));
assert(qqGroupCard.includes("VisualTeX 交流群"));

console.log("Localized update notes and QQ community card smoke test passed");
