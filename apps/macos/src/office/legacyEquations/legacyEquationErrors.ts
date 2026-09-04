export type LegacyLocale = "vi" | "en";

export function legacyEquationError(error: unknown, locale: LegacyLocale) {
  const raw = error instanceof Error ? error.message : typeof error === "string" ? error : "";
  const key = raw.toLowerCase();
  const messages: Array<[string, string, string]> = [
    ["output already exists", "Tệp đầu ra đã tồn tại. Hãy chọn tên khác.", "The output already exists. Choose another name."],
    ["input and output", "Tệp nguồn và tệp đầu ra phải khác nhau.", "Source and output must be different files."],
    ["worker", "Không thể hoàn tất tiến trình chuyển đổi an toàn.", "The conversion worker could not complete safely."],
    ["malformed", "Dữ liệu chuyển đổi không hợp lệ và đã bị chặn.", "Invalid conversion data was blocked."],
    ["cancel", "Đã hủy chuyển đổi.", "Conversion was cancelled."],
  ];
  const found = messages.find(([needle]) => key.includes(needle));
  if (found) return locale === "en" ? found[2] : found[1];
  return locale === "en" ? "The operation failed safely. No source file was changed." : "Thao tác thất bại an toàn. Tệp nguồn không bị thay đổi.";
}
