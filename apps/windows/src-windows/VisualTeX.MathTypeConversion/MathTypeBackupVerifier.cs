using System.Security.Cryptography;

namespace VisualTeX.MathTypeConversion;

public static class MathTypeBackupVerifier
{
    public static void CopyAndVerify(
        string sourcePath,
        string backupPath,
        int expectedOleCount,
        string? expectedSourceFingerprint = null)
    {
        if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
            throw new FileNotFoundException("Không tìm thấy file Word nguồn đã lưu.", sourcePath);
        if (expectedOleCount < 0)
            throw new ArgumentOutOfRangeException(nameof(expectedOleCount));
        var extension = Path.GetExtension(sourcePath);
        if (!extension.Equals(".docx", StringComparison.OrdinalIgnoreCase)
            && !extension.Equals(".docm", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Chỉ có thể tạo backup cho DOCX hoặc DOCM.");
        if (File.Exists(backupPath))
            throw new IOException("File backup đích đã tồn tại.");

        var copied = false;
        try
        {
            File.Copy(sourcePath, backupPath, overwrite: false);
            copied = true;
            // Word keeps the saved document open with write access. Sharing only
            // reads here conflicts with that existing writer even though we only
            // read the file ourselves. Allow Word's write/delete handles for the
            // source; the newly-created backup remains exclusively controlled by
            // VisualTeX while it is verified.
            var sourceHash = ComputeSha256(
                sourcePath,
                FileShare.ReadWrite | FileShare.Delete);
            if (!string.IsNullOrWhiteSpace(expectedSourceFingerprint)
                && !string.Equals(
                    sourceHash,
                    expectedSourceFingerprint,
                    StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException(
                    "File nguồn đã thay đổi sau khi quét MathType.");
            var backupHash = ComputeSha256(backupPath, FileShare.Read);
            if (!string.Equals(sourceHash, backupHash, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("SHA-256 của backup không khớp file nguồn.");
            var actualOleCount = MathTypeDocumentScanner.ValidateAndCountMathTypeOleObjects(backupPath);
            if (actualOleCount != expectedOleCount)
                throw new InvalidDataException(
                    $"Backup chứa {actualOleCount} MathType OLE, dự kiến {expectedOleCount}.");
        }
        catch
        {
            if (copied)
            {
                try { File.Delete(backupPath); } catch { }
            }
            throw;
        }
    }

    private static string ComputeSha256(string path, FileShare share)
    {
        using var stream = File.Open(path, FileMode.Open, FileAccess.Read, share);
        using var sha = SHA256.Create();
        return string.Concat(sha.ComputeHash(stream).Select(value => value.ToString("x2")));
    }
}
