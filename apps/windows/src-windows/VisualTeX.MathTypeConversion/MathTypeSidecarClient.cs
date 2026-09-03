using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;

namespace VisualTeX.MathTypeConversion;

public sealed class MathTypeSidecarRequest { public string FormulaId { get; set; } = string.Empty; public byte[] OleBytes { get; set; } = Array.Empty<byte>(); }
public sealed class MathTypeSidecarResult
{
    public string FormulaId { get; set; } = string.Empty; public string Status { get; set; } = string.Empty;
    public string Risk { get; set; } = "blocked"; public int? MtefVersion { get; set; }
    public string Fingerprint { get; set; } = string.Empty; public string MathMl { get; set; } = string.Empty;
    public string Omml { get; set; } = string.Empty; public List<string> Warnings { get; set; } = new();
    public List<string> Errors { get; set; } = new(); public string ReasonCode { get; set; } = string.Empty;
    public MathTypeParseStatus ParseStatus => Status == "convertible" ? MathTypeParseStatus.Convertible : Status == "unsupported" ? MathTypeParseStatus.Unsupported : MathTypeParseStatus.Corrupt;
    public bool CanConvert => ParseStatus == MathTypeParseStatus.Convertible && MathMl.Length != 0 && Omml.Length != 0 && Errors.Count == 0;
}

public static class MathTypeSidecarClient
{
    private const int ProtocolVersion = 1, TimeoutMilliseconds = 180_000;
    private static readonly JsonSerializerOptions JsonOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, PropertyNameCaseInsensitive = true };
    public static IReadOnlyList<MathTypeSidecarResult> ConvertBatch(IReadOnlyList<MathTypeSidecarRequest> requests, int maxWorkers = 4)
    {
        if (requests.Count == 0) return Array.Empty<MathTypeSidecarResult>();
        var assemblyDirectory = Path.GetDirectoryName(typeof(MathTypeSidecarClient).Assembly.Location);
        if (string.IsNullOrWhiteSpace(assemblyDirectory)) assemblyDirectory = AppContext.BaseDirectory;
        var runtime = Path.Combine(assemblyDirectory, "mathtype-runtime"); var python = Path.Combine(runtime, "python.exe"); var worker = Path.Combine(runtime, "worker.py");
        if (!File.Exists(python) || !File.Exists(worker)) throw new FileNotFoundException("Thiếu engine MathType của VisualTeX. Hãy chạy Tích hợp sửa chữa.");
        var operationId = Guid.NewGuid().ToString("N"); var root = Path.Combine(Path.GetTempPath(), "VisualTeX", "mathtype", operationId); Directory.CreateDirectory(root);
        try
        {
            var inputItems = new List<object>(); var expected = new Dictionary<string, string>(StringComparer.Ordinal);
            for (var index = 0; index < requests.Count; index++)
            {
                var request = requests[index];
                if (string.IsNullOrWhiteSpace(request.FormulaId) || expected.ContainsKey(request.FormulaId)) throw new InvalidDataException("Formula ID MathType không hợp lệ hoặc bị trùng.");
                if (request.OleBytes.Length == 0 || request.OleBytes.Length > MathTypeLimits.MaximumPayloadBytes) throw new InvalidDataException("OLE MathType vượt giới hạn an toàn.");
                var fingerprint = ComputeFingerprint(request.OleBytes); expected.Add(request.FormulaId, fingerprint); var name = $"formula-{index:D4}.bin";
                File.WriteAllBytes(Path.Combine(root, name), request.OleBytes); inputItems.Add(new { formulaId = request.FormulaId, olePath = name, fingerprint });
            }
            var manifestPath = Path.Combine(root, "request.json"); var outputPath = Path.Combine(root, "response.json");
            File.WriteAllText(manifestPath, JsonSerializer.Serialize(new { protocolVersion = ProtocolVersion, operationId, officeXsltPath = FindOfficeXslt(), maxWorkers = Math.Max(1, Math.Min(4, maxWorkers)), items = inputItems }, JsonOptions));
            using var process = Process.Start(new ProcessStartInfo { FileName = python, Arguments = $"-I -X utf8 \"{worker}\" \"{manifestPath}\" \"{outputPath}\"", WorkingDirectory = runtime, UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden, RedirectStandardError = true, RedirectStandardOutput = true }) ?? throw new InvalidOperationException("Không khởi động được engine MathType.");
            var stderrTask = process.StandardError.ReadToEndAsync();
            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            if (!process.WaitForExit(TimeoutMilliseconds)) { try { process.Kill(); } catch { } throw new TimeoutException("Engine MathType vượt thời hạn 3 phút; tài liệu chưa bị thay đổi."); }
            var stderr = stderrTask.GetAwaiter().GetResult();
            _ = stdoutTask.GetAwaiter().GetResult();
            if (process.ExitCode != 0 || !File.Exists(outputPath)) throw new InvalidOperationException($"Engine MathType thất bại (mã {process.ExitCode}). {TrimDiagnostic(stderr)}");
            var response = JsonSerializer.Deserialize<SidecarResponse>(File.ReadAllText(outputPath), JsonOptions) ?? throw new InvalidDataException("Phản hồi engine MathType trống.");
            if (response.ProtocolVersion != ProtocolVersion || response.OperationId != operationId || response.Items.Count != requests.Count) throw new InvalidDataException("Phản hồi engine MathType không khớp yêu cầu.");
            foreach (var item in response.Items)
            {
                if (!expected.TryGetValue(item.FormulaId, out var fingerprint)
                    || !string.Equals(fingerprint, item.Fingerprint, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("Fingerprint hoặc formula ID từ engine MathType không hợp lệ.");
                expected.Remove(item.FormulaId);
            }
            if (expected.Count != 0) throw new InvalidDataException("Engine MathType bỏ thiếu công thức.");
            return response.Items;
        }
        finally { try { Directory.Delete(root, recursive: true); } catch { } }
    }
    public static string ComputeFingerprint(byte[] value) { using var sha = SHA256.Create(); return string.Concat(sha.ComputeHash(value).Select(item => item.ToString("x2"))); }
    private static string FindOfficeXslt() { foreach (var path in new[] { @"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL", @"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL" }) if (File.Exists(path)) return path; throw new FileNotFoundException("Không tìm thấy MML2OMML.XSL của Microsoft Office."); }
    private static string TrimDiagnostic(string value) { var text = value?.Trim() ?? ""; return text.Substring(0, Math.Min(300, text.Length)); }
    private sealed class SidecarResponse { public int ProtocolVersion { get; set; } public string OperationId { get; set; } = string.Empty; public List<MathTypeSidecarResult> Items { get; set; } = new(); }
}
