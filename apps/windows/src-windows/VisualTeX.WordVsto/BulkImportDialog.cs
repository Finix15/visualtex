using System.Drawing;
using System.Text;
using System.Windows.Forms;
using VisualTeX.WindowsOffice.VstoShared;

namespace VisualTeX.WordVsto;

internal sealed class BulkImportDialog : Form
{
    private readonly ComboBox _sourceFormat = new();
    private readonly ComboBox _objectMode = new();
    private readonly TextBox _source = new();
    private readonly Label _summary = new();
    private readonly TextBox _warnings = new();
    private readonly Button _insert = new();
    private WordBulkImportDocument? _parsed;

    internal BulkImportDialog()
    {
        Text = "VisualTeX nhập hàng loạt LaTeX / Markdown";
        StartPosition = FormStartPosition.CenterParent;
        MinimumSize = new Size(820, 600);
        Size = new Size(980, 720);
        Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        MinimizeBox = false;
        ShowIcon = false;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(14),
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 82));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        var title = new Label
        {
            AutoSize = true,
            Text = "Nhập thành văn bản gốc của Word và công thức độc lập",
            Font = new Font(Font, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 5),
        };
        var description = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(900, 0),
            Text = "Văn bản thường trở thành đoạn Word gốc; mỗi công thức cùng dòng hoặc riêng dòng trở thành công thức VisualTeX độc lập, có thể sửa và đổi cỡ riêng.",
            ForeColor = Color.FromArgb(80, 80, 80),
            Margin = new Padding(0, 0, 0, 10),
        };
        var heading = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
        };
        heading.Controls.Add(title);
        heading.Controls.Add(description);
        root.Controls.Add(heading, 0, 0);

        var options = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Margin = new Padding(0, 0, 0, 8),
        };
        options.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "Định dạng nguồn:",
            Padding = new Padding(0, 7, 0, 0),
        });
        _sourceFormat.DropDownStyle = ComboBoxStyle.DropDownList;
        _sourceFormat.Width = 145;
        _sourceFormat.Items.AddRange(new object[] { "Tự động nhận diện", "Markdown", "LaTeX" });
        _sourceFormat.SelectedIndex = 0;
        options.Controls.Add(_sourceFormat);
        options.Controls.Add(new Label
        {
            AutoSize = true,
            Text = "Định dạng công thức:",
            Padding = new Padding(18, 7, 0, 0),
        });
        _objectMode.DropDownStyle = ComboBoxStyle.DropDownList;
        _objectMode.Width = 210;
        _objectMode.Items.AddRange(new object[]
        {
            "OMML gốc của Word (khuyên dùng)",
            "VisualTeX OLE",
        });
        _objectMode.SelectedIndex = 0;
        options.Controls.Add(_objectMode);
        var open = new Button
        {
            AutoSize = true,
            Text = "Mở tệp…",
            Margin = new Padding(18, 0, 0, 0),
        };
        open.Click += (_, _) => OpenFile();
        options.Controls.Add(open);
        var preview = new Button
        {
            AutoSize = true,
            Text = "Phân tích và xem trước",
            Margin = new Padding(8, 0, 0, 0),
        };
        preview.Click += (_, _) => ParseAndPreview(showError: true);
        options.Controls.Add(preview);
        root.Controls.Add(options, 0, 1);

        _source.Dock = DockStyle.Fill;
        _source.Multiline = true;
        _source.AcceptsReturn = true;
        _source.AcceptsTab = true;
        _source.ScrollBars = ScrollBars.Both;
        _source.WordWrap = false;
        _source.Font = new Font("Consolas", 10.5f, FontStyle.Regular, GraphicsUnit.Point);
        _source.Text = "# Ví dụ\r\n\r\nĐây là văn bản và công thức cùng dòng $E=mc^2$。\r\n\r\n$$\r\n\\int_0^1 x^2\\,\\mathrm{d}x=\\frac13\r\n$$";
        _source.TextChanged += (_, _) =>
        {
            _parsed = null;
            _summary.Text = "Nội dung đã thay đổi; hãy phân tích lại bản xem trước.";
            _warnings.Clear();
        };
        root.Controls.Add(_source, 0, 2);

        var status = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = new Padding(0, 8, 0, 8),
        };
        status.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
        status.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
        _summary.Dock = DockStyle.Fill;
        _summary.Padding = new Padding(8);
        _summary.Text = "Nhấn “Phân tích và xem trước” để xem cấu trúc nhập.";
        _summary.BorderStyle = BorderStyle.FixedSingle;
        status.Controls.Add(_summary, 0, 0);
        _warnings.Dock = DockStyle.Fill;
        _warnings.Multiline = true;
        _warnings.ReadOnly = true;
        _warnings.ScrollBars = ScrollBars.Vertical;
        _warnings.BackColor = SystemColors.Window;
        status.Controls.Add(_warnings, 1, 0);
        root.Controls.Add(status, 0, 3);

        var actions = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
        };
        var cancel = new Button
        {
            Text = "Hủy",
            DialogResult = DialogResult.Cancel,
            AutoSize = true,
            Padding = new Padding(10, 3, 10, 3),
        };
        _insert.Text = "Chèn vào Word";
        _insert.AutoSize = true;
        _insert.Padding = new Padding(10, 3, 10, 3);
        _insert.Click += (_, _) =>
        {
            if (!ParseAndPreview(showError: true)) return;
            DialogResult = DialogResult.OK;
            Close();
        };
        actions.Controls.Add(cancel);
        actions.Controls.Add(_insert);
        root.Controls.Add(actions, 0, 4);
        AcceptButton = _insert;
        CancelButton = cancel;
    }

    internal WordBulkImportDocument ParsedDocument =>
        _parsed ?? throw new InvalidOperationException("Nội dung nhập hàng loạt chưa được phân tích.");

    internal string SourceText
    {
        get => _source.Text;
        set => _source.Text = value ?? string.Empty;
    }

    internal WordBulkSourceFormat SelectedSourceFormat
    {
        get => _sourceFormat.SelectedIndex switch
        {
            1 => WordBulkSourceFormat.Markdown,
            2 => WordBulkSourceFormat.Latex,
            _ => WordBulkSourceFormat.Auto,
        };
        set => _sourceFormat.SelectedIndex = value switch
        {
            WordBulkSourceFormat.Markdown => 1,
            WordBulkSourceFormat.Latex => 2,
            _ => 0,
        };
    }

    internal WordBulkFormulaObjectMode SelectedObjectMode
    {
        get => _objectMode.SelectedIndex == 1
            ? WordBulkFormulaObjectMode.Ole
            : WordBulkFormulaObjectMode.Omml;
        set => _objectMode.SelectedIndex = value == WordBulkFormulaObjectMode.Ole ? 1 : 0;
    }

    private bool ParseAndPreview(bool showError)
    {
        try
        {
            _parsed = WordBulkImportParser.Parse(
                _source.Text,
                SelectedSourceFormat,
                SelectedObjectMode);
            _summary.Text =
                $"Đã nhận diện là {_parsed.SourceFormat}; tổng cộng {_parsed.Blocks.Count} khối, " +
                $"{_parsed.TextCharacterCount} ký tự văn bản, " +
                $"{_parsed.InlineFormulaCount} công thức cùng dòng, " +
                $"{_parsed.DisplayFormulaCount} công thức riêng dòng.";
            _warnings.Text = _parsed.Warnings.Count == 0
                ? "Không có cảnh báo phân tích."
                : string.Join(Environment.NewLine, _parsed.Warnings.Select((warning, index) => $"{index + 1}. {warning}"));
            return true;
        }
        catch (Exception error)
        {
            _parsed = null;
            _summary.Text = "Không thể phân tích nội dung hiện tại.";
            _warnings.Text = error.Message;
            if (showError)
            {
                MessageBox.Show(
                    this,
                    error.Message,
                    "Nhập hàng loạt bằng VisualTeX",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            return false;
        }
    }

    private void OpenFile()
    {
        using var dialog = new OpenFileDialog
        {
            Title = "Mở tệp Markdown hoặc LaTeX",
            Filter = "Markdown / LaTeX (*.md;*.markdown;*.tex;*.txt)|*.md;*.markdown;*.tex;*.txt|Tất cả tệp (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
        };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var file = new FileInfo(dialog.FileName);
        if (file.Length > 5_000_000)
        {
            MessageBox.Show(
                this,
                "Tệp lớn hơn 5 MB nên không thể nhập hàng loạt.",
                "Nhập hàng loạt bằng VisualTeX",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }
        _source.Text = File.ReadAllText(file.FullName, DetectEncoding(file.FullName));
        if (file.Extension.Equals(".tex", StringComparison.OrdinalIgnoreCase))
            SelectedSourceFormat = WordBulkSourceFormat.Latex;
        else if (file.Extension.Equals(".md", StringComparison.OrdinalIgnoreCase)
                 || file.Extension.Equals(".markdown", StringComparison.OrdinalIgnoreCase))
            SelectedSourceFormat = WordBulkSourceFormat.Markdown;
        ParseAndPreview(showError: false);
    }

    private static Encoding DetectEncoding(string path)
    {
        using var stream = File.OpenRead(path);
        var prefix = new byte[Math.Min(4, (int)stream.Length)];
        _ = stream.Read(prefix, 0, prefix.Length);
        if (prefix.Length >= 3 && prefix[0] == 0xEF && prefix[1] == 0xBB && prefix[2] == 0xBF)
            return new UTF8Encoding(true, true);
        if (prefix.Length >= 2 && prefix[0] == 0xFF && prefix[1] == 0xFE)
            return Encoding.Unicode;
        if (prefix.Length >= 2 && prefix[0] == 0xFE && prefix[1] == 0xFF)
            return Encoding.BigEndianUnicode;
        return new UTF8Encoding(false, true);
    }
}
