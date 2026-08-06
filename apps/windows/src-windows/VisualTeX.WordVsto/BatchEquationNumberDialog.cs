using System.Drawing;
using System.Windows.Forms;

namespace VisualTeX.WordVsto;

internal sealed class WordBatchEquationNumberingStats
{
    public int TotalCount { get; set; }
    public int NumberedCount { get; set; }
    public int UnnumberedCount => Math.Max(0, TotalCount - NumberedCount);
    public string CurrentFormatId { get; set; } = EquationNumberFormat.ContinuousId;
}

internal sealed class WordBatchEquationNumberingResult
{
    public int TotalCount { get; set; }
    public int AddedCount { get; set; }
    public int RedrawnCount { get; set; }
    public int NumberedCount { get; set; }
    public string FormatId { get; set; } = EquationNumberFormat.ContinuousId;
}

internal sealed class BatchEquationNumberDialog : Form
{
    private sealed class FormatOption
    {
        internal FormatOption(string id)
        {
            Format = EquationNumberFormat.Resolve(id);
        }

        internal EquationNumberFormat Format { get; }
        public override string ToString() => Format.DisplayName;
    }

    private readonly ComboBox _formatBox = new();

    internal string SelectedFormatId =>
        (_formatBox.SelectedItem as FormatOption)?.Format.Id
        ?? EquationNumberFormat.ContinuousId;

    internal BatchEquationNumberDialog(
        WordBatchEquationNumberingStats stats,
        bool redrawExisting)
    {
        Text = "VisualTeX 批量编号";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        AutoScaleMode = AutoScaleMode.Font;
        Font = new Font("Microsoft YaHei UI", 9f);
        MinimumSize = new Size(620, 620);
        Size = new Size(700, 650);
        Padding = new Padding(18);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            GrowStyle = TableLayoutPanelGrowStyle.FixedSize,
            Padding = new Padding(0),
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var title = new Label
        {
            Text = redrawExisting ? "重绘全部行间公式编号" : "为行间公式批量补充编号",
            AutoSize = true,
            Dock = DockStyle.Fill,
            Font = new Font(Font.FontFamily, 12f, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 14),
        };

        var summaryBox = CreateSectionPanel();
        var summaryTitle = CreateSectionTitle("检测结果");
        var summary = CreateBodyLabel(
            $"共检测到 {stats.TotalCount} 个 VisualTeX 行间公式\r\n"
            + $"已有编号：{stats.NumberedCount} 个    未编号：{stats.UnnumberedCount} 个");
        summaryBox.Controls.Add(summary, 0, 1);
        summaryBox.Controls.Add(summaryTitle, 0, 0);
        summaryBox.Margin = new Padding(0, 0, 0, 12);

        var actionBox = CreateSectionPanel();
        var actionTitle = CreateSectionTitle("本次处理");
        var action = CreateBodyLabel(
            redrawExisting
                ? "重建全部公式编号。公式 ID、编号书签和已有交叉引用保持不变。"
                : "只给未编号公式补充编号。已有编号结构保持不变；选择新格式后，显示与交叉引用会同步更新。");
        actionBox.Controls.Add(action, 0, 1);
        actionBox.Controls.Add(actionTitle, 0, 0);
        actionBox.Margin = new Padding(0, 0, 0, 12);

        var formatBox = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 2,
            RowCount = 2,
            Margin = new Padding(0, 0, 0, 8),
        };
        formatBox.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        formatBox.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        formatBox.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        formatBox.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var formatLabel = new Label
        {
            Text = "编号格式",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Font = new Font(Font, FontStyle.Bold),
            Margin = new Padding(0, 5, 14, 5),
        };

        _formatBox.DropDownStyle = ComboBoxStyle.DropDownList;
        _formatBox.Dock = DockStyle.Fill;
        _formatBox.IntegralHeight = true;
        _formatBox.Margin = new Padding(0, 0, 0, 8);
        _formatBox.Items.AddRange(new object[]
        {
            new FormatOption(EquationNumberFormat.ContinuousId),
            new FormatOption(EquationNumberFormat.Heading1DotId),
            new FormatOption(EquationNumberFormat.Heading1DashId),
            new FormatOption(EquationNumberFormat.Heading2DotId),
            new FormatOption(EquationNumberFormat.Heading2DashId),
        });
        var selectedIndex = 0;
        for (var index = 0; index < _formatBox.Items.Count; index++)
        {
            if ((_formatBox.Items[index] as FormatOption)?.Format.Id == stats.CurrentFormatId)
            {
                selectedIndex = index;
                break;
            }
        }
        _formatBox.SelectedIndex = selectedIndex;

        var note = CreateBodyLabel(
            "按章或按节编号时，VisualTeX 会读取 Word“标题 1 / 标题 2”的实际章节号。"
            + "编号采用原生 SEQ、书签和 REF 字段，后续转换 OLE / OMML 不会丢失。",
            SystemColors.GrayText);
        note.Margin = new Padding(0, 0, 0, 0);
        formatBox.Controls.Add(formatLabel, 0, 0);
        formatBox.Controls.Add(_formatBox, 1, 0);
        formatBox.Controls.Add(note, 1, 1);

        var spacer = new Panel
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(0),
        };

        var buttonBar = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlowDirection = FlowDirection.RightToLeft,
            WrapContents = false,
            Margin = new Padding(0, 16, 0, 0),
        };
        var cancelButton = new Button
        {
            Text = "取消",
            DialogResult = DialogResult.Cancel,
            AutoSize = true,
            MinimumSize = new Size(92, 34),
            Margin = new Padding(8, 0, 0, 0),
        };
        var confirmButton = new Button
        {
            Text = redrawExisting ? "开始重绘" : "开始编号",
            DialogResult = DialogResult.OK,
            AutoSize = true,
            MinimumSize = new Size(104, 34),
            Margin = new Padding(8, 0, 0, 0),
        };
        buttonBar.Controls.Add(cancelButton);
        buttonBar.Controls.Add(confirmButton);

        root.Controls.Add(title, 0, 0);
        root.Controls.Add(summaryBox, 0, 1);
        root.Controls.Add(actionBox, 0, 2);
        root.Controls.Add(formatBox, 0, 3);
        root.Controls.Add(spacer, 0, 4);
        root.Controls.Add(buttonBar, 0, 5);
        Controls.Add(root);

        AcceptButton = confirmButton;
        CancelButton = cancelButton;
    }

    private TableLayoutPanel CreateSectionPanel()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = SystemColors.ControlLight,
            Padding = new Padding(14, 11, 14, 11),
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        return panel;
    }

    private Label CreateSectionTitle(string text) => new()
    {
        Text = text,
        AutoSize = true,
        Dock = DockStyle.Fill,
        Font = new Font(Font, FontStyle.Bold),
        Margin = new Padding(0, 0, 0, 6),
    };

    private Label CreateBodyLabel(string text, Color? color = null) => new()
    {
        Text = text,
        AutoSize = true,
        Dock = DockStyle.Fill,
        ForeColor = color ?? SystemColors.ControlText,
        Margin = new Padding(0),
        UseMnemonic = false,
    };
}
