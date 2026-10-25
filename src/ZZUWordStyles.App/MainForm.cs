using System.Diagnostics; // 仅用于打开用户主动选择的模板、目录和官方网页。
using System.Security.Cryptography; // 恢复确认绑定到用户看到的实际文件版本。
using ZZUWordStyles.Core; // 安装、恢复与 DOCX 副本导入。
using ZZUWordStyles.Windows; // Word 生命周期与字体检查。

namespace ZZUWordStyles.App;

internal sealed class MainForm : Form
{
    private readonly TextBox report = new() { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle }; // 日志可以选择复制，默认不会保存或上传。
    private readonly Label status = new() { AutoSize = true, Text = "就绪 · 先检查环境，再选择配置方式", ForeColor = Color.FromArgb(45, 80, 74), Padding = new Padding(0, 8, 0, 8) }; // 显示当前动作是否完成。
    private readonly ListBox backups = new() { Dock = DockStyle.Fill, IntegralHeight = false }; // 备份列表不自动选中并恢复。
    private readonly List<Button> buttons = []; // 操作期间统一禁用，避免重入。
    private bool busy; // 防止在事务执行中关闭窗口。

    public MainForm()
    {
        Text = "郑大学位论文样式助手 · " + AssetStore.Version; // 标明预览版身份。
        Font = new Font("Microsoft YaHei UI", 10); // 界面字体与论文排版字体独立。
        MinimumSize = new Size(900, 700); // 保证主要按钮与结果区可见。
        Size = new Size(1060, 800); // 普通屏幕可完整显示。
        StartPosition = FormStartPosition.CenterScreen; // 由系统定位当前显示器。
        BackColor = Color.FromArgb(247, 249, 250); // 使用低干扰的浅色背景。
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(24), ColumnCount = 1, RowCount = 5 }; // 自适应窗口和 DPI。
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 54));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 260));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 38));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.Controls.Add(new Label { Text = "郑大学位论文样式助手", Font = new Font(Font.FontFamily, 22, FontStyle.Bold), AutoSize = true, ForeColor = Color.FromArgb(25, 52, 67) }); // 清楚说明工具用途。
        layout.Controls.Add(new Label { Text = "26 项论文样式 · 12 项常用入口 · 自动备份与恢复\n修改默认模板不会自动修改旧文档；已有论文请使用“导入 DOCX 副本”。", Dock = DockStyle.Fill, AutoSize = false }); // 将最容易误解的规则放在首页。
        var tabs = new TabControl { Dock = DockStyle.Fill }; // 将主要操作、恢复和说明分开显示。
        var home = new TabPage("开始配置") { Padding = new Padding(12) };
        var restore = new TabPage("备份与恢复") { Padding = new Padding(12) };
        var help = new TabPage("说明与下载") { Padding = new Padding(12) };
        tabs.TabPages.AddRange([home, restore, help]);
        var grid = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, RowCount = 2 }; // 六项动作采用同等宽度的可读按钮。
        for (var index = 0; index < 3; index++) grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 33.333F));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
        grid.Controls.Add(ActionButton("1  检查环境\n只读检查 Word、字体与样式", Check));
        grid.Controls.Add(ActionButton("2  配置 Word 默认样式\n先备份，再安全写入", Install));
        grid.Controls.Add(ActionButton("导入 DOCX 副本\n保留原文件和直接格式", Import));
        grid.Controls.Add(ActionButton("验证新建文档\n重启 Word 两次检查结果", Verify));
        grid.Controls.Add(ActionButton("导入所需字体\n选择校方 ZIP 或 TTF", ImportFont));
        grid.Controls.Add(ActionButton("打开样式示例\n查看标题、正文与图表题", () => OpenFile(true)));
        home.Controls.Add(grid);
        var backupLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1 };
        backupLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));
        backupLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
        backupLayout.Controls.Add(backups, 0, 0);
        var backupActions = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 }; // 固定单列，防止窗口缩放时第三个按钮被换列裁切。
        backupActions.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        for (var index = 0; index < 3; index++) backupActions.RowStyles.Add(new RowStyle(SizeType.Percent, 33.333F));
        backupActions.Controls.Add(ActionButton("刷新备份列表", RefreshBackups));
        backupActions.Controls.Add(ActionButton("恢复选中的备份", Restore));
        backupActions.Controls.Add(ActionButton("打开备份目录", OpenBackups));
        backupLayout.Controls.Add(backupActions, 1, 0);
        restore.Controls.Add(backupLayout);
        var helpLayout = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true };
        helpLayout.Controls.Add(new Label { Text = "非官方工具，依据研究生院 2024 年发布附件。\n仅配置文字与段落样式，保留你的页面和页眉页脚布局。\n字体未打包；缺失字体会单独提示。文档内容只在本机处理。", AutoSize = true, MaximumSize = new Size(850, 0), Margin = new Padding(6, 6, 6, 18) });
        helpLayout.SetFlowBreak(helpLayout.Controls[0], true);
        helpLayout.Controls.Add(ActionButton("使用独立模板", () => OpenFile(false)));
        helpLayout.Controls.Add(ActionButton("字体官方下载页", () => OpenUrl("https://www5.zzu.edu.cn/student/info/1003/13037.htm")));
        helpLayout.Controls.Add(ActionButton("使用说明与更新", () => OpenUrl("https://github.com/pancake-github/zzu-thesis-word-styles")));
        help.Controls.Add(helpLayout);
        layout.Controls.Add(tabs);
        layout.Controls.Add(status);
        layout.Controls.Add(report);
        Controls.Add(layout);
        report.Text = "选择“检查环境”开始。\r\n\r\n配置默认样式及实际验证前，请先保存并关闭 Word。工具不会代替你关闭文档。\r\n\r\n已有文档保留自己的样式；导入副本后，手工设置的字体、字号仍可能覆盖样式。\r\n\r\n预览版：具体 Windows / Office 测试范围请查看仓库兼容性记录。";
        FormClosing += (_, args) => { if (busy) { args.Cancel = true; MessageBox.Show(this, "当前操作尚未结束，请等待结果后关闭。", "操作进行中"); } }; // 不在写入中断开流程。
    }

    private Button ActionButton(string text, Action action)
    {
        var button = new Button { Text = text, AutoSize = false, Size = new Size(225, 52), Dock = DockStyle.Fill, Margin = new Padding(5), BackColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand, MinimumSize = new Size(175, 48) }; // 保持键盘可访问的原生按钮。
        button.FlatAppearance.BorderColor = Color.FromArgb(207, 217, 222);
        button.Click += (_, _) => { if (!busy) action(); }; // 操作入口统一限制重入。
        buttons.Add(button);
        return button;
    }

    private async void Run(string title, Func<string> operation)
    {
        busy = true; // 后台工作开始前关闭其他入口。
        foreach (var button in buttons) button.Enabled = false;
        status.Text = title + "，请稍候…";
        try
        {
            var result = await RunSta(operation); // Word COM 始终在专门 STA 线程，界面仍可响应。
            report.Text = result;
            status.Text = title + "已结束，请查看下方结果"; // 完成不等于全部通过，实际结论写在报告中。
        }
        catch (Exception exception)
        {
            status.Text = "操作未完成";
            report.Text = exception.Message; // 只展示当前用户本机的错误，不发送外部日志。
        }
        finally
        {
            busy = false;
            foreach (var button in buttons) button.Enabled = true;
        }
    }

    private static Task<string> RunSta(Func<string> operation)
    {
        var completion = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously); // 将线程结果传回窗口。
        var thread = new Thread(() => { try { completion.SetResult(operation()); } catch (Exception exception) { completion.SetException(exception); } }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return completion.Task;
    }

    private void Check() => Run("环境检查", () =>
    {
        var environment = WordIntegration.InspectEnvironment(); // 不创建 Word，不写模板。
        var lines = new List<string>(environment.Messages);
        lines.AddRange(environment.Fonts.Select(font => font.Name + "：" + font.Message));
        if (!string.IsNullOrWhiteSpace(environment.CandidateNormalPath) && File.Exists(environment.CandidateNormalPath))
        {
            try
            {
                var check = StylePackage.Inspect(File.ReadAllBytes(environment.CandidateNormalPath), AssetStore.Read());
                lines.Add("候选默认模板样式：" + (check.IsInstalled ? "与本版一致" : "尚未完整匹配"));
                lines.AddRange(check.Messages);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                lines.Add("模板暂时无法读取，请保存并关闭 Word 后重试。其余环境检查已保留。原因：" + exception.Message); // 单个锁冲突不能抹去完整环境报告。
            }
        }
        lines.Add("此检查中的路径来自本机配置；安装时由 Word 确认实际路径。缺失字体会影响对应文字显示。");
        return string.Join(Environment.NewLine, lines);
    });

    private static void RequireClosedWord()
    {
        if (WordIntegration.IsWordRunning()) throw new InvalidOperationException("请先保存并关闭所有 Word 窗口，再重试。工具不会关闭你的文档。"); // 由用户处理未保存内容。
    }

    private void Install()
    {
        if (MessageBox.Show(this, "将配置以后新建空白文档的 26 项样式，并精简为 12 项常用入口。\n\n会先完整备份当前默认模板，保留页面设置、宏、快捷键和其他自定义样式。请先保存并关闭 Word。", "配置 Word 默认样式", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) != DialogResult.OK) return; // 安装者在自己的电脑上确认具体修改。
        Run("配置默认样式", () =>
        {
            RequireClosedWord();
            var target = WordIntegration.ResolveNormalPath(); // 不使用硬编码的用户目录。
            RequireClosedWord();
            var result = FileTransactions.Install(target, AssetStore.Materialize(), AssetStore.BackupRoot, RequireClosedWord);
            var fonts = FontService.CheckFonts();
            return WordIntegration.LastResolutionMessage + "\r\n" + Format(result) + "\r\n\r\n" + string.Join("\r\n", fonts.Where(font => !font.Available || !font.Verified).Select(font => font.Name + "：" + font.Message)) + (result.Success ? "\r\n\r\n下一步：点击“验证新建文档”。已有 DOCX 需单独导入副本。" : "\r\n\r\n请先处理报告中的问题，再重试。");
        });
    }

    private void Import()
    {
        using var dialog = new OpenFileDialog { Title = "选择要导入样式的 DOCX，原文件会保留", Filter = "Word 文档 (*.docx)|*.docx", CheckFileExists = true };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var selected = dialog.FileName;
        Run("导入 DOCX 副本", () => Format(FileTransactions.ImportCopy(selected, AssetStore.Materialize())) + "\r\n\r\n只替换对应样式，未清除直接格式。使用自定义样式的段落需在 Word 中选择对应论文样式。");
    }

    private void Verify() => Run("验证新建文档", () =>
    {
        RequireClosedWord();
        var result = WordIntegration.VerifyNewDocuments(); // 创建并关闭自己的空白测试文档。
        return (result.Success ? "Word 实际验证通过" : "Word 实际验证未完全通过") + "\r\n" + string.Join("\r\n", result.Messages);
    });

    private void ImportFont()
    {
        using var dialog = new OpenFileDialog { Title = "选择校方附件 ZIP 或楷体_GB2312.ttf", Filter = "字体或校方附件 (*.ttf;*.zip)|*.ttf;*.zip", CheckFileExists = true };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var selected = dialog.FileName;
        Run("导入所需字体", () =>
        {
            var result = FontService.InstallVerifiedFont(selected); // 仅接受已核对的指定字体字节，不解压其他附件。
            return result.Message + "\r\n" + result.InstalledPath + "\r\n\r\n若 Word 已打开，请保存并重新启动 Word 后检查字体。";
        });
    }

    private void RefreshBackups()
    {
        try
        {
            backups.Items.Clear();
            foreach (var receipt in FileTransactions.ListBackups(AssetStore.BackupRoot)) backups.Items.Add(new BackupItem(receipt)); // 仅读取本程序回执。
            status.Text = "找到 " + backups.Items.Count + " 个备份记录";
        }
        catch (Exception exception) { report.Text = exception.Message; }
    }

    private void Restore()
    {
        if (backups.SelectedItem is not BackupItem item) { MessageBox.Show(this, "请刷新列表并选择一个备份。", "选择备份"); return; }
        try
        {
            RequireClosedWord(); // 先排除未保存文档和文件占用。
            var currentHash = Convert.ToHexStringLower(SHA256.HashData(File.ReadAllBytes(item.Receipt.TargetPath))); // 确认时冻结当前版本。
            var changed = currentHash != item.Receipt.AfterSha256 && currentHash != item.Receipt.BeforeSha256;
            var message = changed
                ? "检测到默认模板在所选安装记录之后又发生了修改。\n恢复会替换这些后续修改；工具会先完整备份当前版本。\n\n确实要恢复所选旧版本吗？"
                : "将恢复所选默认模板备份。恢复前会另外完整保存当前版本。";
            if (MessageBox.Show(this, message, "恢复默认模板", MessageBoxButtons.YesNo, changed ? MessageBoxIcon.Warning : MessageBoxIcon.Information, MessageBoxDefaultButton.Button2) != DialogResult.Yes) return;
            Run("恢复备份", () => { RequireClosedWord(); return Format(FileTransactions.Restore(item.Receipt.ReceiptPath, changed, RequireClosedWord, currentHash)); }); // 确认后若再次变化，核心拒绝写入。
        }
        catch (Exception exception) { report.Text = exception.Message; }
    }

    private void OpenBackups()
    {
        if (!Directory.Exists(AssetStore.BackupRoot)) { report.Text = "尚无本工具创建的备份。完成一次安装后可在这里查看。"; return; }
        OpenUrl(AssetStore.BackupRoot); // 使用系统文件管理器显示实际目录。
    }
    private void OpenFile(bool example) { try { OpenUrl(AssetStore.Materialize(example)); } catch (Exception exception) { report.Text = exception.Message; } }
    private void OpenUrl(string target) { try { Process.Start(new ProcessStartInfo(target) { UseShellExecute = true }); } catch (Exception exception) { report.Text = exception.Message; } }
    private static string Format(OperationResult result) => (result.Success ? "完成：" : "未完成：") + result.Message + (result.OutputPath is null ? "" : "\r\n文件：" + result.OutputPath) + (result.ReceiptPath is null ? "" : "\r\n回执：" + result.ReceiptPath) + "\r\n" + string.Join("\r\n", result.Warnings); // 真实状态由核心事务返回。
    private sealed record BackupItem(BackupReceipt Receipt)
    {
        public override string ToString() => Receipt.CreatedAt.LocalDateTime.ToString("yyyy-MM-dd HH:mm:ss") + " · " + (Receipt.Operation == "restore" ? "恢复前备份" : "安装前备份"); // 日期和中文用途便于用户选择版本。
    }
}
