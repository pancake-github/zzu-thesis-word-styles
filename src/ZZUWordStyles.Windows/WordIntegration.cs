using System.Diagnostics;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using Microsoft.Win32;

namespace ZZUWordStyles.Windows;

/// <summary>桌面 Word 的独立自动化会话；从不接管或强制终止用户进程。</summary>
public static class WordIntegration
{
    private static readonly object AutomationGate = new(); // 避免本助手同时开启两个 Word 操作。
    public static string LastResolutionMessage { get; private set; } = ""; // 最近一次路径确认的初始化说明，不包含论文内容。

    public static bool IsWordRunning()
    {
        Process[] processes = Process.GetProcessesByName("WINWORD"); // 只枚举进程，不发送任何消息。
        try { return processes.Length != 0; }
        finally { foreach (Process process in processes) process.Dispose(); }
    }

    /// <summary>被动检查；候选路径未经过 Word 查询时不会标记为实际模板路径。</summary>
    public static EnvironmentReport InspectEnvironment()
    {
        var office = ReadOfficeInformation(); // 注册表读取不会激活 Word。
        var report = new EnvironmentReport
        {
            WordInstalled = Type.GetTypeFromProgID("Word.Application", false) is not null,
            WordRunning = IsWordRunning(),
            WindowsVersion = RuntimeInformation.OSDescription,
            OfficeVersion = office.Version,
            OfficeBitness = office.Bitness,
            CandidateNormalPath = FindCandidateNormalPath(),
            Fonts = FontService.CheckFonts()
        };
        report.Messages.Add(report.WordInstalled ? "已找到桌面 Microsoft Word。" : "未找到桌面 Microsoft Word；请先安装或修复 Office。");
        report.Messages.Add(report.WordRunning ? "Word 正在运行；配置或恢复前请保存文档并关闭所有 Word 窗口。" : "Word 当前未运行。");
        report.Messages.Add("模板候选路径：" + report.CandidateNormalPath + "；配置时将向 Word 查询真实路径。");
        report.Messages.Add("此检查没有启动 Word，也没有写入模板或字体。");
        return report;
    }

    /// <summary>安装入口使用；首次运行 Word 时允许生成其自己的初始默认模板。</summary>
    public static string ResolveNormalPath() => RunSta(() =>
    {
        LastResolutionMessage = ""; // 不复用上一次操作的状态。
        EnsureWordClosed(); // 不为已打开的用户文档切换默认模板。
        string path = WithWord((app, scope) =>
        {
            dynamic normal = scope.Track((object)app.NormalTemplate); // 使用 Word 返回的真实模板对象。
            string actual = Path.GetFullPath((string)normal.FullName);
            if (!Path.GetExtension(actual).Equals(".dotm", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("当前默认模板不是 DOTM，暂不支持自动配置：" + actual);
            if (!File.Exists(actual))
            {
                normal.Save(); // 只在用户选择配置时初始化缺失的 Normal。
                LastResolutionMessage = "此前没有默认模板，已由 Word 保存初始 Normal.dotm；安装器将备份这个初始版本。";
            }
            else LastResolutionMessage = "已向 Word 确认真正的默认模板路径。";
            if (!File.Exists(actual)) throw new IOException("Word 未能保存初始默认模板。");
            return actual;
        });
        EnsureWordClosed(); // 返回前确认自建会话已退出，下一步才能修改 ZIP 文件。
        return path;
    });

    /// <summary>两个独立 Word 进程依次新建空白文档，不指定 DOTX，验证默认模板持久生效。</summary>
    public static VerificationReport VerifyNewDocuments(string? outputDirectory = null) => RunSta(() =>
    {
        var report = new VerificationReport { OfficeBitness = ReadOfficeInformation().Bitness };
        string? expectedNormalHash = null; // 第一轮冻结模板内容供第二轮比较。
        string? normalPath = null;
        try
        {
            for (int round = 1; round <= 2; round++)
            {
                EnsureWordClosed(); // 每轮必须是真正的新会话，不附加现有 Word。
                int currentRound = round;
                WithWord((app, scope) =>
                {
                    report.WordVersion = Convert.ToString(app.Version) ?? "未知";
                    dynamic normal = scope.Track((object)app.NormalTemplate);
                    string currentPath = Path.GetFullPath((string)normal.FullName);
                    normalPath ??= currentPath;
                    Require(string.Equals(normalPath, currentPath, StringComparison.OrdinalIgnoreCase), "两次启动的默认模板路径不同。");
                    string hash = HashShared(currentPath); // Word 打开模板时允许共享读取。
                    expectedNormalHash ??= hash;
                    Require(hash == expectedNormalHash, "两次启动之间默认模板内容发生了变化。");
                    dynamic documents = scope.Track((object)app.Documents);
                    dynamic document = scope.Track((object)documents.Add()); // 不传 Template 参数，避免验收独立模板而误报默认安装成功。
                    try
                    {
                        dynamic attached = scope.Track((object)document.AttachedTemplate);
                        Require(string.Equals(Path.GetFullPath((string)attached.FullName), currentPath, StringComparison.OrdinalIgnoreCase), "空白文档未使用实际 Normal 模板。");
                        dynamic range = scope.Track((object)document.Content);
                        Require((string)range.Text == "\r", "新建空白文档中出现额外内容。");
                        dynamic styles = scope.Track((object)document.Styles);
                        var body = SnapshotStyle(styles, -1); // 内置编号不依赖 Office 界面语言。
                        CheckBody(body, "正文样式");
                        var input = SnapshotRange(range);
                        CheckBody(input, "空白文档初始输入位置"); // 防止模板中的直接格式覆盖样式。

                        int[] sizes = [16, 14, 13, 12];
                        int[] before = [24, 24, 12, 12];
                        int[] after = [18, 6, 6, 6];
                        var expectedQuick = new HashSet<string>(StringComparer.Ordinal);
                        expectedQuick.Add(body.Name);
                        for (int level = 1; level <= 4; level++)
                        {
                            var heading = SnapshotStyle(styles, -1 - level);
                            Require(heading.EastAsia == "黑体" && Near(heading.Size, sizes[level - 1]) && heading.Bold == 0, "标题 " + level + " 的字体、字号或加粗不符。");
                            Require(heading.Outline == level && Near(heading.Before, before[level - 1]) && Near(heading.After, after[level - 1]), "标题 " + level + " 的大纲级别或间距不符。");
                            Require(heading.LineRule == 4 && Near(heading.Line, 20), "标题 " + level + " 的固定行距不符。");
                            expectedQuick.Add(heading.Name);
                        }
                        foreach (object id in new object[] { "中文摘要标题", "英文摘要标题", -267, "图题", "表题", "附属正文", "致谢正文" })
                            expectedQuick.Add(SnapshotStyle(styles, id).Name);
                        var abstractTitle = SnapshotStyle(styles, "英文摘要标题");
                        Require(abstractTitle.Latin == "Arial" && Near(abstractTitle.Size, 18) && abstractTitle.Bold == -1, "英文摘要标题应为 Arial 18 磅加粗。");
                        Require(Near(SnapshotStyle(styles, "图题").After, 12) && Near(SnapshotStyle(styles, "表题").After, 6), "图题和表题的不同段后间距未保留。");

                        // 这些低频段落与字符样式也必须存在，合计与上方十二项为二十六项。
                        foreach (object id in new object[] { -20, -21, -22, -36, -32, -33, "英文页眉", "局部加粗", "封面信息四号", "封面信息三号", "封面题目二号", "封面校名初号", "封面学位小初", "英文封面三号" })
                        {
                            dynamic style = styles.Item(id);
                            try { Require(!string.IsNullOrWhiteSpace((string)style.NameLocal), "缺少预期样式。"); }
                            finally { Release(style); }
                        }
                        dynamic cover = styles.Item((object)"封面题目二号");
                        dynamic? coverFont = null;
                        try
                        {
                            coverFont = cover.Font;
                            Require((string)coverFont.NameFarEast == "楷体_GB2312", "封面未使用已核验的中文字体家族名。");
                        }
                        finally { Release(coverFont); Release(cover); }

                        var actualQuick = new HashSet<string>(StringComparer.Ordinal);
                        int count = styles.Count;
                        for (int index = 1; index <= count; index++)
                        {
                            dynamic style = styles.Item((object)index);
                            try { if (Convert.ToBoolean(style.QuickStyle)) actualQuick.Add((string)style.NameLocal); }
                            finally { Release(style); }
                        }
                        Require(actualQuick.SetEquals(expectedQuick) && actualQuick.Count == 12, "常用样式库并非约定的十二项。");
                        report.QuickStyles = actualQuick.OrderBy(name => name, StringComparer.Ordinal).ToList(); // 显示名称由 Word 实际读取，不从模板文件推断。
                        report.Messages.Add($"第 {currentRound} 次独立启动通过：正文、初始输入位置、26 项样式、12 项常用入口、四级标题及图表题间距。");
                    }
                    finally { document.Close(0); } // 仅关闭本次创建的空白测试文档，不保存到最近文件。
                    return true;
                });
                Require(normalPath is not null && HashShared(normalPath) == expectedNormalHash, "退出 Word 后默认模板内容发生变化。");
            }
            report.Success = true;
            report.Messages.Add("两次启动后模板字节校验值保持一致；未保存或修改已有文档。");
        }
        catch (Exception error)
        {
            report.Success = false;
            report.Messages.Add("验证未通过：" + error.Message); // 检查不通过不继续输出“已生效”。
        }
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory); // 用户显式指定输出时才保存验收报告。
            string destination = Path.Combine(outputDirectory, "Word验证_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N")[..8] + ".json");
            using var file = new FileStream(destination, FileMode.CreateNew, FileAccess.Write, FileShare.None);
            JsonSerializer.Serialize(file, report, new JsonSerializerOptions { WriteIndented = true });
        }
        return report;
    });

    private static void CheckBody(StyleSnapshot body, string label)
    {
        Require(body.EastAsia == "宋体" && body.Latin == "Times New Roman", label + "的中英文字体不符。");
        Require(Near(body.Size, 12) && body.LineRule == 4 && Near(body.Line, 20), label + "的字号或固定行距不符。");
        Require(Near(body.FirstCharacters, 2) && Near(body.Before, 0) && Near(body.After, 0), label + "的缩进或段间距不符。");
    }

    private static StyleSnapshot SnapshotStyle(dynamic styles, object index)
    {
        dynamic style = styles.Item(index);
        try
        {
            StyleSnapshot result = SnapshotRange(style);
            return result with { Name = (string)style.NameLocal };
        }
        finally { Release(style); }
    }

    private static StyleSnapshot SnapshotRange(dynamic range)
    {
        dynamic? font = null;
        dynamic? paragraph = null;
        try
        {
            font = range.Font;
            paragraph = range.ParagraphFormat;
            return new StyleSnapshot("", (string)font.NameFarEast, (string)font.NameAscii, Convert.ToDouble(font.Size),
                Convert.ToInt32(font.Bold), Convert.ToInt32(paragraph.LineSpacingRule), Convert.ToDouble(paragraph.LineSpacing),
                Convert.ToDouble(paragraph.SpaceBefore), Convert.ToDouble(paragraph.SpaceAfter),
                Convert.ToDouble(paragraph.CharacterUnitFirstLineIndent), Convert.ToInt32(paragraph.OutlineLevel));
        }
        finally { Release(paragraph); Release(font); }
    }

    private sealed record StyleSnapshot(string Name, string EastAsia, string Latin, double Size, int Bold,
        int LineRule, double Line, double Before, double After, double FirstCharacters, int Outline);

    private static T WithWord<T>(Func<dynamic, ComScope, T> operation)
    {
        EnsureWordClosed(); // 紧邻实例创建再次检查，减小与用户操作竞争的窗口。
        Type type = Type.GetTypeFromProgID("Word.Application", false)
            ?? throw new InvalidOperationException("未找到桌面 Microsoft Word；此功能不支持 WPS 或网页版 Word。");
        dynamic? app = null;
        int? security = null;
        int? alerts = null;
        int[] ownProcessIds = [];
        using var scope = new ComScope();
        try
        {
            app = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Word 自动化实例未能创建。");
            ownProcessIds = GetWordProcessIds(); // 之前没有 Word；只等待这些进程，不强制终止。
            security = Convert.ToInt32(app.AutomationSecurity); // 必须恢复真实旧值，不能写死为某一策略。
            alerts = Convert.ToInt32(app.DisplayAlerts);
            app.AutomationSecurity = 3; // 本次自动化期间不运行宏。
            app.DisplayAlerts = 0;
            app.Visible = false;
            return operation(app, scope);
        }
        finally
        {
            scope.Dispose(); // 先释放 Document/Style/Template 等子引用。
            bool preserveUnexpectedDocuments = false; // 用户在验收期间打开的文档必须保留。
            if (app is not null)
            {
                try
                {
                    try { if (security.HasValue) app.AutomationSecurity = security.Value; }
                    catch (COMException) { /* Word 已退出时没有可恢复的活实例。 */ }
                    try { if (alerts.HasValue) app.DisplayAlerts = alerts.Value; }
                    catch (COMException) { /* 同上；不更改其他 Word 实例。 */ }
                    dynamic? documents = null;
                    try
                    {
                        documents = app.Documents;
                        preserveUnexpectedDocuments = Convert.ToInt32(documents.Count) > 0; // 本库自己的测试文档此时已在内层 finally 关闭。
                    }
                    finally { Release(documents); }
                    ownProcessIds = ownProcessIds.Concat(GetWordProcessIds()).Distinct().ToArray(); // Word 的启动器可能已退出，退出前补充实际服务进程。
                    if (!preserveUnexpectedDocuments) app.Quit(0); // 只在没有任何文档时退出自建 Word。
                    else app.Visible = true; // 若用户或加载项打开了额外文档，保留并显示它，而不是放弃其内容。
                }
                finally { Release(app); }
            }
            GC.Collect(); // 清理动态绑定可能留下的可回收 COM 包装。
            GC.WaitForPendingFinalizers();
            if (preserveUnexpectedDocuments)
                throw new InvalidOperationException("自建 Word 会话中出现额外文档，已保留该窗口并停止后续操作。请保存并关闭 Word 后重试。");
            var exitWait = Stopwatch.StartNew(); // 为整个退出阶段设置统一上限，避免多个进程累积长时间等待。
            foreach (int pid in ownProcessIds)
            {
                try
                {
                    using Process process = Process.GetProcessById(pid);
                    if (!process.WaitForExit(Math.Max(0, 15_000 - (int)exitWait.ElapsedMilliseconds)))
                        throw new InvalidOperationException("自建 Word 会话尚未正常退出；已停止后续文件写入，请关闭 Word 后重试。");
                }
                catch (ArgumentException) { /* 进程已正常退出。 */ }
            }
            // Office 可能短暂产生晚于第一次 PID 快照的退出进程；等待它们结束，但从不终止进程。
            while (IsWordRunning() && exitWait.ElapsedMilliseconds < 15_000) Thread.Sleep(100);
            if (IsWordRunning())
                throw new InvalidOperationException("Word 尚未退出或又有窗口打开；已停止后续操作，请保存并关闭 Word 后重试。");
        }
    }

    private static T RunSta<T>(Func<T> action)
    {
        lock (AutomationGate)
        {
            T? value = default;
            Exception? failure = null;
            var thread = new Thread(() =>
            {
                try { value = action(); }
                catch (Exception error) { failure = error; }
            }) { IsBackground = true, Name = "ZZU Word STA" };
            thread.SetApartmentState(ApartmentState.STA); // Word 自动化只在单线程单元执行。
            thread.Start();
            thread.Join(); // UI 应从后台任务调用本入口，以保持窗口响应。
            if (failure is not null) ExceptionDispatchInfo.Capture(failure).Throw();
            return value!;
        }
    }

    private static void EnsureWordClosed()
    {
        if (IsWordRunning()) throw new InvalidOperationException("请先保存文档并关闭所有 Word 窗口，再运行此操作。助手不会关闭或强制终止用户文档。");
    }

    private static int[] GetWordProcessIds()
    {
        Process[] processes = Process.GetProcessesByName("WINWORD");
        try { return processes.Select(process => process.Id).ToArray(); }
        finally { foreach (Process process in processes) process.Dispose(); }
    }

    private static string FindCandidateNormalPath()
    {
        foreach (string version in new[] { "16.0", "15.0", "14.0" })
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + version + @"\Common\General");
            if (key?.GetValue("UserTemplates") is string path && !string.IsNullOrWhiteSpace(path))
            {
                string expanded = Environment.ExpandEnvironmentVariables(path);
                if (Path.IsPathFullyQualified(expanded)) return Path.Combine(expanded, "Normal.dotm");
            }
        }
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates", "Normal.dotm");
    }

    private static (string Version, string Bitness) ReadOfficeInformation()
    {
        foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
            using var key = root.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            if (key is not null)
                return (key.GetValue("VersionToReport") as string ?? "未知", key.GetValue("Platform") as string ?? "未知");
        }
        return ("未知（未发现 Click-to-Run 版本记录）", "未知"); // MSI 安装仍可通过 COM 正常使用，不凭运行器位数猜测 Office 位数。
    }

    private static string HashShared(string path)
    {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        return Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant();
    }

    private static bool Near(double left, double right) => Math.Abs(left - right) < 0.01;
    private static void Require(bool success, string message) { if (!success) throw new InvalidOperationException(message); }
    private static void Release(object? value)
    {
        if (value is null || !Marshal.IsComObject(value)) return;
        try { Marshal.FinalReleaseComObject(value); }
        catch (InvalidComObjectException) { /* Word 可能对同一 COM 对象返回同一个已经释放的包装。 */ }
    }

    private sealed class ComScope : IDisposable
    {
        private readonly List<object> owned = []; // 按创建顺序保留引用，最后反序释放。
        public dynamic Track(object value)
        {
            if (!owned.Any(existing => ReferenceEquals(existing, value))) owned.Add(value); // 相同 RCW 只释放一次。
            return value;
        }
        public void Dispose()
        {
            for (int index = owned.Count - 1; index >= 0; index--) Release(owned[index]);
            owned.Clear(); // Dispose 可由 finally 和 using 安全重复调用。
        }
    }
}
