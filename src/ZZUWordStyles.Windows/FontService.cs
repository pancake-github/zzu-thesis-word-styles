using System.Drawing.Text;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace ZZUWordStyles.Windows;

/// <summary>只核验已知字体；不会下载字体或更改系统的字体替代规则。</summary>
public static class FontService
{
    public const string VerifiedKaiTiSha256 = "99092cbb0df301625f46509e85854db8685556551742097ab3fbbc2e2ca0778b";
    private const int VerifiedKaiTiLength = 4_135_804; // 对应校方附件内已经检查过的 Version 2.00。
    private const string FontRegistryPath = @"Software\Microsoft\Windows NT\CurrentVersion\Fonts";
    private const string FontRegistryName = "KaiTi_GB2312 (TrueType)";
    private static readonly (string Name, string[] Aliases, string Usage)[] RequiredFonts =
    [
        ("宋体", ["宋体", "SimSun"], "正文、目录条目、图表题、附属正文、中文页眉与页码"), // 常规中文正文和图表题。
        ("黑体", ["黑体", "SimHei"], "四级标题、中文摘要标题、目录标题与封面学位名称"), // 各级标题。
        ("仿宋", ["仿宋", "FangSong"], "致谢正文"), // 致谢正文。
        ("Arial", ["Arial"], "英文摘要标题"), // 原规范指定的英文摘要标题。
        ("Times New Roman", ["Times New Roman"], "西文正文、英文封面与英文页眉"), // 西文正文与英文封面。
        ("华文行楷", ["华文行楷", "STXingkai"], "封面校名"), // 封面校名。
        ("楷体_GB2312", ["楷体_GB2312", "KaiTi_GB2312"], "封面信息与论文题目") // 按实际字库内容单独核验。
    ];

    /// <summary>枚举字体并核验中文楷体名称实际选中的字库；不改变任何配置。</summary>
    public static IReadOnlyList<FontStatus> CheckFonts()
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // 同时接受中英文家族名。
        using (var fonts = new InstalledFontCollection()) // GDI+ 枚举只读取当前字体集合。
        {
            foreach (var family in fonts.Families)
            {
                using (family) // 及时释放原生字体家族资源。
                {
                    names.Add(family.Name); // 记录当前系统语言下的家族名。
                    names.Add(family.GetName(0x0804)); // 记录中文家族名。
                    names.Add(family.GetName(0x0409)); // 记录英文字体家族名。
                }
            }
        }

        var result = new List<FontStatus>(); // 一种字体对应一个可显示的结果。
        foreach (var required in RequiredFonts)
        {
            bool available = required.Aliases.Any(names.Contains); // 枚举结果用于普通字体存在性检查。
            if (required.Name != "楷体_GB2312")
            {
                result.Add(new FontStatus(required.Name, available, available,
                    available ? "已找到字体家族。用于：" + required.Usage + "。" : "缺少字体；受影响样式：" + required.Usage + "。Word 可能使用替代字体。"));
                continue;
            }

            try
            {
                var selected = ReadSelectedFont("楷体_GB2312"); // 使用模板实际采用的中文名称。
                bool matches = selected.Sha256 == VerifiedKaiTiSha256; // 名称相同并不代表没有字体替代。
                result.Add(new FontStatus(required.Name, available || matches, matches,
                    matches ? "中文家族名实际选中的字库与已核验版本一致。用于：" + required.Usage + "。" : "未选中已核验的楷体_GB2312 字库；受影响样式：" + required.Usage + "，封面字体尚未验收。"));
            }
            catch (Exception error)
            {
                result.Add(new FontStatus(required.Name, available, false, "无法核验实际字库：" + error.Message + "；受影响样式：" + required.Usage + "。"));
            }
        }
        return result; // 缺少字体不触发下载、替代或安装。
    }

    /// <summary>安装用户主动选择的已知 TTF 或 ZIP，仅接受经过核验的字节内容。</summary>
    public static FontInstallResult InstallVerifiedFont(string ttfOrZipPath)
    {
        string target = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft", "Windows", "Fonts", "KaiTi_GB2312_ZZU_99092cbb.ttf"); // 唯一文件名保留现有现代楷体。
        bool createdFile = false; // 只允许在失败时删除本次新建的文件。
        bool changedRegistry = false; // 只撤销本次新写入的注册项。
        bool resourceAdded = false; // 只撤销本次成功添加的会话资源。
        try
        {
            byte[] data = ReadVerifiedSource(ttfOrZipPath); // 先验证全部源内容，再执行任何写入。
            using RegistryKey? existingKey = Registry.CurrentUser.OpenSubKey(FontRegistryPath, false); // 不创建注册表键来进行预检。
            object? oldValue = existingKey?.GetValue(FontRegistryName);
            if (oldValue is not null && (oldValue is not string prior ||
                !string.Equals(Path.GetFullPath(prior), Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException("已有不同的楷体_GB2312 注册项；已保留原配置，请先核对。");
            if (File.Exists(target) && Hash(File.ReadAllBytes(target)) != VerifiedKaiTiSha256)
                throw new InvalidOperationException("目标字体文件已经存在且内容不同；不会覆盖。");
            if (File.Exists(target) && oldValue is not null && ReadSelectedFont("楷体_GB2312").Sha256 == VerifiedKaiTiSha256)
                return new FontInstallResult(true, false, target, "所需字体已安装且实际字库核验通过，无需重复安装。");

            Directory.CreateDirectory(Path.GetDirectoryName(target)!); // 仅建立当前用户字体目录。
            if (!File.Exists(target))
            {
                using var output = new FileStream(target, FileMode.CreateNew, FileAccess.Write, FileShare.None);
                createdFile = true; // 创建成功后立即标记，以覆盖中途写入失败。
                output.Write(data); // 已知字节内容写入新文件。
                output.Flush(true); // 将字体内容同步到磁盘。
            }
            using (var key = Registry.CurrentUser.CreateSubKey(FontRegistryPath, true))
            {
                if (key.GetValue(FontRegistryName) is object concurrent &&
                    !string.Equals(concurrent as string, target, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidOperationException("字体注册项在安装期间发生变化，已停止。");
                if (key.GetValue(FontRegistryName) is null)
                {
                    key.SetValue(FontRegistryName, target, RegistryValueKind.String); // 不覆盖现有注册项。
                    changedRegistry = true; // 仅新建的值需要失败撤回。
                }
            }
            if (Native.AddFontResourceEx(target, 0, IntPtr.Zero) <= 0)
                throw new InvalidOperationException("Windows 未能加载该字体资源。");
            resourceAdded = true; // 资源必须在失败路径配对释放。
            BroadcastFontChange(); // 通知应用更新字体列表，不要求用户重新登录。
            if (ReadSelectedFont("楷体_GB2312").Sha256 != VerifiedKaiTiSha256)
                throw new InvalidOperationException("安装后实际字体仍被替代，未达到封面字体验收要求。");
            return new FontInstallResult(true, createdFile || changedRegistry, target,
                "已按当前用户安装楷体_GB2312，实际字库核验通过；现有楷体及系统替代规则未改变。");
        }
        catch (Exception error)
        {
            var cleanup = new List<string>(); // 清理失败必须明确呈现，不能声称已经回滚。
            if (resourceAdded && !Native.RemoveFontResourceEx(target, 0, IntPtr.Zero))
                cleanup.Add("会话字体资源未能撤回");
            if (changedRegistry)
            {
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey(FontRegistryPath, true);
                    if (string.Equals(key?.GetValue(FontRegistryName) as string, target, StringComparison.OrdinalIgnoreCase))
                        key!.DeleteValue(FontRegistryName, false); // 仅删除仍指向本次目标的新值。
                }
                catch (Exception cleanupError) { cleanup.Add("注册项撤回失败：" + cleanupError.Message); }
            }
            if (createdFile)
            {
                try { File.Delete(target); } // 此路径必须是本次排他创建的字体文件。
                catch (Exception cleanupError) { cleanup.Add("新文件清理失败：" + cleanupError.Message); }
            }
            if (resourceAdded) BroadcastFontChange(); // 撤回后刷新列表。
            string suffix = cleanup.Count == 0 ? "" : "；" + string.Join("；", cleanup);
            return new FontInstallResult(false, cleanup.Count > 0, target, error.Message + suffix);
        }
    }

    private static byte[] ReadVerifiedSource(string source)
    {
        if (!File.Exists(source)) throw new FileNotFoundException("未找到所选字体或压缩包。", source);
        if (Path.GetExtension(source).Equals(".ttf", StringComparison.OrdinalIgnoreCase))
        {
            if (new FileInfo(source).Length != VerifiedKaiTiLength)
                throw new InvalidDataException("字体大小与已核验版本不同，不能安装。");
            byte[] bytes = File.ReadAllBytes(source); // 大小已受限。
            if (Hash(bytes) != VerifiedKaiTiSha256)
                throw new InvalidDataException("字体校验值与已核验版本不同，不能安装。");
            return bytes;
        }
        if (!Path.GetExtension(source).Equals(".zip", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("请选择 TTF 字体或 ZIP 附件。");
        using var archive = ZipFile.OpenRead(source); // 直接读取条目，不把 ZIP 解压到磁盘。
        if (archive.Entries.Count > 5000) throw new InvalidDataException("ZIP 条目过多，已停止检查。");
        int candidates = 0; // 限制恶意压缩包的候选扫描总量。
        foreach (var entry in archive.Entries)
        {
            if (!entry.FullName.EndsWith(".ttf", StringComparison.OrdinalIgnoreCase) || entry.Length != VerifiedKaiTiLength) continue;
            if (++candidates > 20) throw new InvalidDataException("ZIP 中候选字体过多，已停止检查。");
            using var stream = entry.Open(); // 不使用条目路径创建任何文件，避免路径穿越。
            byte[] bytes = new byte[VerifiedKaiTiLength]; // 使用固定上限，拒绝虚假 ZIP 长度导致无限解压。
            stream.ReadExactly(bytes); // 短于已知字库大小即视为损坏。
            if (stream.ReadByte() != -1) throw new InvalidDataException("ZIP 字体条目长度不一致。");
            if (Hash(bytes) == VerifiedKaiTiSha256) return bytes;
        }
        throw new InvalidDataException("ZIP 中没有找到已核验的楷体_GB2312 字体。");
    }

    private static string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();

    private static (string Face, string Sha256) ReadSelectedFont(string family)
    {
        IntPtr dc = Native.CreateCompatibleDC(IntPtr.Zero); // 使用内存绘图上下文，不创建可见窗口。
        IntPtr font = Native.CreateFont(-24, 0, 0, 0, 400, 0, 0, 0, 134, 7, 0, 0, 0, family);
        IntPtr previous = IntPtr.Zero; // 保存原字体，释放前必须恢复。
        try
        {
            if (dc == IntPtr.Zero || font == IntPtr.Zero) throw new InvalidOperationException("无法创建字体验证对象。");
            previous = Native.SelectObject(dc, font);
            var face = new StringBuilder(128); // 获取名称仅用于诊断，验收仍以字节哈希为准。
            Native.GetTextFace(dc, face.Capacity, face);
            uint length = Native.GetFontData(dc, 0, 0, null, 0);
            if (length == 0 || length == uint.MaxValue || length > 100_000_000)
                throw new InvalidOperationException("无法读取实际选中的字库。");
            byte[] bytes = new byte[length]; // 字库大小已设定上限。
            if (Native.GetFontData(dc, 0, 0, bytes, length) != length)
                throw new InvalidOperationException("实际字库未完整读取。");
            return (face.ToString(), Hash(bytes));
        }
        finally
        {
            if (previous != IntPtr.Zero && dc != IntPtr.Zero) Native.SelectObject(dc, previous);
            if (font != IntPtr.Zero) Native.DeleteObject(font);
            if (dc != IntPtr.Zero) Native.DeleteDC(dc);
        }
    }

    private static void BroadcastFontChange() => Native.SendMessageTimeout(new IntPtr(0xffff), 0x001d,
        UIntPtr.Zero, IntPtr.Zero, 2, 1000, out _); // 设置超时，避免其他应用不响应导致挂起。

    private static class Native
    {
        [DllImport("gdi32.dll", EntryPoint = "AddFontResourceExW", CharSet = CharSet.Unicode)]
        internal static extern int AddFontResourceEx(string path, uint flags, IntPtr reserved);
        [DllImport("gdi32.dll", EntryPoint = "RemoveFontResourceExW", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool RemoveFontResourceEx(string path, uint flags, IntPtr reserved);
        [DllImport("gdi32.dll")]
        internal static extern IntPtr CreateCompatibleDC(IntPtr source);
        [DllImport("gdi32.dll", EntryPoint = "CreateFontW", CharSet = CharSet.Unicode)]
        internal static extern IntPtr CreateFont(int height, int width, int escape, int orientation, int weight,
            uint italic, uint underline, uint strikeout, uint charset, uint outputPrecision, uint clipPrecision,
            uint quality, uint pitch, string family);
        [DllImport("gdi32.dll")]
        internal static extern IntPtr SelectObject(IntPtr dc, IntPtr value);
        [DllImport("gdi32.dll")]
        internal static extern uint GetFontData(IntPtr dc, uint table, uint offset, [Out] byte[]? data, uint length);
        [DllImport("gdi32.dll", EntryPoint = "GetTextFaceW", CharSet = CharSet.Unicode)]
        internal static extern int GetTextFace(IntPtr dc, int length, StringBuilder face);
        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DeleteObject(IntPtr value);
        [DllImport("gdi32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool DeleteDC(IntPtr dc);
        [DllImport("user32.dll", EntryPoint = "SendMessageTimeoutW", CharSet = CharSet.Unicode)]
        internal static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, UIntPtr wParam, IntPtr lParam,
            uint flags, uint timeout, out UIntPtr result);
    }
}
