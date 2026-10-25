namespace ZZUWordStyles.Windows;

/// <summary>被动环境检查结果；此操作不会启动 Word。</summary>
public sealed class EnvironmentReport
{
    public bool WordInstalled { get; set; } // 是否找到桌面 Word 的自动化注册。
    public bool WordRunning { get; set; } // 是否需要先保存并关闭用户的 Word。
    public string CandidateNormalPath { get; set; } = ""; // 候选路径仍需安装时向 Word 确认。
    public string WindowsVersion { get; set; } = ""; // 记录实际运行的 Windows 版本。
    public string OfficeVersion { get; set; } = "未知"; // 从注册表读取的 Office 版本。
    public string OfficeBitness { get; set; } = "未知"; // 从注册表读取的 Office 架构。
    public IReadOnlyList<FontStatus> Fonts { get; set; } = Array.Empty<FontStatus>(); // 每种字体的独立结果。
    public List<string> Messages { get; set; } = []; // 可直接呈现给用户的中文说明。
}

/// <summary>独立启动 Word 后的功能验证结果。</summary>
public sealed class VerificationReport
{
    public bool Success { get; set; } // 两轮启动的所有检查均通过时为真。
    public string WordVersion { get; set; } = "未知"; // 通过 Word 自动化读取的版本。
    public string OfficeBitness { get; set; } = "未知"; // 当前 Office 的注册架构。
    public List<string> QuickStyles { get; set; } = []; // 记录 Word 实际返回的名称，便于检查别名显示与本地化。
    public List<string> Messages { get; set; } = []; // 不含论文正文的验收说明。
}

/// <summary>字体存在性与已知字库内容核验。</summary>
public sealed record FontStatus(string Name, bool Available, bool Verified, string Message);

/// <summary>当前用户字体安装结果。</summary>
public sealed record FontInstallResult(bool Success, bool Changed, string InstalledPath, string Message);
