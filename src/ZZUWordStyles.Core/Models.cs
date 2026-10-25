namespace ZZUWordStyles.Core;

/// <summary>界面和自动化测试共用的操作结果；不把异常堆栈作为用户提示。</summary>
public sealed record OperationResult
{
    public bool Success { get; init; }
    public bool Changed { get; init; }
    public string Message { get; init; } = "";
    public string? OutputPath { get; init; }
    public string? ReceiptPath { get; init; }
    public List<string> Warnings { get; init; } = [];
}

/// <summary>只读检查结果；检查不会创建 Word 实例或写入文件。</summary>
public sealed record InspectionResult
{
    public bool IsInstalled { get; init; }
    public List<string> Messages { get; init; } = [];
}

/// <summary>每批独立回执；绝对路径仅保存在当前电脑，不能用于公开分发。</summary>
public sealed record BackupReceipt
{
    public string TargetPath { get; init; } = "";
    public string BackupPath { get; init; } = "";
    public string BeforeSha256 { get; init; } = "";
    public string AfterSha256 { get; init; } = "";
    public DateTimeOffset CreatedAt { get; init; }
    public string Version { get; init; } = "0.1.0-beta.1";
    public string ReceiptPath { get; init; } = "";
    public string Operation { get; init; } = "install";
    public string Status { get; init; } = "Prepared";
    public string? RecoveryBackupPath { get; init; }
}
