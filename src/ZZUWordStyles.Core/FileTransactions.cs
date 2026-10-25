using System.Text;
using System.Text.Json;

namespace ZZUWordStyles.Core;

/// <summary>备份优先的本地事务；调用方须在安装/恢复前确认全部 Word 进程已关闭。</summary>
public static class FileTransactions
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    /// <summary>合并真实目标模板；不存在的 Normal 应先由 Windows 层用 Word 初始化。</summary>
    public static OperationResult Install(string target, string templatePath, string backupRoot, Action? beforeCommitCheck = null)
    {
        try
        {
            target = Path.GetFullPath(target);
            if (!Path.GetExtension(target).Equals(".dotm", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("默认模板安装仅支持 DOTM。旧 DOT 模板需要先由 Word 转换。");
            if (!File.Exists(target)) throw new FileNotFoundException("尚未生成 Normal.dotm，请先通过助手初始化 Word 默认模板。", target);
            var original = ReadStable(target);
            var candidate = StylePackage.Merge(original, ReadStable(templatePath));
            if (candidate.SequenceEqual(original)) return new() { Success = true, Changed = false, Message = "当前样式已匹配，无需重新安装或新增备份。", OutputPath = target };
            return ReplaceWithBackup(target, original, candidate, backupRoot, "install", beforeCommitCheck);
        }
        catch (Exception ex) when (IsExpected(ex)) { return Failed(ex); }
    }

    /// <summary>回执里的备份先验哈希；发现安装后的修改时必须由用户明确允许恢复。</summary>
    public static OperationResult Restore(string receiptPath, bool allowChanged = false, Action? beforeCommitCheck = null, string? expectedCurrentSha256 = null)
    {
        try
        {
            receiptPath = Path.GetFullPath(receiptPath);
            var receipt = JsonSerializer.Deserialize<BackupReceipt>(ReadStable(receiptPath), JsonOptions) ?? throw new InvalidDataException("恢复回执为空。");
            if (receipt.Status != "Completed") throw new InvalidDataException("回执不是已完成状态，请先检查备份文件，不自动恢复未完成事务。");
            ValidateReceiptPaths(receiptPath, receipt);
            var backup = ReadStable(receipt.BackupPath);
            if (StylePackage.Sha256(backup) != receipt.BeforeSha256) throw new InvalidDataException("备份校验值不符，已停止恢复。");
            var current = ReadStable(receipt.TargetPath);
            var currentHash = StylePackage.Sha256(current);
            if (expectedCurrentSha256 is not null && !currentHash.Equals(expectedCurrentSha256, StringComparison.OrdinalIgnoreCase))
                return new() { Success = false, Changed = false, Message = "确认恢复之后，模板又发生变化。已停止，请重新检查当前版本。", OutputPath = receipt.TargetPath, ReceiptPath = receiptPath, Warnings = ["TargetChangedSinceConfirmation"] };
            if (currentHash == receipt.BeforeSha256) return new() { Success = true, Changed = false, Message = "目标已与所选备份相同，无需恢复。", OutputPath = receipt.TargetPath };
            if (currentHash != receipt.AfterSha256 && !allowChanged)
                return new() { Success = false, Changed = false, Message = "安装之后模板又有修改。需要用户明确选择恢复；恢复前会另外保存当前版本。", OutputPath = receipt.TargetPath, ReceiptPath = receiptPath, Warnings = ["TargetChangedSinceInstall"] };
            // 恢复本身也完整备份，因此可以从恢复记录撤回这次恢复操作。
            var backupRoot = Directory.GetParent(Path.GetDirectoryName(receiptPath)!)?.FullName ?? throw new InvalidDataException("备份目录结构不正确。");
            return ReplaceWithBackup(receipt.TargetPath, current, backup, backupRoot, "restore", beforeCommitCheck);
        }
        catch (Exception ex) when (IsExpected(ex)) { return Failed(ex); }
    }

    /// <summary>已有 DOCX 仅生成新副本；原文档字节不被覆盖，文件名冲突使用递增编号。</summary>
    public static OperationResult ImportCopy(string sourceDocx, string templatePath)
    {
        string? createdPath = null;
        try
        {
            sourceDocx = Path.GetFullPath(sourceDocx);
            if (!Path.GetExtension(sourceDocx).Equals(".docx", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("首版只为 DOCX 导入样式副本；DOC、DOCM 和受保护文件请先另行处理。");
            var original = ReadStable(sourceDocx);
            var originalHash = StylePackage.Sha256(original);
            var candidate = StylePackage.Merge(original, ReadStable(templatePath));
            var warnings = StylePackage.AnalyzeDirectFormatting(original);
            var directory = Path.GetDirectoryName(sourceDocx)!;
            var basis = Path.GetFileNameWithoutExtension(sourceDocx) + "_郑大样式";
            // CreateNew 排他创建文件，避免检查文件名与写入之间覆盖别人的同名文件。
            for (var number = 1; number <= 10000; number++)
            {
                var path = Path.Combine(directory, basis + (number == 1 ? "" : "_" + number) + ".docx");
                FileStream output;
                try { output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None); }
                catch (IOException) when (File.Exists(path)) { continue; }
                createdPath = path;
                using (output) { output.Write(candidate); output.Flush(flushToDisk: true); }
                break;
            }
            if (createdPath is null) throw new IOException("同名副本过多，请先整理目标文件夹。");
            if (StylePackage.Sha256(ReadStable(createdPath)) != StylePackage.Sha256(candidate)) throw new IOException("输出副本校验失败。");
            if (StylePackage.Sha256(ReadStable(sourceDocx)) != originalHash)
                throw new IOException("处理期间原文档发生变化。请保存并关闭该文档后重试；本次副本已撤回。");
            return new() { Success = true, Changed = true, Message = "已生成样式副本，原文档内容未改写。直接格式和自定义样式仍保留，请根据报告检查。", OutputPath = createdPath, Warnings = warnings };
        }
        catch (Exception ex) when (IsExpected(ex))
        {
            // 只清理本次排他创建的文件；从不删除用户原件或已有同名副本。
            if (createdPath is not null) TryDelete(createdPath);
            return Failed(ex);
        }
    }

    /// <summary>只展示完成且具有有效目录结构的回执；损坏回执不影响其他可恢复版本。</summary>
    public static List<BackupReceipt> ListBackups(string backupRoot)
    {
        var result = new List<BackupReceipt>();
        if (!Directory.Exists(backupRoot)) return result;
        foreach (var directory in Directory.EnumerateDirectories(backupRoot))
        {
            var file = Path.Combine(directory, "receipt.json");
            if (!File.Exists(file)) continue;
            try
            {
                var receipt = JsonSerializer.Deserialize<BackupReceipt>(ReadStable(file), JsonOptions);
                if (receipt?.Status != "Completed") continue;
                ValidateReceiptPaths(file, receipt);
                result.Add(receipt with { ReceiptPath = file });
            }
            catch (Exception ex) when (IsExpected(ex)) { /* 只忽略当前坏回执；不对任何文件进行修复。 */ }
        }
        return result.OrderByDescending(x => x.CreatedAt).ToList();
    }

    private static OperationResult ReplaceWithBackup(string target, byte[] expected, byte[] candidate, string backupRoot, string operation, Action? beforeCommitCheck)
    {
        var beforeHash = StylePackage.Sha256(expected);
        var afterHash = StylePackage.Sha256(candidate);
        var directory = Path.Combine(Path.GetFullPath(backupRoot), DateTimeOffset.Now.ToString("yyyyMMdd_HHmmss_fff") + "_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var backupPath = Path.Combine(directory, "before.dotm");
        var receiptPath = Path.Combine(directory, "receipt.json");
        var temporary = Path.Combine(Path.GetDirectoryName(target)!, ".zzu-" + Guid.NewGuid().ToString("N") + ".tmp");
        // 原子交换的临时备份必须留在目标同卷，支持 Normal 重定向到另一块磁盘。
        var swapBackup = Path.Combine(Path.GetDirectoryName(target)!, ".zzu-at-swap-" + Guid.NewGuid().ToString("N") + ".tmp");
        var published = false;
        var concurrencyDetected = false;
        var receipt = new BackupReceipt { TargetPath = Path.GetFullPath(target), BackupPath = backupPath, BeforeSha256 = beforeHash, AfterSha256 = afterHash, CreatedAt = DateTimeOffset.Now, ReceiptPath = receiptPath, Operation = operation };
        try
        {
            // 独立操作锁只协调本工具，Word 的文件锁仍由目标文件句柄负责。
            using var transactionLock = new FileStream(target + ".zzu.lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.DeleteOnClose);
            WriteNew(backupPath, expected);
            if (StylePackage.Sha256(ReadStable(backupPath)) != beforeHash) throw new IOException("原模板备份校验失败。");
            WriteNew(receiptPath, JsonSerializer.SerializeToUtf8Bytes(receipt, JsonOptions));
            WriteNew(temporary, candidate);
            if (StylePackage.Sha256(ReadStable(temporary)) != afterHash) throw new IOException("候选模板写入校验失败。");
            beforeCommitCheck?.Invoke(); // Windows 层在此再次确认 Word 未重新打开。

            // 文件句柄拒绝其他读写者，保留 DELETE 共享以允许 Windows 原子替换。
            // File.Replace 另外捕获交换瞬间的实际原件，校验后才能确认发布成功。
            using (var targetLock = new FileStream(target, FileMode.Open, FileAccess.Read, FileShare.Delete))
            {
                using var snapshot = new MemoryStream();
                targetLock.CopyTo(snapshot);
                if (StylePackage.Sha256(snapshot.ToArray()) != beforeHash) throw new IOException("准备期间目标模板已改变，已停止安装。");
                File.Replace(temporary, target, swapBackup, ignoreMetadataErrors: false);
                published = true;
            }
            if (StylePackage.Sha256(ReadStable(swapBackup)) != beforeHash)
            {
                // 发生 rename 竞争时不继续覆盖；持久保存交换时捕获的真实原件供用户检查。
                concurrencyDetected = true;
                receipt = receipt with { RecoveryBackupPath = PreserveRecovery(swapBackup, directory, "exchange-before.dotm") };
                throw new IOException("原子交换时检测到并发替换，已停止自动恢复并保留交换前的实际文件。");
            }
            if (StylePackage.Sha256(ReadStable(target)) != afterHash)
            {
                concurrencyDetected = true;
                throw new IOException("模板写入后内容又发生变化，已停止自动覆盖，完整备份已保留。");
            }
            receipt = receipt with { Status = "Completed" };
            WriteReceipt(receiptPath, receipt);
            TryDelete(swapBackup); // 与 before.dotm 相同的交换备份不重复占用空间。
            return new() { Success = true, Changed = true, Message = operation == "restore" ? "已恢复所选备份；恢复前的当前版本也已另存。" : "已配置默认样式并保存完整备份。请运行 Word 实际验证。", OutputPath = target, ReceiptPath = receiptPath };
        }
        catch (Exception ex) when (IsExpected(ex))
        {
            // 候选已写入但后续验证/回执失败时，自动回滚前先确认仍为本次候选，避免覆盖后续改动。
            var recovery = "原模板未改写。";
            if (published && concurrencyDetected)
                recovery = "检测到并发变动，请检查备份及当前模板：" + (receipt.RecoveryBackupPath ?? backupPath);
            else if (published)
            {
                try
                {
                    if (TryRollback(target, expected, afterHash, directory, out var rollbackMessage))
                    {
                        recovery = "本次候选已自动回滚到安装前版本。";
                        published = false;
                    }
                    else recovery = rollbackMessage;
                }
                catch (Exception rollbackError) when (IsExpected(rollbackError)) { recovery = "自动回滚未完成，完整备份位于：" + backupPath; }
            }
            try { WriteReceipt(receiptPath, receipt with { Status = published ? "NeedsReview" : "Failed" }); } catch (Exception logError) when (IsExpected(logError)) { /* 原件或备份优先，回执失败不抹去最初错误。 */ }
            return new() { Success = false, Changed = published, Message = ex.Message + " " + recovery, OutputPath = target, ReceiptPath = receiptPath };
        }
        finally { TryDelete(temporary); if (!published) TryDelete(swapBackup); }
    }

    private static bool TryRollback(string target, byte[] expected, string candidateHash, string backupDirectory, out string message)
    {
        var folder = Path.GetDirectoryName(target)!;
        var rollback = Path.Combine(folder, ".zzu-rollback-" + Guid.NewGuid().ToString("N") + ".tmp");
        var capture = Path.Combine(folder, ".zzu-rollback-capture-" + Guid.NewGuid().ToString("N") + ".tmp");
        message = "目标出现后续变化，未自动覆盖；请检查完整备份。";
        try
        {
            WriteNew(rollback, expected);
            using (var targetLock = new FileStream(target, FileMode.Open, FileAccess.Read, FileShare.Delete))
            {
                using var snapshot = new MemoryStream();
                targetLock.CopyTo(snapshot);
                if (StylePackage.Sha256(snapshot.ToArray()) != candidateHash) return false;
                File.Replace(rollback, target, capture);
            }
            // 交换再次遇到 rename 竞争时保留双方内容，停止任何进一步的自动替换。
            if (StylePackage.Sha256(ReadStable(capture)) != candidateHash || StylePackage.Sha256(ReadStable(target)) != StylePackage.Sha256(expected))
            {
                var preserved = PreserveRecovery(capture, backupDirectory, "rollback-exchange.dotm");
                message = "回滚时又检测到并发变动，已停止；交换捕获的实际文件保存在：" + preserved;
                return false;
            }
            TryDelete(capture);
            return true;
        }
        finally { TryDelete(rollback); }
    }

    private static string PreserveRecovery(string source, string directory, string fileName)
    {
        var destination = Path.Combine(directory, fileName);
        var bytes = ReadStable(source);
        WriteNew(destination, bytes);
        if (StylePackage.Sha256(ReadStable(destination)) != StylePackage.Sha256(bytes)) throw new IOException("并发恢复备份校验失败，实际捕获文件仍保留在：" + source);
        TryDelete(source);
        return destination;
    }

    private static void ValidateReceiptPaths(string receiptPath, BackupReceipt receipt)
    {
        if (receipt.TargetPath.Length == 0 || receipt.BackupPath.Length == 0 || !Path.IsPathFullyQualified(receipt.TargetPath) || !Path.IsPathFullyQualified(receipt.BackupPath))
            throw new InvalidDataException("回执缺少完整目标或备份路径。");
        var receiptDirectory = Path.GetDirectoryName(Path.GetFullPath(receiptPath))!;
        if (!Path.GetFullPath(receipt.BackupPath).Equals(Path.Combine(receiptDirectory, "before.dotm"), StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("备份路径不在所选回执目录中。");
        if (!Path.GetExtension(receipt.TargetPath).Equals(".dotm", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("回执目标不是 DOTM 模板。");
        if (receipt.BeforeSha256.Length != 64 || receipt.AfterSha256.Length != 64) throw new InvalidDataException("回执校验值不完整。");
    }

    private static byte[] ReadStable(string path)
    {
        using var input = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        if (input.Length > 512L * 1024 * 1024) throw new IOException("首版不处理大于 512 MB 的文件。");
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static void WriteNew(string path, byte[] bytes)
    {
        using var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        output.Write(bytes);
        output.Flush(flushToDisk: true);
    }

    private static void WriteReceipt(string path, BackupReceipt receipt)
    {
        var temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try
        {
            WriteNew(temporary, JsonSerializer.SerializeToUtf8Bytes(receipt, JsonOptions));
            File.Move(temporary, path, overwrite: true);
        }
        finally { TryDelete(temporary); }
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { /* 清理失败仅留下本次临时文件，不触碰其他文件。 */ }
    }

    private static bool IsExpected(Exception ex) => ex is IOException or UnauthorizedAccessException or InvalidDataException or ArgumentException or JsonException or System.Xml.XmlException or NotSupportedException or InvalidOperationException;
    private static OperationResult Failed(Exception ex) => new() { Success = false, Changed = false, Message = ex.Message };
}
