using System.IO.Compression;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using ZZUWordStyles.Core;

// 无额外测试框架依赖；失败时返回非零退出码，CI 与本地运行使用相同测试。
var source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "assets", "郑大论文样式.dotx"));
var results = new List<object>();
var failures = 0;
var work = Path.Combine(Path.GetTempPath(), "ZZUWordStyles-tests-" + Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(work);
try
{
    Test("localized_ids_and_non_style_parts_preserved", () =>
    {
        var original = Fixture();
        var merged = StylePackage.Merge(original, source);
        var parts = Parts(merged);
        var beforeParts = Parts(original);
        Check(beforeParts.Keys.ToHashSet().SetEquals(parts.Keys), "部件数量变化");
        foreach (var part in beforeParts.Where(x => x.Key != "word/styles.xml")) Check(part.Value.SequenceEqual(parts[part.Key]), "非样式部件变化：" + part.Key);
        var xml = XDocument.Parse(Encoding.UTF8.GetString(parts["word/styles.xml"]));
        Check(Style(xml, "a").Element(W + "name")?.Attribute(W + "val")?.Value == "正文", "中文正文身份丢失");
        Check(Style(xml, "h1local").Element(W + "name")?.Attribute(W + "val")?.Value == "标题 1", "中文内置标题身份丢失");
        Check(Style(xml, "a").Descendants(W + "rFonts").First().Attribute(W + "eastAsia")?.Value == "宋体", "正文中文字体错误");
        Check(Style(xml, "a").Descendants(W + "rFonts").First().Attribute(W + "ascii")?.Value == "Times New Roman", "正文西文字体错误");
        Check(Style(xml, "personal").Element(W + "pPr")?.Element(W + "spacing")?.Attribute(W + "after")?.Value == "55", "原有自定义格式丢失");
        Check(Style(xml, "personal").Element(W + "qFormat") is null, "个人常用入口未精简");
        Check(xml.Descendants(W + "style").Count(x => x.Element(W + "qFormat") is not null) == 12, "常用入口不是12个");
        Check(xml.Descendants(W + "lsdException").Count() == 2, "原潜在样式例外丢失");
        Check(xml.Descendants(W + "lang").Any(x => x.Attribute(W + "eastAsia")?.Value == "zh-CN"), "默认语言丢失");
        Check(xml.Descendants(W + "kern").Any(x => x.Attribute(W + "val")?.Value == "2"), "默认字距丢失");
        var aliases = xml.Descendants(W + "aliases").Select(x => x.Attribute(W + "val")?.Value ?? "").ToList();
        Check(aliases.All(x => !x.Contains("ZZUWordStyles.Managed")), "样式库被内部管理别名污染");
        Check(Style(xml, "h1local").Element(W + "aliases")?.Attribute(W + "val")?.Value == "原有章标题", "用户内置样式别名未保留");
        Check(Style(xml, "h1local").Element(W + "pPr")?.Element(W + "numPr")?.Element(W + "numId")?.Attribute(W + "val")?.Value == "7", "标题关联多级编号丢失");
    });
    Test("idempotent_and_read_only_inspection", () =>
    {
        var first = StylePackage.Merge(Fixture(), source);
        var second = StylePackage.Merge(first, source);
        Check(first.SequenceEqual(second), "重复合并改写文件");
        Check(StylePackage.Inspect(first, source).IsInstalled, "检查未识别已安装状态");
        Check(!StylePackage.Inspect(Fixture(), source).IsInstalled, "检查把原始模板误判为已安装");
    });
    Test("missing_defaults_and_id_collision", () =>
    {
        var original = Fixture(missingDefaults: true, idCollision: true);
        var merged = StylePackage.Merge(original, source);
        var xml = XDocument.Parse(Encoding.UTF8.GetString(Parts(merged)["word/styles.xml"]));
        Check(xml.Descendants(W + "docDefaults").Count() == 1, "默认属性未补齐");
        Check(Style(xml, "Heading2").Element(W + "name")?.Attribute(W + "val")?.Value == "自定义旧样式", "ID 冲突覆盖了其他样式");
        Check(Style(xml, "ZZU_Heading2").Element(W + "name")?.Attribute(W + "val")?.Value == "heading 2", "新增标题未安全分配ID");
    });
    Test("preserve_user_aliases_remove_only_internal_candidate_aliases", () =>
    {
        var first = Parts(StylePackage.Merge(Fixture(), source));
        var xml = XDocument.Parse(Encoding.UTF8.GetString(first["word/styles.xml"]));
        Style(xml, "图题").Element(W + "name")!.AddAfterSelf(new XElement(W + "aliases", new XAttribute(W + "val", "用户图题别名,ZZUWordStyles.Managed.图题")));
        first["word/styles.xml"] = Encoding.UTF8.GetBytes(xml.ToString());
        var merged = Parts(StylePackage.Merge(Pack(first), source));
        var checkedXml = XDocument.Parse(Encoding.UTF8.GetString(merged["word/styles.xml"]));
        Check(Style(checkedXml, "图题").Element(W + "aliases")?.Attribute(W + "val")?.Value == "用户图题别名", "自定义用户别名未保留或内部标记未移除");
        Style(xml, "图题").Element(W + "rPr")!.Element(W + "sz")!.SetAttributeValue(W + "val", "44");
        first["word/styles.xml"] = Encoding.UTF8.GetBytes(xml.ToString());
        Reject(() => StylePackage.Merge(Pack(first), source));
    });
    Test("reject_ambiguous_custom_name_and_duplicate_ids", () =>
    {
        Reject(() => StylePackage.Merge(Fixture(extraStyle: "<w:style w:type='paragraph' w:customStyle='1' w:styleId='oldFigure'><w:name w:val='图题'/><w:rPr><w:sz w:val='44'/></w:rPr></w:style>"), source));
        Reject(() => StylePackage.Merge(Fixture(extraStyle: "<w:style w:type='paragraph' w:styleId='a'><w:name w:val='another'/></w:style>"), source));
        Reject(() => StylePackage.Merge(Fixture(extraStyle: "<w:style w:type='paragraph' w:styleId='otherNormal'><w:name w:val='Normal'/></w:style>"), source));
    });
    Test("reject_signed_corrupt_and_external_entities", () =>
    {
        var signed = Parts(Fixture()); signed["_xmlsignatures/sig1.xml"] = "signature"u8.ToArray();
        Reject(() => StylePackage.Merge(Pack(signed), source));
        Reject(() => StylePackage.Merge("not-a-zip"u8.ToArray(), source));
        var entity = Parts(Fixture()); entity["word/styles.xml"] = Encoding.UTF8.GetBytes("<!DOCTYPE x [<!ENTITY ext SYSTEM 'file:///never-read'>]><w:styles xmlns:w='" + W + "'>&ext;</w:styles>");
        Reject(() => StylePackage.Merge(Pack(entity), source));
    });
    Test("direct_format_report_has_locations_not_content", () =>
    {
        var warnings = StylePackage.AnalyzeDirectFormatting(Fixture());
        Check(warnings.Any(x => x.Contains("word/document.xml") && x.Contains("1 个段落")), "未发现直接格式");
        Check(warnings.All(x => !x.Contains("PRIVATE_RESEARCH_TEXT")), "报告包含论文原文");
    });
    var template = Path.Combine(work, "source.dotx");
    File.WriteAllBytes(template, source);
    Test("transaction_install_repeat_restore_and_changed_guard", () =>
    {
        var target = Path.Combine(work, "中文 空格", "normal.dotm"); Directory.CreateDirectory(Path.GetDirectoryName(target)!);
        var backupRoot = Path.Combine(work, "backups");
        var original = Fixture(); File.WriteAllBytes(target, original);
        var installed = FileTransactions.Install(target, template, backupRoot);
        Check(installed.Success && installed.Changed && installed.ReceiptPath is not null, "安装失败：" + installed.Message);
        var repeated = FileTransactions.Install(target, template, backupRoot);
        Check(repeated.Success && !repeated.Changed, "重复安装未幂等：" + repeated.Message);
        Check(FileTransactions.ListBackups(backupRoot).Count == 1, "重复安装产生新备份");
        var changedParts = Parts(File.ReadAllBytes(target)); changedParts["word/customizations.xml"] = "later-user-change"u8.ToArray();
        var changed = Pack(changedParts); File.WriteAllBytes(target, changed);
        var denied = FileTransactions.Restore(installed.ReceiptPath!);
        Check(!denied.Success && denied.Warnings.Contains("TargetChangedSinceInstall"), "没有拦截后续用户修改");
        Check(File.ReadAllBytes(target).SequenceEqual(changed), "拒绝恢复时改写了目标");
        var restored = FileTransactions.Restore(installed.ReceiptPath!, allowChanged: true);
        Check(restored.Success && File.ReadAllBytes(target).SequenceEqual(original), "明确恢复失败：" + restored.Message);
        var receipts = FileTransactions.ListBackups(backupRoot);
        Check(receipts.Count == 2 && receipts.Any(x => x.Operation == "restore" && File.ReadAllBytes(x.BackupPath).SequenceEqual(changed)), "恢复前未备份当前版本");
    });
    Test("locked_target_and_tampered_backup_do_not_overwrite", () =>
    {
        var target = Path.Combine(work, "locked.dotm"); var original = Fixture(); File.WriteAllBytes(target, original);
        using (var locked = new FileStream(target, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            var result = FileTransactions.Install(target, template, Path.Combine(work, "locked-backups"));
            Check(!result.Success && !result.Changed, "锁定模板仍被写入");
        }
        Check(File.ReadAllBytes(target).SequenceEqual(original), "失败安装改写目标");
        var installed = FileTransactions.Install(target, template, Path.Combine(work, "tamper-backups"));
        Check(installed.Success, installed.Message);
        var receipt = FileTransactions.ListBackups(Path.Combine(work, "tamper-backups")).Single();
        File.WriteAllBytes(receipt.BackupPath, "tamper"u8.ToArray());
        var before = File.ReadAllBytes(target);
        var restore = FileTransactions.Restore(receipt.ReceiptPath);
        Check(!restore.Success && File.ReadAllBytes(target).SequenceEqual(before), "损坏备份未阻止恢复");
    });
    Test("docx_copy_preserves_original_and_body_parts", () =>
    {
        var originalPath = Path.Combine(work, "已有论文.docx"); var original = Fixture(); File.WriteAllBytes(originalPath, original);
        var result = FileTransactions.ImportCopy(originalPath, template);
        Check(result.Success && result.OutputPath is not null, "副本导入失败：" + result.Message);
        Check(File.ReadAllBytes(originalPath).SequenceEqual(original), "改写原始 DOCX");
        var parts = Parts(File.ReadAllBytes(result.OutputPath!));
        foreach (var pair in Parts(original).Where(x => x.Key != "word/styles.xml")) Check(parts[pair.Key].SequenceEqual(pair.Value), "导入改写其他部件：" + pair.Key);
        var second = FileTransactions.ImportCopy(originalPath, template);
        Check(second.Success && second.OutputPath != result.OutputPath, "覆盖同名副本");
    });
    Test("missing_target_is_explicit_initialization_branch", () =>
    {
        var path = Path.Combine(work, "missing.dotm");
        var result = FileTransactions.Install(path, template, Path.Combine(work, "missing-backups"));
        Check(!result.Success && !File.Exists(path), "核心偷偷创建了默认模板");
    });
    Test("concurrent_change_before_commit_is_preserved", () =>
    {
        var target = Path.Combine(work, "concurrent.dotm"); var original = Fixture(); File.WriteAllBytes(target, original);
        var changedParts = Parts(original); changedParts["word/customizations.xml"] = "newer-changes"u8.ToArray(); var changed = Pack(changedParts);
        var result = FileTransactions.Install(target, template, Path.Combine(work, "concurrent-backups"), () => File.WriteAllBytes(target, changed));
        Check(!result.Success && !result.Changed, "没有阻止准备期间的并发变化");
        Check(File.ReadAllBytes(target).SequenceEqual(changed), "覆盖了用户并发修改");
    });
    Test("guard_failure_preserves_original", () =>
    {
        var target = Path.Combine(work, "guard.dotm"); var original = Fixture(); File.WriteAllBytes(target, original);
        var result = FileTransactions.Install(target, template, Path.Combine(work, "guard-backups"), () => throw new InvalidOperationException("模拟 Word 在准备期间重新打开"));
        Check(!result.Success && !result.Changed && File.ReadAllBytes(target).SequenceEqual(original), "提交前停止未保护原件");
    });
    Test("restore_confirmation_is_bound_to_exact_current_hash", () =>
    {
        var target = Path.Combine(work, "confirmed.dotm"); File.WriteAllBytes(target, Fixture());
        var installed = FileTransactions.Install(target, template, Path.Combine(work, "confirmed-backups")); Check(installed.Success, installed.Message);
        var confirmedHash = StylePackage.Sha256(File.ReadAllBytes(target));
        var changed = Parts(File.ReadAllBytes(target)); changed["word/customizations.xml"] = "after-confirmation"u8.ToArray(); var changedBytes = Pack(changed); File.WriteAllBytes(target, changedBytes);
        var result = FileTransactions.Restore(installed.ReceiptPath!, true, expectedCurrentSha256: confirmedHash);
        Check(!result.Success && result.Warnings.Contains("TargetChangedSinceConfirmation"), "恢复确认未绑定具体文件版本");
        Check(File.ReadAllBytes(target).SequenceEqual(changedBytes), "恢复覆盖确认后的更新");
    });
}
finally
{
    // 测试目录由本进程以随机名字创建，只清理该绝对目录；不触碰真实 Word 模板。
    var resolved = Path.GetFullPath(work);
    if (resolved.StartsWith(Path.GetFullPath(Path.GetTempPath()), StringComparison.OrdinalIgnoreCase) && Path.GetFileName(resolved).StartsWith("ZZUWordStyles-tests-", StringComparison.Ordinal))
        Directory.Delete(resolved, recursive: true);
}
Console.WriteLine(JsonSerializer.Serialize(new { Passed = results.Count - failures, Failed = failures, Tests = results }, new JsonSerializerOptions { WriteIndented = true }));
return failures == 0 ? 0 : 1;

void Test(string name, Action body)
{
    try { body(); results.Add(new { Name = name, Result = "PASS" }); }
    catch (Exception ex) { failures++; results.Add(new { Name = name, Result = "FAIL", Error = ex.ToString() }); }
}
void Check(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
void Reject(Action action)
{
    try { action(); }
    catch (Exception ex) when (ex is InvalidDataException or System.Xml.XmlException or ArgumentException) { return; }
    throw new InvalidOperationException("应拒绝不安全输入，但操作成功。");
}
XElement Style(XDocument document, string id) => document.Descendants(W + "style").Single(x => x.Attribute(W + "styleId")?.Value == id);

byte[] Fixture(bool missingDefaults = false, bool idCollision = false, string extraStyle = "")
{
    var defaults = missingDefaults ? "" : "<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:eastAsia='等线'/><w:kern w:val='2'/><w:lang w:eastAsia='zh-CN'/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after='160'/></w:pPr></w:pPrDefault></w:docDefaults>";
    var collision = idCollision ? "<w:style w:type='paragraph' w:customStyle='1' w:styleId='Heading2'><w:name w:val='自定义旧样式'/><w:basedOn w:val='a'/></w:style>" : "";
    var styles = "<w:styles xmlns:w='" + W + "'>" + defaults + "<w:latentStyles w:count='376'><w:lsdException w:name='heading 1' w:qFormat='1'/><w:lsdException w:name='Title' w:qFormat='1'/></w:latentStyles><w:style w:type='paragraph' w:default='1' w:styleId='a'><w:name w:val='正文'/><w:rPr><w:rFonts w:eastAsia='等线'/></w:rPr></w:style><w:style w:type='character' w:default='1' w:styleId='a0'><w:name w:val='默认段落字体'/></w:style><w:style w:type='paragraph' w:styleId='h1local'><w:name w:val='标题 1'/><w:aliases w:val='原有章标题'/><w:basedOn w:val='a'/><w:pPr><w:numPr><w:ilvl w:val='0'/><w:numId w:val='7'/></w:numPr></w:pPr></w:style><w:style w:type='paragraph' w:customStyle='1' w:styleId='personal'><w:name w:val='个人样式'/><w:basedOn w:val='a'/><w:qFormat/><w:pPr><w:spacing w:after='55'/></w:pPr></w:style>" + collision + extraStyle + "</w:styles>";
    return Pack(new Dictionary<string, byte[]>
    {
        ["[Content_Types].xml"] = "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'/>"u8.ToArray(),
        ["word/styles.xml"] = Encoding.UTF8.GetBytes(styles),
        ["word/document.xml"] = Encoding.UTF8.GetBytes("<w:document xmlns:w='" + W + "'><w:body><w:p><w:pPr><w:pStyle w:val='h1local'/><w:ind w:firstLine='240'/></w:pPr><w:r><w:rPr><w:rFonts w:ascii='Arial'/></w:rPr><w:t>PRIVATE_RESEARCH_TEXT</w:t></w:r></w:p><w:sectPr><w:pgSz w:w='11906' w:h='16838'/></w:sectPr></w:body></w:document>"),
        ["word/vbaProject.bin"] = [0, 1, 2, 3, 254, 255],
        ["word/vbaProjectSignature.bin"] = [9, 8, 7],
        ["word/customizations.xml"] = "private-keyboard-mappings"u8.ToArray(),
        ["word/glossary/document.xml"] = "private-building-blocks"u8.ToArray(),
        ["word/media/image1.png"] = [137, 80, 78, 71],
        ["word/embeddings/oleObject1.bin"] = [1, 3, 5, 7],
        ["word/numbering.xml"] = Encoding.UTF8.GetBytes("<w:numbering xmlns:w='" + W + "'><w:num w:numId='7'><w:abstractNumId w:val='0'/></w:num></w:numbering>"),
    });
}
byte[] Pack(Dictionary<string, byte[]> parts)
{
    using var stream = new MemoryStream();
    using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        foreach (var part in parts) { using var write = zip.CreateEntry(part.Key).Open(); write.Write(part.Value); }
    return stream.ToArray();
}
Dictionary<string, byte[]> Parts(byte[] bytes)
{
    using var stream = new MemoryStream(bytes);
    using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
    return zip.Entries.ToDictionary(e => e.FullName, e => { using var read = e.Open(); using var output = new MemoryStream(); read.CopyTo(output); return output.ToArray(); });
}

// 把命名空间放在同一个辅助静态类型中，便于顶层语句内共享。
partial class Program
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
}
