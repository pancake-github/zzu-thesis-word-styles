using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace ZZUWordStyles.Core;

/// <summary>限定修改 word/styles.xml 的样式迁入器，不使用 Office，不运行包内代码。</summary>
public static class StylePackage
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string StylesPart = "word/styles.xml";
    private const string ManagedAlias = "ZZUWordStyles.Managed";
    private static readonly string[] Approved = ["Normal", "Heading 1", "Heading 2", "Heading 3", "Heading 4", "中文摘要标题", "英文摘要标题", "TOC Heading", "图题", "表题", "附属正文", "致谢正文", "TOC 1", "TOC 2", "TOC 3", "Table of Figures", "Header", "Footer", "英文页眉", "局部加粗", "封面信息四号", "封面信息三号", "封面题目二号", "封面校名初号", "封面学位小初", "英文封面三号"];
    private static readonly HashSet<string> Quick = new(Approved.Take(12), StringComparer.OrdinalIgnoreCase);
    private static readonly string[] References = ["basedOn", "next", "link"];
    private static readonly string[] DefaultProperties = ["rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "color"];
    private static readonly Dictionary<string, string> Builtins = CreateBuiltinAliases();

    /// <summary>只在内存中生成候选包；未改变时直接返回原字节，保证重复安装不改写。</summary>
    public static byte[] Merge(byte[] target, byte[] sourceTemplate)
    {
        var destination = ReadPackage(target);
        var source = ReadPackage(sourceTemplate);
        var root = ParseXml(destination[StylesPart]).Root!;
        var before = new XElement(root);
        var sourceRoot = ParseXml(source[StylesPart]).Root!;
        ValidateStyles(root);
        ValidateStyles(sourceRoot);
        var sourceByName = sourceRoot.Elements(W + "style").ToDictionary(Name, StringComparer.OrdinalIgnoreCase);
        foreach (var required in Approved)
            if (!sourceByName.ContainsKey(required)) throw new InvalidDataException($"发布模板缺少样式：{required}");

        var targetStyles = root.Elements(W + "style").ToList();
        var sourceMap = new Dictionary<string, string>(StringComparer.Ordinal);
        var targetBySource = new Dictionary<string, XElement?>(StringComparer.Ordinal);
        var usedIds = targetStyles.Select(Id).ToHashSet(StringComparer.Ordinal);

        // 先映射全部依赖；现有样式 ID 被正文、编号和其它部件引用，不能改名。
        var selected = Approved.Select(name => sourceByName[name]).ToList();
        var defaultCharacter = sourceRoot.Elements(W + "style").SingleOrDefault(x => Canonical(Name(x)) == "Default Paragraph Font");
        if (defaultCharacter is not null && selected.Any(x => References.Any(r => Value(x.Element(W + r)) == Id(defaultCharacter))))
            selected.Add(defaultCharacter);
        foreach (var style in selected)
        {
            var existing = FindTarget(style, targetStyles);
            targetBySource.Add(Id(style), existing);
            if (existing is not null) sourceMap.Add(Id(style), Id(existing));
            else
            {
                var candidate = Id(style);
                // 单纯 ID 冲突可分配项目 ID，不接管占用该 ID 的其他自定义样式。
                if (usedIds.Contains(candidate)) candidate = "ZZU_" + candidate;
                var suffix = 2;
                var basis = candidate;
                while (usedIds.Contains(candidate)) candidate = basis + "_" + suffix++;
                usedIds.Add(candidate);
                sourceMap.Add(Id(style), candidate);
            }
        }

        foreach (var sourceStyle in selected)
        {
            var old = targetBySource[Id(sourceStyle)];
            var isDependency = !Approved.Contains(Name(sourceStyle), StringComparer.OrdinalIgnoreCase);
            if (isDependency && old is not null) continue; // 已有默认字符基类完全保留。
            var replacement = new XElement(sourceStyle);
            replacement.SetAttributeValue(W + "styleId", sourceMap[Id(sourceStyle)]);
            foreach (var local in References)
            {
                var reference = replacement.Element(W + local);
                if (reference is null) continue;
                var sourceId = Value(reference);
                if (!sourceMap.TryGetValue(sourceId, out var mapped))
                    throw new InvalidDataException($"样式 {Name(sourceStyle)} 的依赖无法安全映射：{sourceId}");
                reference.SetAttributeValue(W + "val", mapped);
            }
            if (old is not null)
            {
                if (Attribute(old, "type") != Attribute(replacement, "type"))
                    throw new InvalidDataException($"样式类型冲突：{Name(sourceStyle)}。请先处理同名样式。");
                // 内置名称必须保留原语言名称，才能继续被本地 Word 识别为同一内置样式。
                if (IsBuiltin(sourceStyle)) replacement.Element(W + "name")!.SetAttributeValue(W + "val", Name(old));
                // 旧论文的多级编号常通过标题样式连接；源规范未重建编号，必须保留这一连接。
                var oldNumbering = old.Element(W + "pPr")?.Element(W + "numPr");
                if (oldNumbering is not null && replacement.Element(W + "pPr")?.Element(W + "numPr") is null)
                {
                    var paragraph = replacement.Element(W + "pPr");
                    if (paragraph is null)
                    {
                        paragraph = new XElement(W + "pPr");
                        var runProperties = replacement.Element(W + "rPr");
                        if (runProperties is null) replacement.Add(paragraph); else runProperties.AddBeforeSelf(paragraph);
                    }
                    var preceding = paragraph.Elements().LastOrDefault(x => x.Name.LocalName is "pStyle" or "keepNext" or "keepLines" or "pageBreakBefore" or "framePr" or "widowControl");
                    if (preceding is null) paragraph.AddFirst(new XElement(oldNumbering)); else preceding.AddAfterSelf(new XElement(oldNumbering));
                }
                if (!IsBuiltin(sourceStyle) && !IsManaged(old) && !EquivalentFormat(old, replacement))
                    throw new InvalidDataException($"已有自定义样式“{Name(old)}”与项目样式同名但格式不同。已停止，原文件未更改。");
                var defaultFlag = old.Attribute(W + "default");
                if (defaultFlag is not null) replacement.SetAttributeValue(W + "default", defaultFlag.Value);
            }
            if (!IsBuiltin(sourceStyle) && !isDependency) MarkManaged(replacement);
            if (old is null) root.Add(replacement); else old.ReplaceWith(replacement);
        }

        // 仅覆盖已经授权的默认字符属性；原语言、字距、连字和默认段落属性保留。
        var sourceDefaults = sourceRoot.Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr");
        if (sourceDefaults is null) throw new InvalidDataException("发布模板缺少默认字符定义。");
        var defaults = Ensure(root, "docDefaults", first: true);
        var defaultRun = Ensure(Ensure(defaults, "rPrDefault", first: true), "rPr");
        foreach (var property in DefaultProperties)
        {
            defaultRun.Elements(W + property).Remove();
            var value = sourceDefaults.Element(W + property);
            if (value is not null) defaultRun.Add(new XElement(value));
        }
        CurateGallery(root);
        ValidateStyles(root);

        // XML 语义一致时不重新压缩，避免产生无意义的备份与修改时间变化。
        if (XNode.DeepEquals(before, root)) return target;
        destination[StylesPart] = Serialize(root);
        return Rewrite(target, destination[StylesPart]);
    }

    /// <summary>采用同一个合并算法比较差异，检查过程无文件和 Office 副作用。</summary>
    public static InspectionResult Inspect(byte[] target, byte[] sourceTemplate)
    {
        try
        {
            var merged = Merge(target, sourceTemplate);
            var installed = ReadPackage(target)[StylesPart].SequenceEqual(ReadPackage(merged)[StylesPart]);
            return new() { IsInstalled = installed, Messages = [installed ? "26 项样式和 12 项常用入口已匹配当前发布版本。" : "模板尚未完整匹配当前版本，可以配置或更新。"] };
        }
        catch (Exception ex) when (ex is InvalidDataException or XmlException or ArgumentException)
        {
            return new() { IsInstalled = false, Messages = ["无法安全检查：" + ex.Message] };
        }
    }

    /// <summary>报告可能覆盖样式的直接格式位置，只给部件和段落序号，不泄露论文文字。</summary>
    public static List<string> AnalyzeDirectFormatting(byte[] docx)
    {
        var parts = ReadPackage(docx);
        var report = new List<string>();
        foreach (var part in parts.Where(p => p.Key.StartsWith("word/", StringComparison.Ordinal) && (p.Key == "word/document.xml" || p.Key.StartsWith("word/header", StringComparison.Ordinal) || p.Key.StartsWith("word/footer", StringComparison.Ordinal) || p.Key is "word/footnotes.xml" or "word/endnotes.xml")))
        {
            var paragraphs = ParseXml(part.Value).Descendants(W + "p").ToList();
            var locations = new List<int>();
            for (var i = 0; i < paragraphs.Count; i++)
            {
                var p = paragraphs[i];
                var paragraphOverride = p.Element(W + "pPr")?.Elements().Any(x => x.Name.LocalName is "spacing" or "ind" or "jc" or "outlineLvl") == true;
                var runOverride = p.Descendants(W + "rPr").SelectMany(x => x.Elements()).Any(x => x.Name.LocalName is "rFonts" or "sz" or "szCs" or "b" or "bCs" or "i" or "iCs" or "color");
                if (paragraphOverride || runOverride) locations.Add(i + 1);
            }
            if (locations.Count > 0) report.Add($"{part.Key}：{locations.Count} 个段落存在直接格式（段落序号：{string.Join("、", locations.Take(30))}{(locations.Count > 30 ? "等" : "")}）。这些格式已保留，请在 Word 中按需检查。");
        }
        if (report.Count == 0) report.Add("未发现所检查的字体、字号、加粗、斜体、颜色、间距、缩进或对齐直接格式；这不代表完整论文排版验收。");
        return report;
    }

    /// <summary>标准 SHA-256 供事务回执和验收使用。</summary>
    public static string Sha256(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();

    private static Dictionary<string, byte[]> ReadPackage(byte[] data)
    {
        if (data.Length > 512L * 1024 * 1024) throw new InvalidDataException("文件超过首版支持的 512 MB 大小。");
        try
        {
            using var stream = new MemoryStream(data, writable: false);
            using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
            if (zip.Entries.Count > 10000) throw new InvalidDataException("压缩包部件数量异常。");
            var parts = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            long total = 0;
            foreach (var entry in zip.Entries)
            {
                if (entry.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("此文档具有包级数字签名，不能自动修改样式。");
                total += entry.Length;
                if (total > 1024L * 1024 * 1024 || entry.Length > 256L * 1024 * 1024)
                    throw new InvalidDataException("压缩包展开大小超过安全限额。");
                using var input = entry.Open();
                using var output = new MemoryStream();
                input.CopyTo(output);
                if (!parts.TryAdd(entry.FullName, output.ToArray())) throw new InvalidDataException("压缩包包含重复部件。");
            }
            if (!parts.ContainsKey(StylesPart) || !parts.ContainsKey("[Content_Types].xml") || !parts.ContainsKey("word/document.xml"))
                throw new InvalidDataException("文件不是支持的 DOCX/DOTX/DOTM 包，或缺少样式部件。加密文件、旧 DOC/DOT 不受支持。");
            if (ParseXml(parts[StylesPart]).Root?.Name != W + "styles") throw new InvalidDataException("文件使用不支持的样式 XML 名称空间。");
            return parts;
        }
        catch (InvalidDataException) { throw; }
        catch (Exception ex) when (ex is IOException or XmlException or NotSupportedException)
        {
            throw new InvalidDataException("无法读取 Office 文件，文件可能损坏、被加密或格式不受支持。", ex);
        }
    }

    private static byte[] Rewrite(byte[] original, byte[] styles)
    {
        using var output = new MemoryStream();
        using (var inputStream = new MemoryStream(original, writable: false))
        using (var input = new ZipArchive(inputStream, ZipArchiveMode.Read))
        using (var result = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                var copy = result.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                copy.LastWriteTime = entry.LastWriteTime;
                copy.ExternalAttributes = entry.ExternalAttributes;
                using var write = copy.Open();
                if (entry.FullName == StylesPart) write.Write(styles);
                else { using var read = entry.Open(); read.CopyTo(write); }
            }
        }
        var candidate = output.ToArray();
        var oldParts = ReadPackage(original);
        var newParts = ReadPackage(candidate);
        if (!oldParts.Keys.ToHashSet(StringComparer.Ordinal).SetEquals(newParts.Keys) || oldParts.Any(p => p.Key != StylesPart && !p.Value.SequenceEqual(newParts[p.Key])))
            throw new InvalidDataException("候选校验失败：样式之外的文件内容发生变化。");
        return candidate;
    }

    private static XDocument ParseXml(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = 256L * 1024 * 1024 });
        return XDocument.Load(reader); // 去掉排版空白后比较 XML 语义，不依赖原保存器的缩进。
    }

    private static byte[] Serialize(XElement root)
    {
        using var output = new MemoryStream();
        using (var writer = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false }))
            new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root).Save(writer);
        return output.ToArray();
    }

    private static void ValidateStyles(XElement root)
    {
        var styles = root.Elements(W + "style").ToList();
        if (styles.Any(x => Id(x).Length == 0 || Name(x).Length == 0)) throw new InvalidDataException("存在未命名或无 ID 的样式。");
        if (styles.Select(Id).Distinct(StringComparer.Ordinal).Count() != styles.Count) throw new InvalidDataException("样式 ID 重复。");
        if (styles.Count(x => Attribute(x, "type") == "paragraph" && IsEnabledAttribute(x, "default")) > 1) throw new InvalidDataException("存在多个默认段落样式。");
    }

    private static XElement? FindTarget(XElement source, List<XElement> styles)
    {
        var name = Canonical(Name(source));
        var candidates = styles.Where(x => Canonical(Name(x)).Equals(name, StringComparison.OrdinalIgnoreCase)).ToList();
        if (name == "Normal") candidates.AddRange(styles.Where(x => Attribute(x, "type") == "paragraph" && IsEnabledAttribute(x, "default")));
        if (name == "Default Paragraph Font") candidates.AddRange(styles.Where(x => Attribute(x, "type") == "character" && IsEnabledAttribute(x, "default")));
        candidates = candidates.Distinct().ToList();
        if (candidates.Count > 1) throw new InvalidDataException($"无法唯一匹配样式“{Name(source)}”，请先处理同名或默认样式冲突。");
        if (candidates.Count == 1 && IsBuiltin(source) && IsEnabledAttribute(candidates[0], "customStyle"))
            throw new InvalidDataException($"自定义样式“{Name(candidates[0])}”占用了内置样式名称，不能自动替换。");
        return candidates.SingleOrDefault();
    }

    private static void CurateGallery(XElement root)
    {
        var latent = root.Element(W + "latentStyles");
        if (latent is null)
        {
            latent = new XElement(W + "latentStyles");
            var firstStyle = root.Element(W + "style");
            if (firstStyle is null) root.Add(latent); else firstStyle.AddBeforeSelf(latent);
        }
        latent.SetAttributeValue(W + "defSemiHidden", "1");
        latent.SetAttributeValue(W + "defQFormat", "0");
        latent.SetAttributeValue(W + "defUnhideWhenUsed", "0");
        foreach (var exception in latent.Elements(W + "lsdException"))
        {
            var name = Canonical(Attribute(exception, "name"));
            var approved = Approved.Contains(name, StringComparer.OrdinalIgnoreCase);
            exception.SetAttributeValue(W + "semiHidden", approved ? "0" : "1");
            exception.SetAttributeValue(W + "qFormat", Quick.Contains(name) ? "1" : "0");
            exception.SetAttributeValue(W + "unhideWhenUsed", approved ? "1" : "0");
        }
        foreach (var style in root.Elements(W + "style"))
        {
            var name = Canonical(Name(style));
            if (!Quick.Contains(name)) style.Elements(W + "qFormat").Remove();
            if (!Approved.Contains(name, StringComparer.OrdinalIgnoreCase)) style.Elements(W + "unhideWhenUsed").Remove();
        }
    }

    private static bool EquivalentFormat(XElement left, XElement right)
    {
        var a = new XElement(left);
        var b = new XElement(right);
        foreach (var item in new[] { a, b })
        {
            item.Attribute(W + "styleId")?.Remove();
            foreach (var local in new[] { "aliases", "uiPriority", "qFormat", "semiHidden", "unhideWhenUsed", "rsid" }) item.Elements(W + local).Remove();
        }
        return XNode.DeepEquals(a, b);
    }

    private static void MarkManaged(XElement style)
    {
        var aliases = style.Element(W + "aliases");
        if (aliases is null) { aliases = new XElement(W + "aliases"); style.Element(W + "name")!.AddAfterSelf(aliases); }
        var values = Value(aliases).Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).ToList();
        // Word 别名与样式名称共用命名空间，每项必须有不同的管理标记。
        var marker = ManagedAlias + "." + Id(style);
        if (!values.Contains(marker, StringComparer.Ordinal)) values.Add(marker);
        aliases.SetAttributeValue(W + "val", string.Join(',', values));
    }

    private static bool IsManaged(XElement style) => Value(style.Element(W + "aliases")).Split(',', StringSplitOptions.TrimEntries).Contains(ManagedAlias + "." + Id(style), StringComparer.Ordinal);
    private static bool IsBuiltin(XElement style) => Builtins.ContainsKey(Name(style));
    private static bool IsEnabledAttribute(XElement node, string name) => Attribute(node, name) is "1" or "true" or "on";
    private static string Name(XElement node) => Value(node.Element(W + "name"));
    private static string Id(XElement node) => Attribute(node, "styleId");
    private static string Value(XElement? node) => node?.Attribute(W + "val")?.Value ?? "";
    private static string Attribute(XElement node, string name) => node.Attribute(W + name)?.Value ?? "";
    private static string Canonical(string value) => Builtins.GetValueOrDefault(value, value);
    private static XElement Ensure(XElement parent, string local, bool first = false)
    {
        var node = parent.Element(W + local);
        if (node is not null) return node;
        node = new XElement(W + local);
        if (first) parent.AddFirst(node); else parent.Add(node);
        return node;
    }

    private static Dictionary<string, string> CreateBuiltinAliases()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        void Add(string canonical, params string[] aliases) { map[canonical] = canonical; foreach (var alias in aliases) map[alias] = canonical; }
        Add("Normal", "正文");
        Add("Default Paragraph Font", "DefaultParagraphFont", "默认段落字体");
        for (var i = 1; i <= 9; i++) Add("Heading " + i, "标题 " + i, "标题" + i);
        for (var i = 1; i <= 9; i++) Add("TOC " + i, "目录 " + i, "目录" + i);
        Add("TOC Heading", "目录标题");
        Add("Table of Figures", "图表目录");
        Add("Header", "页眉");
        Add("Footer", "页脚");
        return map;
    }
}
