using System.Reflection; // 读取随程序构建的原始资源，避免依赖用户的下载目录。
using System.Security.Cryptography; // 比较缓存文件和内置资源是否完全相同。

namespace ZZUWordStyles.App;

internal static class AssetStore
{
    public const string Version = "0.1.0-beta.1"; // 此版本与发布标签对应。
    public static string DataRoot => Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ZZUWordStyles"); // 所有运行记录只在当前用户目录。
    public static string BackupRoot => Path.Combine(DataRoot, "backups"); // 完整备份与公开下载目录分开。
    public static byte[] Read(bool example = false)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(example ? "ZZU.Example.docx" : "ZZU.Template.dotx") ?? throw new InvalidOperationException("程序缺少内置模板，请重新下载完整发布包。"); // 缺失资源时不继续写入。
        using var buffer = new MemoryStream(); // 资源只在内存中解压。
        stream.CopyTo(buffer); // 复制完整资源，不能使用截断文件。
        return buffer.ToArray(); // 调用者得到独立字节数组。
    }
    public static string Materialize(bool example = false)
    {
        var data = Read(example); // 只有用户选择具体操作时才写缓存。
        var directory = Path.Combine(DataRoot, "templates", Version); // 不覆盖其他版本。
        Directory.CreateDirectory(directory); // 创建本程序的用户目录。
        var path = Path.Combine(directory, example ? "郑大论文样式示例.docx" : "郑大论文样式.dotx"); // 文件名便于用户辨认。
        if (File.Exists(path))
        {
            if (!SHA256.HashData(File.ReadAllBytes(path)).SequenceEqual(SHA256.HashData(data))) // 不覆盖用户修改过的缓存模板。
                throw new IOException("缓存模板已被修改，请保留该文件并更名后重试：" + path);
            return path; // 相同内容无需重复写入。
        }
        using var output = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None); // 拒绝覆盖竞争创建的文件。
        output.Write(data); // 写入内置模板的完整内容。
        return path; // 输出仅供当前用户使用。
    }
}
