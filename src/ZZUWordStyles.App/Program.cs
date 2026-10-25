using System.Text.Json; // 维护者的显式验收模式输出结构化结果。
using ZZUWordStyles.Core; // 包内样式的只读验收。
using ZZUWordStyles.Windows; // Word 与字体环境检查。

namespace ZZUWordStyles.App;

internal static class Program
{
    [STAThread] // Word COM 与 Windows 窗口均使用单线程单元。
    private static int Main(string[] args)
    {
        ApplicationConfiguration.Initialize(); // 使用系统 DPI 与可视样式。
        if (args.Length == 2 && (args[0] == "--smoke-json" || args[0] == "--verify-word-json")) // 这两个显式模式只用于发行包验收。
        {
            try
            {
                object result = args[0] == "--smoke-json"
                    ? new { Version = AssetStore.Version, Environment = WordIntegration.InspectEnvironment(), Template = StylePackage.Inspect(StylePackage.Merge(AssetStore.Read(), AssetStore.Read()), AssetStore.Read()), SelfContained = File.Exists(Path.Combine(AppContext.BaseDirectory, "coreclr.dll")) }
                    : WordIntegration.VerifyNewDocuments(); // 后者明确会启动两次自己的 Word 测试实例。
                File.WriteAllText(Path.GetFullPath(args[1]), JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true })); // 不将内容上传网络。
                return result is VerificationReport verification && !verification.Success ? 1 : 0; // 验证未通过不能返回成功退出码。
            }
            catch (Exception exception)
            {
                File.WriteAllText(Path.GetFullPath(args[1]), JsonSerializer.Serialize(new { Success = false, Error = exception.Message })); // 失败也提供可读回执。
                return 1;
            }
        }
        Application.Run(new MainForm()); // 普通双击始终进入中文界面。
        return 0;
    }
}
