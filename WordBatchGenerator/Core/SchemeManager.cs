using System.IO;
using System.Text.Json;

namespace WordBatchGenerator.Core;

/// <summary>
/// 方案管理器：保存、加载、导入、导出方案
/// </summary>
public class SchemeManager
{
    private static readonly string SchemesBaseDir = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "WordBatchGenerator",
        "Schemes"
    );

    static SchemeManager()
    {
        if (!Directory.Exists(SchemesBaseDir))
        {
            Directory.CreateDirectory(SchemesBaseDir);
        }
    }

    /// <summary>
    /// 获取方案根目录
    /// </summary>
    public static string GetSchemesDirectory() => SchemesBaseDir;

    /// <summary>
    /// 保存方案
    /// </summary>
    public static void SaveScheme(Scheme scheme)
    {
        if (string.IsNullOrWhiteSpace(scheme.Name))
        {
            throw new ArgumentException("方案名称不能为空");
        }

        var schemeDir = Path.Combine(SchemesBaseDir, scheme.Name);
        if (Path.GetFullPath(schemeDir).TrimEnd('\\') == Path.GetFullPath(SchemesBaseDir).TrimEnd('\\'))
        {
            throw new InvalidOperationException("无法保存到方案根目录");
        }

        if (!Directory.Exists(schemeDir))
        {
            Directory.CreateDirectory(schemeDir);
        }

        // 复制模板文件并更新 TemplatePath
        if (!string.IsNullOrEmpty(scheme.TemplatePath) && File.Exists(scheme.TemplatePath))
        {
            var templateDestPath = Path.Combine(schemeDir, "template.docx");
            if (Path.GetFullPath(scheme.TemplatePath).ToLower() != Path.GetFullPath(templateDestPath).ToLower())
            {
                File.Copy(scheme.TemplatePath, templateDestPath, true);
            }
            scheme.TemplatePath = templateDestPath;
        }

        // 保存配置文件 (此时 TemplatePath 已经更新为正确的本地路径)
        var configPath = Path.Combine(schemeDir, "config.json");
        var json = JsonSerializer.Serialize(scheme, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(configPath, json);
    }

    /// <summary>
    /// 加载方案
    /// </summary>
    public static Scheme? LoadScheme(string schemeName)
    {
        if (string.IsNullOrWhiteSpace(schemeName))
            return null;

        var schemeDir = Path.Combine(SchemesBaseDir, schemeName);
        var configPath = Path.Combine(schemeDir, "config.json");

        if (!File.Exists(configPath))
            return null;

        var json = File.ReadAllText(configPath);
        var scheme = JsonSerializer.Deserialize<Scheme>(json);
        if (scheme != null)
        {
            scheme.Name = schemeName;
        }
        return scheme;
    }

    /// <summary>
    /// 获取所有方案列表
    /// </summary>
    public static List<Scheme> GetAllSchemes()
    {
        var schemes = new List<Scheme>();

        if (!Directory.Exists(SchemesBaseDir))
            return schemes;

        foreach (var dir in Directory.GetDirectories(SchemesBaseDir))
        {
            var schemeName = Path.GetFileName(dir);
            var scheme = LoadScheme(schemeName);
            if (scheme != null)
            {
                schemes.Add(scheme);
            }
        }

        return schemes;
    }

    /// <summary>
    /// 删除方案
    /// </summary>
    public static void DeleteScheme(string schemeName)
    {
        if (string.IsNullOrWhiteSpace(schemeName))
        {
            return;
        }

        var schemeDir = Path.Combine(SchemesBaseDir, schemeName);
        if (Path.GetFullPath(schemeDir).TrimEnd('\\') == Path.GetFullPath(SchemesBaseDir).TrimEnd('\\'))
        {
            return;
        }

        if (Directory.Exists(schemeDir))
        {
            Directory.Delete(schemeDir, true);
        }
    }

    /// <summary>
    /// 导出方案为 .wsp 文件（ZIP 格式）
    /// </summary>
    public static void ExportScheme(string schemeName, string outputPath)
    {
        var schemeDir = Path.Combine(SchemesBaseDir, schemeName);
        if (!Directory.Exists(schemeDir))
            throw new DirectoryNotFoundException($"方案 {schemeName} 不存在");

        System.IO.Compression.ZipFile.CreateFromDirectory(schemeDir, outputPath);
    }

    /// <summary>
    /// 导入 .wsp 文件
    /// </summary>
    public static void ImportScheme(string wspFilePath, string? newSchemeName = null)
    {
        if (!File.Exists(wspFilePath))
            throw new FileNotFoundException("方案文件不存在", wspFilePath);

        // 如果没有指定新名称，使用文件名
        if (string.IsNullOrEmpty(newSchemeName))
        {
            newSchemeName = Path.GetFileNameWithoutExtension(wspFilePath);
        }

        if (string.IsNullOrWhiteSpace(newSchemeName))
        {
            throw new ArgumentException("方案名称不能为空", nameof(newSchemeName));
        }

        var targetDir = Path.Combine(SchemesBaseDir, newSchemeName);
        if (Path.GetFullPath(targetDir).TrimEnd('\\') == Path.GetFullPath(SchemesBaseDir).TrimEnd('\\'))
        {
            throw new InvalidOperationException("无法导入到方案根目录");
        }

        // 如果目标已存在，先删除
        if (Directory.Exists(targetDir))
        {
            Directory.Delete(targetDir, true);
        }

        System.IO.Compression.ZipFile.ExtractToDirectory(wspFilePath, targetDir);
    }

    public class AppState
    {
        public string LastScheme { get; set; } = string.Empty;
        public string LastPage { get; set; } = string.Empty;
        public double? WindowLeft { get; set; }
        public double? WindowTop { get; set; }
        public double? WindowWidth { get; set; }
        public double? WindowHeight { get; set; }
        public bool IsMaximized { get; set; } = false;
    }

    /// <summary>
    /// 获取完整的应用程序状态数据
    /// </summary>
    public static AppState GetAppStateData()
    {
        var statePath = Path.Combine(SchemesBaseDir, "last_state.json");
        if (!File.Exists(statePath))
            return new AppState();

        try
        {
            var json = File.ReadAllText(statePath);
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var state = new AppState();

            if (root.TryGetProperty("last_scheme", out var pScheme))
                state.LastScheme = pScheme.GetString() ?? string.Empty;

            if (root.TryGetProperty("last_page", out var pPage))
                state.LastPage = pPage.GetString() ?? string.Empty;

            if (root.TryGetProperty("window_left", out var pLeft) && pLeft.ValueKind == JsonValueKind.Number)
                state.WindowLeft = pLeft.GetDouble();

            if (root.TryGetProperty("window_top", out var pTop) && pTop.ValueKind == JsonValueKind.Number)
                state.WindowTop = pTop.GetDouble();

            if (root.TryGetProperty("window_width", out var pWidth) && pWidth.ValueKind == JsonValueKind.Number)
                state.WindowWidth = pWidth.GetDouble();

            if (root.TryGetProperty("window_height", out var pHeight) && pHeight.ValueKind == JsonValueKind.Number)
                state.WindowHeight = pHeight.GetDouble();

            if (root.TryGetProperty("window_maximized", out var pMax))
                state.IsMaximized = pMax.GetBoolean();

            return state;
        }
        catch
        {
            return new AppState();
        }
    }

    /// <summary>
    /// 保存完整的应用程序状态数据
    /// </summary>
    public static void SaveAppStateData(AppState state)
    {
        try
        {
            var statePath = Path.Combine(SchemesBaseDir, "last_state.json");
            var dto = new
            {
                last_scheme = state.LastScheme,
                last_page = state.LastPage,
                window_left = state.WindowLeft,
                window_top = state.WindowTop,
                window_width = state.WindowWidth,
                window_height = state.WindowHeight,
                window_maximized = state.IsMaximized
            };
            var json = JsonSerializer.Serialize(dto, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(statePath, json);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"保存应用状态数据失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 保存最后使用的方案名称 (保留兼容性)
    /// </summary>
    public static void SaveLastScheme(string schemeName)
    {
        try
        {
            var state = GetAppStateData();
            state.LastScheme = schemeName;
            SaveAppStateData(state);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"保存最后使用的方案失败: {ex.Message}");
        }
    }

    /// <summary>
    /// 获取最后使用的方案名称 (保留兼容性)
    /// </summary>
    public static string? GetLastScheme()
    {
        return GetAppStateData().LastScheme;
    }
}

/// <summary>
/// 方案模型
/// </summary>
public class Scheme
{
    public string Name { get; set; } = string.Empty;
    public string TemplatePath { get; set; } = string.Empty;
    public List<string> Variables { get; set; } = new();
    public string DefaultOutputDir { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    public string FileNameTemplate { get; set; } = "文档-{{序号}}.docx";
    public string DefaultSubfolderVar { get; set; } = string.Empty;
    public int LastTabIndex { get; set; } = 0;
    public Dictionary<string, string> PastedTexts { get; set; } = new();
    public List<string> FixedVariables { get; set; } = new();
    public DateTime CreatedAt { get; set; } = DateTime.Now;
    public DateTime UpdatedAt { get; set; } = DateTime.Now;
}
