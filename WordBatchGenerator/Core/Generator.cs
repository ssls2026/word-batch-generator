using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text;

namespace WordBatchGenerator.Core;

/// <summary>
/// Word 文档批量生成器
/// </summary>
public class Generator
{
    /// <summary>
    /// 批量生成 Word 文档
    /// </summary>
    /// <param name="templatePath">模板文件路径</param>
    /// <param name="dataList">数据列表（每行为一个字典）</param>
    /// <param name="outputDir">输出目录</param>
    /// <param name="fileNameTemplate">文件名模板（例如：授权书-{{公司名称}}.docx）</param>
    /// <param name="progress">进度回调</param>
    /// <returns>生成的文件数量</returns>
    public static int GenerateBatch(
        string templatePath,
        List<Dictionary<string, string>> dataList,
        string outputDir,
        string fileNameTemplate,
        string subfolderVar = "",
        Action<int, int>? progress = null)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("模板文件不存在", templatePath);

        int successCount = 0;

        for (int i = 0; i < dataList.Count; i++)
        {
            try
            {
                var data = dataList[i];
                var fileName = ReplaceVariables(fileNameTemplate, data);
                
                var targetOutputDir = outputDir;
                if (!string.IsNullOrEmpty(subfolderVar) && data.TryGetValue(subfolderVar, out var subDirValue) && !string.IsNullOrEmpty(subDirValue))
                {
                    targetOutputDir = Path.Combine(outputDir, SanitizeFileName(subDirValue));
                }

                if (!Directory.Exists(targetOutputDir))
                    Directory.CreateDirectory(targetOutputDir);

                var outputPath = Path.Combine(targetOutputDir, SanitizeFileName(fileName));

                GenerateSingle(templatePath, data, outputPath);
                successCount++;

                progress?.Invoke(i + 1, dataList.Count);
            }
            catch (Exception ex)
            {
                // 记录错误但继续处理下一个
                Console.WriteLine($"生成第 {i + 1} 个文件失败: {ex.Message}");
            }
        }

        return successCount;
    }

    /// <summary>
    /// 生成单个 Word 文档
    /// </summary>
    private static void GenerateSingle(string templatePath, Dictionary<string, string> data, string outputPath)
    {
        // 复制模板文件
        File.Copy(templatePath, outputPath, true);

        // 打开并替换变量
        using var doc = WordprocessingDocument.Open(outputPath, true);
        var body = doc.MainDocumentPart?.Document.Body;
        if (body == null) return;

        // 替换所有段落中的变量
        foreach (var element in body.Descendants())
        {
            if (element is Text textElement)
            {
                textElement.Text = ReplaceVariables(textElement.Text, data);
            }
        }

        doc.Save();
    }

    /// <summary>
    /// 替换文本中的变量占位符
    /// </summary>
    public static string ReplaceVariables(string template, Dictionary<string, string> data)
    {
        var result = template;

        foreach (var kvp in data)
        {
            var placeholder = $"{{{{{kvp.Key}}}}}";
            result = result.Replace(placeholder, kvp.Value);
        }

        return result;
    }

    /// <summary>
    /// 清理文件名中的非法字符
    /// </summary>
    public static string SanitizeFileName(string fileName)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = new StringBuilder(fileName);

        foreach (var c in invalidChars)
        {
            sanitized.Replace(c, '_');
        }

        // 确保有 .docx 扩展名
        var result = sanitized.ToString();
        if (!result.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            result += ".docx";
        }

        return result;
    }

    /// <summary>
    /// 验证数据完整性（检查所有必需变量是否都有对应的数据列）
    /// </summary>
    public static List<string> ValidateData(List<string> requiredVariables, List<string> dataHeaders)
    {
        var missingVariables = new List<string>();

        foreach (var variable in requiredVariables)
        {
            if (!dataHeaders.Contains(variable))
            {
                missingVariables.Add(variable);
            }
        }

        return missingVariables;
    }
}
