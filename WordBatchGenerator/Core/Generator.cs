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
                var originalData = dataList[i];
                var data = new Dictionary<string, string>(originalData);
                if (!data.ContainsKey("序号"))
                {
                    data["序号"] = (i + 1).ToString();
                }

                var fileName = ReplaceVariables(fileNameTemplate, data);
                
                var targetOutputDir = outputDir;
                if (!string.IsNullOrEmpty(subfolderVar))
                {
                    var subfolder = ResolveSubfolder(subfolderVar, data);
                    if (!string.IsNullOrEmpty(subfolder))
                    {
                        targetOutputDir = Path.Combine(outputDir, subfolder);
                    }
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

    /// <summary>
    /// 合并多个 Word 文档为一个文档
    /// </summary>
    public static void MergeWordFiles(List<string> sourceFiles, string destPath)
    {
        if (sourceFiles == null || sourceFiles.Count == 0) return;

        // 1. 复制第一个文件作为基础文件
        File.Copy(sourceFiles[0], destPath, true);

        // 2. 依次追加其它文件内容
        using var destDoc = WordprocessingDocument.Open(destPath, true);
        var mainPart = destDoc.MainDocumentPart;
        if (mainPart == null) return;

        var body = mainPart.Document.Body;
        if (body == null) return;

        int chunkId = 1;
        // 从第二个文件开始合并
        for (int i = 1; i < sourceFiles.Count; i++)
        {
            var srcFile = sourceFiles[i];
            if (!File.Exists(srcFile)) continue;

            // 添加分页符
            body.AppendChild(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

            string altChunkId = $"AltChunkId{chunkId++}";
            var chunkPart = mainPart.AddAlternativeFormatImportPart(
                AlternativeFormatImportPartType.WordprocessingML, 
                altChunkId
            );

            using (var fileStream = File.OpenRead(srcFile))
            {
                chunkPart.FeedData(fileStream);
            }

            var altChunk = new AltChunk { Id = altChunkId };
            body.AppendChild(altChunk);
        }

        destDoc.Save();
    }

    /// <summary>
    /// 将 Word 文档使用本地 Word 软件转换为 PDF 并保存
    /// </summary>
    public static bool ConvertDocxToPdfDynamic(string docxPath, string pdfPath)
    {
        Type? wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null) return false;

        object? wordApp = null;
        object? doc = null;
        try
        {
            wordApp = Activator.CreateInstance(wordType);
            if (wordApp == null) return false;

            // wordApp.Visible = false
            wordType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, wordApp, new object[] { false });

            // object docs = wordApp.Documents
            object? docs = wordType.InvokeMember("Documents", System.Reflection.BindingFlags.GetProperty, null, wordApp, null);
            if (docs == null) return false;

            // object doc = docs.Open(docxPath)
            doc = docs.GetType().InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, docs, new object[] { docxPath });
            if (doc == null) return false;

            // wdExportFormatPDF = 17
            // doc.ExportAsFixedFormat(pdfPath, 17)
            doc.GetType().InvokeMember("ExportAsFixedFormat", System.Reflection.BindingFlags.InvokeMethod, null, doc, new object[] { pdfPath, 17 });

            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Word 转 PDF 失败: {ex.Message}");
            return false;
        }
        finally
        {
            if (doc != null)
            {
                try
                {
                    doc.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, doc, new object[] { 0 /* wdDoNotSaveChanges = 0 */ });
                }
                catch { }
            }
            if (wordApp != null)
            {
                try
                {
                    wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, wordApp, null);
                }
                catch { }
            }
        }
    }

    /// <summary>
    /// 解析并动态生成子目录路径，支持多级规则 (如：{{年份}}/{{客户}})
    /// </summary>
    public static string ResolveSubfolder(string subfolderRule, Dictionary<string, string> data)
    {
        if (string.IsNullOrEmpty(subfolderRule)) return string.Empty;

        string rule = subfolderRule;
        // 兼容老版本：如果不含双花括号，说明是单一变量名称，自动包装为新格式
        if (!rule.Contains("{{") && !rule.Contains("}}"))
        {
            rule = "{{" + rule + "}}";
        }

        // 替换所有变量占位符
        string resolved = ReplaceVariables(rule, data);

        // 清理非法字符并拼接
        return SanitizeSubfolderPath(resolved);
    }

    /// <summary>
    /// 替换子目录中非法的字符，并保持路径层级（/ 或 \）
    /// </summary>
    public static string SanitizeSubfolderPath(string path)
    {
        if (string.IsNullOrEmpty(path)) return string.Empty;

        var parts = path.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);
        var sanitizedParts = parts.Select(p => {
            // 对每个目录段进行文件名合法性清理
            var clean = SanitizeFileName(p);
            // 去除可能产生的扩展名（如果是子目录的话）
            if (clean.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
            {
                clean = clean.Substring(0, clean.Length - 5);
            }
            return clean.Trim();
        });

        return string.Join(Path.DirectorySeparatorChar.ToString(), sanitizedParts);
    }

    /// <summary>
    /// 打印指定文档，对于 docx 使用 Word COM，对于其他或回退使用 Windows Shell 打印
    /// </summary>
    public static void PrintDocument(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("待打印的文件不存在", filePath);

        string ext = Path.GetExtension(filePath).ToLower();
        if (ext == ".docx")
        {
            Type? wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType != null)
            {
                object? wordApp = null;
                object? doc = null;
                try
                {
                    wordApp = Activator.CreateInstance(wordType);
                    if (wordApp != null)
                    {
                        wordType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, wordApp, new object[] { false });
                        object? docs = wordType.InvokeMember("Documents", System.Reflection.BindingFlags.GetProperty, null, wordApp, null);
                        if (docs != null)
                        {
                            doc = docs.GetType().InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, docs, new object[] { filePath });
                            if (doc != null)
                            {
                                doc.GetType().InvokeMember("PrintOut", System.Reflection.BindingFlags.InvokeMethod, null, doc, null);
                                return;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Word COM 打印失败，将回退到 Shell: {ex.Message}");
                }
                finally
                {
                    if (doc != null)
                    {
                        try
                        {
                            doc.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, doc, new object[] { 0 /* wdDoNotSaveChanges = 0 */ });
                        }
                        catch { }
                    }
                    if (wordApp != null)
                    {
                        try
                        {
                            wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, wordApp, null);
                        }
                        catch { }
                    }
                }
            }
        }

        // 回退逻辑 / PDF 打印逻辑：直接调用系统的 print 动词发送到默认打印机
        System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo
        {
            FileName = filePath,
            Verb = "print",
            UseShellExecute = true,
            CreateNoWindow = true,
            WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
        };
        using (var p = System.Diagnostics.Process.Start(info))
        {
            p?.WaitForExit(3000);
        }
    }

    /// <summary>
    /// 设置指定打印机的单双面打印模式
    /// </summary>
    public static void SetPrinterDuplex(string printerName, string duplexMode)
    {
        try
        {
            using (var printServer = new System.Printing.LocalPrintServer())
            {
                System.Printing.PrintQueue? queue = null;
                try
                {
                    queue = new System.Printing.PrintQueue(printServer, printerName, System.Printing.PrintSystemDesiredAccess.AdministratePrinter);
                }
                catch
                {
                    try
                    {
                        queue = new System.Printing.PrintQueue(printServer, printerName, System.Printing.PrintSystemDesiredAccess.UsePrinter);
                    }
                    catch
                    {
                        // 忽略
                    }
                }

                if (queue != null)
                {
                    using (queue)
                    {
                        System.Printing.Duplexing targetDuplex = System.Printing.Duplexing.Unknown;
                        switch (duplexMode.ToLower())
                        {
                            case "onesided":
                            case "simplex":
                                targetDuplex = System.Printing.Duplexing.OneSided;
                                break;
                            case "duplexlongedge":
                            case "twosidedlongedge":
                                targetDuplex = System.Printing.Duplexing.TwoSidedLongEdge;
                                break;
                            case "duplexshortedge":
                            case "twosidedshortedge":
                                targetDuplex = System.Printing.Duplexing.TwoSidedShortEdge;
                                break;
                        }

                        if (targetDuplex != System.Printing.Duplexing.Unknown)
                        {
                            if (queue.UserPrintTicket != null)
                            {
                                queue.UserPrintTicket.Duplexing = targetDuplex;
                            }
                            if (queue.DefaultPrintTicket != null)
                            {
                                queue.DefaultPrintTicket.Duplexing = targetDuplex;
                            }
                            queue.Commit();
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"修改打印机双面模式失败: {ex.Message}");
            // 即使修改失败，也继续打印，不阻断主流程
        }
    }

    /// <summary>
    /// 自定义打印机和双面设置打印 Word 或 PDF 文档
    /// </summary>
    public static void PrintFileWithSettings(string filePath, string printerName, string duplexMode)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException("待打印的文件不存在", filePath);

        // 1. 尝试修改打印机双面模式
        SetPrinterDuplex(printerName, duplexMode);

        string ext = Path.GetExtension(filePath).ToLower();
        if (ext == ".docx")
        {
            Type? wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType != null)
            {
                object? wordApp = null;
                object? doc = null;
                try
                {
                    wordApp = Activator.CreateInstance(wordType);
                    if (wordApp != null)
                    {
                        wordType.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, wordApp, new object[] { false });
                        
                        // 设置当前活动打印机
                        wordType.InvokeMember("ActivePrinter", System.Reflection.BindingFlags.SetProperty, null, wordApp, new object[] { printerName });

                        object? docs = wordType.InvokeMember("Documents", System.Reflection.BindingFlags.GetProperty, null, wordApp, null);
                        if (docs != null)
                        {
                            doc = docs.GetType().InvokeMember("Open", System.Reflection.BindingFlags.InvokeMethod, null, docs, new object[] { filePath });
                            if (doc != null)
                            {
                                doc.GetType().InvokeMember("PrintOut", System.Reflection.BindingFlags.InvokeMethod, null, doc, null);
                                return;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Word COM 打印失败: {ex.Message}");
                    throw new Exception($"Word 打印失败: {ex.Message}");
                }
                finally
                {
                    if (doc != null)
                    {
                        try
                        {
                            doc.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, doc, new object[] { 0 /* wdDoNotSaveChanges = 0 */ });
                        }
                        catch { }
                    }
                    if (wordApp != null)
                    {
                        try
                        {
                            wordType.InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, wordApp, null);
                        }
                        catch { }
                    }
                }
            }
        }

        // 2. 对于 PDF 或在没有 Word COM 的情况下，调用默认或 Microsoft Edge 进行静默打印
        try
        {
            // 尝试检测本地 Edge 路径并以静默参数执行
            string edgePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft\Edge\Application\msedge.exe");
            if (!File.Exists(edgePath))
            {
                edgePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft\Edge\Application\msedge.exe");
            }

            if (File.Exists(edgePath))
            {
                // Edge headless print to printer
                System.Diagnostics.ProcessStartInfo edgeInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = edgePath,
                    Arguments = $"--headless --print-to-printer --printer-name=\"{printerName}\" \"{filePath}\"",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                };
                using (var p = System.Diagnostics.Process.Start(edgeInfo))
                {
                    p?.WaitForExit(5000);
                }
            }
            else
            {
                // 回退到系统 print 动词（使用默认打印机设置，因为可能无法设置非默认打印机）
                System.Diagnostics.ProcessStartInfo info = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    Verb = "print",
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                };
                using (var p = System.Diagnostics.Process.Start(info))
                {
                    p?.WaitForExit(3000);
                }
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"PDF/其它文档打印失败: {ex.Message}，可能是系统没有默认的 PDF 关联程序。");
        }
    }
}
