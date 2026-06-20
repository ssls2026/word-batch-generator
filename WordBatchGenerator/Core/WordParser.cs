using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace WordBatchGenerator.Core;

/// <summary>
/// Word 文档解析器：提取段落、表格、变量占位符
/// </summary>
public class WordParser
{
    /// <summary>
    /// 解析 Word 文档，提取所有段落和表格内容
    /// </summary>
    public static List<ParagraphInfo> ParseDocument(string filePath)
    {
        using var doc = WordprocessingDocument.Open(filePath, false);
        var body = doc.MainDocumentPart?.Document.Body;
        if (body == null) return new List<ParagraphInfo>();
        return ParseDocument(body);
    }

    /// <summary>
    /// 从已打开的 Body 中解析段落和表格内容
    /// </summary>
    public static List<ParagraphInfo> ParseDocument(Body body)
    {
        var paragraphs = new List<ParagraphInfo>();
        int index = 0;
        foreach (var element in body.Elements())
        {
            if (element is Paragraph para)
            {
                paragraphs.Add(new ParagraphInfo
                {
                    Index = index++,
                    Text = GetParagraphText(para),
                    Element = para,
                    IsInTable = false
                });
            }
            else if (element is Table table)
            {
                foreach (var row in table.Elements<TableRow>())
                {
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        foreach (var cellPara in cell.Elements<Paragraph>())
                        {
                            paragraphs.Add(new ParagraphInfo
                            {
                                Index = index++,
                                Text = GetParagraphText(cellPara),
                                Element = cellPara,
                                IsInTable = true
                            });
                        }
                    }
                }
            }
        }
        return paragraphs;
    }

    /// <summary>
    /// 获取段落纯文本
    /// </summary>
    private static string GetParagraphText(Paragraph paragraph)
    {
        var sb = new StringBuilder();
        foreach (var run in paragraph.Elements<Run>())
        {
            foreach (var text in run.Elements<Text>())
            {
                sb.Append(text.Text);
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// 提取文档中的所有变量占位符（格式：{{变量名}}）
    /// </summary>
    public static List<string> ExtractVariables(string filePath)
    {
        var variables = new HashSet<string>();
        var paragraphs = ParseDocument(filePath);

        foreach (var para in paragraphs)
        {
            var matches = System.Text.RegularExpressions.Regex.Matches(para.Text, @"\{\{(.+?)\}\}");
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                variables.Add(match.Groups[1].Value);
            }
        }

        return variables.ToList();
    }

    /// <summary>
    /// 在段落中/跨段落替换文本为变量占位符，完美保留其余部分的格式并支持段落合并
    /// </summary>
    public static void ReplaceTextAcrossParagraphs(
        string filePath,
        int startPIndex,
        int endPIndex,
        int startIdx,
        int endIdx,
        string variableName)
    {
        using var doc = WordprocessingDocument.Open(filePath, true);
        var body = doc.MainDocumentPart?.Document.Body;
        if (body == null) return;

        var paragraphs = ParseDocument(body);

        if (startPIndex < 0 || startPIndex >= paragraphs.Count ||
            endPIndex < 0 || endPIndex >= paragraphs.Count ||
            startPIndex > endPIndex)
        {
            throw new ArgumentException("段落索引无效");
        }

        var startParaInfo = paragraphs[startPIndex];
        var endParaInfo = paragraphs[endPIndex];

        if (startPIndex == endPIndex)
        {
            var para = startParaInfo.Element;
            var text = startParaInfo.Text;

            if (startIdx < 0 || endIdx >= text.Length || startIdx > endIdx)
                throw new ArgumentException("字符索引无效");

            var charMapping = new List<(char Value, RunProperties? Properties)>();
            foreach (var run in para.Elements<Run>())
            {
                var runProps = run.RunProperties;
                var runText = GetRunText(run);
                foreach (var c in runText)
                {
                    charMapping.Add((c, runProps));
                }
            }

            var newRunsData = new List<(string Text, RunProperties? Properties)>();

            // 1. Before selection
            var currentText = new StringBuilder();
            RunProperties? lastProps = null;
            for (int i = 0; i < startIdx; i++)
            {
                var (c, props) = charMapping[i];
                if (props != lastProps && currentText.Length > 0)
                {
                    newRunsData.Add((currentText.ToString(), lastProps));
                    currentText.Clear();
                }
                currentText.Append(c);
                lastProps = props;
            }
            if (currentText.Length > 0)
            {
                newRunsData.Add((currentText.ToString(), lastProps));
            }

            // 2. Selection replacement
            var replacementText = $"{{{{{variableName}}}}}";
            RunProperties? selectedProps = charMapping.Count > startIdx ? charMapping[startIdx].Properties : null;
            newRunsData.Add((replacementText, selectedProps));

            // 3. After selection
            currentText.Clear();
            lastProps = null;
            for (int i = endIdx + 1; i < text.Length; i++)
            {
                var (c, props) = charMapping[i];
                if (props != lastProps && currentText.Length > 0)
                {
                    newRunsData.Add((currentText.ToString(), lastProps));
                    currentText.Clear();
                }
                currentText.Append(c);
                lastProps = props;
            }
            if (currentText.Length > 0)
            {
                newRunsData.Add((currentText.ToString(), lastProps));
            }

            para.RemoveAllChildren<Run>();

            foreach (var runData in newRunsData)
            {
                var newRun = new Run();
                if (runData.Properties != null)
                {
                    newRun.AppendChild(runData.Properties.CloneNode(true));
                }
                newRun.AppendChild(new Text(runData.Text));
                para.AppendChild(newRun);
            }
        }
        else
        {
            var startPara = startParaInfo.Element;
            var endPara = endParaInfo.Element;

            var startText = startParaInfo.Text;
            var endText = endParaInfo.Text;

            var startCharMapping = new List<(char Value, RunProperties? Properties)>();
            foreach (var run in startPara.Elements<Run>())
            {
                var runProps = run.RunProperties;
                var runText = GetRunText(run);
                foreach (var c in runText)
                {
                    startCharMapping.Add((c, runProps));
                }
            }

            var endCharMapping = new List<(char Value, RunProperties? Properties)>();
            foreach (var run in endPara.Elements<Run>())
            {
                var runProps = run.RunProperties;
                var runText = GetRunText(run);
                foreach (var c in runText)
                {
                    endCharMapping.Add((c, runProps));
                }
            }

            var newStartRuns = new List<(string Text, RunProperties? Properties)>();
            var currentText = new StringBuilder();
            RunProperties? lastProps = null;
            for (int i = 0; i < startIdx && i < startCharMapping.Count; i++)
            {
                var (c, props) = startCharMapping[i];
                if (props != lastProps && currentText.Length > 0)
                {
                    newStartRuns.Add((currentText.ToString(), lastProps));
                    currentText.Clear();
                }
                currentText.Append(c);
                lastProps = props;
            }
            if (currentText.Length > 0)
            {
                newStartRuns.Add((currentText.ToString(), lastProps));
            }

            var newEndRuns = new List<(string Text, RunProperties? Properties)>();
            currentText.Clear();
            lastProps = null;
            for (int i = endIdx + 1; i < endCharMapping.Count; i++)
            {
                var (c, props) = endCharMapping[i];
                if (props != lastProps && currentText.Length > 0)
                {
                    newEndRuns.Add((currentText.ToString(), lastProps));
                    currentText.Clear();
                }
                currentText.Append(c);
                lastProps = props;
            }
            if (currentText.Length > 0)
            {
                newEndRuns.Add((currentText.ToString(), lastProps));
            }

            RunProperties? selectedProps = null;
            if (startIdx < startCharMapping.Count)
                selectedProps = startCharMapping[startIdx].Properties;
            else if (startCharMapping.Count > 0)
                selectedProps = startCharMapping[^1].Properties;

            startPara.RemoveAllChildren<Run>();

            foreach (var runData in newStartRuns)
            {
                var newRun = new Run();
                if (runData.Properties != null)
                {
                    newRun.AppendChild(runData.Properties.CloneNode(true));
                }
                newRun.AppendChild(new Text(runData.Text));
                startPara.AppendChild(newRun);
            }

            var varRun = new Run();
            if (selectedProps != null)
            {
                varRun.AppendChild(selectedProps.CloneNode(true));
            }
            varRun.AppendChild(new Text($"{{{{{variableName}}}}}"));
            startPara.AppendChild(varRun);

            foreach (var runData in newEndRuns)
            {
                var newRun = new Run();
                if (runData.Properties != null)
                {
                    newRun.AppendChild(runData.Properties.CloneNode(true));
                }
                newRun.AppendChild(new Text(runData.Text));
                startPara.AppendChild(newRun);
            }

            for (int i = startPIndex + 1; i <= endPIndex; i++)
            {
                var paraToDelete = paragraphs[i].Element;
                paraToDelete.Remove();
            }
        }

        doc.Save();
    }

    /// <summary>
    /// 将 Word 文档转换为 HTML（用于预览）
    /// </summary>
    public static string ConvertToHtml(string filePath)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<!DOCTYPE html>");
        sb.AppendLine("<html><head><meta charset='utf-8'>");
        sb.AppendLine("<style>");
        sb.AppendLine("body { font-family: '微软雅黑', 'Microsoft YaHei', '宋体', SimSun; font-size: 15px; color: #1e293b; padding: 35px; line-height: 1.15; white-space: pre-wrap; word-break: break-all; }");
        sb.AppendLine("p { margin: 0; padding: 0; }");
        sb.AppendLine("table { border-collapse: collapse; width: 100%; margin: 15px 0; background-color: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }");
        sb.AppendLine("td { border: 1px solid #cbd5e1; padding: 10px 12px; vertical-align: middle; }");
        sb.AppendLine(".variable { font-weight: bold; color: #4f46e5; background-color: #f0f3ff; border: 1px solid #c7d2fe; border-radius: 4px; padding: 2px 6px; display: inline-block; }");
        sb.AppendLine("</style>");
        sb.AppendLine("</head><body>");

        using var doc = WordprocessingDocument.Open(filePath, false);
        var body = doc.MainDocumentPart?.Document.Body;
        if (body == null) return sb.ToString();

        int pIndex = 0;
        foreach (var element in body.Elements())
        {
            if (element is Paragraph para)
            {
                sb.AppendLine(RenderParagraphWithStyles(para, pIndex++));
            }
            else if (element is Table table)
            {
                sb.AppendLine(RenderTableHtml(table, ref pIndex));
            }
        }

        sb.AppendLine("</body></html>");
        return sb.ToString();
    }

    /// <summary>
    /// 渲染段落为带完整样式的 HTML（高保真还原 Word 格式）
    /// </summary>
    private static string RenderParagraphWithStyles(Paragraph paragraph, int pIndex)
    {
        var styles = new List<string>();

        // 对齐方式
        var alignment = paragraph.ParagraphProperties?.Justification?.Val?.Value;
        var alignStyle = "text-align: justify;";
        if (alignment != null)
        {
            alignStyle = alignment.ToString() switch
            {
                "left" => "text-align: left;",
                "center" => "text-align: center;",
                "right" => "text-align: right;",
                "both" => "text-align: justify;",
                _ => "text-align: justify;"
            };
        }
        styles.Add(alignStyle);

        var fmt = paragraph.ParagraphProperties;
        if (fmt != null)
        {
            // 首行缩进
            if (fmt.Indentation?.FirstLine != null)
            {
                var pt = ConvertTwipsToPoints(fmt.Indentation.FirstLine.Value);
                var px = (int)(pt * 1.33);
                if (px >= 0)
                    styles.Add($"text-indent: {px}px;");
                else
                    styles.Add($"text-indent: {px}px; padding-left: {-px}px;");
            }

            // 左右缩进
            if (fmt.Indentation?.Left != null)
            {
                var pt = ConvertTwipsToPoints(fmt.Indentation.Left.Value);
                styles.Add($"margin-left: {(int)(pt * 1.33)}px;");
            }
            if (fmt.Indentation?.Right != null)
            {
                var pt = ConvertTwipsToPoints(fmt.Indentation.Right.Value);
                styles.Add($"margin-right: {(int)(pt * 1.33)}px;");
            }

            // 段前段后间距
            var marginTop = fmt.SpacingBetweenLines?.Before != null
                ? $"{(int)(ConvertTwipsToPoints(fmt.SpacingBetweenLines.Before.Value) * 1.33)}px"
                : "0px";
            var marginBottom = fmt.SpacingBetweenLines?.After != null
                ? $"{(int)(ConvertTwipsToPoints(fmt.SpacingBetweenLines.After.Value) * 1.33)}px"
                : "2px";
            styles.Add($"margin-top: {marginTop}; margin-bottom: {marginBottom};");

            // 行距
            if (fmt.SpacingBetweenLines?.Line != null)
            {
                var lineValue = fmt.SpacingBetweenLines.Line.Value;
                var lineRule = fmt.SpacingBetweenLines.LineRule?.Value;
                var lineRaw = lineValue; // string or int from Open XML

                if (lineRule == LineSpacingRuleValues.Exact || lineRule == LineSpacingRuleValues.AtLeast)
                {
                    var pt = ConvertTwipsToPoints(lineRaw);
                    styles.Add($"line-height: {(int)(pt * 1.33)}px;");
                }
                else
                {
                    // 按倍数计算（Word 中 240 = 1.0 倍行距）
                    if (int.TryParse(lineRaw, out var raw))
                    {
                        var multiplier = raw / 240.0;
                        styles.Add($"line-height: {multiplier:F2};");
                    }
                    else
                    {
                        styles.Add("line-height: 1.15;");
                    }
                }
            }
            else
            {
                styles.Add("line-height: 1.15;");
            }
        }
        else
        {
            styles.Add("margin-top: 0px; margin-bottom: 2px; line-height: 1.15;");
        }

        styles.Add("font-size: 15px; color: #1e293b; font-family: 微软雅黑, Microsoft YaHei, 宋体, SimSun;");

        var styleStr = string.Join(" ", styles);
        var sb = new StringBuilder();
        sb.Append($"<p data-p-index='{pIndex}' style='{styleStr}'>");

        // 渲染 runs
        var runs = paragraph.Elements<Run>().ToList();
        if (runs.Count == 0 || string.IsNullOrEmpty(GetParagraphText(paragraph)))
        {
            // 空段落添加零宽空格防止块崩塌
            sb.Append("<span>&#8203;</span>");
        }
        else
        {
            foreach (var run in runs)
            {
                var runStyles = new List<string>();
                var runProps = run.RunProperties;

                if (runProps != null)
                {
                    // 粗体
                    if (runProps.Bold != null)
                        runStyles.Add("font-weight: bold;");

                    // 斜体
                    if (runProps.Italic != null)
                        runStyles.Add("font-style: italic;");

                    // 颜色
                    if (runProps.Color?.Val != null)
                    {
                        var color = ConvertHexToRgb(runProps.Color.Val.Value);
                        runStyles.Add($"color: {color};");
                    }

                    // 字体大小
                    if (runProps.FontSize?.Val != null)
                    {
                        if (int.TryParse(runProps.FontSize.Val.Value, out var halfPoints))
                        {
                            var px = (int)(halfPoints / 2.0 * 1.3);
                            runStyles.Add($"font-size: {px}px;");
                        }
                    }

                    // 字体名称
                    if (runProps.RunFonts?.Ascii != null)
                    {
                        runStyles.Add($"font-family: {runProps.RunFonts.Ascii.Value};");
                    }
                }

                var text = GetRunText(run);
                // 转义 HTML 敏感字符
                text = text.Replace("&", "&amp;")
                           .Replace("<", "&lt;")
                           .Replace(">", "&gt;")
                           .Replace("\n", "<br>");

                // 变量占位符高亮
                text = System.Text.RegularExpressions.Regex.Replace(
                    text,
                    @"\{\{([^}]+)\}\}",
                    "<span class='variable'>{{$1}}</span>"
                );

                if (runStyles.Count > 0)
                {
                    var runStyleStr = string.Join(" ", runStyles);
                    sb.Append($"<span style=\"{runStyleStr}\">{text}</span>");
                }
                else
                {
                    sb.Append($"<span>{text}</span>");
                }
            }
        }

        sb.Append("</p>");
        return sb.ToString();
    }

    private static string RenderParagraphHtml(Paragraph paragraph, int pIndex)
    {
        return RenderParagraphWithStyles(paragraph, pIndex);
    }

    private static string GetRunText(Run run)
    {
        var sb = new StringBuilder();
        foreach (var text in run.Elements<Text>())
        {
            sb.Append(text.Text);
        }
        return sb.ToString();
    }

    /// <summary>
    /// 将 Twips 转换为 Points（1 point = 20 twips）
    /// </summary>
    private static double ConvertTwipsToPoints(string? twips)
    {
        if (int.TryParse(twips, out var value))
            return value / 20.0;
        return 0;
    }

    /// <summary>
    /// 将十六进制颜色转换为 RGB
    /// </summary>
    private static string ConvertHexToRgb(string? hex)
    {
        if (hex == null || hex == "auto" || hex.Length != 6)
            return "rgb(30, 41, 59)"; // 默认颜色

        try
        {
            var r = Convert.ToInt32(hex.Substring(0, 2), 16);
            var g = Convert.ToInt32(hex.Substring(2, 2), 16);
            var b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return $"rgb({r}, {g}, {b})";
        }
        catch
        {
            return "rgb(30, 41, 59)";
        }
    }

    /// <summary>
    /// 渲染表格为 HTML（支持合并单元格）
    /// </summary>
    private static string RenderTableHtml(Table table, ref int pIndex)
    {
        var sb = new StringBuilder();
        sb.Append("<table style='border-collapse: collapse; width: 100%; margin: 15px 0; background-color: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.05);'>");

        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0)
        {
            sb.Append("</table>");
            return sb.ToString();
        }

        var R = rows.Count;
        var C = rows[0].Elements<TableCell>().Count();

        // 构建单元格网格（用于跟踪合并）
        var grid = new TableCell[R, C];
        for (int r = 0; r < R; r++)
        {
            var cells = rows[r].Elements<TableCell>().ToList();
            for (int c = 0; c < cells.Count && c < C; c++)
            {
                grid[r, c] = cells[c];
            }
        }

        var visited = new HashSet<TableCell>();

        for (int r = 0; r < R; r++)
        {
            sb.Append("<tr>");
            var cells = rows[r].Elements<TableCell>().ToList();

            for (int c = 0; c < cells.Count && c < C; c++)
            {
                var cell = grid[r, c];
                if (cell == null || visited.Contains(cell))
                    continue;

                visited.Add(cell);

                // 计算 colspan
                int colspan = 1;
                while (c + colspan < C && grid[r, c + colspan] == cell)
                {
                    colspan++;
                }

                // 计算 rowspan
                int rowspan = 1;
                while (r + rowspan < R && grid[r + rowspan, c] == cell)
                {
                    bool allMatch = true;
                    for (int dc = 0; dc < colspan; dc++)
                    {
                        if (grid[r + rowspan, c + dc] != cell)
                        {
                            allMatch = false;
                            break;
                        }
                    }

                    if (!allMatch)
                        break;

                    // 标记为已访问
                    for (int dc = 0; dc < colspan; dc++)
                    {
                        visited.Add(grid[r + rowspan, c + dc]);
                    }
                    rowspan++;
                }

                // 构建单元格属性
                var tdAttrs = new List<string>();
                if (colspan > 1)
                    tdAttrs.Add($"colspan='{colspan}'");
                if (rowspan > 1)
                    tdAttrs.Add($"rowspan='{rowspan}'");

                var tdStyle = "border: 1px solid #cbd5e1; padding: 10px 12px; vertical-align: middle; background-color: #ffffff;";
                var attrsStr = tdAttrs.Count > 0 ? " " + string.Join(" ", tdAttrs) : "";
                sb.Append($"<td style='{tdStyle}'{attrsStr}>");

                // 渲染单元格内段落
                foreach (var para in cell.Elements<Paragraph>())
                {
                    sb.Append(RenderParagraphWithStyles(para, pIndex++));
                }

                sb.Append("</td>");
            }
            sb.Append("</tr>");
        }

        sb.Append("</table>");
        return sb.ToString();
    }
}

/// <summary>
/// 段落信息
/// </summary>
public class ParagraphInfo
{
    public int Index { get; set; }
    public string Text { get; set; } = string.Empty;
    public Paragraph Element { get; set; } = null!;
    public bool IsInTable { get; set; }
}
