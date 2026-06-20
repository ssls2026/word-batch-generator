using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.IO;

namespace WordBatchGenerator.Core;

/// <summary>
/// Excel 数据处理器：读取数据源
/// </summary>
public class ExcelHandler
{
    static ExcelHandler()
    {
        // 设置 EPPlus 授权上下文（非商业用途）
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    }

    /// <summary>
    /// 从 Excel 文件读取数据
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <param name="sheetIndex">工作表索引（从 1 开始）</param>
    /// <returns>数据列表，每行为一个字典（列名 -> 值）</returns>
    public static List<Dictionary<string, string>> ReadExcel(string filePath, int sheetIndex = 1)
    {
        var dataList = new List<Dictionary<string, string>>();

        using var package = new ExcelPackage(new FileInfo(filePath));

        if (package.Workbook.Worksheets.Count < sheetIndex)
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出范围");

        var worksheet = package.Workbook.Worksheets[sheetIndex - 1];
        var rowCount = worksheet.Dimension?.Rows ?? 0;
        var colCount = worksheet.Dimension?.Columns ?? 0;

        if (rowCount < 2) return dataList; // 至少需要表头 + 1 行数据

        // 读取表头（第一行）
        var headers = new List<string>();
        for (int col = 1; col <= colCount; col++)
        {
            var headerValue = worksheet.Cells[1, col].Text.Trim();
            headers.Add(string.IsNullOrEmpty(headerValue) ? $"列{col}" : headerValue);
        }

        // 读取数据行（从第二行开始）
        for (int row = 2; row <= rowCount; row++)
        {
            var rowData = new Dictionary<string, string>();
            bool isEmptyRow = true;

            for (int col = 1; col <= colCount; col++)
            {
                var cellValue = worksheet.Cells[row, col].Text.Trim();
                rowData[headers[col - 1]] = cellValue;

                if (!string.IsNullOrEmpty(cellValue))
                    isEmptyRow = false;
            }

            // 跳过空行
            if (!isEmptyRow)
                dataList.Add(rowData);
        }

        return dataList;
    }

    /// <summary>
    /// 获取 Excel 文件的所有工作表名称
    /// </summary>
    public static List<string> GetSheetNames(string filePath)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        return package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
    }

    /// <summary>
    /// 获取指定工作表的列名（表头）
    /// </summary>
    public static List<string> GetHeaders(string filePath, int sheetIndex = 1)
    {
        var headers = new List<string>();

        using var package = new ExcelPackage(new FileInfo(filePath));

        if (package.Workbook.Worksheets.Count < sheetIndex)
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出范围");

        var worksheet = package.Workbook.Worksheets[sheetIndex - 1];
        var colCount = worksheet.Dimension?.Columns ?? 0;

        for (int col = 1; col <= colCount; col++)
        {
            var headerValue = worksheet.Cells[1, col].Text.Trim();
            headers.Add(string.IsNullOrEmpty(headerValue) ? $"列{col}" : headerValue);
        }

        return headers;
    }

    /// <summary>
    /// 清洗 Excel 日期格式（将数字日期转换为字符串）
    /// </summary>
    public static string CleanDateValue(object? cellValue)
    {
        if (cellValue == null) return string.Empty;

        // 如果是 DateTime 类型，格式化为字符串
        if (cellValue is DateTime dateTime)
            return dateTime.ToString("yyyy-MM-dd");

        // 如果是数字（Excel 日期序列号），转换为日期
        if (double.TryParse(cellValue.ToString(), out double oaDate))
        {
            try
            {
                var date = DateTime.FromOADate(oaDate);
                return date.ToString("yyyy-MM-dd");
            }
            catch
            {
                return cellValue.ToString() ?? string.Empty;
            }
        }

        return cellValue.ToString() ?? string.Empty;
    }

    /// <summary>
    /// 生成美化的 Excel 模板文件
    /// </summary>
    /// <param name="variables">变量名列表（作为表头）</param>
    /// <param name="filePath">输出文件路径</param>
    public static void GenerateExcelTemplate(List<string> variables, string filePath)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("批量授权数据录入");

        // 样式配置：Office 经典蓝主题
        var headerFont = new Font("微软雅黑", 11, FontStyle.Bold);
        var headerColor = Color.White;
        var headerFillColor = ColorTranslator.FromHtml("#2B579A"); // Office 经典蓝
        var borderColor = ColorTranslator.FromHtml("#D1D5DB");

        // 写入表头（第一行）
        for (int col = 1; col <= variables.Count; col++)
        {
            var cell = worksheet.Cells[1, col];
            cell.Value = variables[col - 1];

            // 表头样式
            cell.Style.Font.Name = "微软雅黑";
            cell.Style.Font.Size = 11;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Color.SetColor(headerColor);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(headerFillColor);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

            // 边框
            cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            cell.Style.Border.Left.Color.SetColor(borderColor);
            cell.Style.Border.Right.Color.SetColor(borderColor);
            cell.Style.Border.Top.Color.SetColor(borderColor);
            cell.Style.Border.Bottom.Color.SetColor(borderColor);

            // 自动调整列宽
            int maxLen = Math.Max(variables[col - 1].Length, 12);
            worksheet.Column(col).Width = maxLen + 5;
        }

        // 写入 4 行空白示例行（带边框引导用户填写）
        for (int row = 2; row <= 5; row++)
        {
            for (int col = 1; col <= variables.Count; col++)
            {
                var cell = worksheet.Cells[row, col];
                cell.Style.Font.Name = "微软雅黑";
                cell.Style.Font.Size = 10;

                // 边框
                cell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                cell.Style.Border.Left.Color.SetColor(borderColor);
                cell.Style.Border.Right.Color.SetColor(borderColor);
                cell.Style.Border.Top.Color.SetColor(borderColor);
                cell.Style.Border.Bottom.Color.SetColor(borderColor);
            }
        }

        // 设置行高
        worksheet.Row(1).Height = 28;
        for (int row = 2; row <= 5; row++)
        {
            worksheet.Row(row).Height = 20;
        }

        // 保存文件
        package.SaveAs(new FileInfo(filePath));
    }

    /// <summary>
    /// 增强版 Excel 读取：支持日期转换、跳过空行、表头验证
    /// </summary>
    /// <param name="filePath">Excel 文件路径</param>
    /// <param name="expectedHeaders">期望的表头列表（用于验证）</param>
    /// <param name="sheetIndex">工作表索引（从 1 开始）</param>
    /// <returns>元组：(数据列表, 缺失的表头列表)</returns>
    public static (List<Dictionary<string, string>> data, List<string> missingHeaders)
        ReadExcelWithValidation(string filePath, List<string>? expectedHeaders = null, int sheetIndex = 1)
    {
        var dataList = new List<Dictionary<string, string>>();
        var missingHeaders = new List<string>();

        using var package = new ExcelPackage(new FileInfo(filePath));

        if (package.Workbook.Worksheets.Count < sheetIndex)
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出范围");

        var worksheet = package.Workbook.Worksheets[sheetIndex - 1];
        var rowCount = worksheet.Dimension?.Rows ?? 0;
        var colCount = worksheet.Dimension?.Columns ?? 0;

        if (rowCount < 2) return (dataList, missingHeaders); // 至少需要表头 + 1 行数据

        // 读取表头（第一行）
        var headers = new List<string>();
        for (int col = 1; col <= colCount; col++)
        {
            var headerValue = worksheet.Cells[1, col].Text.Trim();
            if (!string.IsNullOrEmpty(headerValue))
                headers.Add(headerValue);
        }

        // 验证表头（如果提供了期望表头）
        if (expectedHeaders != null && expectedHeaders.Count > 0)
        {
            missingHeaders = expectedHeaders.Except(headers).ToList();
        }

        // 读取数据行（从第二行开始）
        for (int row = 2; row <= rowCount; row++)
        {
            var rowData = new Dictionary<string, string>();
            bool isEmptyRow = true;

            for (int col = 1; col <= headers.Count; col++)
            {
                var cell = worksheet.Cells[row, col];
                string cellValue;

                // 日期格式转换
                if (cell.Value is DateTime dateTime)
                {
                    // 判断是否包含时间部分
                    if (dateTime.Hour == 0 && dateTime.Minute == 0 && dateTime.Second == 0)
                        cellValue = dateTime.ToString("yyyy-MM-dd");
                    else
                        cellValue = dateTime.ToString("yyyy-MM-dd HH:mm:ss");
                }
                else if (cell.Value is double doubleValue && doubleValue > 1 && doubleValue < 2958466)
                {
                    // Excel 日期序列号转换
                    try
                    {
                        var date = DateTime.FromOADate(doubleValue);
                        cellValue = date.ToString("yyyy-MM-dd");
                    }
                    catch
                    {
                        cellValue = cell.Text.Trim();
                    }
                }
                else
                {
                    cellValue = cell.Text.Trim();
                }

                rowData[headers[col - 1]] = cellValue;

                if (!string.IsNullOrEmpty(cellValue))
                    isEmptyRow = false;
            }

            // 跳过空行
            if (!isEmptyRow)
                dataList.Add(rowData);
        }

        return (dataList, missingHeaders);
    }
}
