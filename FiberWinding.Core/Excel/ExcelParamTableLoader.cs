using ClosedXML.Excel;
using FiberWinding.Core.Models;

namespace FiberWinding.Core.Excel;

public sealed class ExcelParamTableLoader
{
    public IReadOnlyList<ParameterRow> Load(string xlsxPath, string sheetName = "Sheet1")
    {
        using var wb = new XLWorkbook(xlsxPath);
        var ws = wb.Worksheet(sheetName);

        // 找到最后一行（按第一列有内容判断）
        var lastRow = ws.LastRowUsed().RowNumber();
        var rows = new List<ParameterRow>();

        // 从第2行开始读（第1行是表头）
        for (int r = 2; r <= lastRow; r++)
        {
            var group = ws.Cell(r, 1).GetString().Trim(); // A Group
            if (string.IsNullOrWhiteSpace(group))
                continue; // 空行跳过

            var ioRaw = ws.Cell(r, 4).GetString().Trim(); // D IO
            if (!TryParseIo(ioRaw, out var io))
                throw new InvalidOperationException($"无法识别 IO：行 {r}, 值='{ioRaw}'");

            rows.Add(new ParameterRow
            {
                ExcelRow = r,
                Group = group,
                Item = ws.Cell(r, 2).GetString().Trim(),          // B Item
                RequiredRaw = ws.Cell(r, 3).GetString().Trim(),   // C 必填
                Io = io,
                NameCn = ws.Cell(r, 5).GetString().Trim(),        // E 参数名
                Symbol = ws.Cell(r, 6).GetString().Trim(),        // F 符号
                TypeRaw = ws.Cell(r, 7).GetString().Trim(),       // G 类型
                DefaultRaw = ws.Cell(r, 8).GetString().Trim(),    // H 默认值
                ExprRaw = GetExprOrValue(ws.Cell(r, 9)).Trim(),   // I 函数关系（公式或值）
                Unit = ws.Cell(r, 10).GetString().Trim(),         // J 单位
                LookupRaw = ws.Cell(r, 11).GetString().Trim(),    // K 是否查表
            });
        }

        return rows;
    }

    private static string GetExprOrValue(IXLCell cell)
    {
        // 如果单元格有公式，取公式字符串（ClosedXML 里 FormulaA1 不含前导 '='，这里我们补上）
        if (!string.IsNullOrWhiteSpace(cell.FormulaA1))
            return "=" + cell.FormulaA1;

        // 否则取显示值（注意百分号会被显示成 0.2 或 20% 取决于原始格式）
        // 这里取文本形式，后面表达式引擎会做兼容处理
        return cell.GetFormattedString();
    }

    private static bool TryParseIo(string raw, out IoKind io)
    {
        raw = raw.Trim();
        if (raw.Equals("Input", StringComparison.OrdinalIgnoreCase)) { io = IoKind.Input; return true; }
        if (raw.Equals("Intermediate", StringComparison.OrdinalIgnoreCase)) { io = IoKind.Intermediate; return true; }
        if (raw.Equals("Output", StringComparison.OrdinalIgnoreCase)) { io = IoKind.Output; return true; }

        io = default;
        return false;
    }
}