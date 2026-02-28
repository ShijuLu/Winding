using System.Globalization;
using System.Text.Json;
using ClosedXML.Excel;
using FiberWinding.Core.Models;

namespace FiberWinding.AppLogic.Services;

public sealed class DefinitionExporter
{
    // 你那张表的列名（按你截图）
    private const string ColGroup = "Group";
    private const string ColItem = "Item";
    private const string ColRequired = "必填";
    private const string ColIO = "IO";
    private const string ColNameCn = "参数名";
    private const string ColSymbol = "符号";
    private const string ColDefault = "默认值";
    private const string ColExpr = "函数关系";
    private const string ColUnit = "单位";

    public ExportResult Export(string xlsxPath, string outputDir, string sheetName = "")
    {
        Directory.CreateDirectory(outputDir);

        using var wb = new XLWorkbook(xlsxPath);

        var ws = string.IsNullOrWhiteSpace(sheetName)
            ? wb.Worksheets.First()
            : wb.Worksheet(sheetName);

        var headerRow = FindHeaderRow(ws);
        var colMap = BuildColumnMap(ws, headerRow);

        var lastRow = ws.LastRowUsed().RowNumber();

        var paramDefs = new List<ParamDefinition>(capacity: 256);
        var formulas = new List<FormulaDefinition>(capacity: 256);

        // 默认算例：把 Input 行当前显示的数值当作输入
        var defaultCase = new CaseDefinition { CaseName = "算例1" };

        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            var nameCn = GetCellString(ws, r, colMap, ColNameCn);
            if (string.IsNullOrWhiteSpace(nameCn))
                continue;

            var group = GetCellString(ws, r, colMap, ColGroup);
            var item = GetCellString(ws, r, colMap, ColItem);
            var io = NormalizeIO(GetCellString(ws, r, colMap, ColIO));
            var unit = GetCellString(ws, r, colMap, ColUnit);
            var symbol = GetCellString(ws, r, colMap, ColSymbol);

            // ✅ 必填只对 Input 生效
            var required = ParseYesNo(GetCellString(ws, r, colMap, ColRequired)) && io == "Input";

            var defaultVal = TryParseDouble(GetCellString(ws, r, colMap, ColDefault), out var dv) ? dv : (double?)null;

            var key = $"I{r}";

            var p = new ParamDefinition
            {
                Key = key,
                ExcelRow = r,
                Group = group,
                Item = item,
                NameCn = nameCn.Trim(),
                DisplaySymbol = string.IsNullOrWhiteSpace(symbol) ? null : symbol.Trim(),
                IO = io,
                Unit = (unit ?? "").Trim(),
                Required = required,
                DefaultValue = defaultVal,
                ValueType = "number"
            };
            paramDefs.Add(p);

            // =========================
            // ✅ 关键修复：导出“公式文本”，而不是显示值
            // =========================
            var exprCell = GetCell(ws, r, colMap, ColExpr);

            // 如果单元格有公式：返回 "="+FormulaA1；否则返回 GetString
            var exprRaw = GetFormulaOrString(exprCell);
            var exprNorm = NormalizeExpr(exprRaw);

            formulas.Add(new FormulaDefinition
            {
                Key = key,
                ExcelRow = r,
                ExprRaw = exprRaw,
                ExprNormalized = exprNorm
            });

            // =========================
            // ✅ 算例生成：Input 行读“数值”（缓存值/显示值）
            // =========================
            if (io == "Input")
            {
                if (TryReadCellAsDouble(exprCell, out var vFromCell))
                {
                    defaultCase.Inputs[key] = vFromCell;
                }
                else
                {
                    if (TryParseDouble(exprRaw, out var vFromExpr))
                        defaultCase.Inputs[key] = vFromExpr;
                    else if (defaultVal.HasValue)
                        defaultCase.Inputs[key] = defaultVal.Value;
                }
            }
        }

        var cases = new List<CaseDefinition> { defaultCase };

        var opt = new JsonSerializerOptions { WriteIndented = true };

        var paramsPath = Path.Combine(outputDir, "params.json");
        var formulasPath = Path.Combine(outputDir, "formulas.json");
        var casesPath = Path.Combine(outputDir, "cases.json");

        File.WriteAllText(paramsPath, JsonSerializer.Serialize(paramDefs, opt));
        File.WriteAllText(formulasPath, JsonSerializer.Serialize(formulas, opt));
        File.WriteAllText(casesPath, JsonSerializer.Serialize(cases, opt));

        return new ExportResult(paramsPath, formulasPath, casesPath, paramDefs.Count);
    }

    private static int FindHeaderRow(IXLWorksheet ws)
    {
        for (int r = 1; r <= Math.Min(50, ws.LastRowUsed().RowNumber()); r++)
        {
            var row = ws.Row(r);
            var cells = row.CellsUsed().Select(c => c.GetString().Trim()).ToList();
            if (cells.Contains(ColGroup) && cells.Contains(ColNameCn))
                return r;
        }
        throw new InvalidOperationException("未找到表头行：请确认Excel里包含列名 Group / 参数名 等。");
    }

    private static Dictionary<string, int> BuildColumnMap(IXLWorksheet ws, int headerRow)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        var row = ws.Row(headerRow);
        foreach (var cell in row.CellsUsed())
        {
            var name = cell.GetString().Trim();
            if (string.IsNullOrWhiteSpace(name)) continue;
            if (!map.ContainsKey(name))
                map[name] = cell.Address.ColumnNumber;
        }

        var required = new[] { ColGroup, ColItem, ColIO, ColNameCn, ColExpr };
        foreach (var col in required)
        {
            if (!map.ContainsKey(col))
                throw new InvalidOperationException($"缺少列：{col}");
        }

        return map;
    }

    private static IXLCell GetCell(IXLWorksheet ws, int row, Dictionary<string, int> colMap, string colName)
    {
        if (!colMap.TryGetValue(colName, out var c))
            return ws.Cell(row, 1);
        return ws.Cell(row, c);
    }

    private static string GetCellString(IXLWorksheet ws, int row, Dictionary<string, int> colMap, string colName)
    {
        if (!colMap.TryGetValue(colName, out var c))
            return "";
        return ws.Cell(row, c).GetString().Trim();
    }

    private static string GetFormulaOrString(IXLCell cell)
    {
        if (cell.HasFormula && !string.IsNullOrWhiteSpace(cell.FormulaA1))
            return "=" + cell.FormulaA1.Trim();

        return (cell.GetString() ?? "").Trim();
    }

    private static bool TryReadCellAsDouble(IXLCell cell, out double v)
    {
        // 1) 优先用 GetDouble（对数值/可转换数值最稳）
        try
        {
            v = cell.GetDouble();
            return true;
        }
        catch
        {
            // ignore
        }

        // 2) 用 Value.ToString()（ClosedXML 的 Value 是 struct，不能用 ?.）
        var txt = cell.Value.ToString().Trim();
        return TryParseDouble(txt, out v);
    }

    private static string NormalizeIO(string io)
    {
        io = (io ?? "").Trim();
        if (io.Equals("Input", StringComparison.OrdinalIgnoreCase)) return "Input";
        if (io.Equals("Output", StringComparison.OrdinalIgnoreCase)) return "Output";
        return "Intermediate";
    }

    private static bool ParseYesNo(string s)
    {
        s = (s ?? "").Trim();
        return s == "是"
               || s.Equals("yes", StringComparison.OrdinalIgnoreCase)
               || s == "Y"
               || s == "1"
               || s.Equals("true", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParseDouble(string s, out double v)
    {
        s = (s ?? "").Trim();
        if (s.StartsWith("=")) s = s[1..].Trim();

        return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v)
               || double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out v);
    }

    private static string NormalizeExpr(string exprRaw)
    {
        if (string.IsNullOrWhiteSpace(exprRaw))
            return "";

        var s = exprRaw.Trim();

        // 纯数字不作为表达式
        if (TryParseDouble(s, out _))
            return "";

        if (s.StartsWith("="))
            s = s[1..].Trim();

        s = s.Replace("（", "(").Replace("）", ")");

        s = RewritePowFunc(s, "SINPOW", "SIN");
        s = RewritePowFunc(s, "COSPOW", "COS");

        return s;
    }

    private static string RewritePowFunc(string input, string powFunc, string baseFunc)
    {
        while (true)
        {
            var idx = input.IndexOf(powFunc + "(", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) break;

            var start = idx + powFunc.Length + 1;
            int depth = 1;
            int i = start;
            for (; i < input.Length; i++)
            {
                if (input[i] == '(') depth++;
                else if (input[i] == ')')
                {
                    depth--;
                    if (depth == 0) break;
                }
            }
            if (depth != 0) break;

            var inside = input.Substring(start, i - start);
            var comma = FindTopLevelComma(inside);
            if (comma < 0) break;

            var arg1 = inside[..comma].Trim();
            var arg2 = inside[(comma + 1)..].Trim();

            var replaced = $"(({baseFunc}({arg1}))^{arg2})";
            input = input.Substring(0, idx) + replaced + input.Substring(i + 1);
        }

        return input;
    }

    private static int FindTopLevelComma(string s)
    {
        int depth = 0;
        for (int i = 0; i < s.Length; i++)
        {
            if (s[i] == '(') depth++;
            else if (s[i] == ')') depth--;
            else if (s[i] == ',' && depth == 0) return i;
        }
        return -1;
    }
}

public readonly record struct ExportResult(
    string ParamsJsonPath,
    string FormulasJsonPath,
    string CasesJsonPath,
    int RowCount
);