using System.Globalization;
using System.Text.RegularExpressions;
using FiberWinding.Core.Models;
using NCalc;

namespace FiberWinding.AppLogic.Services;

public sealed class ComputeService
{
    private readonly DefinitionExporter _exporter = new();
    private readonly JsonDefinitionLoader _jsonLoader = new();

    // 匹配百分号常量：80% / 0.2% 等
    private static readonly Regex PercentRegex =
        new(@"(\d+(?:\.\d+)?)\s*%", RegexOptions.Compiled);

    // 匹配 I2 / I31 这样的引用（忽略大小写）
    private static readonly Regex RefRegex =
        new(@"(?<![A-Z0-9_])I(\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>
    /// 1) 从Excel导出 params/formulas/cases 三个JSON
    /// </summary>
    public ExportResult ExportJsonFromExcel(string xlsxPath, string outputDir, string sheetName = "")
        => _exporter.Export(xlsxPath, outputDir, sheetName);

    /// <summary>
    /// 2) 从JSON目录计算（不再依赖Excel）
    /// </summary>
    public ComputeResult ComputeFromJsonDir(string defDir, string caseName = "算例1")
    {
        var defs = _jsonLoader.LoadFromDir(defDir);

        var c = defs.Cases.FirstOrDefault(x => x.CaseName == caseName)
                ?? defs.Cases.FirstOrDefault()
                ?? new CaseDefinition { CaseName = "Empty" };

        return ComputeFromDefinitions(defs.Params, defs.Formulas, c.Inputs);
    }

    /// <summary>
    /// 3) 核心：给定 Params/Formulas/Inputs，跑计算链
    /// - Key 永远用 Ixx
    /// - 表达式变量也用 Ixx（例如 I31）
    ///
    /// ✅ 修复：不再简单按 ExcelRow 顺序算，而是按依赖拓扑顺序算
    /// 解决 I67 = I71 + I76 这种“前行依赖后行”导致的 0 值问题
    /// </summary>
    public ComputeResult ComputeFromDefinitions(
        IReadOnlyList<ParamDefinition> paramDefs,
        IReadOnlyList<FormulaDefinition> formulas,
        IReadOnlyDictionary<string, double> inputs
    )
    {
        // row -> value（以 ExcelRow 为索引）
        var valuesByRow = new Dictionary<int, double>();
        var warnings = new List<string>();

        // 索引
        var paramByRow = paramDefs.ToDictionary(x => x.ExcelRow, x => x);
        var formulaByRow = formulas.ToDictionary(x => x.ExcelRow, x => x);

        // 先把所有 Input 写入（输入优先）
        foreach (var p in paramDefs.Where(p => string.Equals(p.IO, "Input", StringComparison.OrdinalIgnoreCase)))
        {
            if (inputs.TryGetValue(p.Key, out var v))
                valuesByRow[p.ExcelRow] = v;
            else if (p.DefaultValue.HasValue)
                valuesByRow[p.ExcelRow] = p.DefaultValue.Value;
            else if (p.Required)
                warnings.Add($"缺少必填输入：{p.NameCn}");
        }

        // ✅ 构建依赖图（基于 ExprNormalized 里的 Ixx 引用）
        var nodeRows = new HashSet<int>(formulas.Select(f => f.ExcelRow));
        var deps = new Dictionary<int, HashSet<int>>();
        foreach (var r in nodeRows) deps[r] = new HashSet<int>();

        foreach (var f in formulas)
        {
            var row = f.ExcelRow;

            // Input 行不需要计算
            if (paramByRow.TryGetValue(row, out var p) &&
                string.Equals(p.IO, "Input", StringComparison.OrdinalIgnoreCase))
                continue;

            if (string.IsNullOrWhiteSpace(f.ExprNormalized))
                continue;

            foreach (Match m in RefRegex.Matches(f.ExprNormalized))
            {
                if (!int.TryParse(m.Groups[1].Value, out var depRow))
                    continue;

                if (depRow == row) continue;

                // 依赖只指向“表里存在的行”
                // 这里不要求 depRow 也一定是公式行：Input/常量行也允许
                // 但拓扑排序只对公式节点排序，Input/常量由 valuesByRow 提供
                if (nodeRows.Contains(depRow))
                    deps[row].Add(depRow);
            }
        }

        // ✅ 拓扑排序（在公式节点集合上）
        var order = TopoSort(nodeRows, deps);

        // ✅ 按拓扑顺序计算（保留你原来的预处理和函数注入）
        foreach (var row in order)
        {
            if (!formulaByRow.TryGetValue(row, out var f))
                continue;

            paramByRow.TryGetValue(row, out var p);

            // Input 行不覆盖
            if (p != null && string.Equals(p.IO, "Input", StringComparison.OrdinalIgnoreCase))
                continue;

            // ExprNormalized 为空：尝试把 ExprRaw 当数值
            if (string.IsNullOrWhiteSpace(f.ExprNormalized))
            {
                if (TryParseDouble(f.ExprRaw, out var num))
                    valuesByRow[row] = num;
                continue;
            }

            // 计算前归一化（% / ^）
            var exprText = PreNormalizeForNCalc(f.ExprNormalized);

            try
            {
                var expr = new Expression(exprText, EvaluateOptions.IgnoreCase);

                // 变量注入：I2、I31...
                expr.EvaluateParameter += (name, args) =>
                {
                    if (string.IsNullOrWhiteSpace(name))
                        return;

                    if (name.Length >= 2 && (name[0] == 'I' || name[0] == 'i'))
                    {
                        if (int.TryParse(name[1..], out var depRow))
                        {
                            if (valuesByRow.TryGetValue(depRow, out var vv))
                                args.Result = vv;
                            else
                                args.Result = 0.0; // 未定义默认0（开发期策略）
                        }
                    }
                };

                // 函数注入：PI/SIN/COS/TAN/ASIN/ACOS/ATAN/ABS/MAX/MIN/POW
                expr.EvaluateFunction += (name, args) =>
                {
                    var n = (name ?? "").ToUpperInvariant();

                    if (n == "PI")
                    {
                        args.Result = Math.PI;
                        return;
                    }

                    if (n == "POW")
                    {
                        var a = Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture);
                        var b = Convert.ToDouble(args.Parameters[1].Evaluate(), CultureInfo.InvariantCulture);
                        args.Result = Math.Pow(a, b);
                        return;
                    }

                    if (n == "COS")
                    {
                        args.Result = Math.Cos(Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture));
                        return;
                    }

                    if (n == "SIN")
                    {
                        args.Result = Math.Sin(Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture));
                        return;
                    }

                    if (n == "TAN")
                    {
                        args.Result = Math.Tan(Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture));
                        return;
                    }

                    if (n == "ACOS")
                    {
                        args.Result = Math.Acos(Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture));
                        return;
                    }

                    if (n == "ASIN")
                    {
                        args.Result = Math.Asin(Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture));
                        return;
                    }

                    if (n == "ATAN")
                    {
                        args.Result = Math.Atan(Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture));
                        return;
                    }

                    if (n == "ABS")
                    {
                        args.Result = Math.Abs(Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture));
                        return;
                    }

                    if (n == "MAX")
                    {
                        var vals = args.Parameters
                            .Select(p2 => Convert.ToDouble(p2.Evaluate(), CultureInfo.InvariantCulture))
                            .ToArray();
                        args.Result = vals.Max();
                        return;
                    }

                    if (n == "MIN")
                    {
                        var vals = args.Parameters
                            .Select(p2 => Convert.ToDouble(p2.Evaluate(), CultureInfo.InvariantCulture))
                            .ToArray();
                        args.Result = vals.Min();
                        return;
                    }

                    // 等价实现 Excel 公式：
                    // --LEFT(SUBSTITUTE(x, ".", ""), 1) + 1
                    // 用于 I52/I53/I55/I56 等“取首位数字+1”的层数规则
                    if (n is "FIRSTDIGITPLUS1" or "FIRSTDIGIT_PLUS1")
                    {
                        var x = Convert.ToDouble(args.Parameters[0].Evaluate(), CultureInfo.InvariantCulture);
                        args.Result = FirstDigitPlus1(x);
                        return;
                    }
                };

                var result = expr.Evaluate();
                var dv = Convert.ToDouble(result, CultureInfo.InvariantCulture);
                valuesByRow[row] = dv;
            }
            catch (Exception ex)
            {
                // 不在 UI 中暴露单元格 Key（Ixx），避免干扰使用者。
                warnings.Add($"公式计算失败：{f.ExprRaw}；原因：{ex.Message}");
            }
        }

        return new ComputeResult(valuesByRow, warnings);
    }

    private static double FirstDigitPlus1(double x)
    {
        // 你的 Excel 写法会先去掉小数点、取第 1 位字符再转数字。
        // 这里做一个稳健实现：
        // 1) 取绝对值（避免负号影响首字符）
        // 2) 用不带科学计数法的格式输出（尽量保留有效数字）
        // 3) 去掉小数点
        // 4) 找到第 1 个数字字符并 +1；找不到则返回 1
        var s = Math.Abs(x).ToString("0.############################", CultureInfo.InvariantCulture);
        s = s.Replace(".", "");

        foreach (var ch in s)
        {
            if (ch is >= '0' and <= '9')
                return (ch - '0') + 1;
        }

        return 1;
    }

    private static bool TryParseDouble(string s, out double v)
    {
        s = (s ?? "").Trim();
        if (s.StartsWith("=")) s = s[1..].Trim();

        return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out v)
               || double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out v);
    }

    /// <summary>
    /// ✅ 计算前归一化：
    /// 1) 去掉分号
    /// 2) 80% -> (80/100)
    /// 3) a^b -> POW(a,b)
    /// </summary>
    private static string PreNormalizeForNCalc(string expr)
    {
        var s = (expr ?? "").Trim();

        s = s.Replace(";", "");

        s = PercentRegex.Replace(s, m => $"({m.Groups[1].Value}/100)");

        s = ReplacePowerWithPow(s);

        return s;
    }

    private static string ReplacePowerWithPow(string s)
    {
        int i = 0;
        while (i < s.Length)
        {
            if (s[i] != '^') { i++; continue; }

            int expStart = i + 1;
            while (expStart < s.Length && char.IsWhiteSpace(s[expStart])) expStart++;

            int expEnd = expStart;
            bool hasDigit = false;
            while (expEnd < s.Length && (char.IsDigit(s[expEnd]) || s[expEnd] == '.'))
            {
                hasDigit = true;
                expEnd++;
            }

            if (!hasDigit) { i++; continue; }

            string exp = s.Substring(expStart, expEnd - expStart);

            int baseEnd = i - 1;
            while (baseEnd >= 0 && char.IsWhiteSpace(s[baseEnd])) baseEnd--;
            if (baseEnd < 0) { i++; continue; }

            int baseStart;

            if (s[baseEnd] == ')')
            {
                int depth = 0;
                int j = baseEnd;
                for (; j >= 0; j--)
                {
                    if (s[j] == ')') depth++;
                    else if (s[j] == '(')
                    {
                        depth--;
                        if (depth == 0) break;
                    }
                }
                if (j < 0) { i++; continue; }

                baseStart = j;

                int k = baseStart - 1;
                while (k >= 0 && char.IsWhiteSpace(s[k])) k--;

                if (k >= 0 && (char.IsLetter(s[k]) || s[k] == '_'))
                {
                    int funcEnd = k;
                    int funcStart = funcEnd;
                    while (funcStart >= 0 && (char.IsLetterOrDigit(s[funcStart]) || s[funcStart] == '_'))
                        funcStart--;
                    funcStart++;
                    baseStart = funcStart;
                }
            }
            else
            {
                int j = baseEnd;
                while (j >= 0 && (char.IsLetterOrDigit(s[j]) || s[j] == '_' || s[j] == '.'))
                    j--;
                baseStart = j + 1;
            }

            if (baseStart > baseEnd) { i++; continue; }

            string bas = s.Substring(baseStart, baseEnd - baseStart + 1);

            string left = s.Substring(0, baseStart);
            string right = s.Substring(expEnd);
            string mid = $"POW({bas},{exp})";
            s = left + mid + right;

            i = left.Length + mid.Length;
        }

        return s;
    }

    private static List<int> TopoSort(IEnumerable<int> nodes, Dictionary<int, HashSet<int>> deps)
    {
        var nodeList = nodes.Distinct().ToList();
        var inDeg = nodeList.ToDictionary(n => n, _ => 0);

        foreach (var (n, ds) in deps)
        {
            if (!inDeg.ContainsKey(n)) continue;
            foreach (var _ in ds) inDeg[n]++;
        }

        var users = nodeList.ToDictionary(n => n, _ => new List<int>());
        foreach (var (n, ds) in deps)
        {
            foreach (var d in ds)
            {
                if (users.TryGetValue(d, out var list))
                    list.Add(n);
            }
        }

        var q = new Queue<int>(inDeg.Where(kv => kv.Value == 0).Select(kv => kv.Key));
        var res = new List<int>(capacity: nodeList.Count);

        while (q.Count > 0)
        {
            var x = q.Dequeue();
            res.Add(x);

            foreach (var u in users[x])
            {
                inDeg[u]--;
                if (inDeg[u] == 0) q.Enqueue(u);
            }
        }

        if (res.Count != nodeList.Count)
            throw new InvalidOperationException("依赖图存在循环，无法拓扑排序（请检查 formulas 是否有循环引用）。");

        return res;
    }
}

public sealed record ComputeResult(
    IReadOnlyDictionary<int, double> ValuesByRow,
    IReadOnlyList<string> Warnings
);