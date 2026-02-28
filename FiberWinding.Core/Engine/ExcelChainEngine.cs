using System.Globalization;
using System.Text.RegularExpressions;
using FiberWinding.Core.Models;
using NCalc;

namespace FiberWinding.Core.Engine;

public sealed class ExcelChainEngine
{
    // 匹配 I2 / I31 这样的引用（不允许前面紧跟字母，避免匹配到别的单词中间）
    private static readonly Regex RefRegex =
        new(@"(?<![A-Z])I(\d+)", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    // 匹配百分号常量：80% / 0.2% 等
    private static readonly Regex PercentRegex =
        new(@"(\d+(?:\.\d+)?)\s*%", RegexOptions.Compiled);

    public sealed class ComputeResult
    {
        public Dictionary<int, double> ValuesByRow { get; } = new(); // key=ExcelRow（I列对应的行号）
        public List<string> Warnings { get; } = new();
    }

    public ComputeResult Compute(IReadOnlyList<ParameterRow> table)
    {
        // 1) 每行都有一个节点 I{row}
        var exprByRow = table.ToDictionary(x => x.ExcelRow, x => x.ExprRaw);

        // 2) 依赖图（row -> deps）
        var deps = new Dictionary<int, HashSet<int>>();
        foreach (var row in table)
        {
            deps[row.ExcelRow] = new HashSet<int>();
            var expr = exprByRow[row.ExcelRow];

            if (IsFormula(expr))
            {
                foreach (Match m in RefRegex.Matches(expr))
                {
                    if (!int.TryParse(m.Groups[1].Value, out int depRow))
                        continue;

                    if (exprByRow.ContainsKey(depRow) && depRow != row.ExcelRow)
                        deps[row.ExcelRow].Add(depRow);
                }
            }
        }

        // 3) 拓扑排序
        var order = TopoSort(exprByRow.Keys, deps);

        // 4) 逐个求值
        var result = new ComputeResult();

        foreach (var rowNum in order)
        {
            var expr = exprByRow[rowNum];

            if (!IsFormula(expr))
            {
                // Input 或“无公式”的情况：当成数值解析
                if (TryParseNumber(expr, out var v))
                {
                    result.ValuesByRow[rowNum] = v;
                }
                else
                {
                    result.ValuesByRow[rowNum] = double.NaN;
                    result.Warnings.Add($"I{rowNum} 无法解析为数值：'{expr}'");
                }
                continue;
            }

            // 去掉前导 '='，并做归一化（百分号、分号、幂）
            var formula = expr.TrimStart().Substring(1);
            formula = NormalizeFormula(formula);

            var e = new Expression(formula, EvaluateOptions.IgnoreCase);

            // 注入函数与常量
            e.EvaluateFunction += (name, args) =>
            {
                var n = name.ToUpperInvariant();

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

                // 三角函数（弧度制）
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
                        .Select(p => Convert.ToDouble(p.Evaluate(), CultureInfo.InvariantCulture))
                        .ToArray();
                    args.Result = vals.Max();
                    return;
                }

                if (n == "MIN")
                {
                    var vals = args.Parameters
                        .Select(p => Convert.ToDouble(p.Evaluate(), CultureInfo.InvariantCulture))
                        .ToArray();
                    args.Result = vals.Min();
                    return;
                }
            };

            // 注入变量：I2、I3...
            e.EvaluateParameter += (name, args) =>
            {
                var m = RefRegex.Match(name);
                if (m.Success && int.TryParse(m.Groups[1].Value, out int depRow))
                {
                    if (result.ValuesByRow.TryGetValue(depRow, out var v))
                    {
                        args.Result = v;
                        return;
                    }
                }

                // 未定义变量默认 0
                args.Result = 0.0;
            };

            try
            {
                var raw = e.Evaluate();
                var v = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                result.ValuesByRow[rowNum] = v;
            }
            catch (Exception ex)
            {
                result.ValuesByRow[rowNum] = double.NaN;
                result.Warnings.Add($"I{rowNum} 公式计算失败：={formula}；原因：{ex.Message}");
            }
        }

        return result;
    }

    private static bool IsFormula(string s) => s.TrimStart().StartsWith("=");

    /// <summary>
    /// Excel 表达式 -> NCalc 友好表达式
    /// 1) 去分号
    /// 2) 80% -> (80/100)
    /// 3) a^b -> POW(a,b)（支持嵌套括号底数，支持 SIN(x)^2 这种）
    /// </summary>
    private static string NormalizeFormula(string formula)
    {
        var s = formula.Trim();

        // 1) 去掉分号
        s = s.Replace(";", "");

        // 2) 百分号常量：80% -> (80/100)
        s = PercentRegex.Replace(s, m => $"({m.Groups[1].Value}/100)");

        // 3) 幂：a^b -> POW(a,b)
        s = ReplacePowerWithPow(s);

        return s;
    }

    /// <summary>
    /// 把 Excel 的 a^b 替换为 POW(a,b)，支持 a 为嵌套括号表达式。
    /// 修复：支持 SIN(x)^2 / COS(x)^2，输出 POW(SIN(x),2) 而不是 SINPOW(...)
    /// 仅处理指数 b 为数字（你这份表主要是 ^2）。
    /// </summary>
    private static string ReplacePowerWithPow(string s)
    {
        int i = 0;
        while (i < s.Length)
        {
            if (s[i] != '^') { i++; continue; }

            // 读指数（右侧）
            int expStart = i + 1;
            while (expStart < s.Length && char.IsWhiteSpace(s[expStart])) expStart++;

            int expEnd = expStart;
            bool hasDigit = false;
            while (expEnd < s.Length && (char.IsDigit(s[expEnd]) || s[expEnd] == '.'))
            {
                hasDigit = true;
                expEnd++;
            }

            if (!hasDigit)
            {
                i++;
                continue;
            }

            string exp = s.Substring(expStart, expEnd - expStart);

            // 读底数（左侧）
            int baseEnd = i - 1;
            while (baseEnd >= 0 && char.IsWhiteSpace(s[baseEnd])) baseEnd--;
            if (baseEnd < 0) { i++; continue; }

            int baseStart;

            if (s[baseEnd] == ')')
            {
                // 找匹配的 '('（支持嵌套）
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

                // ★关键修复：如果 '(' 左边紧挨着是函数名（如 SIN/COS），把函数名也包含进底数
                int k = baseStart - 1;
                while (k >= 0 && char.IsWhiteSpace(s[k])) k--;

                if (k >= 0 && (char.IsLetter(s[k]) || s[k] == '_'))
                {
                    int funcEnd = k;
                    int funcStart = funcEnd;
                    while (funcStart >= 0 && (char.IsLetterOrDigit(s[funcStart]) || s[funcStart] == '_'))
                        funcStart--;
                    funcStart++;

                    // 只有在函数名后面确实紧跟 '(' 才算函数调用
                    // 这里 baseStart 是 '('，funcEnd 在 '(' 左侧，所以成立
                    baseStart = funcStart;
                }
            }
            else
            {
                // token：字母/数字/下划线/点
                int j = baseEnd;
                while (j >= 0 && (char.IsLetterOrDigit(s[j]) || s[j] == '_' || s[j] == '.'))
                    j--;
                baseStart = j + 1;
            }

            if (baseStart > baseEnd) { i++; continue; }

            string bas = s.Substring(baseStart, baseEnd - baseStart + 1);

            // 替换 bas^exp -> POW(bas,exp)
            string left = s.Substring(0, baseStart);
            string right = s.Substring(expEnd);
            string mid = $"POW({bas},{exp})";
            s = left + mid + right;

            i = left.Length + mid.Length;
        }

        return s;
    }

    private static bool TryParseNumber(string raw, out double value)
    {
        raw = raw.Trim();
        if (string.IsNullOrWhiteSpace(raw))
        {
            value = double.NaN;
            return false;
        }

        // 支持 "20%" / "0.2%"
        if (raw.EndsWith("%", StringComparison.Ordinal))
        {
            var n = raw.TrimEnd('%').Trim();
            if (double.TryParse(n, NumberStyles.Any, CultureInfo.InvariantCulture, out var p))
            {
                value = p / 100.0;
                return true;
            }

            if (double.TryParse(n, NumberStyles.Any, CultureInfo.CurrentCulture, out p))
            {
                value = p / 100.0;
                return true;
            }
        }

        if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
            return true;

        if (double.TryParse(raw, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
            return true;

        value = double.NaN;
        return false;
    }

    private static List<int> TopoSort(IEnumerable<int> nodes, Dictionary<int, HashSet<int>> deps)
    {
        var inDeg = nodes.ToDictionary(n => n, _ => 0);

        foreach (var (n, ds) in deps)
        {
            foreach (var _ in ds)
                if (inDeg.ContainsKey(n))
                    inDeg[n]++;
        }

        var q = new Queue<int>(inDeg.Where(kv => kv.Value == 0).Select(kv => kv.Key));
        var res = new List<int>();

        var users = nodes.ToDictionary(n => n, _ => new List<int>());
        foreach (var (n, ds) in deps)
            foreach (var d in ds)
                users[d].Add(n);

        while (q.Count > 0)
        {
            var x = q.Dequeue();
            res.Add(x);

            foreach (var u in users[x])
            {
                inDeg[u]--;
                if (inDeg[u] == 0)
                    q.Enqueue(u);
            }
        }

        if (res.Count != inDeg.Count)
            throw new InvalidOperationException("依赖图存在循环或缺失节点，无法拓扑排序。");

        return res;
    }
}