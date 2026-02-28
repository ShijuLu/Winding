namespace FiberWinding.Core.Models;

public enum IoKind
{
    Input,
    Intermediate,
    Output
}

public sealed class ParameterRow
{
    // Excel 行号（从 2 开始，因为第1行是表头）
    public int ExcelRow { get; init; }

    public string Group { get; init; } = "";
    public string Item { get; init; } = "";
    public string RequiredRaw { get; init; } = "";  // 原表“必填”
    public IoKind Io { get; init; }

    public string NameCn { get; init; } = "";
    public string Symbol { get; init; } = "";
    public string TypeRaw { get; init; } = "";      // 原表“类型”
    public string DefaultRaw { get; init; } = "";   // 原表“默认值”
    public string ExprRaw { get; init; } = "";      // 原表“函数关系”列：可能是值，也可能是公式
    public string Unit { get; init; } = "";
    public string LookupRaw { get; init; } = "";    // 原表“是否查表”
}