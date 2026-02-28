namespace FiberWinding.Core.Models;

public sealed class FormulaDefinition
{
    /// <summary>
    /// 对应 ParamDefinition.Key（Ixx）
    /// </summary>
    public required string Key { get; init; }

    public required int ExcelRow { get; init; }

    /// <summary>
    /// 原始表达式：可能是数值，也可能以 "=" 开头
    /// </summary>
    public string ExprRaw { get; init; } = "";

    /// <summary>
    /// 规范化后表达式（用于 NCalc），可能为空（表示直接用数值）
    /// </summary>
    public string ExprNormalized { get; init; } = "";
}