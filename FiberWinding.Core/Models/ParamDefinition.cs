using System.Text.Json.Serialization;

namespace FiberWinding.Core.Models;

public sealed class ParamDefinition
{
    /// <summary>
    /// 唯一Key：I{ExcelRow}，例如 I31
    /// </summary>
    public required string Key { get; init; }

    /// <summary>
    /// Excel行号（1-based，与 Ixx 对应）
    /// </summary>
    public required int ExcelRow { get; init; }

    public string Group { get; init; } = "";
    public string Item { get; init; } = "";

    public string NameCn { get; init; } = "";

    /// <summary>
    /// 原Excel“符号”列，用于展示，可空、可重复
    /// </summary>
    public string? DisplaySymbol { get; init; }

    /// <summary>
    /// Input / Intermediate / Output
    /// </summary>
    public string IO { get; init; } = "Intermediate";

    public string Unit { get; init; } = "";

    /// <summary>
    /// 仅对 Input 生效
    /// </summary>
    public bool Required { get; init; }

    /// <summary>
    /// 默认值（如果有）
    /// </summary>
    public double? DefaultValue { get; init; }

    /// <summary>
    /// 类型：number/int/percent/enum... 先预留
    /// </summary>
    public string ValueType { get; init; } = "number";
}