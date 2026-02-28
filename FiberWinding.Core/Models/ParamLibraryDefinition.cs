using System.Text.Json.Serialization;

namespace FiberWinding.Core.Models;

/// <summary>
/// B 方案：按“参数”建立材料/取值库。
/// 一个参数 Key(Ixx) 对应一组可选项（名称 -> 数值）。
/// </summary>
public sealed class ParamLibraryDefinition
{
    /// <summary>
    /// 对应参数 Key，例如 "I11"。
    /// </summary>
    public string ParamKey { get; set; } = "";

    /// <summary>
    /// 可选项列表。
    /// </summary>
    public List<ParamLibraryItem> Items { get; set; } = new();
}

public sealed class ParamLibraryItem
{
    /// <summary>
    /// 下拉显示名称，例如 "T700"。
    /// </summary>
    public string Name { get; set; } = "";

    /// <summary>
    /// 对应数值。
    /// </summary>
    public double Value { get; set; }
}
