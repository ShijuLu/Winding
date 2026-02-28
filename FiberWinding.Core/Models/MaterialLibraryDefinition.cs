using System.Text.Json.Serialization;

namespace FiberWinding.Core.Models;

/// <summary>
/// 内置材料库：一个库包含若干材料，每个材料可提供多个参数值（按参数 Key 映射）。
/// </summary>
public sealed class MaterialLibraryDefinition
{
    [JsonPropertyName("name")]
    public string Name { get; set; } = "";

    [JsonPropertyName("items")]
    public List<MaterialItemDefinition> Items { get; set; } = new();
}

public sealed class MaterialItemDefinition
{
    [JsonPropertyName("material")]
    public string Material { get; set; } = "";

    /// <summary>
    /// 参数值：Key => 数值。Key 使用 Ixx（推荐）或业务字段名（由绑定决定）。
    /// </summary>
    [JsonPropertyName("values")]
    public Dictionary<string, double> Values { get; set; } = new();
}
