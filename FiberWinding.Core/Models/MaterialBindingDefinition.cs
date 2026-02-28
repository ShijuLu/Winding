using System.Text.Json.Serialization;

namespace FiberWinding.Core.Models;

/// <summary>
/// 将某个 Input 参数 Key 绑定到某个材料库，并指定从材料条目中取哪个 valueKey。
/// </summary>
public sealed class MaterialBindingDefinition
{
    [JsonPropertyName("paramKey")]
    public string ParamKey { get; set; } = "";

    [JsonPropertyName("library")]
    public string Library { get; set; } = "";

    /// <summary>
    /// 取值字段名；为空时默认等于 ParamKey。
    /// </summary>
    [JsonPropertyName("valueKey")]
    public string? ValueKey { get; set; }
}
