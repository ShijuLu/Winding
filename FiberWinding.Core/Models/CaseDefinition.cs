namespace FiberWinding.Core.Models;

public sealed class CaseDefinition
{
    public string CaseName { get; init; } = "Default";

    /// <summary>
    /// 只存 Input 的值：Key(Ixx) -> value
    /// </summary>
    public Dictionary<string, double> Inputs { get; init; } = new();
}