using System.Text.Json;
using FiberWinding.Core.Models;

namespace FiberWinding.AppLogic.Services;

public sealed class JsonDefinitionLoader
{
    public LoadedDefinitions LoadFromDir(string defDir)
    {
        var paramsPath = Path.Combine(defDir, "params.json");
        var formulasPath = Path.Combine(defDir, "formulas.json");
        var casesPath = Path.Combine(defDir, "cases.json");

        if (!File.Exists(paramsPath)) throw new FileNotFoundException(paramsPath);
        if (!File.Exists(formulasPath)) throw new FileNotFoundException(formulasPath);
        if (!File.Exists(casesPath)) throw new FileNotFoundException(casesPath);

        var opt = new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        };

        var paramDefs = JsonSerializer.Deserialize<List<ParamDefinition>>(File.ReadAllText(paramsPath), opt) ?? [];
        var formulas = JsonSerializer.Deserialize<List<FormulaDefinition>>(File.ReadAllText(formulasPath), opt) ?? [];
        var cases = JsonSerializer.Deserialize<List<CaseDefinition>>(File.ReadAllText(casesPath), opt) ?? [];

        // 可选：B 方案 - 按参数建立取值库
        var paramLibPath = Path.Combine(defDir, "param_libraries.json");
        var paramLibraries = File.Exists(paramLibPath)
            ? JsonSerializer.Deserialize<List<ParamLibraryDefinition>>(File.ReadAllText(paramLibPath), opt) ?? []
            : [];

        // 可选：材料库与挂载关系
        var materialsPath = Path.Combine(defDir, "materials.json");
        var bindingsPath = Path.Combine(defDir, "material_bindings.json");

        var materialLibs = File.Exists(materialsPath)
            ? JsonSerializer.Deserialize<List<MaterialLibraryDefinition>>(File.ReadAllText(materialsPath), opt) ?? []
            : [];

        var materialBindings = File.Exists(bindingsPath)
            ? JsonSerializer.Deserialize<List<MaterialBindingDefinition>>(File.ReadAllText(bindingsPath), opt) ?? []
            : [];

        return new LoadedDefinitions(paramDefs, formulas, cases, paramLibraries, materialLibs, materialBindings);
    }
}

public sealed record LoadedDefinitions(
    IReadOnlyList<ParamDefinition> Params,
    IReadOnlyList<FormulaDefinition> Formulas,
    IReadOnlyList<CaseDefinition> Cases,
    IReadOnlyList<ParamLibraryDefinition> ParamLibraries,
    IReadOnlyList<MaterialLibraryDefinition> MaterialLibraries,
    IReadOnlyList<MaterialBindingDefinition> MaterialBindings
);