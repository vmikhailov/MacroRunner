using System.IO;
using System.Reflection;
using MacroRunner.Helpers;

namespace MacroRunner.Compiler;

public class CompiledCode
{
    public CompiledCode(string name, Stream stream)
    {
        Name = name;
        Assembly = Assembly.Load(stream.ReadAllBytes());
    }

    public Assembly Assembly { get; set; }

    public string Name { get; }

    public static CompilationResult Load(Stream stream)
    {
        return null;
    }

    public CompiledMacroInstance Create<T>()
        where T : CompiledMacroInstance, new()
    {
        var instance = new T();
        instance.Init(this);

        return instance;
    }

    public void Save(Stream stream)
    {
    }
}