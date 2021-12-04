namespace MacroRunner.Compiler;

public class CodeBlock
{
    public CodeBlock(string name, string body)
    {
        Name = name;
        Body = body;
    }

    public string Name { get; }

    public string Body { get; }
}