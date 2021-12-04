using System;
using System.IO;
using MacroRunner.Compiler;

namespace MacroRunner.Runtime;

public class ExcelMacroInstance : CompiledMacroInstance<ExcelGlobals>
{
    public override void LoadFile(Stream file)
    {
        throw new NotImplementedException();
    }

    public override void InitGlobals()
    {
        throw new NotImplementedException();
    }
}