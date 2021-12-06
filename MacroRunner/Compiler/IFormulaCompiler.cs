using System;
using MacroRunner.Runtime;

namespace MacroRunner.Compiler;

public interface IFormulaCompiler
{
    Func<IExecutionContext, object> Compile(string text);
}