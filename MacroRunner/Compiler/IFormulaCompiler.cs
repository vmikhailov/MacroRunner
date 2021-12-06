using System;
using MacroRunner.Runtime.Excel;
using MacroRunner.Runtime.Excel.Cells;

namespace MacroRunner.Compiler;

public interface IFormulaCompiler
{
    Func<IExecutionContext, object> Compile(string text);
}