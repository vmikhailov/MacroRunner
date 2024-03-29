using System;

namespace MacroRunner.Compiler.Formulas;

public static partial class ExcelFormulaFunctions
{
    public static int Len(string? s) => s?.Length ?? 0;

    public static string Trim(string s) => s.Trim();
        
    public static string Concat(string a, string b) => a + b;
        
    public static string Concatenate(string a, string b) => a + b;
}