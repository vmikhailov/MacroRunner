using System;

namespace MacroRunner.Compiler.Formulas;

public static partial class ExcelFormulaOperators
{
    public static dynamic Add(dynamic a, dynamic b) => a + b;
    public static dynamic Subtract(dynamic a, dynamic b) => a - b;
    public static dynamic Multiply(dynamic a, dynamic b) => a * b;
    public static dynamic Divide(dynamic a, dynamic b) => a / b;
    public static dynamic Modulo(dynamic a, dynamic b) => a % b;
    public static dynamic Power(dynamic a, dynamic b) => Math.Pow(a, b);
    public static bool Equal(dynamic a, dynamic b) => a == b;
    
    public static string Concat(string a, string b) => a + b;
}