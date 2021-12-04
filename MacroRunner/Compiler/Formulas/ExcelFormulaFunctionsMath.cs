using System;

namespace MacroRunner.Compiler.Formulas;

public static partial class ExcelFormulaFunctions
{
    public static double Abs(double a) => Math.Abs(a);

    public static double Sqrt(double a) => (double)Math.Sqrt((double)a);
}