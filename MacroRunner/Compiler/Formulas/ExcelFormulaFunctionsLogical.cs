namespace MacroRunner.Compiler.Formulas;

public static partial class ExcelFormulaFunctions
{
    public static bool And(bool a, bool b) => a && b;

    public static bool Or(bool a, bool b) => a || b;

    public static bool Not(bool a) => !a;
}