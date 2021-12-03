using System;
using Range = MacroRunner.Runtime.Excel.Range;

namespace MacroRunner.Compiler.Formulas
{
    public static partial class ExcelFormulaFunctionMath
    {
        public static decimal Abs(decimal a) => Math.Abs(a);

        public static decimal Sqrt(decimal a) => (decimal)Math.Sqrt((double)a);
    }

    public static partial class ExcelFormulaFunctionLogical
    {
        public static bool And(bool a, bool b) => a && b;

        public static bool Or(bool a, bool b) => a || b;

        public static bool Not(bool a) => !a;

        public static bool True = true;
        
        public static bool False = false;
    }
}