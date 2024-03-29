using Range = MacroRunner.Runtime.Excel.Range;

namespace MacroRunner.Compiler.Formulas;

public static partial class ExcelFormulaFunctions
{
    public static double Sum(Range r)
    {
        return (r.To.RowNumber - r.From.RowNumber + 1) * (r.To.ColumnNumber - r.From.ColumnNumber + 1);
    }
        
    public static string Text(double a)
    {
        return a.ToString();
    }

    public static string Text(double a, string format)
    {
        return a.ToString(format);
    }
        
    public static dynamic Vlookup(int a, int b)
    {
        return null;
    }
        
    public static dynamic Vlookup(int a, int b, int c)
    {
        return null;
    }
        
    public static dynamic Vlookup(int a, string b, int c)
    {
        return a + b + c;
    }
        
    public static dynamic Vlookup2(int a, string b, int c)
    {
        return "Vlookup2:" + a + b + c;
    }
        
    public static dynamic Vlookup2(Range a, string b, int c)
    {
        return "Range: " + a + b + c;
    }
        
    public static dynamic Vlookup(int a, int b, string c)
    {
        return null;
    }
        
    public static dynamic Vlookup(int a, int b, int c, int d)
    {
        return null;
    }
}