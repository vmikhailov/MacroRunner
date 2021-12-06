using System;
using MacroRunner.Runtime;

namespace MacroRunner.Compiler.Formulas;

public static partial class ExcelFormulaFunctions
{
    public static object Let(object variable, Func<IExecutionContext, object> init, Func<IExecutionContext, object> body)
    {
        var ec = new FormulaExecutionContext();
        ec.SetNamedValue("x", init(ec));
        return body(ec);
    }

    public static object Call(Func<IExecutionContext, object> body)
    {
        var ec = new FormulaExecutionContext();
        ec.SetNamedValue("x", 1);
        return body(ec);
    }
    
    
    public static object Call(Func<IExecutionContext, object> f1, Func<IExecutionContext, object> f2)
    {
        var ec = new FormulaExecutionContext();
        ec.SetNamedValue("x", f1(ec));
        return f2(ec);
    }
}