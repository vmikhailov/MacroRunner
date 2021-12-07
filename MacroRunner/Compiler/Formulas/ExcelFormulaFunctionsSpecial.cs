using System;
using System.Linq;
using System.Linq.Expressions;
using MacroRunner.Runtime;

namespace MacroRunner.Compiler.Formulas;

public static partial class ExcelFormulaFunctions
{
    public static object Let(Expression variable, Func<IExecutionContext, object> init, Func<IExecutionContext, object> body)
    {
        var name = GetParameterName(variable);
        var ec = new FormulaExecutionContext();
        ec.SetNamedValue(name, init(ec));
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

    private static string GetParameterName(Expression exp)
    {
        var arg = ((exp as MethodCallExpression)?.Arguments.First() as ConstantExpression)?.Value as string;
        return arg;
    }
}