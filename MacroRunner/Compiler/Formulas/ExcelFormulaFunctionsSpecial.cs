using System;
using System.Linq.Expressions;
using MacroRunner.Runtime;

namespace MacroRunner.Compiler.Formulas;

public static partial class ExcelFormulaFunctions
{
    public static object Let(
        IExecutionContext context,
        Expression variable,
        Func<IExecutionContext, object> init,
        Func<IExecutionContext, object> body)
    {
        var name = GetParameterName(variable);
        var value = init(context);
        var scoped = context.GetScopedWithValue(name, value);
        return body(scoped);
    }

    public static object Call(
        IExecutionContext context,
        Func<IExecutionContext, object> body)
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
        var arg = (exp as IArgumentProvider)?.GetArgument(0);
        var name = (arg as ConstantExpression)?.Value as string;
        return name;
    }
}