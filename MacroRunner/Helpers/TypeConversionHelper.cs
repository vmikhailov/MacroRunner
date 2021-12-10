using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
using MacroRunner.Compiler;
using MacroRunner.Runtime;

namespace MacroRunner.Helpers;

public class TypeConversionHelper
{
    public static bool FindMinimalMatchingType(
        IParserContext context,
        ExpressionType op,
        ref Expression left,
        ref Expression right)
    {
        var leftTypeId = FindTypeId(left.Type);
        var rightTypeId = FindTypeId(right.Type);
        if (!MinimumTypeForOperation.TryGetValue(op, out var minType))
        {
            throw new($"Operation {op} is not supported yet");
        }

        var minTypeId = FindTypeId(minType);

        var typeId = Math.Max(Math.Max(leftTypeId, rightTypeId), minTypeId);
        var leftExp = GetTypeConversion(context, leftTypeId, typeId, left);
        var rightExp = GetTypeConversion(context, rightTypeId, typeId, right);
        if (leftExp == null || rightExp == null)
        {
            return false;
            //throw new($"Unable to find conversion between {left.Type} and {right.Type}");
        }

        left = leftExp;
        right = rightExp;
        return true;
    }

    public static Expression? GetTypeConversion(IParserContext context, Type sourceType, Type targetType, Expression exp)
    {
        var sourceTypeId = FindTypeId(sourceType);
        var targetTypeId = FindTypeId(targetType);
        return GetTypeConversion(context, sourceTypeId, targetTypeId, exp);
    }

    private static Expression? GetTypeConversion(IParserContext context, int sourceTypeId, int targetTypeId, Expression exp) =>
        TypeConversionMap[sourceTypeId, targetTypeId](exp, context);

    private static int FindTypeId(Type type)
    {
        for (var i = TypeOrder.Length - 1; i >= 0; i--)
        {
            var x = TypeOrder[i];
            if (x == type || x.IsAssignableFrom(type))
            {
                return i;
            }
        }

        return -1;
    }

    private static Type[] TypeOrder = new[]
    {
        typeof(object),
        typeof(string),
        typeof(bool),
        typeof(int),
        typeof(double),
        typeof(Func<IExecutionContext, object>),
        typeof(Expression)
    };

    private static Func<Expression, IParserContext, Expression?>[,] TypeConversionMap =
        new Func<Expression, IParserContext, Expression?>[,]
        {
            // from object
            {
                (e, c) => e,
                (e, c) => null,
                (e, c) => null,
                (e, c) => null,
                (e, c) => null,
                (e, c) => ToLambda<object>(e, c),
                (e, c) => Expression.Constant(e)
            },

            //from string
            {
                (e, c) => Convert<object>(e),
                (e, c) => e,
                (e, c) => Parse<bool>(e),
                (e, c) => Parse<int>(e),
                (e, c) => Parse<double>(e),
                (e, c) => ToLambda<object>(e, c),
                (e, c) => Expression.Constant(e)
            },

            // from bool
            {
                (e, c) => Convert<object>(e),
                (e, c) => ToString<bool>(e),
                (e, c) => e,
                (e, c) => Convert<int>(e),
                (e, c) => Convert<double>(e),
                (e, c) => ToLambda<object>(e, c),
                (e, c) => Expression.Constant(e)
            },

            // from int
            {
                (e, c) => Convert<object>(e),
                (e, c) => ToString<int>(e),
                (e, c) => GreaterThanZero<int>(e),
                (e, c) => e,
                (e, c) => Convert<double>(e),
                (e, c) => ToLambda<object>(e, c),
                (e, c) => Expression.Constant(e)
            },

            // from double
            {
                (e, c) => Convert<object>(e),
                (e, c) => ToString<double>(e),
                (e, c) => GreaterThanZero<double>(e),
                (e, c) => Convert<int>(e),
                (e, c) => e,
                (e, c) => ToLambda<object>(e, c),
                (e, c) => Expression.Constant(e)
            },

            {
                (e, c) => null,
                (e, c) => null,
                (e, c) => null,
                (e, c) => null,
                (e, c) => null,
                (e, c) => null,
                (e, c) => Expression.Constant(e)
            }
        };

    private static Expression Parse<T>(Expression e) =>
        Expression.Call(typeof(T), "Parse", null, e, Expression.Constant(CultureInfo.InvariantCulture));

    private static Expression Convert<T>(Expression e) => Expression.Convert(e, typeof(T));

    private static Expression Convert2<T>(Expression e)
    {
        var genericMethod = typeof(TypeConversionHelper)
            .GetMethod(nameof(_), BindingFlags.Static | BindingFlags.NonPublic)!;
        var method = genericMethod.MakeGenericMethod(e.Type, typeof(T));
        return Expression.Call(null, method, e);
    }
        
    private static T2 _<T1, T2>(T1 val)
    {
        var v = System.Convert.ChangeType(val, typeof(T2));
        return (T2)v;
    }

    private static Expression TypeAs<T>(Expression e) => Expression.TypeAs(e, typeof(T));

    private static Expression ToString<T>(Expression e) => Expression.Call(e, "ToString", Type.EmptyTypes);

    private static Expression GreaterThanZero<T>(Expression e) => Expression.GreaterThan(e, Expression.Constant(default(T)));

    private static Expression ToLambda<T>(Expression e, IParserContext context)
    {
        var body = Expression.Convert(e, typeof(object));
        var lambda = Expression.Lambda<Func<IExecutionContext, object>>(body, context.ExecutionContextParameter);
        return lambda;
        //var func = lambda.Compile();
        //return Expression.Constant(func);
    }

    private static IDictionary<ExpressionType, Type> MinimumTypeForOperation = new Dictionary<ExpressionType, Type>()
    {
        { ExpressionType.Add, typeof(int) },
        { ExpressionType.Divide, typeof(double) },
        { ExpressionType.Modulo, typeof(int) },
        { ExpressionType.Multiply, typeof(int) },
        { ExpressionType.Subtract, typeof(int) },
        { ExpressionType.Power, typeof(double) },
        { ExpressionType.GreaterThan, typeof(int) },
        { ExpressionType.GreaterThanOrEqual, typeof(int) },
        { ExpressionType.LessThan, typeof(int) },
        { ExpressionType.LessThanOrEqual, typeof(int) },
        { ExpressionType.Equal, typeof(string) },
        { ExpressionType.NotEqual, typeof(string) },
    };
}