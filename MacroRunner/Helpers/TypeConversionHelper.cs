using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace MacroRunner.Helpers;

public class TypeConversionHelper
{
    public static (Expression left, Expression right) FindBestMinimalMatchingType(Expression left, Expression right, ExpressionType op)
    {
        var leftTypeId = FindTypeId(left.Type);
        var rightTypeId = FindTypeId(right.Type);
        if (!MinimumTypeForOperation.TryGetValue(op, out var minType))
        {
            throw new($"Operation {op} is not supported yet");
        }
        var minTypeId = FindTypeId(minType);

        var typeId = Math.Max(Math.Max(leftTypeId, rightTypeId), minTypeId);
        var leftExp = GetTypeConversion(leftTypeId, typeId, left);
        var rightExp = GetTypeConversion(rightTypeId, typeId, right);
        if (leftExp == null || rightExp == null)
        {
            throw new($"Unable to find conversion between {left.Type} and {right.Type}");
        }

        return (leftExp, rightExp);
    }

    public static Expression? GetTypeConversion(Type sourceType, Type targetType, Expression exp)
    {
        var sourceTypeId = FindTypeId(sourceType);
        var targetTypeId = FindTypeId(targetType);
        return GetTypeConversion(sourceTypeId, targetTypeId, exp);
    }

    private static Expression? GetTypeConversion(int sourceTypeId, int targetTypeId, Expression exp) => 
        TypeConversionMap[sourceTypeId, targetTypeId](exp);

    private static int FindTypeId(Type type)
    {
        for (var i = 0; i < TypeOrder.Length; i++)
        {
            var x = TypeOrder[i];
            if (x == type || x.IsAssignableFrom(type))
            {
                return i;
            }
        }

        return -1;
    }

    private static Type[] TypeOrder = new[] { typeof(string), typeof(bool), typeof(int), typeof(double), typeof(object) };

    private static Func<Expression, Expression?>[,] TypeConversionMap = new Func<Expression, Expression?>[,]
    {
        //from string
        { e => e, e => Parse<bool>(e), e => Parse<int>(e), e => Parse<double>(e), e => TypeAs<object>(e) },
        // from bool
        { e => ToString<bool>(e), e => e, e => Convert<int>(e), e => Convert<double>(e), e => TypeAs<object>(e) },
        // from int
        { e => ToString<int>(e), e => GreaterThanZero<int>(e), e => e, e => Convert<double>(e), e => TypeAs<object>(e) },
        // from double
        { e => ToString<double>(e), e => GreaterThanZero<double>(e), e => Convert<int>(e), e => e, e => TypeAs<object>(e) },
        // from object
        { e => ToString<object>(e), e => Convert<bool>(e), e => Convert<int>(e), e => Convert<double>(e), e => e },
    };

    private static Expression Parse<T>(Expression e) => Expression.Call(typeof(T), "Parse", null, e);

    private static Expression Convert<T>(Expression e) => Expression.Convert(e, typeof(T));
    private static Expression TypeAs<T>(Expression e) => Expression.TypeAs(e, typeof(T));

    private static Expression ToString<T>(Expression e) => Expression.Call(e, "ToString", Type.EmptyTypes);

    private static Expression GreaterThanZero<T>(Expression e) => Expression.GreaterThan(e, Expression.Constant(default(T)));

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