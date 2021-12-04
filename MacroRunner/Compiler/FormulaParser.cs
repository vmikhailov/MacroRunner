using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using MacroRunner.Compiler.Formulas;
using MacroRunner.Helpers;
using MacroRunner.Runtime;
using MacroRunner.Runtime.Excel;
using Sprache;
using ExcelRange = MacroRunner.Runtime.Excel.Range;

namespace MacroRunner.Compiler
{
    public class FormulaParser
    {
        private readonly FormulaParserSettings _settings;

        public FormulaParser(FormulaParserSettings settings)
        {
            _settings = settings;
        }

        public Expression<Func<T>> ParseExpression<T>(string text) => GetLambda<T>().Parse(text);

        public static Parser<Address> Address =>
            from fixedColumn in Parse.String("$").Token().Optional()
            from columnName in Parse.Letter.Repeat(1, 3).Text()
            from fixedRow in Parse.String("$").Token().Optional()
            from rowName in Parse.Digit.Repeat(1, 8).Text()
            select new Address()
            {
                ColumnName = columnName,
                RowNumber = int.Parse(rowName),
                FixedRow = fixedRow.IsDefined,
                FixedColumn = fixedColumn.IsDefined
            };

        public static Parser<ExcelRange> Range =>
            from a in Address.Once()
            from b in Parse.Char(':').Then(_ => Address).Repeat(0, 1)
            select RangeFactory.Create(a.Concat(b));

        private static Parser<string> Identifier =>
            from first in Parse.Letter.AtLeastOnce().Text()
            from second in Parse.LetterOrDigit.Many().Text()
            select first + second;

        private static Parser<OperatorType> Add = Operator("+", OperatorType.Add);
        private static Parser<OperatorType> Subtract = Operator("-", OperatorType.Subtract);
        private static Parser<OperatorType> Multiply = Operator("*", OperatorType.Multiply);
        private static Parser<OperatorType> Divide = Operator("/", OperatorType.Divide);
        private static Parser<OperatorType> Modulo = Operator("%", OperatorType.Modulo);
        private static Parser<OperatorType> Power = Operator("^", OperatorType.Power);
        private static Parser<OperatorType> GreaterThan = Operator(">", OperatorType.GreaterThan);
        private static Parser<OperatorType> GreaterThanOrEqual = Operator(">=", OperatorType.GreaterThanOrEqual);
        private static Parser<OperatorType> LessThan = Operator("<", OperatorType.LessThan);
        private static Parser<OperatorType> LessThanOrEqual = Operator("<=", OperatorType.LessThanOrEqual);
        private static Parser<OperatorType> Equal = Operator("=", OperatorType.Equal);
        private static Parser<OperatorType> NotEqual = Operator("<>", OperatorType.NotEqual);
        private static Parser<OperatorType> Concatenate = Operator("&", OperatorType.StringConcat);

        private Parser<Expression> Function =>
            from name in Identifier
            from lparen in Parse.Char('(')
            from expr in Parse.Ref(() => MathExpression).DelimitedBy(Parse.Char(_settings.ParametersDelimiter).Token())
            from rparen in Parse.Char(')')
            select MakeFunctionCall(name, expr.ToArray());

        private Parser<Expression> StringConstant =>
            from open in Parse.Char('"')
            from content in Parse.CharExcept('"').Many().Text()
            from close in Parse.Char('"')
            select Expression.Constant(content);

        private Parser<Expression> Constant =>
            Parse.Decimal
                 .Select(x => Expression.Constant(ParseConstant(x)))
                 .Or(StringConstant);

        private Parser<Expression> Variable =>
            from name in Identifier.Text()
            select MakeVariableAccess(name);

        private Parser<Expression> RangeExp =>
            from range in Range
            select Expression.Constant(range);

        private Parser<Expression> ExpressionInBraces =>
            (from lparen in Parse.Char('(')
             from expr in Parse.Ref(() => LogicalExpression)
             from rparen in Parse.Char(')')
             select expr).Named("expression");

        private Parser<Expression> Factor =>
            ExpressionInBraces.XOr(Constant.Or(Function).Or(Variable).Or(RangeExp));

        private Parser<Expression> Operand =>
            (from sign in Parse.Char('-')
             from factor in Factor
             select Expression.Negate(factor)
            ).XOr(Factor).Token();

        private Parser<Expression> InnerTerm => Parse.ChainOperator(Power, Operand, MakeBinary);

        private Parser<Expression> Term => Parse.ChainOperator(Multiply.Or(Divide).Or(Modulo), InnerTerm, MakeBinary);

        private Parser<Expression> MathExpression => Parse.ChainOperator(Add.Or(Subtract).Or(Concatenate), Term, MakeBinary);

        private Parser<Expression> LogicalExpression => Comparision.Or(MathExpression);

        private Parser<Expression> Comparision =>
            from exp1 in MathExpression
            from op in NotEqual.Or(LessThanOrEqual).Or(LessThan).XOr(GreaterThanOrEqual.Or(GreaterThan).Or(Equal))
            from exp2 in MathExpression
            select MakeBinary(op, exp1, exp2);

        private Parser<Expression<Func<T>>> GetLambda<T>() => LogicalExpression.End().Select(TypeCast<T>);

        private static Expression<Func<T>> TypeCast<T>(Expression body)
        {
            if (typeof(T) != body.Type)
            {
                body = TypeConversionHelper.GetTypeConversion(body.Type, typeof(T), body)!;
            }

            return Expression.Lambda<Func<T>>(body);
        }

        private static object ParseConstant(string str)
        {
            if (int.TryParse(str, out var intValue)) return intValue;
            if (double.TryParse(str, out var decValue)) return decValue;
            return str;
        }

        private static Parser<OperatorType> Operator(string op, OperatorType opType) =>
            Parse.String(op).Token().Return(opType);

        private static Parser<Func<T, T, T>> Operator<T>(string op, Func<T, T, T> opMethod) =>
            Parse.String(op).Token().Return(opMethod);

        private Expression MakeBinary(OperatorType operatorType, Expression left, Expression right)
        {
            if(OperatorToExpressionMap.TryGetValue(operatorType, out var exp))
            {
                var (leftArg, rightArg) = TypeConversionHelper.FindBestMinimalMatchingType(left, right, exp.type);
                return exp.factory(leftArg, rightArg);
            }
            
            if(OperatorToFunctionMap.TryGetValue(operatorType, out var name))
            {
                return MakeFunctionCall(name, left, right);
            }

            throw new ParseException($"Unknown operator {operatorType}");
        }

        private static Expression MakeVariableAccess(string name)
        {
            var prop = typeof(ExcelFormulaConstants).GetPublicStaticProperty(name);
            if (prop == null)
            {
                throw new ParseException(string.Format("Variable or constant '{0}' does not exist.", name));
            }

            return Expression.Constant(prop.GetValue(null));
        }

        private Expression MakeFunctionCall(string name, params Expression[] args) => MakeCall(typeof(ExcelFormulaFunctions), name, args);

        private Expression MakeCall(Type impl, string name, params Expression[] args)
        {
            var methodWithArgs = impl.GetPublicStaticMethods(name, args.Length)
                                     .Select(x => new { Method = x, Match = MatchArgs(x.GetParameters(), args) })
                                     .Where(x => x.Match.Score > 0)
                                     .OrderBy(x => x.Match.Score)
                                     .Select(x => new { x.Method, x.Match.Args })
                                     .FirstOrDefault();

            if (methodWithArgs == null)
            {
                throw new ParseException(
                    string.Format(
                        "Function '{0}({1})' does not exist.",
                        name,
                        string.Join(", ", args.Select(e => e.Type.Name))));
            }

            return Expression.Call(methodWithArgs.Method, methodWithArgs.Args);
        }

        private (int Score, Expression[] Args) MatchArgs(ParameterInfo[] parameters, Expression[] args)
        {
            var newArgs = parameters.Zip(args)
                                    .Select(x => MatchArgs(x.First.ParameterType, x.Second))
                                    .ToList();

            if (newArgs.Any(x => x.Score == 0))
            {
                return (0, args);
            }

            return (newArgs.Select(x => x.Score).Sum(), newArgs.Select(x => x.Arg).ToArray());
        }

        private (int Score, Expression Arg) MatchArgs(Type paramType, Expression arg)
        {
            var argType = arg.Type;
            if (paramType == argType)
            {
                return (1, arg);
            }

            var conversion = TypeConversionHelper.GetTypeConversion(argType, paramType, arg);
            return conversion != null ? (2, conversion) : (0, arg);
        }

        private static IDictionary<OperatorType, (ExpressionType type, Func<Expression, Expression, Expression> factory) > OperatorToExpressionMap = 
            new Dictionary<OperatorType, (ExpressionType, Func<Expression, Expression, Expression>)>()
        {
            { OperatorType.Add, (ExpressionType.Add, (l, r) => Expression.Add(l, r)) },
            { OperatorType.Divide, (ExpressionType.Divide, (l, r) => Expression.Divide(l, r)) },
            { OperatorType.Modulo, (ExpressionType.Modulo, (l, r) => Expression.Modulo(l, r))  },
            { OperatorType.Multiply, (ExpressionType.Multiply, (l, r) => Expression.Multiply(l, r))  },
            { OperatorType.Subtract, (ExpressionType.Subtract, (l, r) => Expression.Subtract(l, r))  },
            { OperatorType.Power, (ExpressionType.Power, (l, r) => Expression.Power(l, r))  },
            { OperatorType.GreaterThan, (ExpressionType.GreaterThan, (l, r) => Expression.GreaterThan(l, r)) },
            { OperatorType.GreaterThanOrEqual, (ExpressionType.GreaterThanOrEqual, (l, r) => Expression.GreaterThanOrEqual(l, r)) },
            { OperatorType.LessThan, (ExpressionType.LessThan, (l, r) => Expression.LessThan(l, r)) },
            { OperatorType.LessThanOrEqual, (ExpressionType.LessThanOrEqual, (l, r) => Expression.LessThanOrEqual(l, r)) },
            { OperatorType.Equal, (ExpressionType.Equal, (l, r) => Expression.Equal(l, r)) },
            { OperatorType.NotEqual, (ExpressionType.NotEqual, (l, r) => Expression.NotEqual(l, r)) },
        };

        private static IDictionary<OperatorType, string> OperatorToFunctionMap = new Dictionary<OperatorType, string>()
        {
            { OperatorType.StringConcat, "Concat" },
        };
    }
}