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

        private static Parser<ExpressionType> Add = Operator("+", ExpressionType.Add);
        private static Parser<ExpressionType> Subtract = Operator("-", ExpressionType.Subtract);
        private static Parser<ExpressionType> Multiply = Operator("*", ExpressionType.Multiply);
        private static Parser<ExpressionType> Divide = Operator("/", ExpressionType.Divide);
        private static Parser<ExpressionType> Modulo = Operator("%", ExpressionType.Modulo);
        private static Parser<ExpressionType> Power = Operator("^", ExpressionType.Power);
        private static Parser<ExpressionType> GreaterThan = Operator(">", ExpressionType.GreaterThan);
        private static Parser<ExpressionType> LessThan = Operator("<", ExpressionType.LessThan);
        private static Parser<ExpressionType> Equal = Operator("=", ExpressionType.Equal);
        private static Parser<ExpressionType> NotEqual = Operator("<>", ExpressionType.NotEqual);

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
            ((from sign in Parse.Char('-')
              from factor in Factor
              select Expression.Negate(factor)
                ).XOr(Factor)).Token();

        private Parser<Expression> InnerTerm => Parse.ChainOperator(Power, Operand, MakeBinary);

        private Parser<Expression> Term => Parse.ChainOperator(Multiply.Or(Divide).Or(Modulo), InnerTerm, MakeBinary);

        private Parser<Expression> MathExpression => Parse.ChainOperator(Add.Or(Subtract), Term, MakeBinary);

        private Parser<Expression> LogicalExpression => Comparision.Or(MathExpression);

        private Parser<Expression> Comparision =>
            from exp1 in MathExpression
            from op in NotEqual.Or(LessThan).XOr(GreaterThan.Or(Equal))
            from exp2 in MathExpression
            select MakeBinary(op, exp1, exp2);

        private Parser<Expression<Func<T>>> GetLambda<T>() => LogicalExpression.End().Select(TypeCast<T>);

        private static Expression<Func<T>> TypeCast<T>(Expression body)
        {
            var type = GetExpressionType(body);
            if (typeof(T) != type)
            {
                body = Expression.Convert(body, typeof(T));
            }

            return Expression.Lambda<Func<T>>(body);
        }

        private static object ParseConstant(string str)
        {
            if (int.TryParse(str, out var intValue)) return intValue;
            if (double.TryParse(str, out var decValue)) return decValue;
            return str;
        }

        private static Parser<ExpressionType> Operator(string op, ExpressionType opType)
        {
            return Parse.String(op).Token().Return(opType);
        }

        private Expression MakeBinary(ExpressionType expressionType, Expression arg1, Expression arg2)
        {
            switch (expressionType)
            {
                case ExpressionType.Add:
                case ExpressionType.Divide:
                case ExpressionType.Modulo:
                case ExpressionType.Multiply:
                case ExpressionType.Subtract:
                case ExpressionType.GreaterThan:
                case ExpressionType.LessThan:
                case ExpressionType.Equal:
                case ExpressionType.NotEqual:
                    return MakeOperationCall(expressionType.ToString(), arg1, arg2);

                default:
                    throw new NotImplementedException($"Expression of type {expressionType} is not supported yet");
            }
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

        private Expression MakeFunctionCall(string name, Expression[] args) => MakeCall(typeof(ExcelFormulaFunctions), name, args);

        private Expression MakeOperationCall(string name, params Expression[] args) => MakeCall(typeof(ExcelFormulaOperations), name, args);

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
            var argType = GetExpressionType(arg);
            if (paramType == argType)
            {
                return (1, arg);
            }

            if (paramType == typeof(double) && argType == typeof(int))
            {
                return (2, Expression.Convert(arg, paramType));
            }

            if (paramType == typeof(bool) && argType == typeof(int))
            {
                return (2, Expression.GreaterThan(arg, Expression.Constant(0)));
            }

            if (paramType == typeof(bool) && argType == typeof(double))
            {
                return (2, Expression.GreaterThan(arg, Expression.Constant(0m)));
            }
            
            if (paramType == typeof(int) && argType == typeof(bool))
            {
                return (2, Expression.Convert(arg, paramType));
            }
            
            if (paramType == typeof(double) && argType == typeof(bool))
            {
                return (2, Expression.Convert(arg, paramType));
            }

            if (paramType.IsAssignableFrom(argType)) return (3, Expression.Convert(arg, paramType));

            return (0, arg);
        }

        private static Type GetExpressionType(Expression exp) =>
            exp switch
            {
                MethodCallExpression mce => mce.Type,
                ConstantExpression ce => ce.Type,
                _ => throw new ParseException("Unsupported expression")
            };
    }
}