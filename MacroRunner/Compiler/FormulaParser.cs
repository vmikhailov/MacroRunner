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
    public class FormulaParserSettings
    {
        public char ParametersDelimiter = ',';
    }

    public class FormulaParser<T>
    {
        private readonly FormulaParserSettings _settings;

        public FormulaParser(FormulaParserSettings settings)
        {
            _settings = settings;
        }

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

        private Parser<Expression> Function =>
            from name in Identifier
            from lparen in Parse.Char('(')
            from expr in Parse.Ref(() => Expr).DelimitedBy(Parse.Char(_settings.ParametersDelimiter).Token())
            from rparen in Parse.Char(')')
            select CallFunction(name, expr.ToArray());

        private static Parser<Expression> StringConstant =
            from open in Parse.Char('"')
            from content in Parse.CharExcept('"').Many().Text()
            from close in Parse.Char('"')
            select Expression.Constant(content);

        private static Parser<Expression> Constant =
            Parse.Decimal
                 .Select(x => Expression.Constant(ParseConstant(x)))
                 .Or(StringConstant);

        private Parser<Expression> ExpressionInBraces =>
            (from lparen in Parse.Char('(')
             from expr in Parse.Ref(() => Expr)
             from rparen in Parse.Char(')')
             select expr).Named("expression");

        private Parser<Expression> Factor =>
            ExpressionInBraces.XOr(Constant.Or(Function).Or(Range.Select(x => Expression.Constant(x))));

        private Parser<Expression> Operand =>
            ((from sign in Parse.Char('-')
              from factor in Factor
              select Expression.Negate(factor)
                ).XOr(Factor)).Token();

        private Parser<Expression> InnerTerm => Parse.ChainOperator(Power, Operand, MakeBinary);

        private Parser<Expression> Term => Parse.ChainOperator(Multiply.Or(Divide).Or(Modulo), InnerTerm, MakeBinary);

        private Parser<Expression> Expr => Parse.ChainOperator(Add.Or(Subtract), Term, MakeBinary);

        private Parser<Expression<Func<T>>> Lambda => Expr.End().Select(TypeCast);

        public Expression<Func<T>> ParseExpression(string text) => Lambda.Parse(text);

        private static Expression<Func<T>> TypeCast(Expression body)
        {
            var type = GetExpressionType(body);
            if (typeof(T) != type)
            {
                body = Expression.Convert(body, typeof(T));
            }

            return Expression.Lambda<Func<T>>(body);
        }

        private static Expression CallFunction(string name, Expression[] parameters)
        {
            var parameterTypes = parameters.Select(x => x.Type).ToArray();
            var methodInfo = FindFunction(name, parameterTypes);
            if (methodInfo == null)
            {
                throw new ParseException(
                    string.Format(
                        "Function '{0}({1})' does not exist.",
                        name,
                        string.Join(",", parameters.Select(e => e.Type.Name))));
            }

            return Expression.Call(methodInfo, parameters);
        }

        private static object ParseConstant(string str)
        {
            //if (int.TryParse(str, out var intValue)) return intValue;
            if (decimal.TryParse(str, out var decValue)) return decValue;
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
                    return MakeOperationCall(expressionType.ToString(), arg1, arg2);

                default:
                    throw new NotImplementedException($"Expression of type {expressionType} is not supported yet");
            }
        }

        private Expression MakeOperationCall(string name, params Expression[] args)
        {
            var argTypes = args.Select(GetExpressionType).ToList();

            var methodInfo = FindOperation(name, args, argTypes);

            if (methodInfo == null)
            {
                throw new ParseException(
                    string.Format(
                        "Function '{0}({1})' does not exist.",
                        name,
                        string.Join(",", args.Select(e => e.Type.Name))));
            }

            var arguments = new Expression[args.Length];
            var parameters = methodInfo.GetParameters();
            for (var i = 0; i < arguments.Length; i++)
            {
                var paramType = parameters[i].ParameterType;
                if (paramType != argTypes[i])
                {
                    arguments[i] = Expression.Convert(args[i], paramType);
                }
                else
                {
                    arguments[i] = args[i];
                }
            }

            return Expression.Call(typeof(ExcelFormulaOperations), name, null, arguments);
        }

        private MethodInfo? FindOperation(string name, Expression[] args, List<Type> argTypes)
        {
            var methodInfo = typeof(ExcelFormulaOperations)
                             .GetPublicStaticMethods(name, argTypes.Count)
                             .OrderBy(x => GetMatchingScore(x.GetParameters(), argTypes))
                             .FirstOrDefault();

            return methodInfo;
        }

        private static MethodInfo? FindFunction(string name, Type[] parameterTypes)
        {
            var methodInfo = GetFunctionClasses()
                              .SelectMany(
                                  t => t.GetPublicStaticMethods(name, parameterTypes.Length)
                                        .Where(x => x.GetParameters().Select(y => y.ParameterType).SequenceEqual(parameterTypes)))
                              .FirstOrDefault();
            
            return methodInfo;
        }

        private static IEnumerable<Type> GetFunctionClasses()
        {
            return Assembly.GetExecutingAssembly()
                           .GetTypes()
                           .Where(x => x.Name.StartsWith("ExcelFormulaFunction"));
        }

        private int GetMatchingScore(ParameterInfo[] parameters, List<Type> argTypes)
        {
            return parameters.Zip(argTypes)
                             .Select(x => GetMatchingScore(x.First.ParameterType, x.Second))
                             .Sum();
        }

        private int GetMatchingScore(Type type1, Type type2)
        {
            if (type1 == type2) return 0;
            if (type1 == typeof(decimal) && type2 == typeof(int)) return 1;
            if (type1.IsAssignableFrom(type2)) return 10;
            return 1000;
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