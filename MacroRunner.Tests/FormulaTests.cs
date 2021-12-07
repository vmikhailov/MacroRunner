using System;
using System.Linq.Expressions;
using FluentAssertions;
using MacroRunner.Compiler;
using Xunit;

namespace MacroRunner.Tests;

public class FormulaTests
{
    [Theory]
    [InlineData("1+2*(3+4)", 15)]
    [InlineData("-5+1", -4)]
    [InlineData("(1 > 0) + (3 = 3)", 2)]
    [InlineData("126 + 4.5", 130)]
    [InlineData("5^2", 25)]
    [InlineData("5^2+1", 26)]
    [InlineData("5^(2+1)", 125)]
    public void ShouldComputeSimpleIntMath(string exp, int result) => RunTest(exp, result);

    [Theory]
    [InlineData("1.1 + 2.3", 3.4)]
    [InlineData("1 + 2.3", 3.3)]
    [InlineData("-1 + 2.3", 1.3)]
    [InlineData("-(1 + 2.7)", -3.7)]
    [InlineData("sqrt(4)", 2)]
    [InlineData("(1 > 0) + 3.3", 4.3)]
    public void ShouldComputeSimpleDoubleMath(string exp, double result) => RunTest(exp, result, 1e-15);

    [Theory]
    [InlineData("\"5\" + \"3\"", "8")]
    public void ShouldConcatAutoConvertTwoStringsAndCompute(string exp, string result) => RunTest(exp, result);

    [Theory]
    [InlineData("\"5\" & \"3\"", "53")]
    [InlineData("5 & 3", "53")]
    public void ShouldPerformStringConcat(string exp, string result) => RunTest(exp, result);

    
    [Theory]
    [InlineData("len(\"abc\")", 3)]
    public void ShouldCallFunctionAndReturnInt(string exp, int result) => RunTest(exp, result);
    
    [Theory]
    [InlineData("trim(\" abc \")", "abc")]
    [InlineData("text(sqrt(4))", "2")]
    [InlineData("text(sqrt(4), \"0000\")", "0002")]
    public void ShouldCallFunctionAndReturnString(string exp, string result) => RunTest(exp, result);

    [Theory]
    [InlineData("(1 + 3) > 0", true)]
    [InlineData("(1 + 3) < 0", false)]
    [InlineData("1 <> 2", true)]
    [InlineData("1 = 1", true)]
    [InlineData("1 = 1.0", true)]
    [InlineData("5 >= 3", true)]
    [InlineData("5 <= 3", false)]
    [InlineData("or(1, and(1, 0))", true)]
    [InlineData("or(true, 0)", true)]
    [InlineData("or(1 > 0, 3 == 1.0)", true)]
    [InlineData("not(1 = 0)", true)]
    [InlineData("not(1 = 1)", false)]
    [InlineData("not(1 == 1)", false)]
    [InlineData("and(true, 0)", false)]
    [InlineData("or(true, and(true, false))", true)]
    [InlineData("abs(5-6)=1", true)]
    [InlineData("abs(5-6)<0", false)]
    [InlineData("1.3 >= (3 + 2) * 34.1", false)]
    [InlineData("200 >= (3 + 2) * 34.1", true)]
    public void ShouldComputeComparision(string exp, bool result) => RunTest(exp, result);

    [Theory]
    [InlineData("call(1+2)=3", 1)]
    [InlineData("call(1+2)", 3)]
    [InlineData("call(1,x+x+1)", 3)]
    [InlineData("call(x+1)", 2)]
    [InlineData("let(x,2,x+1)", 3)]
    [InlineData("let(x, let(y,1,y+1), x*2)", 4)]
    // [InlineData("LET(x,1,LET(y,2,x)+y)", 3)]
    // [InlineData("LET(x,2,x^3)+LET(x,3,x^4)", 89)]
    public void ShouldComputeLet(string exp, int result) => RunTest(exp, result);

    [Fact]
    public void ShouldNotSubtractTwoStrings()
    {
        var parsed = new FormulaParser().ParseExpression<dynamic>("\"aaaa\" - \"bbb\"");
        var func = parsed.Compile();
        var exception = Assert.Throws<FormatException>(() => func());

        Assert.Contains("Input string was not in a correct format", exception.Message);
    }

    #region helpers

    private static void RunTest<T>(string exp, T result)
    {
        var arr = new[] { 1 };
        Expression<Func<object>> f = () => arr[0];
        
        var parsed = new FormulaParser().ParseExpression<T>(exp);
        var func = parsed.Compile();

        func().Should().Be(result);
    }

    private static void RunTest(string exp, double result, double precision)
    {
        var parsed = new FormulaParser().ParseExpression<double>(exp);
        var func = parsed.Compile();

        func().Should().BeApproximately(result, precision);
    }

    #endregion
}