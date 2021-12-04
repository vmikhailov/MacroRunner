using FluentAssertions;
using MacroRunner.Compiler;
using Sprache;
using Xunit;

namespace MacroRunner.Tests;

public class FormulaTests
{
    [Theory]
    [InlineData("1+2*(3+4)", 15)]
    [InlineData("-5+1", -4)]
    [InlineData("(1 > 0) + (3 = 3)", 2)]
    [InlineData("126 + 4.5", 130)]
    public void ShouldComputeSimpleIntMath(string exp, int result) => RunTest(exp, result);

    [Theory]
    [InlineData("1.1 + 2.3", 3.4)]
    [InlineData("1 + 2.3", 3.3)]
    [InlineData("-1 + 2.3", 1.3)]
    [InlineData("-(1 + 2.7)", -3.7)]
    [InlineData("sqrt(4)", 2)]
    [InlineData("(1 > 0) + 3.3", 4.3)]
    public void ShouldComputeSimpleDoubleMath(string exp, double result) => RunTest(exp, result, 1e-15);

    [Fact]
    public void ShouldConcatTwoStrings() => RunTest("\"aaaa\" + \"bbb\"", "aaaabbb");

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
    [InlineData("5 >= 3", true)]
    [InlineData("5 <= 3", false)]
    [InlineData("or(1, and(1, 0))", true)]
    [InlineData("or(true, 0)", true)]
    [InlineData("and(true, 0)", false)]
    [InlineData("or(true, and(true, false))", true)]
    [InlineData("abs(5-6)=1", true)]
    [InlineData("abs(5-6)<0", false)]
    [InlineData("1.3 >= (3 + 2) * 34.1", false)]
    [InlineData("200 >= (3 + 2) * 34.1", true)]
    public void ShouldComputeComparision(string exp, bool result) => RunTest(exp, result);

    [Fact]
    public void ShouldNotSubtractTwoStrings()
    {
        var exception = Assert.Throws<ParseException>(() => CreateParser().ParseExpression<dynamic>("\"aaaa\" - \"bbb\""));

        Assert.Contains("Function 'Subtract(String, String)' does not exist.", exception.Message);
    }

    #region helpers

    private static FormulaParser CreateParser() => new(new() { ParametersDelimiter = ',' });

    private static void RunTest<T>(string exp, T result)
    {
        var parsed = CreateParser().ParseExpression<T>(exp);
        var func = parsed.Compile();

        func().Should().Be(result);
    }

    private static void RunTest(string exp, double result, double precision)
    {
        var parsed = CreateParser().ParseExpression<double>(exp);
        var func = parsed.Compile();

        func().Should().BeApproximately(result, precision);
    }

    #endregion
}