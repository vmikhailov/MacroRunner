using System;
using FluentAssertions;
using MacroRunner.Compiler;
using Sprache;
using Xunit;

namespace MacroRunner.Tests;

public class FormulaTests
{
    [Fact]
    public void ShouldComputeSimpleIntMath()
    {
        var parsed = CreateParser().ParseExpression<int>("1+2*(3+4)");
        var func = parsed.Compile();

        func().Should().Be(15);
    }

    [Fact]
    public void ShouldComputeSimpleDoubleMath()
    {
        var parsed = CreateParser().ParseExpression<double>("1.1 + 2.3");
        var func = parsed.Compile();

        func().Should().Be(3.4);
    }

    [Fact]
    public void ShouldComputeSimpleMixedMath()
    {
        var parsed = CreateParser().ParseExpression<double>("1 + 2.3");
        var func = parsed.Compile();

        func().Should().Be(3.3);
    }

    [Fact]
    public void ShouldConcatTwoStrings()
    {
        var parsed = CreateParser().ParseExpression<string>("\"aaaa\" + \"bbb\"");
        var func = parsed.Compile();

        func().Should().Be("aaaabbb");
    }

    [Fact]
    public void ShouldComputeScalarFunction()
    {
        var parsed = CreateParser().ParseExpression<double>("sqrt(4)");
        var func = parsed.Compile();

        func().Should().Be(2);
    }

    [Fact]
    public void ShouldComputeScalarFunctionAndConvertToString()
    {
        var parsed = CreateParser().ParseExpression<string>("text(sqrt(4))");
        var func = parsed.Compile();

        func().Should().Be("2");
    }

    [Fact]
    public void ShouldComputeScalarFunctionAndConvertToStringWithFormatting()
    {
        var parsed = CreateParser().ParseExpression<string>("text(sqrt(4), \"0000\")");
        var func = parsed.Compile();

        func().Should().Be("0002");
    }

    [Fact]
    public void ShouldComputeLogicalFunction()
    {
        var parsed = CreateParser().ParseExpression<bool>("or(true, and(true, false))");
        var func = parsed.Compile();

        func().Should().Be(true);
    }

    [Fact]
    public void ShouldComputeLogicalFunctionAutoCastToBool()
    {
        var parsed = CreateParser().ParseExpression<bool>("or(1, and(1, 0))");
        var func = parsed.Compile();

        func().Should().Be(true);
    }

    [Fact]
    public void ShouldComputeComparisionGreaterThan()
    {
        var parsed = CreateParser().ParseExpression<bool>("(1 + 3) > 0");
        var func = parsed.Compile();

        func().Should().Be(true);
    }


    [Fact]
    public void ShouldComputeComparisionLessThan()
    {
        var parsed = CreateParser().ParseExpression<bool>("(1 + 3) < 0");
        var func = parsed.Compile();

        func().Should().Be(false);
    }

    [Fact]
    public void ShouldComputeComparisionNotEqual()
    {
        var parsed = CreateParser().ParseExpression<bool>("1 <> 2");
        var func = parsed.Compile();

        func().Should().Be(true);
    }

    [Fact]
    public void ShouldComputeMathFunctionAutoCastFromBool()
    {
        var parsed = CreateParser().ParseExpression<int>("(1 > 0) + (3 = 3)");
        var func = parsed.Compile();

        func().Should().Be(2);
    }

    [Fact]
    public void ShouldComputeMathFunctionAutoCastFromBoolMixed()
    {
        var parsed = CreateParser().ParseExpression<double>("(1 > 0) + 3.3");
        var func = parsed.Compile();

        func().Should().Be(4.3);
    }

    private static FormulaParser CreateParser() => new (new() { ParametersDelimiter = ',' });


    [Fact]
    public void ShouldNotSubtractTwoStrings()
    {
        var exception = Assert.Throws<ParseException>(() => CreateParser().ParseExpression<dynamic>("\"aaaa\" - \"bbb\""));

        Assert.Contains("Function 'Subtract(String, String)' does not exist.", exception.Message);
    }
}