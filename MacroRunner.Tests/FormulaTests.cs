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
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<int>("1+2*(3+4)");
        var func = parsed.Compile();
        
        func().Should().Be(15);
    }
    
    [Fact]
    public void ShouldComputeSimpleDoubleMath()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<double>("1.1 + 2.3");
        var func = parsed.Compile();
        
        func().Should().Be(3.4);
    }
    
    [Fact]
    public void ShouldComputeSimpleMixedMath()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<double>("1 + 2.3");
        var func = parsed.Compile();
        
        func().Should().Be(3.3);
    }
    
    [Fact]
    public void ShouldConcatTwoStrings()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<string>("\"aaaa\" + \"bbb\"");
        var func = parsed.Compile();
        
        func().Should().Be("aaaabbb");
    }

    [Fact]
    public void ShouldComputeScalarFunction()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<double>("sqrt(4)");
        var func = parsed.Compile();
        
        func().Should().Be(2);
    }
    
    [Fact]
    public void ShouldComputeScalarFunctionAndConvertToString()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<string>("text(sqrt(4))");
        var func = parsed.Compile();
        
        func().Should().Be("2");
    }
    
    [Fact]
    public void ShouldComputeScalarFunctionAndConvertToStringWithFormatting()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<string>("text(sqrt(4), \"0000\")");
        var func = parsed.Compile();
        
        func().Should().Be("0002");
    }
    
    [Fact]
    public void ShouldComputeLogicalFunction()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<bool>("or(true, and(true, false))");
        var func = parsed.Compile();
        
        func().Should().Be(true);
    }
    
    [Fact]
    public void ShouldComputeLogicalFunction01()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression<bool>("or(1, and(1, 0))");
        var func = parsed.Compile();
        
        func().Should().Be(true);
    }
   
    [Fact]
    public void ShouldNotSubtractTwoStrings()
    {
        var parser = new FormulaParser(new() { ParametersDelimiter = ',' });
        var exception = Assert.Throws<ParseException>(() => parser.ParseExpression<dynamic>("\"aaaa\" - \"bbb\""));

        Assert.Contains("Function 'Subtract(String, String)' does not exist.", exception.Message);
    }
}