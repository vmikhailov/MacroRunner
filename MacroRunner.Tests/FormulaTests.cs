using System;
using FluentAssertions;
using MacroRunner.Compiler;
using Xunit;

namespace MacroRunner.Tests;

public class FormulaTests
{
    [Fact]
    public void ShouldComputeSimpleIntMath()
    {
        var parser = new FormulaParser<decimal>(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression("1+2*(3+4)");
        var func = parsed.Compile();
        
        ((decimal)func()).Should().Be(15);
    }
    
    [Fact]
    public void ShouldComputeSimpleDecimalMath()
    {
        var parser = new FormulaParser<decimal>(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression("1.1 + 2.3");
        var func = parsed.Compile();
        
        ((decimal)func()).Should().Be(3.4m);
    }
    
    [Fact]
    public void ShouldComputeSimpleMixedMath()
    {
        var parser = new FormulaParser<decimal>(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression("1 + 2.3");
        var func = parsed.Compile();
        
        ((decimal)func()).Should().Be(3.3m);
    }
    
    [Fact]
    public void ShouldConcatTwoStrings()
    {
        var parser = new FormulaParser<string>(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression("\"aaaa\" + \"bbb\"");
        var func = parsed.Compile();
        
        func().Should().Be("aaaabbb");
    }

    [Fact]
    public void ShouldUseScalarFunction()
    {
        var parser = new FormulaParser<string>(new() { ParametersDelimiter = ',' });
        var parsed = parser.ParseExpression("\"aaaa\" + \"bbb\"");
        var func = parsed.Compile();
        
        func().Should().Be("aaaabbb");
    }
   
    [Fact]
    public void ShouldNotSubtractTwoStrings()
    {
        var parser = new FormulaParser<dynamic>(new() { ParametersDelimiter = ',' });
        var exception = Assert.Throws<InvalidOperationException>(() => parser.ParseExpression("\"aaaa\" - \"bbb\""));

        Assert.Contains("No method", exception.Message);
    }
}