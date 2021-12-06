using System.Collections.Generic;
using System.Linq.Expressions;

namespace MacroRunner.Compiler;

public interface IParserContext
{
    IDictionary<string, ParameterExpression> Parameters { get; }
    ParameterExpression ExecutionContextParameter { get; set; }
}