using System.Collections.Generic;
using Microsoft.CodeAnalysis;

namespace MacroRunner.Compiler
{
    public class CompilationResult
    {
        public CompilationResult(IReadOnlyCollection<Diagnostic> diagnostics, CompiledCode compiled)
        {
            Diagnostics = diagnostics;
            CompiledCode = compiled;
        }

        public IReadOnlyCollection<Diagnostic> Diagnostics { get; }

        public CompiledCode CompiledCode { get; }
    }
}