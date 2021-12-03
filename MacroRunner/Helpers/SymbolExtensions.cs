using System.Linq;
using Microsoft.CodeAnalysis;

namespace MacroRunner.Helpers
{
    public static class SymbolExtensions
    {
        public static string FullName(this ITypeSymbol symbol)
        {
            var nmsp = symbol.ContainingNamespace.FullName();
            if (symbol.Name != "Void")
            {
                return string.IsNullOrEmpty(nmsp) ? symbol.Name : nmsp + "." + symbol.Name;
            }
            else
            {
                return string.Empty;
            }
        }

        public static string FullName(this INamespaceSymbol symbol)
        {
            if (symbol == null)
            {
                return null;
            }

            if (!string.IsNullOrEmpty(symbol.ContainingNamespace?.Name))
            {
                return symbol.ContainingNamespace.FullName() + "." + symbol.Name;
            }
            else
            {
                return symbol.Name;
            }
        }

        public static string FullName(this IPropertySymbol symbol)
        {
            return $"{symbol.ContainingType.Name}.{symbol.Name} : {symbol.Type.FullName()}";
        }

        public static string FullName(this IFieldSymbol symbol)
        {
            return $"{symbol.ContainingType.Name}.{symbol.Name} : {symbol.Type.FullName()}";
        }

        public static string FullName(this IParameterSymbol symbol)
        {
            return $"{symbol.Name} : {symbol.Type.FullName()}";
        }

        public static string FullName(this IMethodSymbol symbol)
        {
            var argsList = string.Join(", ", symbol.TypeArguments.Select(a => a.Name));

            return $"{symbol.ContainingType.Name}.{symbol.Name}({argsList}) : {symbol.ReturnType.FullName()}";
        }

        public static string FullName(this ISymbol symbol)
        {
            var staticFlag = symbol.IsStatic ? " *" : "";
            return symbol switch
            {
                INamedTypeSymbol typeSymbol => typeSymbol.FullName() + staticFlag,
                IFieldSymbol filedSymbol => filedSymbol.FullName() + staticFlag,
                IPropertySymbol propertySymbol => propertySymbol.FullName() + staticFlag,
                IMethodSymbol methodSymbol => methodSymbol.FullName() + staticFlag,
                IParameterSymbol parameterSymbol => parameterSymbol.FullName() + staticFlag,
                _ => $"<unknown>"
            };
        }
    }
}