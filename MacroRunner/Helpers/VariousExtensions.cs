using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;

namespace MacroRunner.Helpers;

public static class VariousExtensions
{
    public static void PrintExposedTypes(this Assembly assembly)
    {
        Console.WriteLine($"Exposed types {assembly.GetTypes().Count()} are:");
        foreach (var type in assembly.GetTypes())
        {
            Console.WriteLine($" {type.FullName}");
            foreach (var member in type.GetFields())
            {
                Console.WriteLine($"\t{member.MemberType} {member.Name} : {member.FieldType}");
            }
            foreach (var member in type.GetProperties())
            {
                Console.WriteLine($"\t{member.MemberType} {member.Name} : {member.PropertyType}");
            }
            foreach (var member in type.GetMethods())
            {
                Console.WriteLine($"\t{member.MemberType} {member.Name} : {member.ReturnType}");
            }
        }
    }

    public static void PrintErrors(this IEnumerable<Diagnostic> diags)
    {
        diags.Where(x => x.IsWarningAsError || x.Severity == DiagnosticSeverity.Error)
             .PrintDiagnostics();
    }

    public static void PrintWarnings(this IEnumerable<Diagnostic> diags)
    {
        diags.Where(x => !x.IsWarningAsError && x.Severity != DiagnosticSeverity.Error)
             .PrintDiagnostics();
    }

    public static void PrintDiagnostics(this IEnumerable<Diagnostic> diags)
    {
        foreach (var diagnostic in diags)
        {
            var pos = diagnostic.Location.GetLineSpan().Span.Start;
            Console.Error.WriteLine("{0}, {1}: {2} at {3}/{4} {5}",
                diagnostic.Location.SourceTree.FilePath,
                diagnostic.Id,
                diagnostic.Severity,
                pos.Line,
                pos.Character,
                diagnostic.GetMessage());
        }
    }

    public static void PrintSymbols(this SyntaxTree tree, SemanticModel model, Func<SyntaxNode, bool>? filter = null)
    {
        var nodes = tree.GetRoot().DescendantNodesAndSelf();
        if (filter != null)
        {
            nodes = nodes.Where(filter);
        }
        var nodeList = nodes.ToList();
        Console.WriteLine($"Symbols of {tree.FilePath}: {nodeList.Count}");
            
        foreach (var node in nodeList)
        {
            var symbol = model.GetSymbolInfo(node).Symbol;
            var indent = GetNodeIndent(node);
            var indentStr = new string(' ', indent);
            var hasError = !model.GetConversion(node).Exists;
            var nodeText = node.Kind().ToString();
                
            Console.Write($"{indentStr}{nodeText.PadRight(45 - indent)} | ");

            if (symbol != null)
            {
                var symbolText = symbol.FullName();
                if (hasError)
                {
                    symbolText += "?";
                }

                Console.WriteLine($"{node} | {symbol.Kind} | {symbolText}");
            }
            else
            {
                // if (s.node.ChildNodes().Any())
                if (!IsSingleLine(node.WithoutTrivia().ToFullString()))
                {
                    Console.WriteLine($"complex");
                }
                else
                {
                    Console.WriteLine($"{node}");
                }
            }
        }

        Console.WriteLine();
    }

    public static string RemoveCRLF(this string str)
    {
        return str.Replace("\n", @"\n").Replace("\r", @"\r");
    }

    private static int GetNodeIndent(SyntaxNode node)
    {
        var indent = 0;
        while (node.Parent != null)
        {
            indent++;
            node = node.Parent;
        }

        return indent;
    }

    private static bool IsSingleLine(string s)
    {
        s = s.Trim('\n', '\r');

        return !s.Contains('\n');
    }

}