using System;
using System.Collections.Generic;
using System.Linq;
using MacroRunner.Compiler.VBA;
using Microsoft.CodeAnalysis;

namespace MacroRunner.Helpers
{
    public static class SyntaxTreeExtensions
    {
        public static void Print(this IEnumerable<SyntaxTree> trees)
        {
            foreach (var tree in trees)
            {
                tree.Print();
            }
        }

        public static void Print(this SyntaxTree tree)
        {
            var printer = new VBASyntaxPrinter();
            Console.WriteLine(new string('*', 30));

            Console.WriteLine($"{tree.FilePath} has {tree.GetRoot().DescendantNodesAndSelf().Count()} elements in syntax tree.");
            printer.Visit(tree.GetRoot());
        }
    }
}
