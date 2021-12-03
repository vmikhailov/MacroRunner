using System;
using System.Linq;
using MacroRunner.Helpers;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace MacroRunner.Compiler.VBA
{
    public class VBASyntaxPrinter : VisualBasicSyntaxWalker
    {
        private int _depth = 0;

        public override void Visit(SyntaxNode node)
        {
            var indent = new string(' ', _depth * 2);
            //Console.WriteLine($"{indent}{node.Kind()}");
            Console.Write($"{indent}{node.Kind()} ");
            var mbs = node as ModuleBlockSyntax;
            if (!node.ChildNodes().Any())
            {
                ConsolePlus.Write(ConsoleColor.Yellow, $"\"{node.ToFullString().RemoveCRLF()}\" ");
            }

            if (node.ChildTokens().Any())
            {
                var tokens = string.Join(", ", node.ChildTokens().Select(x => x.Kind()));
                Console.WriteLine($"({tokens})");
            }
            else
            {
                Console.WriteLine();
            }

            _depth++;
            base.Visit(node);
            _depth--;
        }
    }
}