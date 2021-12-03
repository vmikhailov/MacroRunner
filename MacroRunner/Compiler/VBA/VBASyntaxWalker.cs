using System;
using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;

namespace MacroRunner.Compiler.VBA
{
    public class VBASyntaxWalker : VisualBasicSyntaxWalker
    {
        private bool _print = false;
        private int _depth = 0;

        public override void Visit(SyntaxNode node)
        {
            var text = node.ToFullString();

            if (_print)
            {
                var indent = new string(' ', _depth * 2);
                var tokens = string.Join(", ", node.ChildTokens().Select(x => x.Kind()));
                if (node.ChildNodes().Any())
                {
                    Console.WriteLine($"{indent}{node.Kind()} '{tokens}'");
                }
                else
                {
                    Console.WriteLine($"{indent}{node.Kind()} '{tokens}' '{node}'");
                }

            }

            if (_print)
            {
                _depth++;
            }

            base.Visit(node);

            if (_print)
            {
                _depth--;
            }

            if (_depth == 0 && _print)
            {
                _print = false;
            }
        }
    }
}