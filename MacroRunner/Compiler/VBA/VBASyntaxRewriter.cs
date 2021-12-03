using System.Linq;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;

namespace MacroRunner.Compiler.VBA
{
    public class VBASyntaxRewriter : VisualBasicSyntaxRewriter
    {
        public override SyntaxNode VisitOptionStatement(OptionStatementSyntax node)
        {
            return SyntaxFactory.EmptyStatement();
        }

        public override SyntaxNode VisitEmptyStatement(EmptyStatementSyntax node)
        {
            bool IsLetOrSet(SyntaxTrivia trivia) 
            {
                var triviaToLower = trivia.ToFullString().ToLower();
                return triviaToLower == "let" || triviaToLower == "set";
            }

            var text = node.ToFullString();
            if (!string.IsNullOrWhiteSpace(text))
            {
                var token = node.ChildTokens().First();
                var trivia = token.GetAllTrivia().ToList();
                var hasLetSet = trivia.Any(IsLetOrSet);

                if (hasLetSet)
                {
                    var otherTrivials = trivia.Where(x => !IsLetOrSet(x)).Select(x => x.ToFullString()).ToList();
                    var newText = string.Join("", otherTrivials);
                    var newNode = SyntaxFactory.ParseExecutableStatement(newText);

                    return newNode;
                }
            }

            return base.VisitEmptyStatement(node);
        }

        public override SyntaxNode VisitArgumentList(ArgumentListSyntax node)
        {
            var tokens = node.ChildTokens().ToList();
            if (tokens.First().Text == "(" && tokens.Last().Text == ")")
            {
                return base.VisitArgumentList(node);
            }

            return node.WithOpenParenToken(SyntaxFactory.Token(SyntaxKind.OpenParenToken, "("))
                       .WithCloseParenToken(SyntaxFactory.Token(SyntaxKind.CloseParenToken, ")"));

        }
    }
}