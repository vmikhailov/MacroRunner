using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using MacroRunner.Compiler.VBA;
using MacroRunner.Helpers;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using MoreLinq;
using RazorLight;

namespace MacroRunner.Compiler
{
    public class MacroCompiler
    {
        private static readonly string VBAClassTemplateName = $"Templates.VbaClassTemplate.cshtml";
        private static readonly string VBAModuleTemplateName = $"Templates.VbaModuleTemplate.cshtml";

        private static readonly MetadataReference[] References =
        {
            MetadataReference.CreateFromFile(typeof(object).Assembly.Location),
            MetadataReference.CreateFromFile(Path.Combine(Path.GetDirectoryName(typeof(object).Assembly.Location), "System.Runtime.dll")),
            MetadataReference.CreateFromFile(typeof(Enumerable).Assembly.Location),
            MetadataReference.CreateFromFile(typeof(IComponent).Assembly.Location),
            MetadataReference.CreateFromFile(Assembly.GetExecutingAssembly().Location)
        };

        private readonly List<SyntaxTree> _syntaxTrees;
        private readonly List<GlobalImport> _imports;
        private readonly ISet<Type> _globals;
        private readonly RazorLightEngine _razor;

        public MacroCompiler(string name)
        {
            Name = name;
            _syntaxTrees = new List<SyntaxTree>();
            _imports = new List<GlobalImport> { GlobalImport.Parse("System") };
            _globals = new HashSet<Type>();
            _razor = new RazorLightEngineBuilder()
                     .UseEmbeddedResourcesProject(typeof(MacroCompiler))
                     .UseMemoryCachingProvider()
                     .Build();
        }

        public IEnumerable<Type> Globals => _globals;

        public IReadOnlyCollection<GlobalImport> Imports => _imports;

        public string Name { get; }

        public IReadOnlyCollection<SyntaxTree> SyntaxTrees => _syntaxTrees;

        public MacroCompiler AddClass(string className, string classText)
        {
            Add(className, classText, VBAClassTemplateName);

            return this;
        }

        public MacroCompiler AddClasses(IEnumerable<TextResource> resource, string prefix = null)
        {
            resource.ForEach(x => AddClass(prefix + x.Name, x.Text));

            return this;
        }

        public MacroCompiler AddImport(string import)
        {
            _imports.Add(GlobalImport.Parse(import));

            return this;
        }

        public MacroCompiler AddModule(string moduleName, string moduleText)
        {
            Add(moduleName, moduleText, VBAModuleTemplateName);

            return this;
        }

        public MacroCompiler AddModules(IEnumerable<TextResource> resource, string prefix = null)
        {
            resource.ForEach(x => AddModule(prefix + x.Name, x.Text));

            return this;
        }

        public CompilationResult Compile()
        {
            if (!SyntaxTrees.Any())
            {
                throw new Exception("At least one syntax tree must be provided");
            }

            var compilationOptions =
                new VisualBasicCompilationOptions(OutputKind.DynamicallyLinkedLibrary)
                    .WithModuleName(Name)
                    .WithGlobalImports(Imports)
                    .WithPlatform(Platform.AnyCpu)
                    .WithEmbedVbCoreRuntime(true)
                    .WithOptionStrict(OptionStrict.Off)
                    .WithMetadataImportOptions(MetadataImportOptions.Public);

            var compilation = VisualBasicCompilation.Create(
                Name,
                SyntaxTrees,
                References,
                compilationOptions);

            foreach (var tree in SyntaxTrees)
            {
                var newTree = PatchVBAtoBecomeVBNet(tree, compilation.GetSemanticModel(tree));
                if (newTree != null)
                {
                    compilation = compilation.ReplaceSyntaxTree(tree, newTree);
                }
            }

            return Emit(compilation);
        }

        public MacroCompiler WithGlobals<T>()
        {
            _globals.Add(typeof(T));

            return this;
        }

        private void Add(string name, string body, string templateName)
        {
            var preProcessedText = _razor.CompileRenderAsync(
                                             templateName,
                                             new CodeBlock { Name = name, Body = body })
                                         .Result;

            var tree = VisualBasicSyntaxTree.ParseText(
                preProcessedText,
                new VisualBasicParseOptions()
                    .WithKind(SourceCodeKind.Regular),
                name);
            tree.Print();
            var postProcessedTree = PostProcess(tree);
            postProcessedTree.Print();

            _syntaxTrees.Add(postProcessedTree);
        }

        private CompilationResult Emit(Compilation compilation)
        {
            using (var ms = new MemoryStream())
            {
                var result = compilation.Emit(ms);

                ms.Seek(0, SeekOrigin.Begin);
                var compiled = result.Success ? new CompiledCode(Name, ms) : null;

                return new CompilationResult(result.Diagnostics, compiled);
            }
        }

        private SyntaxTree PatchVBAtoBecomeVBNet(SyntaxTree tree, SemanticModel model)
        {
            SyntaxNode CreateTimeSpanConversionSyntax(BinaryExpressionSyntax x)
            {
                // ToDate(x)
                return SyntaxFactory.InvocationExpression(
                    SyntaxFactory.IdentifierName("ToDate"),
                    SyntaxFactory.ArgumentList(
                        SyntaxFactory.SeparatedList<ArgumentSyntax>()
                                     .Add(SyntaxFactory.SimpleArgument(x.WithoutTrivia()))));
            }

            SyntaxNode CreateGlobalsMemberAccessSyntax(IdentifierNameSyntax name)
            {
                return SyntaxFactory.MemberAccessExpression(
                                        SyntaxKind.SimpleMemberAccessExpression,
                                        SyntaxFactory.IdentifierName(@"Globals"),
                                        SyntaxFactory.Token(SyntaxKind.DotToken),
                                        SyntaxFactory.IdentifierName(name.Identifier.ToString()))
                                    .WithLeadingTrivia(name.GetLeadingTrivia())
                                    .WithTrailingTrivia(name.GetTrailingTrivia());
            }

            //tree.PrintSymbols(model);
            //Console.WriteLine("Static");
            //tree.PrintSymbols(model, n => model.GetSymbolInfo(n).Symbol?.IsStatic ?? false);
            //tree.PrintSymbols(model, n => n.Kind() == SyntaxKind.IdentifierName && model.GetSymbolInfo(n).Symbol == null);

            var globals = tree.GetRoot()
                              .DescendantNodes()
                              .OfType<IdentifierNameSyntax>()
                              .Select(x => new { symbol = model.GetSymbolInfo(x).Symbol, syntax = x })
                              .Where(x => x.symbol?.ContainingType?.TypeKind == TypeKind.Module)
                              .ToList();

            if (globals.Any())
            {
                var toConvert = globals.Select(x => x.syntax);
                var newRoot = tree.GetRoot().ReplaceNodes(toConvert, (x, y) => CreateGlobalsMemberAccessSyntax(x));

                return newRoot.SyntaxTree;
            }

            var allNodes = tree.GetRoot().DescendantNodes().ToList();

            var staticVarsToReplace = allNodes.OfType<IdentifierNameSyntax>()
                                              .Where(x => model.GetSymbolInfo(x).Symbol?.IsStatic ?? false)
                                              .ToList();

            // check all nodes which return TimeSpan
            // VBA expects them to be autoconverted to DateTime
            var expToConvert = allNodes.OfType<BinaryExpressionSyntax>()
                                       .Where(x => !model.GetConversion(x).Exists) // check what conversion is needed
                                       .Select(
                                           x => new
                                           {
                                               node = x,
                                               symbol = (IMethodSymbol)model.GetSymbolInfo(x).Symbol
                                           })
                                       .Where(x => x.symbol?.ReturnType.FullName() == typeof(TimeSpan).FullName)
                                       .Select(x => x.node)
                                       .ToList();

            if (expToConvert.Any())
            {
                var newRoot = tree.GetRoot().ReplaceNodes(expToConvert, (x, y) => CreateTimeSpanConversionSyntax(x));

                return newRoot.SyntaxTree;
            }

            //var text = newRoot.NormalizeWhitespace().ToFullString();

            return null;
        }

        private SyntaxTree PostProcess(SyntaxTree tree)
        {
            var rewriter = new VBASyntaxRewriter();
            var root = rewriter.Visit(tree.GetRoot());

            return root.SyntaxTree.WithFilePath(tree.FilePath);
        }
    }
}