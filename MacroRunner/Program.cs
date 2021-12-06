using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using MacroRunner.Compiler;
using MacroRunner.Helpers;
using MacroRunner.Runtime;

namespace MacroRunner;

internal class Program
{
    private static void Main(string[] args)
    {
        ParseExcel();
        var parser = new FormulaParser();
            
        // var addr1 = FormulaParser.Range.Parse("A1");
        // var addr2 = FormulaParser.Range.Parse("AAA1");
        // var addr3 = FormulaParser.Range.Parse("$B1");
        // var addr4 = FormulaParser.Range.Parse("C$4");
        // var addr5 = FormulaParser.Range.Parse("$D$47");
        // var addr6 = FormulaParser.Range.Parse("A1:B4");
        //
        // var raddr1 = FormulaParser.RangeAddress.Parse("A1:$D$47");


        //var parsed = FormulaParser.ParseExpression("1 + 2");

        //var parsed1 = FormulaParser.ParseExpression("vlookup(1, \"value\", 3)");
        var sw = Stopwatch.StartNew();
        int run = 1;
        // var parsed1 = parser.ParseExpression("A1");
        // Console.WriteLine($"Run {run++}, Elapsed {sw.Elapsed}");
        // var parsed2 = parser.ParseExpression("A1:B3");
        // Console.WriteLine($"Run {run++}, Elapsed {sw.Elapsed}");

        // var parsed3 = parser.ParseExpression("vlookup2(A1, \"value\", 3)");
        // Console.WriteLine($"Run {run++}, Elapsed {sw.Elapsed}");
        //
        // var parsed4 = parser.ParseExpression("vlookup2(22, \"value\", 3)");
        // Console.WriteLine($"Run {run++}, Elapsed {sw.Elapsed}");

        //var parsed5 = parser.ParseExpression("sum(A1:B1) * sum(A2:ZZ3332) * 19 + 7");
        var parsed5 = parser.ParseExpression<int>("1 + 2");
        Console.WriteLine($"Run {run++}, Elapsed {sw.Elapsed}");
            
        //var parsed6 = parser.ParseExpression("100/(10 % 4) - 47.0 + 1");
        var parsed6 = parser.ParseExpression<double>("100.0 / (10 % 4) - (49 + (2 - 1 - 2 + 1 + 2 + 3 - 3))");
        Console.WriteLine($"Run {run++}, Elapsed {sw.Elapsed}");

        var fnc = parsed6.Compile();
        Console.WriteLine($"Compilation, Elapsed {sw.Elapsed}");

        var xx = fnc();
        var n = 0L;
        var ss = sw.Elapsed;
        var nn = 100_000_000;
        for (var i = 0; i < nn; i++)
        {
            var x = fnc();
            n += (long)x;
        }
        Console.WriteLine($"Run {nn}, {n}, Mln Per second {nn/(sw.Elapsed-ss).TotalMilliseconds/1000:F5}");

        return;
        string NameExtract(string x)
        {
            return x.Split('.').TakeLast(2).First();
        }

        var macro = new MacroCompiler("test")
                    .AddModules(ResourceHelper.GetResourcesByMask(@"Tests\.Test1\.Modules\..+\.vba", NameExtract), "Module_")
                    .AddClasses(ResourceHelper.GetResourcesByMask(@"Tests\.Test1\.Classes\..+\.vba", NameExtract), "Class_")
                    .WithGlobals<ExcelGlobals>();

        Console.WriteLine();

        //macro.SyntaxTrees.Print();

        var result = macro.Compile();

        var compiled = result.CompiledCode;

        if (result.CompiledCode == null)
        {
            result.Diagnostics.PrintErrors();
            result.Diagnostics.PrintWarnings();
        }
        else
        {
            // entry point
            var testClass = compiled.Assembly.GetTypes()
                                    .SingleOrDefault(x => x.Name.EndsWith("TestClass"));

            var inst1 = Activator.CreateInstance(testClass);
            inst1.SetField("Globals", new ExcelGlobals());
            var g1 = inst1.GetField<ExcelGlobals>("Globals");

            var inst2 = Activator.CreateInstance(testClass);
            inst2.SetField("Globals", new ExcelGlobals());
            var g2 = inst2.GetField<ExcelGlobals>("Globals");

            inst1.CallVoid("Test");
            var gg = g2.TestInt != g1.TestInt;

            // init globals
            var macroInstance = compiled.Create<ExcelMacroInstance>();
            var file = File.Open("", FileMode.Open);
            macroInstance.LoadFile(file);
            macroInstance.Run("method_name");

            // entry point
            var testGlobal = compiled.Assembly.GetTypes()
                                     .Select(x => x.GetField("TestValue"))
                                     .Where(x => x != null)
                                     .ToList();
            var i = 1;
            foreach (var t in testGlobal)
            {
                t.SetValue(null, i++);
            }

            foreach (var t in testGlobal)
            {
                i = (int)t.GetValue(null);
                Console.WriteLine(i);
            }

            // entry point
            var main = compiled.Assembly.GetTypes()
                               .Select(x => x.GetMethod("Main"))
                               .Single(x => x != null);

            main.Invoke(null, null);
            foreach (var t in testGlobal)
            {
                i = (int)t.GetValue(null);
                Console.WriteLine(i);
            }

            //var mainSet = result.Assembly.GetTypes()
            //                    .Select(x => x.GetMethod("MainSetValue"))
            //                    .Single(x => x != null);
            //var mainGet = result.Assembly.GetTypes()
            //                    .Select(x => x.GetMethod("MainGetValue"))
            //                    .Single(x => x != null);

            //mainSet.Invoke(null, null);
            //var r1 = mainGet.Invoke(null, null);

            //var main2 = result2.Assembly.GetTypes()
            //                   .Select(x => x.GetMethod("MainGetValue"))
            //                   .Single(x => x != null);

            //var r2 = main2.Invoke(null, null);
        }

        Console.ReadLine();
    }

    private static void ParseExcel()
    {
        var exStream = ResourceHelper.GetResourceAsStream("MacroRunner.Samples.Excel.Book1.xlsm");
        var ex = new XLWorkbook(exStream);
        var ws = ex.Worksheets.First();

        var r = ws.Cell(1, 1);
    }
}