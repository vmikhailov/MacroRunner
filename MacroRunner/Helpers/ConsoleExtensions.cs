using System;

namespace MacroRunner.Helpers;

public static class ConsolePlus
{
    public static void Write(ConsoleColor color, string format, params object[] parameters)
    {
        var currentColor = Console.ForegroundColor;
        Console.ForegroundColor = color;
        Console.Write(format, parameters);
        Console.ForegroundColor = currentColor;
    }

    public static void WriteLine(ConsoleColor color, string format, params object[] parameters)
    {
        Write(color, format, parameters);
        Console.WriteLine();
    }
}