using System;

namespace PowerPointToOBSSceneSwitcher;

internal static class Console
{
    public static void Error(Exception exception, string text)
    {
        var c = System.Console.ForegroundColor;
        System.Console.ForegroundColor = ConsoleColor.Red;

        System.Console.WriteLine(text);
        System.Console.ForegroundColor = ConsoleColor.DarkRed;
        System.Console.WriteLine(exception.Message);
        System.Console.WriteLine();

        System.Console.ForegroundColor = c;
    }

    public static void WriteLine(string text) => System.Console.WriteLine(text);
    public static void Write(string text) => System.Console.Write(text);
    public static ConsoleKeyInfo ReadKey() => System.Console.ReadKey();
    public static int CursorLeft
    {
        get => System.Console.CursorLeft;
        set => System.Console.CursorLeft = value;
    }
}