namespace PowerPointToOBSSceneSwitcher;

internal static class Console
{
    public static void Error(Exception exception, string text) => WriteWithColorAndException(exception, text, ConsoleColor.Red, ConsoleColor.DarkRed);
    public static void Error(string text) => WriteWithColor(text, ConsoleColor.Red);

    public static void Warning(Exception exception, string text) => WriteWithColorAndException(exception, text, ConsoleColor.Yellow, ConsoleColor.DarkYellow);
    public static void Warning(string text) => WriteWithColor(text, ConsoleColor.Yellow);

    private static void WriteWithColorAndException(Exception exception, string text, ConsoleColor color1, ConsoleColor color2)
    {
        var c = System.Console.ForegroundColor;
        System.Console.ForegroundColor = color1;

        System.Console.WriteLine(text);
        System.Console.ForegroundColor = color2;
        System.Console.WriteLine(exception.Message);
        System.Console.WriteLine();

        System.Console.ForegroundColor = c;
    }

    private static void WriteWithColor(string text, ConsoleColor color1)
    {
        var c = System.Console.ForegroundColor;
        System.Console.ForegroundColor = color1;
        System.Console.WriteLine(text);
        System.Console.WriteLine();

        System.Console.ForegroundColor = c;
    }

    public static void WriteLine(string text) => System.Console.WriteLine(text);
    public static void Write(string text) => System.Console.Write(text);
    public static ConsoleKeyInfo ReadKey() => System.Console.ReadKey();
}