//
//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.
//

namespace PowerPointToOBSSceneSwitcher;

internal class Program
{
    private static readonly PptLocal Ppt = new();
    private static readonly ObsLocal Obs = new();

    internal static KeyMappings KeyMappings = new();
    internal static DeckOperations DeckOperations = new(Obs, Ppt);

    private static void Main(string[] args)
    {
        ConnectToPowerPoint();
        ConnectToObs(args[0]);

        Obs.GetScenes();

        DeckUi.Start();

        WaitForCommandsV2();
    }

    private static void WaitForCommandsV2()
    {
        while (true)
        {
            var keyInfo = Console.ReadKey();
            if (KeyMappings.TryGetValue(new KeyInfo(keyInfo.Key, keyInfo.Modifiers), out var operation))
            {
                if (DeckOperations.TryGetValue(operation.Op, out var action))
                {
                    action(DeckOperations.ToDeckOperationParameters(operation.Parameters));
                }
            }
        }
    }

    private static void ConnectToObs(string password)
    {
        Console.Write("Connecting to OBS... ");
        Obs.Connect(password);
        Console.WriteLine("connected");
    }

    private static void ConnectToPowerPoint()
    {
        Console.Write("Connecting to PowerPoint... ");
        Ppt.Connect();
        Console.WriteLine("connected");
    }
}

public class KeyMappings : Dictionary<KeyInfo, Operation>
{
    public KeyMappings()
    {
        Add(new KeyInfo(ConsoleKey.F1), new("OBS.ToggleRecording"));
        Add(new KeyInfo(ConsoleKey.F1, ConsoleModifiers.Control), new("OBS.StopRecording"));
        Add(new KeyInfo(ConsoleKey.F2), new("OBS.SwitchScene", "scene=Iz swipe 2"));
        Add(new KeyInfo(ConsoleKey.LeftArrow), new("PPT.PreviousSlide"));
        Add(new KeyInfo(ConsoleKey.RightArrow), new("PPT.NextSlide"));
        Add(new KeyInfo(ConsoleKey.T, ConsoleModifiers.Control), new("PPT.ToggleBridge"));
        Add(new KeyInfo(ConsoleKey.F12), new("FJSD.Sequence", "seq=scripts\\swipe.seq"));
    }
}

public record KeyInfo(ConsoleKey ConsoleKey, ConsoleModifiers Modifiers = 0);
public record Operation(string Op, params string[] Parameters);