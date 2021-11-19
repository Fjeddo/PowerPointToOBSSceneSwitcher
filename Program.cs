using System;
using System.Collections.Generic;

//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.

namespace PowerPointToOBSSceneSwitcher
{
    internal class Program
    {
        private const string Forward = "forward";
        private const string Backwards = "backwards";

        private static readonly PptLocal Ppt = new();
        public static readonly ObsLocal Obs = new();
        private static readonly DeckUi DeckUi = new();

        public static bool PowerPointToObsBridgeOpen;

        private static void Main(string[] args)
        {
            //ConnectToPowerPoint();
            ConnectToObs(args[0]);

            PowerPointToObsBridgeOpen = true;

            Obs.GetScenes();

            SetupDeckUi();

            WaitForCommandsV2();
        }

        private static void SetupDeckUi() => DeckUi.Start();

        private static void WaitForCommandsV2()
        {
            while (true)
            {
                var keyInfo = Console.ReadKey();
                if(DefaultMappings.TryGetValue(new KeyInfo(keyInfo.Key, keyInfo.Modifiers), out var operation))
                {
                    if (DefaultOpertions.TryGetValue(operation.Op, out var action))
                    {
                        action(operation.Op.StartsWith("OBS.") ? Obs : null);
                    }
                }

                if (keyInfo.Key == ConsoleKey.C && keyInfo.Modifiers == ConsoleModifiers.Control)
                {
                    return;
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

        public static KeyMap DefaultMappings = new()
        {
            {new KeyInfo(ConsoleKey.F1), new("OBS.ToggleRecording", 0)},
            {new KeyInfo(ConsoleKey.F1, ConsoleModifiers.Control), new("OBS.StopRecording", 3)},
            {new KeyInfo(ConsoleKey.LeftArrow), new("PPT.PreviousSlide", 1)},
            {new KeyInfo(ConsoleKey.RightArrow), new("PPT.NextSlide", 4)},
        };

        public static Operations DefaultOpertions = new()
        {
            {"OBS.ToggleRecording", obj => (obj as ObsLocal)?.StartPauseResumeRecording(true)},
            {"OBS.StopRecording", obj => (obj as ObsLocal)?.StopRecording()},
            {"PPT.PreviousSlide", obj => (obj as PptLocal)?.SwitchSlide(Backwards)},
            {"PPT.NextSlide", obj => (obj as PptLocal)?.SwitchSlide(Forward)}
        };

        public const string ObsImage = "https://upload.wikimedia.org/wikipedia/commons/7/78/OBS.svg";
        public const string PptImage = "https://upload.wikimedia.org/wikipedia/commons/6/62/Microsoft_Office_PowerPoint_%282013%E2%80%932019%29.svg";
    }

    public class KeyMap : Dictionary<KeyInfo, Operation> {}

    public record Operation(string Op, int Position);

    public class Operations : Dictionary<string, Action<object>> {}

    public record KeyInfo(ConsoleKey ConsoleKey, ConsoleModifiers Modifiers = 0);
}