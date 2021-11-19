//
//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.
//

namespace PowerPointToOBSSceneSwitcher
{
    internal class Program
    {
        internal static readonly PptLocal Ppt = new();
        internal static readonly ObsLocal Obs = new();

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
                        action(operation.Op.StartsWith("OBS.") ? Obs : Ppt);
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

        internal static KeyMappings KeyMappings = new();
        internal static DeckOperations DeckOperations = new();
    }
}