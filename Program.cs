using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

//Thanks to CSharpFritz and EngstromJimmy for their gists, snippets, and thoughts.

namespace PowerPointToOBSSceneSwitcher
{
    internal class Program
    {
        private const string SmallTab = "  ";

        private const string Forward = "forward";
        private const string Backwards = "backwards";

        private static readonly Application Ppt = new();
        private static readonly ObsLocal Obs = new();

        private static bool _powerPointToObsBridgeOpen;

        private static void Main(string[] args)
        {
            ConnectToPowerPoint();
            ConnectToObs(args[0]);

            _powerPointToObsBridgeOpen = true;

            Obs.GetScenes();

            WaitForCommands();
        }

        private static void WaitForCommands()
        {
            while (true)
            {
                var keyInfo = Console.ReadKey();
                if (keyInfo.Key == ConsoleKey.T)
                {
                    _powerPointToObsBridgeOpen = !_powerPointToObsBridgeOpen;

                    Console.CursorLeft--;
                    Console.WriteLine($"{SmallTab}---> Bridge is {(_powerPointToObsBridgeOpen ? "open" : "closed")}");
                }

                if (keyInfo.Key == ConsoleKey.RightArrow)
                {
                    SwitchSlide(Forward);
                }

                if (keyInfo.Key == ConsoleKey.LeftArrow)
                {
                    SwitchSlide(Backwards);
                }

                if (keyInfo.Key == ConsoleKey.Q && keyInfo.Modifiers == ConsoleModifiers.Shift)
                {
                    Console.CursorLeft--;
                    Console.WriteLine("Exiting");
                    return;
                }
            }
        }

        private static void SwitchSlide(string direction)
        {
            var from = $"Switching {direction} from {Ppt.ActivePresentation.SlideShowWindow.View.Slide.SlideNumber}";

            Ppt.ActivePresentation.SlideShowWindow.Activate();
            if (direction == Forward)
            {
                Ppt.ActivePresentation.SlideShowWindow.View.Next();
            }
            else
            {
                Ppt.ActivePresentation.SlideShowWindow.View.Previous();
            }

            Console.WriteLine($"{SmallTab}{from} to {Ppt.ActivePresentation.SlideShowWindow.View.Slide.SlideNumber}");
        }

        private static void ConnectToObs(string password)
        {
            Console.Write("Connecting to OBS...");
            Obs.Connect(password);
            Console.WriteLine("connected");
        }

        private static void ConnectToPowerPoint()
        {
            Console.Write("Connecting to PowerPoint...");
            Ppt.SlideShowNextSlide += App_SlideShowNextSlide;
            Console.WriteLine("connected");
        }

        private static void App_SlideShowNextSlide(SlideShowWindow slideShowWindow)
        {
            if (_powerPointToObsBridgeOpen && slideShowWindow != null)
            {
                Console.WriteLine($"Moved to Slide Number {slideShowWindow.View.Slide.SlideNumber}");

                //Text starts at Index 2 ¯\_(ツ)_/¯
                var note = string.Empty;
                try
                {
                    note = slideShowWindow.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text;
                }
                catch(Exception exception)
                {
                     Console.WriteLine($"ERROR: {exception.Message}");
                }

                var sceneHandled = false;
                
                var noteReader = new StringReader(note);
                string line;

                while ((line = noteReader.ReadLine()) != null)
                {
                    if (line.StartsWith("OBS:"))
                    {
                        line = line[4..].Trim();

                        if (!sceneHandled)
                        {
                            Console.WriteLine($"{SmallTab}Switching to OBS Scene named \"{line}\"");
                            try
                            {
                                sceneHandled = Obs.ChangeScene(line);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"{SmallTab}ERROR: {ex.Message}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"{SmallTab}WARNING: Multiple scene definitions found.  I used the first and have ignored \"{line}\"");
                        }
                    }

                    if (line.StartsWith("OBSDEF:"))
                    {
                        Obs.DefaultScene = line[7..].Trim();
                        Console.WriteLine($"{SmallTab}Setting the default OBS Scene to \"{Obs.DefaultScene}\"");
                    }

                    if (line.StartsWith("**START"))
                    {
                        Obs.StartRecording();
                    }

                    if (line.StartsWith("**STOP"))
                    {
                        Obs.StopRecording();
                    }

                    if (!sceneHandled)
                    {
                        Obs.ChangeScene(Obs.DefaultScene);
                        Console.WriteLine($"{SmallTab}Switching to OBS Default Scene named \"{Obs.DefaultScene}\"");
                    }
                }
            }
        }

    }
}