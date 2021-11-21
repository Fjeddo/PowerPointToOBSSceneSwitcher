using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointToOBSSceneSwitcher;

public class PptLocal
{
    private readonly Application _ppt = new();
    public const string Forward = "forward";
    public const string Backwards = "backwards";

    private bool _powerPointToObsBridgeOpen = true;

    public void Connect() => _ppt.SlideShowNextSlide += App_SlideShowNextSlide;

    private void App_SlideShowNextSlide(SlideShowWindow slideShowWindow)
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
            catch (Exception exception)
            {
                Console.Error(exception, "ERROR");
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
                        Console.WriteLine($"Switching to OBS Scene named \"{line}\"");
                        try
                        {
                            sceneHandled = Program.DeckOperations["OBS.SwitchScene"](DeckOperations.ToDeckOperationParameters($"scene={line}"));
                        }
                        catch (Exception ex)
                        {
                            Console.Error(ex, "ERROR");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"WARNING: Multiple scene definitions found.  I used the first and have ignored \"{line}\"");
                    }
                }

                //if (line.StartsWith("OBSDEF:"))
                //{
                //    Program.Obs.DefaultScene = line[7..].Trim();
                //    Console.WriteLine($"Setting the default OBS Scene to \"{Program.Obs.DefaultScene}\"");
                //}

                if (line.StartsWith("**START"))
                {
                    Program.DeckOperations["OBS.StartRecording"](null);
                }

                if (line.StartsWith("**STOP"))
                {
                    Program.DeckOperations["OBS.StopRecording"](null);
                }

                if (line.StartsWith("**PAUSE"))
                {
                    Program.DeckOperations["OBS.PauseRecording"](null);
                }

                //if (!sceneHandled)
                //{
                //    Program.Obs.ChangeScene(Program.Obs.DefaultScene);
                //    Console.WriteLine($"Switching to OBS Default Scene named \"{Program.Obs.DefaultScene}\"");
                //}
            }
        }
    }

    public bool ToggleBridge()
    {
        _powerPointToObsBridgeOpen = !_powerPointToObsBridgeOpen;
        Console.WriteLine($"Ppt to Obs bridge is {(_powerPointToObsBridgeOpen ? "enabled" : "disabled")}");

        return true;
    }

    public bool SwitchSlide(string direction = Forward)
    {
        try
        {
            var from = $"Switching {direction} from {_ppt.ActivePresentation.SlideShowWindow.View.Slide.SlideNumber}";

            _ppt.ActivePresentation.SlideShowWindow.Activate();
            if (direction == Forward)
            {
                _ppt.ActivePresentation.SlideShowWindow.View.Next();
            }
            else
            {
                _ppt.ActivePresentation.SlideShowWindow.View.Previous();
            }

            Console.WriteLine($"{from} to {_ppt.ActivePresentation.SlideShowWindow.View.Slide.SlideNumber}");

            return true;
        }
        catch (Exception e)
        {
            Console.Error(e, "Exception caught while switching slide");
        }

        return false;
    }
}