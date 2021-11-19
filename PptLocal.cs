using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointToOBSSceneSwitcher;

internal class PptLocal
{
    private static readonly Application Ppt = new();
    private const string Forward = "forward";

    public void Connect()
    {
        Console.Write("Connecting to PowerPoint... ");
        Ppt.SlideShowNextSlide += App_SlideShowNextSlide;
        Console.WriteLine("connected");
    }

    private static void App_SlideShowNextSlide(SlideShowWindow slideShowWindow)
    {
        if (Program.PowerPointToObsBridgeOpen && slideShowWindow != null)
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
                            sceneHandled = Program.Obs.ChangeScene(line);
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

                if (line.StartsWith("OBSDEF:"))
                {
                    Program.Obs.DefaultScene = line[7..].Trim();
                    Console.WriteLine($"Setting the default OBS Scene to \"{Program.Obs.DefaultScene}\"");
                }

                if (line.StartsWith("**START"))
                {
                    Program.Obs.StartRecording();
                }

                if (line.StartsWith("**STOP"))
                {
                    Program.Obs.StopRecording();
                }

                if (line.StartsWith("**PAUSE"))
                {
                    Program.Obs.PauseRecording();
                }

                if (!sceneHandled)
                {
                    Program.Obs.ChangeScene(Program.Obs.DefaultScene);
                    Console.WriteLine($"Switching to OBS Default Scene named \"{Program.Obs.DefaultScene}\"");
                }
            }
        }
    }

    public void SwitchSlide(string direction)
    {
        try
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

            Console.WriteLine($"{from} to {Ppt.ActivePresentation.SlideShowWindow.View.Slide.SlideNumber}");
        }
        catch (Exception e)
        {
            Console.Error(e, "Exception caught while switching slide");
        }
    }
}