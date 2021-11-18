using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.AspNetCore.Builder;
using Microsoft.Office.Interop.PowerPoint;
using Newtonsoft.Json;

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
            //ConnectToPowerPoint();
            ConnectToObs(args[0]);

            _powerPointToObsBridgeOpen = true;

            Obs.GetScenes();

            var builder = WebApplication.CreateBuilder(args);
            var app = builder.Build();

            app.MapGet("/", () => JsonConvert.SerializeObject(DefaultMappings, Formatting.Indented));
            foreach (var mapping in DefaultMappings)
            {
                app.MapPost($"/op/{mapping.Value.Op}", () =>
                {
                    var op = DefaultOpertions[mapping.Value.Op];
                    op(mapping.Value.Op.StartsWith("OBS.") ? Obs : null);
                });
            }

            app.MapGet("/deck", async context =>
            {
                context.Response.ContentType = "text/html";
                await context.Response.Body.WriteAsync(Encoding.UTF8.GetBytes(GetDeckHtml()));
            });

            app.MapGet("/manifest.json", async context =>
            {
                context.Response.ContentType = "application/json";
                await context.Response.Body.WriteAsync(File.ReadAllBytes("deck\\manifest.json"));
            });

            app.MapGet("/sw.js", async context =>
            {
                context.Response.ContentType = "application/javascript";
                await context.Response.Body.WriteAsync(File.ReadAllBytes("deck\\sw.js"));
            });

            app.Run("http://0.0.0.0:5555");

            WaitForCommandsV2();
        }

        private static string GetDeckHtml()
        {
            var head = File.ReadAllText("deck\\head.html");
            var body = new StringBuilder("<body>");

            var buttonMatrix = File.ReadAllLines("deck\\buttonmatrix.html").ToList();

            var buttonIdx = 0;

            for (var i = 0; i < buttonMatrix.Count; i++)
            {
                if (!buttonMatrix[i].Contains($"<!--{buttonIdx}-->"))
                {
                    continue;
                }

                var mappingOp = DefaultMappings.Values.FirstOrDefault(x => x.Position == buttonIdx);
                if (mappingOp != null)
                {
                    buttonMatrix[i] = buttonMatrix[i].Replace("#text#", mappingOp.Op).Replace("#imagesrc#", mappingOp.Op.StartsWith("OBS.") ? OBSImage : PPTImage).Replace("#op#", mappingOp.Op);
                }
                else
                {
                    buttonMatrix[i] = "<div style='display:inline-block;' class='opbtn'></div>";
                }

                buttonIdx++;
            }
            
            body.Append(string.Join(Environment.NewLine, buttonMatrix));
            body.Append("</body>");

            return $"<!DOCTYPE html><html lang=\"en\" xmlns=\"http://www.w3.org/1999/xhtml\">{head}{body}</html>";
        }

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
            }
        }

        private static void SwitchSlide(string direction)
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

                Console.WriteLine($"{SmallTab}{from} to {Ppt.ActivePresentation.SlideShowWindow.View.Slide.SlideNumber}");
            }
            catch (Exception e)
            {
                Console.Error(e, "Exception caught while switching slide");
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
                            Console.WriteLine($"{SmallTab}Switching to OBS Scene named \"{line}\"");
                            try
                            {
                                sceneHandled = Obs.ChangeScene(line);
                            }
                            catch (Exception ex)
                            {
                                Console.Error(ex, "ERROR");
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

                    if (line.StartsWith("**PAUSE"))
                    {
                        Obs.PauseRecording();
                    }

                    if (!sceneHandled)
                    {
                        Obs.ChangeScene(Obs.DefaultScene);
                        Console.WriteLine($"{SmallTab}Switching to OBS Default Scene named \"{Obs.DefaultScene}\"");
                    }
                }
            }
        }

        public static KeyMap DefaultMappings = new()
        {
            {new KeyInfo(ConsoleKey.F1), new("OBS.ToggleRecording", 0)},
            {new KeyInfo(ConsoleKey.F1, ConsoleModifiers.Control), new("OBS.StopRecording", 2)},
            {new KeyInfo(ConsoleKey.LeftArrow), new("PPT.PreviousSlide", 4)},
            {new KeyInfo(ConsoleKey.RightArrow), new("PPT.NextSlide", 8)},
        };

        public static Operations DefaultOpertions = new()
        {
            {"OBS.ToggleRecording", obj => (obj as ObsLocal)?.StartPauseResumeRecording(true)},
            {"OBS.StopRecording", obj => (obj as ObsLocal)?.StopRecording()},
            {"PPT.PreviousSlide", _ => SwitchSlide(Backwards)},
            {"PPT.NextSlide", _ => SwitchSlide(Forward)}
        };

        private const string OBSImage = "https://upload.wikimedia.org/wikipedia/commons/7/78/OBS.svg";
        private const string PPTImage = "https://upload.wikimedia.org/wikipedia/commons/6/62/Microsoft_Office_PowerPoint_%282013%E2%80%932019%29.svg";
    }

    public class KeyMap : Dictionary<KeyInfo, Operation> {}

    public record Operation(string Op, int Position);

    public class Operations : Dictionary<string, Action<object>> {}

    public record KeyInfo(ConsoleKey ConsoleKey, ConsoleModifiers Modifiers = 0);
}