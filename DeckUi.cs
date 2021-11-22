using System.Text;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;

namespace PowerPointToOBSSceneSwitcher;

public class DeckUi
{
    private static WebApplication _app;

    public static void Start()
    {
        var builder = WebApplication.CreateBuilder();
        builder.Logging.SetMinimumLevel(LogLevel.Warning);

        _app = builder.Build();

        _app.MapGet("/", () => JsonConvert.SerializeObject(Program.KeyMappings, Formatting.Indented));

        MapOperations(_app);

        _app.MapGet("/deck", GetDeck());
        _app.MapGet("/manifest.json", GetManifest());
        _app.MapGet("/sw.js", GetServiceWorker());

        _app.RunAsync("http://0.0.0.0:5555");
    }

    private static void MapOperations(WebApplication app)
    {
        foreach (var mapping in Program.KeyMappings)
        {
            app.MapPost($"/op/{mapping.Value.Op}", context =>
            {
                var op = Program.DeckOperations[mapping.Value.Op];
                op(context.Request.Query.ToDictionary(s => s.Key, s => s.Value.ToString()));

                return Task.CompletedTask;
            });
        }
    }

    private static RequestDelegate GetServiceWorker() =>
        async context =>
        {
            context.Response.ContentType = "application/javascript";
            await context.Response.Body.WriteAsync(await File.ReadAllBytesAsync("deck\\sw.js"));
        };

    private static RequestDelegate GetManifest() =>
        async context =>
        {
            context.Response.ContentType = "application/json";
            await context.Response.Body.WriteAsync(await File.ReadAllBytesAsync("deck\\manifest.json"));
        };

    private static RequestDelegate GetDeck() =>
        async context =>
        {
            context.Response.ContentType = "text/html";
            await context.Response.Body.WriteAsync(Encoding.UTF8.GetBytes(GetDeckHtml()));
        };

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

            var mappingOp = DeckUiOperations.FirstOrDefault(x => x.Position == buttonIdx);
            if (mappingOp != null)
            {
                buttonMatrix[i] = buttonMatrix[i].Replace("#text#", mappingOp.Op).Replace("#imagesrc#", GetImageSrc(mappingOp)).Replace("#op#", $"{mappingOp.Op}?{string.Join('&', mappingOp.Parameters)}");
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

    private static string GetImageSrc(DeckUiOperation mappingOp) =>
        mappingOp.Op.Split('.').First() switch
        {
            "OBS" => ObsImage,
            "PPT" => PptImage,
            _ => FjsdImage
        };

    private const string ObsImage = "https://upload.wikimedia.org/wikipedia/commons/7/78/OBS.svg";
    private const string PptImage = "https://upload.wikimedia.org/wikipedia/commons/6/62/Microsoft_Office_PowerPoint_%282013%E2%80%932019%29.svg";
    private const string FjsdImage = "https://upload.wikimedia.org/wikipedia/commons/6/6d/Windows_Settings_app_icon.png";

    private static readonly DeckUiOperation[] DeckUiOperations =
    {
        new(0, "OBS.ToggleRecording"),
        new(1, "OBS.StopRecording"),
        new(2, "OBS.SwitchScene", "scene=Laptop screen ppt slideshow w camera left"),
        new(3, "PPT.NextSlide"),
        new(4, "PPT.PreviousSlide"),
        new(5, "PPT.ToggleBridge"),
        new(6, "FJSD.Sequence", "seq=swipe.seq"),
        //new DeckUiOperation(7, ""),
        //new DeckUiOperation(8, ""),
    };
}

public record DeckUiOperation(int Position, string Op, params string[] Parameters) : Operation(Op, Parameters);