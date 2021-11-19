using System;
using System.IO;
using System.Linq;
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
            app.MapPost($"/op/{mapping.Value.Op}", () =>
            {
                var op = Program.DeckOperations[mapping.Value.Op];
                op(mapping.Value.Op.StartsWith("OBS.") ? Program.Obs : Program.Ppt);
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

            var mappingOp = Program.KeyMappings.Values.FirstOrDefault(x => x.Position == buttonIdx);
            if (mappingOp != null)
            {
                buttonMatrix[i] = buttonMatrix[i].Replace("#text#", mappingOp.Op).Replace("#imagesrc#", mappingOp.Op.StartsWith("OBS.") ? ObsImage : PptImage).Replace("#op#", mappingOp.Op);
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

    private const string ObsImage = "https://upload.wikimedia.org/wikipedia/commons/7/78/OBS.svg";
    private const string PptImage = "https://upload.wikimedia.org/wikipedia/commons/6/62/Microsoft_Office_PowerPoint_%282013%E2%80%932019%29.svg";
}