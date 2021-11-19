using System;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.AspNetCore.Builder;
using Newtonsoft.Json;

namespace PowerPointToOBSSceneSwitcher;

public class DeckUi
{
    public void Start()
    {
        var builder = WebApplication.CreateBuilder();
        var app = builder.Build();

        app.MapGet("/", () => JsonConvert.SerializeObject(Program.DefaultMappings, Formatting.Indented));

        foreach (var mapping in Program.DefaultMappings)
        {
            app.MapPost($"/op/{mapping.Value.Op}", () =>
            {
                var op = Program.DefaultOpertions[mapping.Value.Op];
                op(mapping.Value.Op.StartsWith("OBS.") ? Program.Obs : null);
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

            var mappingOp = Program.DefaultMappings.Values.FirstOrDefault(x => x.Position == buttonIdx);
            if (mappingOp != null)
            {
                buttonMatrix[i] = buttonMatrix[i].Replace("#text#", mappingOp.Op).Replace("#imagesrc#", mappingOp.Op.StartsWith("OBS.") ? Program.ObsImage : Program.PptImage).Replace("#op#", mappingOp.Op);
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
}