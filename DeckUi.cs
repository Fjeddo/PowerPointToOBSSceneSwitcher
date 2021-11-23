using Newtonsoft.Json;

namespace PowerPointToOBSSceneSwitcher;

public class DeckUi
{
    private static WebApplication _app;

    public static void Start()
    {
        var builder = WebApplication.CreateBuilder();
        builder.Logging.SetMinimumLevel(LogLevel.Warning);
        builder.Services.AddRazorPages();
        
        _app = builder.Build();
        _app.UseStaticFiles();
        _app.MapRazorPages();

        _app.MapGet("/", () => JsonConvert.SerializeObject(Program.KeyMappings, Formatting.Indented));

        MapOperations(_app);

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
}