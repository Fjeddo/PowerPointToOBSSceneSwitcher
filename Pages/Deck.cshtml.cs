using Microsoft.AspNetCore.Mvc.RazorPages;

namespace PowerPointToOBSSceneSwitcher.Pages;

public class DeckModel : PageModel
{
    private readonly ILogger<DeckModel> _logger;

    public DeckModel(ILogger<DeckModel> logger)
    {
        _logger = logger;
    }

    public void OnGet()
    {

    }

    public static readonly DeckUiOperation[] DeckUiOperations =
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

    private const string ObsImage = "https://upload.wikimedia.org/wikipedia/commons/7/78/OBS.svg";
    private const string PptImage = "https://upload.wikimedia.org/wikipedia/commons/6/62/Microsoft_Office_PowerPoint_%282013%E2%80%932019%29.svg";
    private const string FjsdImage = "https://upload.wikimedia.org/wikipedia/commons/6/6d/Windows_Settings_app_icon.png";

    public static string GetImageSrc(DeckUiOperation mappingOp) =>
        mappingOp.Op.Split('.').First() switch
        {
            "OBS" => ObsImage,
            "PPT" => PptImage,
            _ => FjsdImage
        };
}

public record DeckUiOperation(int Position, string Op, params string[] Parameters) : Operation(Op, Parameters);