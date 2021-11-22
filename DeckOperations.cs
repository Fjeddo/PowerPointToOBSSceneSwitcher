namespace PowerPointToOBSSceneSwitcher;

public class DeckOperations : Dictionary<string, Func<Dictionary<string, string>, bool>>
{
    private readonly ObsLocal _obsLocal;
    private readonly PptLocal _pptLocal;

    public DeckOperations(ObsLocal obsLocal, PptLocal pptLocal)
    {
        _obsLocal = obsLocal;
        _pptLocal = pptLocal;

        Add("OBS.StartRecording", _ => ObsStartRecording());
        Add("OBS.ToggleRecording", _ => ObsToggleRecording());
        Add("OBS.StopRecording", _ => ObsStopRecording());
        Add("OBS.PauseRecording", _ => ObsPauseRecording());
        Add("OBS.SwitchScene", qc => ObsSwitchScene(qc["scene"]));
        Add("OBS.RenderSource", qc => ObsRenderSource(qc["sourceName"], qc["render"]));
        Add("PPT.PreviousSlide", _ => PptSlideClick(PptLocal.Backwards));
        Add("PPT.NextSlide", _ => PptSlideClick(PptLocal.Forward));
        Add("PPT.ToggleBridge", _ => PptToggleBridge());
        Add("FJSD.Delay", qc => FjsdDelay(qc["ms"]));
        Add("FJSD.Sequence", qc => FjsdSequence(qc["seq"]));
    }

    public static Dictionary<string, string> ToDeckOperationParameters(params string[] parameters) => parameters.Select(x => x.Split('=')).ToDictionary(parts => parts[0], parts => parts[1]);

    private bool ObsStartRecording() => _obsLocal.StartRecording();
    private bool ObsPauseRecording() => _obsLocal.PauseRecording();
    private bool ObsToggleRecording() => _obsLocal.StartPauseResumeRecording(true);
    private bool ObsStopRecording() => _obsLocal.StopRecording();
    private bool ObsSwitchScene(string scene) => _obsLocal.ChangeScene(scene);
    private bool ObsRenderSource(string sourceName, string render) => _obsLocal.RenderSource(sourceName, bool.Parse(render));
    
    private bool PptSlideClick(string direction) => _pptLocal.SwitchSlide(direction);
    private bool PptToggleBridge() => _pptLocal.ToggleBridge();

    private static bool FjsdDelay(string ms)
    {
        Console.WriteLine($"Pause for {ms} ms...");
        Task.Delay(int.Parse(ms)).Wait();
        Console.WriteLine("Continuing");

        return true;
    }

    private bool FjsdSequence(string seq)
    {
        var actions = File.ReadAllLines(seq).Select(op =>
        {
            var opParts = op.Split(',');
            
            if (TryGetValue(opParts[0], out var action))
            {
                var query = opParts.Length > 1
                    ? opParts[1..].Select(x => x.Split('=')).ToDictionary(parts => parts[0], parts => parts[1])
                    : new Dictionary<string, string>();

                return () => action(query);
            }

            throw new Exception("Operation in sequence not valid");
        }).ToList();

        actions.ForEach(x => x());

        return true;
    }
}
