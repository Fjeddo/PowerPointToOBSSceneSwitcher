using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

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
        Add("PPT.PreviousSlide", _ => PptSlideClick(PptLocal.Backwards));
        Add("PPT.NextSlide", _ => PptSlideClick(PptLocal.Forward));
        Add("PPT.ToggleBridge", _ => PptToggleBridge());
        Add("FJSD.Delay", qc => FjsdDelay(qc["ms"]));
        Add("FJSD.Sequence", qc => FjsdSequence(qc["seq"]));
    }

    public static Dictionary<string, string> ToDeckOperationParameters(params string[] parameters)
    {
        var deckOperationParameters = new Dictionary<string, string>();
        parameters.ToList().ForEach(x =>
        {
            var parts = x.Split('=');
            deckOperationParameters.Add(parts[0], parts[1]);
        });

        return deckOperationParameters;
    }

    private bool ObsStartRecording() => _obsLocal.StartRecording();
    private bool ObsPauseRecording() => _obsLocal.PauseRecording();
    private bool ObsToggleRecording() => _obsLocal.StartPauseResumeRecording(true);
    private bool ObsStopRecording() => _obsLocal.StopRecording();
    private bool ObsSwitchScene(string scene) => _obsLocal.ChangeScene(scene);
    
    private bool PptSlideClick(string direction) => _pptLocal.SwitchSlide(direction);
    private bool PptToggleBridge() => _pptLocal.ToggleBridge();

    private static bool FjsdDelay(string ms)
    {
        Console.WriteLine($"Pause for {ms} ms...");
        Task.Delay(int.Parse(ms)).Wait();
        Console.WriteLine("Continuing");

        return true;
    }

    private bool FjsdSequence(string seq) => ExecuteSequence(seq);

    private bool ExecuteSequence(string seq)
    {
        var ops = File.ReadAllLines(seq);

        var actions = new List<Action>();

        foreach (var op in ops)
        {
            var opParts = op.Split(',');
            if (TryGetValue(opParts[0], out var action))
            {
                var query = new Dictionary<string,string>();
                if (opParts.Length > 1)
                {
                    opParts[^1].Split(',').ToList().ForEach(x =>
                    {
                        var p = x.Split('=');
                        query.Add(p[0], p[1]);
                    });
                }

                actions.Add(() => action(query));
            }
        }

        actions.ForEach(x => x());

        return true;
    }
}
