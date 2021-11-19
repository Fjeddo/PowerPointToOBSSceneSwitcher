using System;
using System.Collections.Generic;

namespace PowerPointToOBSSceneSwitcher;

public class DeckOperations : Dictionary<string, Action<object>>
{
    public DeckOperations()
    {
        Add("OBS.ToggleRecording", obj => (obj as ObsLocal)?.StartPauseResumeRecording(true));
        Add("OBS.StopRecording", obj => (obj as ObsLocal)?.StopRecording());
        Add("PPT.PreviousSlide", obj => (obj as PptLocal)?.SwitchSlide(PptLocal.Backwards));
        Add("PPT.NextSlide", obj => (obj as PptLocal)?.SwitchSlide());
        Add("PPT.ToggleBridge", obj => (obj as PptLocal)?.ToggleBridge());
    }
}