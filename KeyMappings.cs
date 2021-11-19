using System;
using System.Collections.Generic;

namespace PowerPointToOBSSceneSwitcher;

public class KeyMappings : Dictionary<KeyInfo, Operation>
{
    public KeyMappings()
    {
        Add(new KeyInfo(ConsoleKey.F1), new("OBS.ToggleRecording", 0));
        Add(new KeyInfo(ConsoleKey.F1, ConsoleModifiers.Control), new("OBS.StopRecording", 1));
        Add(new KeyInfo(ConsoleKey.LeftArrow), new("PPT.PreviousSlide", 3));
        Add(new KeyInfo(ConsoleKey.RightArrow), new("PPT.NextSlide", 4));
        Add(new KeyInfo(ConsoleKey.T, ConsoleModifiers.Control), new("PPT.ToggleBridge", 5));
    }
}

public record KeyInfo(ConsoleKey ConsoleKey, ConsoleModifiers Modifiers = 0);
public record Operation(string Op, int Position);