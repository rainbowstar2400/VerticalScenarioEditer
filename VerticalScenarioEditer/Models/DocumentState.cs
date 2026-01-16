using System.Collections.Generic;

namespace VerticalScenarioEditer.Models;

public sealed class DocumentState
{
    public List<ScriptRecord> Records { get; set; } = new();
    public bool PageNumberEnabled { get; set; } = true;
    public Dictionary<string, string> RoleDictionary { get; set; } = new();

    public static DocumentState CreateDefault()
    {
        return new DocumentState
        {
            PageNumberEnabled = true,
            Records = new List<ScriptRecord>(),
            RoleDictionary = new Dictionary<string, string>()
        };
    }
}
