using System.Collections.Generic;

namespace VerticalScenarioEditer.Models;

public sealed class DocumentState
{
    public List<ScriptRecord> Records { get; set; } = new();
    public bool PageNumberEnabled { get; set; } = true;
    public bool ShowGuides { get; set; } = true;
    public string SummaryText { get; set; } = string.Empty;
    public Dictionary<string, string> RoleDictionary { get; set; } = new();

    public static DocumentState CreateDefault()
    {
        return new DocumentState
        {
            PageNumberEnabled = true,
            ShowGuides = true,
            SummaryText = string.Empty,
            Records = new List<ScriptRecord>(),
            RoleDictionary = new Dictionary<string, string>()
        };
    }
}
