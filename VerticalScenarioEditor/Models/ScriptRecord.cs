namespace VerticalScenarioEditor.Models;

public sealed class ScriptRecord
{
    public string RoleName { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public bool PageBreakBefore { get; set; }
}

