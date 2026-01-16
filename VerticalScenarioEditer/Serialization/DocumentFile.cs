using System.Text.Json.Serialization;
using VerticalScenarioEditer.Models;

namespace VerticalScenarioEditer.Serialization;

public sealed class DocumentFile
{
    public const int CurrentVersion = 1;

    [JsonPropertyName("version")]
    public int Version { get; set; } = CurrentVersion;

    [JsonPropertyName("document")]
    public DocumentState Document { get; set; } = DocumentState.CreateDefault();
}
