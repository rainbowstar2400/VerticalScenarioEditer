using System.Collections.Generic;
using System.Text.Json.Serialization;
using VerticalScenarioEditor.Models;

namespace VerticalScenarioEditor.WebView;

public static class WebMessageFaultPolicy
{
    public const string InvalidMessageHandling = "ignore-and-warn";
}

public sealed class WebMessageEnvelope
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }
}

public sealed class InputPatchWebMessage
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("recordIndex")]
    public int RecordIndex { get; set; }

    [JsonPropertyName("field")]
    public string? Field { get; set; }

    [JsonPropertyName("text")]
    public string? Text { get; set; }
}

public sealed class CommandWebMessage
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("name")]
    public string? Name { get; set; }

    [JsonPropertyName("recordIndex")]
    public int? RecordIndex { get; set; }

    [JsonPropertyName("bodyText")]
    public string? BodyText { get; set; }
}

public sealed class ZoomDeltaWebMessage
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("direction")]
    public int Direction { get; set; }
}

public sealed class LayoutStatusWebMessage
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("overflowRecords")]
    public List<int>? OverflowRecords { get; set; }

    [JsonPropertyName("overflowCount")]
    public int? OverflowCount { get; set; }

    [JsonPropertyName("currentPage")]
    public int? CurrentPage { get; set; }

    [JsonPropertyName("totalPages")]
    public int? TotalPages { get; set; }

    [JsonPropertyName("focusedRecordIndex")]
    public int? FocusedRecordIndex { get; set; }
}

public sealed class SelectionChangedWebMessage
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("startRecordIndex")]
    public int? StartRecordIndex { get; set; }

    [JsonPropertyName("endRecordIndex")]
    public int? EndRecordIndex { get; set; }
}

public sealed class PdfReadyWebMessage
{
    [JsonPropertyName("type")]
    public string? Type { get; set; }

    [JsonPropertyName("hasOverflow")]
    public bool HasOverflow { get; set; }
}

public sealed class InitDocumentHostMessage
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "initDocument";

    [JsonPropertyName("document")]
    public DocumentState Document { get; set; } = DocumentState.CreateDefault();

    [JsonPropertyName("settings")]
    public DocumentRenderSettingsHostMessage Settings { get; set; } = new();
}

public sealed class DocumentRenderSettingsHostMessage
{
    [JsonPropertyName("pageWidthMm")]
    public double PageWidthMm { get; set; }

    [JsonPropertyName("pageHeightMm")]
    public double PageHeightMm { get; set; }

    [JsonPropertyName("marginLeftMm")]
    public double MarginLeftMm { get; set; }

    [JsonPropertyName("marginRightMm")]
    public double MarginRightMm { get; set; }

    [JsonPropertyName("marginTopMm")]
    public double MarginTopMm { get; set; }

    [JsonPropertyName("marginBottomMm")]
    public double MarginBottomMm { get; set; }

    [JsonPropertyName("fontFamily")]
    public string FontFamily { get; set; } = string.Empty;

    [JsonPropertyName("fontSizePt")]
    public double FontSizePt { get; set; }

    [JsonPropertyName("lineSpacing")]
    public double LineSpacing { get; set; }

    [JsonPropertyName("pageGapPx")]
    public double PageGapPx { get; set; }

    [JsonPropertyName("pageNumberEnabled")]
    public bool PageNumberEnabled { get; set; }

    [JsonPropertyName("showGuides")]
    public bool ShowGuides { get; set; }

    [JsonPropertyName("showBreakMarkers")]
    public bool ShowBreakMarkers { get; set; }

    [JsonPropertyName("documentTitle")]
    public string DocumentTitle { get; set; } = string.Empty;

    [JsonPropertyName("roleLabelHeightChars")]
    public double RoleLabelHeightChars { get; set; }

    [JsonPropertyName("zoomScale")]
    public double ZoomScale { get; set; }

    [JsonPropertyName("selectionMode")]
    public bool SelectionMode { get; set; }

    [JsonPropertyName("summaryMode")]
    public bool SummaryMode { get; set; }

    [JsonPropertyName("simpleMode")]
    public bool SimpleMode { get; set; }
}

public sealed class ApplyRoleDictionaryHostMessage
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "applyRoleDictionary";

    [JsonPropertyName("roleDictionary")]
    public Dictionary<string, string> RoleDictionary { get; set; } = new();
}

public sealed class ApplyOverflowAttentionHostMessage
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "applyOverflowAttention";

    [JsonPropertyName("overflowAttentionRecords")]
    public IEnumerable<int> OverflowAttentionRecords { get; set; } = [];
}

public sealed class ApplySelectionModeHostMessage
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "applySelectionMode";

    [JsonPropertyName("selectionMode")]
    public bool SelectionMode { get; set; }
}

public sealed class EnterPdfModeHostMessage
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "enterPdfMode";

    [JsonPropertyName("summaryText")]
    public string SummaryText { get; set; } = string.Empty;

    [JsonPropertyName("documentTitle")]
    public string DocumentTitle { get; set; } = string.Empty;
}

public sealed class ApplyDocumentTitleHostMessage
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "applyDocumentTitle";

    [JsonPropertyName("documentTitle")]
    public string DocumentTitle { get; set; } = string.Empty;
}

public sealed class ExitPdfModeHostMessage
{
    [JsonPropertyName("type")]
    public string Type { get; set; } = "exitPdfMode";
}
