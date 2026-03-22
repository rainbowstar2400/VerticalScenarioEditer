using System;
using System.Collections.Generic;

namespace VerticalScenarioEditor.Editing;

public sealed class BodyPasteApplyResult
{
    public BodyPasteApplyResult(
        string currentBody,
        IReadOnlyList<string> insertedBodies,
        int caretRecordOffset,
        int caretOffset,
        bool usedSplit,
        string normalizedClipboardText,
        IReadOnlyList<string> paragraphSegments)
    {
        CurrentBody = currentBody ?? string.Empty;
        InsertedBodies = insertedBodies ?? [];
        CaretRecordOffset = Math.Max(0, caretRecordOffset);
        CaretOffset = Math.Max(0, caretOffset);
        UsedSplit = usedSplit;
        NormalizedClipboardText = normalizedClipboardText ?? string.Empty;
        ParagraphSegments = paragraphSegments ?? [];
    }

    public string CurrentBody { get; }

    public IReadOnlyList<string> InsertedBodies { get; }

    public int CaretRecordOffset { get; }

    public int CaretOffset { get; }

    public bool UsedSplit { get; }

    public string NormalizedClipboardText { get; }

    public IReadOnlyList<string> ParagraphSegments { get; }
}

public static class BodyPasteProcessor
{
    public static string NormalizeNewlines(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return string.Empty;
        }

        return text
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');
    }

    public static IReadOnlyList<string> ParseParagraphSegments(string? clipboardText)
    {
        var normalized = NormalizeNewlines(clipboardText);
        if (normalized.Length == 0)
        {
            return [];
        }

        var lines = normalized.Split('\n');
        var currentParagraphLines = new List<string>();
        var segments = new List<string>();

        foreach (var line in lines)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                FlushCurrentParagraph(currentParagraphLines, segments);
                continue;
            }

            currentParagraphLines.Add(line);
        }

        FlushCurrentParagraph(currentParagraphLines, segments);
        return segments;
    }

    public static BodyPasteApplyResult Apply(
        string? currentBody,
        int selectionStart,
        int selectionEnd,
        string? clipboardText,
        bool splitByParagraph)
    {
        var normalizedCurrentBody = NormalizeNewlines(currentBody);
        var normalizedClipboardText = NormalizeNewlines(clipboardText);
        var paragraphSegments = ParseParagraphSegments(normalizedClipboardText);

        var clampedStart = Math.Clamp(selectionStart, 0, normalizedCurrentBody.Length);
        var clampedEnd = Math.Clamp(selectionEnd, clampedStart, normalizedCurrentBody.Length);

        var before = normalizedCurrentBody[..clampedStart];
        var after = normalizedCurrentBody[clampedEnd..];

        if (!splitByParagraph || paragraphSegments.Count < 2)
        {
            var mergedBody = before + normalizedClipboardText + after;
            return new BodyPasteApplyResult(
                mergedBody,
                [],
                caretRecordOffset: 0,
                caretOffset: (before + normalizedClipboardText).Length,
                usedSplit: false,
                normalizedClipboardText,
                paragraphSegments);
        }

        var insertedBodies = new List<string>(paragraphSegments.Count - 1);
        for (var index = 1; index < paragraphSegments.Count; index += 1)
        {
            insertedBodies.Add(paragraphSegments[index]);
        }

        var lastInsertedIndex = insertedBodies.Count - 1;
        insertedBodies[lastInsertedIndex] = insertedBodies[lastInsertedIndex] + after;

        var updatedCurrentBody = before + paragraphSegments[0];
        return new BodyPasteApplyResult(
            updatedCurrentBody,
            insertedBodies,
            caretRecordOffset: insertedBodies.Count,
            caretOffset: paragraphSegments[^1].Length,
            usedSplit: true,
            normalizedClipboardText,
            paragraphSegments);
    }

    private static void FlushCurrentParagraph(List<string> currentParagraphLines, List<string> segments)
    {
        if (currentParagraphLines.Count == 0)
        {
            return;
        }

        var paragraph = string.Join('\n', currentParagraphLines);
        if (paragraph.Length > 0)
        {
            segments.Add(paragraph);
        }

        currentParagraphLines.Clear();
    }
}
