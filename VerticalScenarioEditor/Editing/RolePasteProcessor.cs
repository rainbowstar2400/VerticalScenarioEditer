using System;
using System.Collections.Generic;

namespace VerticalScenarioEditor.Editing;

public enum RolePasteShortageMode
{
    AddRecords = 0,
    ExistingOnly = 1,
    Cancel = 2
}

public sealed class RolePasteApplyResult
{
    public RolePasteApplyResult(
        bool canceled,
        bool usedSplit,
        string normalizedClipboardText,
        IReadOnlyList<string> paragraphSegments,
        IReadOnlyList<string> assignedExistingRoles,
        IReadOnlyList<string> appendedRoles,
        bool hasOverwriteTargets,
        int droppedSegmentCount,
        int caretRecordOffset,
        int caretOffset,
        bool hasChanges)
    {
        Canceled = canceled;
        UsedSplit = usedSplit;
        NormalizedClipboardText = normalizedClipboardText ?? string.Empty;
        ParagraphSegments = paragraphSegments ?? [];
        AssignedExistingRoles = assignedExistingRoles ?? [];
        AppendedRoles = appendedRoles ?? [];
        HasOverwriteTargets = hasOverwriteTargets;
        DroppedSegmentCount = Math.Max(0, droppedSegmentCount);
        CaretRecordOffset = Math.Max(0, caretRecordOffset);
        CaretOffset = Math.Max(0, caretOffset);
        HasChanges = hasChanges;
    }

    public bool Canceled { get; }

    public bool UsedSplit { get; }

    public string NormalizedClipboardText { get; }

    public IReadOnlyList<string> ParagraphSegments { get; }

    public IReadOnlyList<string> AssignedExistingRoles { get; }

    public IReadOnlyList<string> AppendedRoles { get; }

    public bool HasOverwriteTargets { get; }

    public int DroppedSegmentCount { get; }

    public int CaretRecordOffset { get; }

    public int CaretOffset { get; }

    public bool HasChanges { get; }
}

public static class RolePasteProcessor
{
    public static RolePasteApplyResult Apply(
        IReadOnlyList<string> currentRoleNames,
        int startRecordIndex,
        int selectionStart,
        int selectionEnd,
        string? clipboardText,
        bool splitByParagraph,
        RolePasteShortageMode shortageMode)
    {
        var safeRoleNames = currentRoleNames ?? [];
        if (safeRoleNames.Count == 0 || startRecordIndex < 0 || startRecordIndex >= safeRoleNames.Count)
        {
            return new RolePasteApplyResult(
                canceled: true,
                usedSplit: splitByParagraph,
                normalizedClipboardText: BodyPasteProcessor.NormalizeNewlines(clipboardText),
                paragraphSegments: [],
                assignedExistingRoles: [],
                appendedRoles: [],
                hasOverwriteTargets: false,
                droppedSegmentCount: 0,
                caretRecordOffset: 0,
                caretOffset: 0,
                hasChanges: false);
        }

        var normalizedClipboardText = BodyPasteProcessor.NormalizeNewlines(clipboardText);
        var paragraphSegments = BodyPasteProcessor.ParseParagraphSegments(normalizedClipboardText);
        var normalizedCurrentRole = BodyPasteProcessor.NormalizeNewlines(safeRoleNames[startRecordIndex]);

        if (!splitByParagraph || paragraphSegments.Count < 2)
        {
            var clampedStart = Math.Clamp(selectionStart, 0, normalizedCurrentRole.Length);
            var clampedEnd = Math.Clamp(selectionEnd, clampedStart, normalizedCurrentRole.Length);
            var before = normalizedCurrentRole[..clampedStart];
            var after = normalizedCurrentRole[clampedEnd..];
            var mergedRole = before + normalizedClipboardText + after;
            var hasChanges = !string.Equals(normalizedCurrentRole, mergedRole, StringComparison.Ordinal);

            return new RolePasteApplyResult(
                canceled: false,
                usedSplit: false,
                normalizedClipboardText,
                paragraphSegments,
                assignedExistingRoles: [mergedRole],
                appendedRoles: [],
                hasOverwriteTargets: false,
                droppedSegmentCount: 0,
                caretRecordOffset: 0,
                caretOffset: (before + normalizedClipboardText).Length,
                hasChanges);
        }

        var availableExistingCount = safeRoleNames.Count - startRecordIndex;
        var requestedSegmentCount = paragraphSegments.Count;
        var appliedSegmentCount = requestedSegmentCount;

        if (requestedSegmentCount > availableExistingCount)
        {
            if (shortageMode == RolePasteShortageMode.Cancel)
            {
                return new RolePasteApplyResult(
                    canceled: true,
                    usedSplit: true,
                    normalizedClipboardText,
                    paragraphSegments,
                    assignedExistingRoles: [],
                    appendedRoles: [],
                    hasOverwriteTargets: false,
                    droppedSegmentCount: requestedSegmentCount,
                    caretRecordOffset: 0,
                    caretOffset: 0,
                    hasChanges: false);
            }

            if (shortageMode == RolePasteShortageMode.ExistingOnly)
            {
                appliedSegmentCount = Math.Max(0, availableExistingCount);
            }
        }

        var existingAssignCount = Math.Min(appliedSegmentCount, availableExistingCount);
        var appendedCount = Math.Max(0, appliedSegmentCount - existingAssignCount);

        var assignedExistingRoles = new List<string>(existingAssignCount);
        var appendedRoles = new List<string>(appendedCount);

        var hasOverwriteTargets = false;
        var hasExistingChanges = false;

        for (var offset = 0; offset < existingAssignCount; offset += 1)
        {
            var nextRole = paragraphSegments[offset];
            assignedExistingRoles.Add(nextRole);

            var currentRole = BodyPasteProcessor.NormalizeNewlines(safeRoleNames[startRecordIndex + offset]);
            if (currentRole.Length > 0)
            {
                hasOverwriteTargets = true;
            }
            if (!string.Equals(currentRole, nextRole, StringComparison.Ordinal))
            {
                hasExistingChanges = true;
            }
        }

        for (var index = existingAssignCount; index < appliedSegmentCount; index += 1)
        {
            appendedRoles.Add(paragraphSegments[index]);
        }

        var droppedSegmentCount = Math.Max(0, requestedSegmentCount - appliedSegmentCount);
        var hasAnyChanges = hasExistingChanges || appendedRoles.Count > 0;
        var caretRole = paragraphSegments[Math.Max(0, appliedSegmentCount - 1)];

        return new RolePasteApplyResult(
            canceled: false,
            usedSplit: true,
            normalizedClipboardText,
            paragraphSegments,
            assignedExistingRoles,
            appendedRoles,
            hasOverwriteTargets,
            droppedSegmentCount,
            caretRecordOffset: Math.Max(0, appliedSegmentCount - 1),
            caretOffset: caretRole.Length,
            hasChanges: hasAnyChanges);
    }
}
