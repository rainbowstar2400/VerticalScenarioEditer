using VerticalScenarioEditor.Editing;
using Xunit;

namespace VerticalScenarioEditor.Tests.Editing;

public sealed class RolePasteProcessorTests
{
    [Fact]
    public void Apply_ShouldReplaceSelectionWithNormalPaste()
    {
        var result = RolePasteProcessor.Apply(
            currentRoleNames: ["ABCDE", "未使用"],
            startRecordIndex: 0,
            selectionStart: 1,
            selectionEnd: 4,
            clipboardText: "甲",
            splitByParagraph: false,
            shortageMode: RolePasteShortageMode.AddRecords);

        Assert.False(result.Canceled);
        Assert.False(result.UsedSplit);
        Assert.Single(result.AssignedExistingRoles);
        Assert.Equal("A甲E", result.AssignedExistingRoles[0]);
        Assert.Empty(result.AppendedRoles);
        Assert.Equal(0, result.CaretRecordOffset);
        Assert.Equal("A甲".Length, result.CaretOffset);
        Assert.True(result.HasChanges);
    }

    [Fact]
    public void Apply_ShouldAssignSplitRolesSequentially()
    {
        var result = RolePasteProcessor.Apply(
            currentRoleNames: ["", "", ""],
            startRecordIndex: 0,
            selectionStart: 0,
            selectionEnd: 0,
            clipboardText: "役A\n\n役B\n\n役C",
            splitByParagraph: true,
            shortageMode: RolePasteShortageMode.AddRecords);

        Assert.False(result.Canceled);
        Assert.True(result.UsedSplit);
        Assert.Equal(3, result.AssignedExistingRoles.Count);
        Assert.Equal("役A", result.AssignedExistingRoles[0]);
        Assert.Equal("役B", result.AssignedExistingRoles[1]);
        Assert.Equal("役C", result.AssignedExistingRoles[2]);
        Assert.Empty(result.AppendedRoles);
        Assert.Equal(2, result.CaretRecordOffset);
        Assert.Equal("役C".Length, result.CaretOffset);
    }

    [Fact]
    public void Apply_ShouldAppendRecordsWhenShortageModeIsAddRecords()
    {
        var result = RolePasteProcessor.Apply(
            currentRoleNames: ["", ""],
            startRecordIndex: 0,
            selectionStart: 0,
            selectionEnd: 0,
            clipboardText: "役A\n\n役B\n\n役C",
            splitByParagraph: true,
            shortageMode: RolePasteShortageMode.AddRecords);

        Assert.False(result.Canceled);
        Assert.Equal(2, result.AssignedExistingRoles.Count);
        Assert.Single(result.AppendedRoles);
        Assert.Equal("役C", result.AppendedRoles[0]);
        Assert.Equal(0, result.DroppedSegmentCount);
    }

    [Fact]
    public void Apply_ShouldDropRemainderWhenShortageModeIsExistingOnly()
    {
        var result = RolePasteProcessor.Apply(
            currentRoleNames: ["", ""],
            startRecordIndex: 0,
            selectionStart: 0,
            selectionEnd: 0,
            clipboardText: "役A\n\n役B\n\n役C",
            splitByParagraph: true,
            shortageMode: RolePasteShortageMode.ExistingOnly);

        Assert.False(result.Canceled);
        Assert.Equal(2, result.AssignedExistingRoles.Count);
        Assert.Empty(result.AppendedRoles);
        Assert.Equal(1, result.DroppedSegmentCount);
        Assert.Equal(1, result.CaretRecordOffset);
        Assert.Equal("役B".Length, result.CaretOffset);
    }

    [Fact]
    public void Apply_ShouldDetectOverwriteTargetsInSplitMode()
    {
        var overwriteResult = RolePasteProcessor.Apply(
            currentRoleNames: ["既存", ""],
            startRecordIndex: 0,
            selectionStart: 0,
            selectionEnd: 0,
            clipboardText: "役A\n\n役B",
            splitByParagraph: true,
            shortageMode: RolePasteShortageMode.AddRecords);

        var nonOverwriteResult = RolePasteProcessor.Apply(
            currentRoleNames: ["", ""],
            startRecordIndex: 0,
            selectionStart: 0,
            selectionEnd: 0,
            clipboardText: "役A\n\n役B",
            splitByParagraph: true,
            shortageMode: RolePasteShortageMode.AddRecords);

        Assert.True(overwriteResult.HasOverwriteTargets);
        Assert.False(nonOverwriteResult.HasOverwriteTargets);
    }

    [Fact]
    public void Apply_ShouldUseSameParagraphSplitRuleAsBodyPaste()
    {
        var result = RolePasteProcessor.Apply(
            currentRoleNames: ["", ""],
            startRecordIndex: 0,
            selectionStart: 0,
            selectionEnd: 0,
            clipboardText: "役A\r\n\u200B\uFEFF\r\n役B",
            splitByParagraph: true,
            shortageMode: RolePasteShortageMode.AddRecords);

        Assert.True(result.UsedSplit);
        Assert.Equal(2, result.ParagraphSegments.Count);
        Assert.Equal("役A", result.ParagraphSegments[0]);
        Assert.Equal("役B", result.ParagraphSegments[1]);
    }
}
