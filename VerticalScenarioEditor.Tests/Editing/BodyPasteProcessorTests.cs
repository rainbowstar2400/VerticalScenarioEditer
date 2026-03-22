using VerticalScenarioEditor.Editing;
using Xunit;

namespace VerticalScenarioEditor.Tests.Editing;

public sealed class BodyPasteProcessorTests
{
    [Fact]
    public void ParseParagraphSegments_ShouldSplitByBlankLinesAndKeepInnerNewlines()
    {
        var input = "一行目\n二行目\n\n三行目\n四行目";

        var segments = BodyPasteProcessor.ParseParagraphSegments(input);

        Assert.Equal(2, segments.Count);
        Assert.Equal("一行目\n二行目", segments[0]);
        Assert.Equal("三行目\n四行目", segments[1]);
    }

    [Fact]
    public void ParseParagraphSegments_ShouldIgnoreLeadingTrailingAndConsecutiveEmptyParagraphs()
    {
        var input = "\n \n一行目\n\n\n二行目\n\t\n";

        var segments = BodyPasteProcessor.ParseParagraphSegments(input);

        Assert.Equal(2, segments.Count);
        Assert.Equal("一行目", segments[0]);
        Assert.Equal("二行目", segments[1]);
    }

    [Fact]
    public void ParseParagraphSegments_ShouldTreatInvisibleOnlyLineAsBlankSeparator()
    {
        var input = "一行目\n\u200B\uFEFF\u2060\n二行目";

        var segments = BodyPasteProcessor.ParseParagraphSegments(input);

        Assert.Equal(2, segments.Count);
        Assert.Equal("一行目", segments[0]);
        Assert.Equal("二行目", segments[1]);
    }

    [Fact]
    public void Apply_ShouldSplitAtCaretAndAppendTailToLastRecord()
    {
        var result = BodyPasteProcessor.Apply(
            currentBody: "前後",
            selectionStart: 1,
            selectionEnd: 1,
            clipboardText: "一行目\n\n二行目",
            splitByParagraph: true);

        Assert.True(result.UsedSplit);
        Assert.Equal("前一行目", result.CurrentBody);
        Assert.Single(result.InsertedBodies);
        Assert.Equal("二行目後", result.InsertedBodies[0]);
        Assert.Equal(1, result.CaretRecordOffset);
        Assert.Equal("二行目".Length, result.CaretOffset);
    }

    [Fact]
    public void Apply_ShouldReplaceSelectionBeforeSplit()
    {
        var result = BodyPasteProcessor.Apply(
            currentBody: "ABCDE",
            selectionStart: 1,
            selectionEnd: 4,
            clipboardText: "甲\n\n乙",
            splitByParagraph: true);

        Assert.True(result.UsedSplit);
        Assert.Equal("A甲", result.CurrentBody);
        Assert.Single(result.InsertedBodies);
        Assert.Equal("乙E", result.InsertedBodies[0]);
        Assert.Equal(1, result.CaretRecordOffset);
        Assert.Equal("乙".Length, result.CaretOffset);
    }

    [Fact]
    public void Apply_ShouldUseNormalPasteWhenSplitIsDisabled()
    {
        var result = BodyPasteProcessor.Apply(
            currentBody: "前後",
            selectionStart: 1,
            selectionEnd: 1,
            clipboardText: "一\n\n二",
            splitByParagraph: false);

        Assert.False(result.UsedSplit);
        Assert.Equal("前一\n\n二後", result.CurrentBody);
        Assert.Empty(result.InsertedBodies);
        Assert.Equal(0, result.CaretRecordOffset);
        Assert.Equal("前一\n\n二".Length, result.CaretOffset);
    }

    [Fact]
    public void Apply_ShouldFallbackToNormalPasteWhenSplitCandidateIsSingleParagraph()
    {
        var result = BodyPasteProcessor.Apply(
            currentBody: "前後",
            selectionStart: 1,
            selectionEnd: 1,
            clipboardText: "一行目\n二行目",
            splitByParagraph: true);

        Assert.False(result.UsedSplit);
        Assert.Equal("前一行目\n二行目後", result.CurrentBody);
        Assert.Empty(result.InsertedBodies);
        Assert.Equal(0, result.CaretRecordOffset);
        Assert.Equal("前一行目\n二行目".Length, result.CaretOffset);
    }

    [Fact]
    public void NormalizeNewlines_ShouldConvertCrLfAndCrToLf()
    {
        var normalized = BodyPasteProcessor.NormalizeNewlines("A\r\nB\rC");

        Assert.Equal("A\nB\nC", normalized);
    }
}
