using System;
using VerticalScenarioEditor.History;
using VerticalScenarioEditor.Models;
using Xunit;

namespace VerticalScenarioEditor.Tests.History;

public sealed class DocumentHistoryServiceTests
{
    [Fact]
    public void PushSnapshot_ShouldKeepLatestEntriesWithinLimit()
    {
        var history = new DocumentHistoryService(maxSnapshots: 2);
        history.PushSnapshot(CreateDocument("A"));
        history.PushSnapshot(CreateDocument("B"));
        history.PushSnapshot(CreateDocument("C"));

        var current = CreateDocument("NOW");
        Assert.True(history.TryUndo(current, out var previous1));
        Assert.Equal("C", previous1.SummaryText);

        current = previous1;
        Assert.True(history.TryUndo(current, out var previous2));
        Assert.Equal("B", previous2.SummaryText);

        current = previous2;
        Assert.False(history.TryUndo(current, out _));
    }

    [Fact]
    public void PushInputSnapshot_ShouldMergeWithinWindow()
    {
        var now = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        var history = new DocumentHistoryService(
            maxSnapshots: 10,
            inputMergeWindow: TimeSpan.FromMilliseconds(500),
            utcNow: () => now);

        history.PushInputSnapshot(CreateDocument("A"), 0, "body");
        now = now.AddMilliseconds(100);
        history.PushInputSnapshot(CreateDocument("B"), 0, "body");

        var current = CreateDocument("NOW");
        Assert.True(history.TryUndo(current, out _));
        Assert.False(history.TryUndo(current, out _));
    }

    [Fact]
    public void PushInputSnapshot_ShouldSplitWhenWindowElapsed()
    {
        var now = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        var history = new DocumentHistoryService(
            maxSnapshots: 10,
            inputMergeWindow: TimeSpan.FromMilliseconds(500),
            utcNow: () => now);

        history.PushInputSnapshot(CreateDocument("A"), 0, "body");
        now = now.AddMilliseconds(600);
        history.PushInputSnapshot(CreateDocument("B"), 0, "body");

        var current = CreateDocument("NOW");
        Assert.True(history.TryUndo(current, out var previous1));
        Assert.Equal("B", previous1.SummaryText);

        current = previous1;
        Assert.True(history.TryUndo(current, out var previous2));
        Assert.Equal("A", previous2.SummaryText);
    }

    [Fact]
    public void UndoRedo_ShouldRoundTripStates()
    {
        var history = new DocumentHistoryService();
        var previous = CreateDocument("Before");
        history.PushSnapshot(previous);

        var current = CreateDocument("After");
        Assert.True(history.TryUndo(current, out var undone));
        Assert.Equal("Before", undone.SummaryText);

        Assert.True(history.TryRedo(undone, out var redone));
        Assert.Equal("After", redone.SummaryText);
    }

    [Fact]
    public void UndoRedo_ShouldPreserveManualPageBreakFlags()
    {
        var history = new DocumentHistoryService();
        var previous = new DocumentState
        {
            SummaryText = "Before",
            Records =
            {
                new ScriptRecord
                {
                    RoleName = "役A",
                    Body = "台詞A",
                    PageBreakBefore = false
                },
                new ScriptRecord
                {
                    RoleName = "役B",
                    Body = "台詞B",
                    PageBreakBefore = true
                }
            }
        };
        history.PushSnapshot(previous);

        var current = CreateDocument("After");
        Assert.True(history.TryUndo(current, out var undone));
        Assert.Equal(2, undone.Records.Count);
        Assert.False(undone.Records[0].PageBreakBefore);
        Assert.True(undone.Records[1].PageBreakBefore);

        Assert.True(history.TryRedo(undone, out var redone));
        Assert.False(redone.Records[0].PageBreakBefore);
    }

    private static DocumentState CreateDocument(string summaryText)
    {
        return new DocumentState
        {
            SummaryText = summaryText,
            Records =
            {
                new ScriptRecord
                {
                    RoleName = "役",
                    Body = "台詞"
                }
            },
            RoleDictionary =
            {
                ["役"] = "#111111"
            }
        };
    }
}
