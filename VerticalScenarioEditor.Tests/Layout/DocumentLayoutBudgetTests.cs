using System;
using Xunit;

namespace VerticalScenarioEditor.Tests.Layout;

public sealed class DocumentLayoutBudgetTests
{
    private const double PageEpsilonPx = 1.0;

    [Fact]
    public void Margins_ShouldMatchDocumentSpec()
    {
        Assert.Equal(17.73, DocumentSettings.MarginLeftMm, 2);
        Assert.Equal(17.73, DocumentSettings.MarginRightMm, 2);
        Assert.Equal(25.4, DocumentSettings.MarginTopMm, 1);
        Assert.Equal(25.4, DocumentSettings.MarginBottomMm, 1);
    }

    [Fact]
    public void ContentWidthBudget_ShouldAllowThirtyNineColumns()
    {
        var contentWidth = DocumentSettings.PageWidthDip - DocumentSettings.MarginLeftDip - DocumentSettings.MarginRightDip;
        var availableWidth = Math.Round(contentWidth);
        var widthForThirtyNineColumns = Math.Ceiling(DocumentSettings.ColumnAdvanceDip * 39);

        Assert.True(widthForThirtyNineColumns <= availableWidth + PageEpsilonPx);
    }

    [Fact]
    public void ContentWidthBudget_ShouldRejectFortyColumns()
    {
        var contentWidth = DocumentSettings.PageWidthDip - DocumentSettings.MarginLeftDip - DocumentSettings.MarginRightDip;
        var availableWidth = Math.Round(contentWidth);
        var widthForFortyColumns = Math.Ceiling(DocumentSettings.ColumnAdvanceDip * 40);

        Assert.True(widthForFortyColumns > availableWidth + PageEpsilonPx);
    }
}
