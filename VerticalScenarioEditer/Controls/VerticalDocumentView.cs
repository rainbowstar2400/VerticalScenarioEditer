using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows;
using System.Windows.Media;
using VerticalScenarioEditer.Models;

namespace VerticalScenarioEditer.Controls;

public sealed class VerticalDocumentView : FrameworkElement
{
    public static readonly DependencyProperty DocumentProperty = DependencyProperty.Register(
        nameof(Document),
        typeof(DocumentState),
        typeof(VerticalDocumentView),
        new FrameworkPropertyMetadata(DocumentState.CreateDefault(), FrameworkPropertyMetadataOptions.AffectsMeasure | FrameworkPropertyMetadataOptions.AffectsRender));

    private DocumentLayout? _layout;

    public DocumentState Document
    {
        get => (DocumentState)GetValue(DocumentProperty);
        set => SetValue(DocumentProperty, value ?? DocumentState.CreateDefault());
    }

    protected override Size MeasureOverride(Size availableSize)
    {
        _layout = BuildLayout();
        return _layout.Size;
    }

    protected override void OnRender(DrawingContext drawingContext)
    {
        var layout = _layout ?? BuildLayout();
        DrawLayout(drawingContext, layout);
    }

    private DocumentLayout BuildLayout()
    {
        var pageWidth = DocumentSettings.PageWidthDip;
        var pageHeight = DocumentSettings.PageHeightDip;
        var contentRect = new Rect(
            DocumentSettings.MarginLeftDip,
            DocumentSettings.MarginTopDip,
            pageWidth - DocumentSettings.MarginLeftDip - DocumentSettings.MarginRightDip,
            pageHeight - DocumentSettings.MarginTopDip - DocumentSettings.MarginBottomDip);

        var columnAdvance = DocumentSettings.ColumnAdvanceDip;
        var bodyHeight = contentRect.Height - DocumentSettings.RoleLabelHeightDip;
        var columnsPerPage = Math.Max(1, (int)Math.Floor(contentRect.Width / columnAdvance));
        var bodyCharsPerColumn = Math.Max(1, (int)Math.Floor(bodyHeight / columnAdvance));

        var pages = new List<PageLayout>();
        var currentPage = new PageLayout(contentRect, pageWidth, pageHeight);
        pages.Add(currentPage);

        var currentColumnIndex = 0;
        var hasOverflow = false;

        foreach (var record in Document?.Records ?? new List<ScriptRecord>())
        {
            var bodyColumns = BuildColumns(record.Body ?? string.Empty, bodyCharsPerColumn);
            var recordColumns = Math.Max(1, bodyColumns.Count);
            var overflow = recordColumns > columnsPerPage;
            hasOverflow |= overflow;

            if (currentColumnIndex > 0 && currentColumnIndex + recordColumns > columnsPerPage)
            {
                currentPage = new PageLayout(contentRect, pageWidth, pageHeight);
                pages.Add(currentPage);
                currentColumnIndex = 0;
            }

            currentPage.Records.Add(new RecordLayout(record, bodyColumns, currentColumnIndex, recordColumns, overflow));
            currentColumnIndex += recordColumns + (int)Math.Round(DocumentSettings.RecordGapChars);
        }

        var totalHeight = pageHeight * pages.Count + DocumentSettings.PageGapDip * Math.Max(0, pages.Count - 1);
        return new DocumentLayout(new Size(pageWidth, totalHeight), pages, hasOverflow);
    }

    private static List<string> BuildColumns(string text, int maxCharsPerColumn)
    {
        var columns = new List<string>();
        var current = new StringBuilder();

        foreach (var ch in text)
        {
            if (ch == '\r')
            {
                continue;
            }

            if (ch == '\n')
            {
                columns.Add(current.ToString());
                current.Clear();
                continue;
            }

            current.Append(ch);
            if (current.Length >= maxCharsPerColumn)
            {
                columns.Add(current.ToString());
                current.Clear();
            }
        }

        if (current.Length > 0 || columns.Count == 0)
        {
            columns.Add(current.ToString());
        }

        return columns;
    }

    private void DrawLayout(DrawingContext dc, DocumentLayout layout)
    {
        var dpi = VisualTreeHelper.GetDpi(this).PixelsPerDip;
        var typeface = new Typeface(DocumentSettings.DefaultFontFamily, FontStyles.Normal, FontWeights.Normal, FontStretches.Normal);
        var columnAdvance = DocumentSettings.ColumnAdvanceDip;
        var roleLabelHeight = DocumentSettings.RoleLabelHeightDip;
        var bodyCharsPerColumn = Math.Max(1, (int)Math.Floor((layout.ContentHeight - roleLabelHeight) / columnAdvance));
        var roleCharsPerColumn = Math.Max(1, (int)Math.Floor(roleLabelHeight / columnAdvance));

        for (var pageIndex = 0; pageIndex < layout.Pages.Count; pageIndex++)
        {
            var pageTop = pageIndex * (layout.PageHeight + DocumentSettings.PageGapDip);
            var pageRect = new Rect(0, pageTop, layout.PageWidth, layout.PageHeight);
            dc.DrawRectangle(Brushes.White, new Pen(Brushes.Gray, 1), pageRect);

            var contentRect = new Rect(
                layout.ContentRect.Left,
                layout.ContentRect.Top + pageTop,
                layout.ContentRect.Width,
                layout.ContentRect.Height);

            foreach (var recordLayout in layout.Pages[pageIndex].Records)
            {
                var recordColumns = recordLayout.RecordColumns;
                var drawColumns = Math.Min(recordColumns, layout.ColumnsPerPage);
                var recordLeft = contentRect.Right - (recordLayout.ColumnStart + drawColumns) * columnAdvance;
                var recordRect = new Rect(recordLeft, contentRect.Top, drawColumns * columnAdvance, contentRect.Height);
                var recordPen = recordLayout.Overflow ? new Pen(Brushes.Red, 1.5) : new Pen(Brushes.LightGray, 1);

                dc.DrawRectangle(null, recordPen, recordRect);

                for (var columnIndex = 0; columnIndex < drawColumns; columnIndex++)
                {
                    var globalColumnIndex = recordLayout.ColumnStart + columnIndex;
                    var columnLeft = contentRect.Right - (globalColumnIndex + 1) * columnAdvance;
                    var columnTop = contentRect.Top;

                    if (columnIndex == 0)
                    {
                        DrawVerticalText(dc, typeface, recordLayout.Record.RoleName ?? string.Empty, columnLeft, columnTop, roleCharsPerColumn, dpi);
                    }

                    var bodyColumnIndex = Math.Min(columnIndex, recordLayout.BodyColumns.Count - 1);
                    var bodyText = recordLayout.BodyColumns.Count == 0 ? string.Empty : recordLayout.BodyColumns[bodyColumnIndex];
                    DrawVerticalText(dc, typeface, bodyText, columnLeft, columnTop + roleLabelHeight, bodyCharsPerColumn, dpi);
                }

                if (recordLayout.Overflow)
                {
                    var warningText = new FormattedText(
                        "警告: 1ページに収まらないレコードがあります",
                        CultureInfo.CurrentCulture,
                        FlowDirection.LeftToRight,
                        typeface,
                        DocumentSettings.FontSizeDip,
                        Brushes.Red,
                        dpi);
                    dc.DrawText(warningText, new Point(contentRect.Left, contentRect.Top - warningText.Height - 4));
                }
            }
        }
    }

    private void DrawVerticalText(DrawingContext dc, Typeface typeface, string text, double columnLeft, double columnTop, int maxChars, double dpi)
    {
        var columnAdvance = DocumentSettings.ColumnAdvanceDip;
        for (var index = 0; index < text.Length && index < maxChars; index++)
        {
            var glyph = text[index].ToString();
            var formatted = new FormattedText(
                glyph,
                CultureInfo.CurrentCulture,
                FlowDirection.LeftToRight,
                typeface,
                DocumentSettings.FontSizeDip,
                Brushes.Black,
                dpi);

            var x = columnLeft + (columnAdvance - formatted.Width) / 2;
            var y = columnTop + index * columnAdvance;
            var charRect = new Rect(columnLeft, y, columnAdvance, columnAdvance);
            if (ShouldRotate(text[index]))
            {
                var center = new Point(charRect.Left + charRect.Width / 2, charRect.Top + charRect.Height / 2);
                dc.PushTransform(new RotateTransform(90, center.X, center.Y));
                dc.DrawText(formatted, new Point(x, y));
                dc.Pop();
            }
            else
            {
                dc.DrawText(formatted, new Point(x, y));
            }
        }
    }

    private static bool ShouldRotate(char ch)
    {
        return ch <= 0x7F;
    }

    private sealed record DocumentLayout(Size Size, List<PageLayout> Pages, bool HasOverflow)
    {
        public Rect ContentRect => Pages[0].ContentRect;
        public double PageWidth => Pages[0].PageWidth;
        public double PageHeight => Pages[0].PageHeight;
        public double ContentHeight => Pages[0].ContentRect.Height;
        public int ColumnsPerPage => Math.Max(1, (int)Math.Floor(Pages[0].ContentRect.Width / DocumentSettings.ColumnAdvanceDip));
    }

    private sealed record PageLayout(Rect ContentRect, double PageWidth, double PageHeight)
    {
        public List<RecordLayout> Records { get; } = new();
    }

    private sealed record RecordLayout(ScriptRecord Record, List<string> BodyColumns, int ColumnStart, int RecordColumns, bool Overflow);
}
