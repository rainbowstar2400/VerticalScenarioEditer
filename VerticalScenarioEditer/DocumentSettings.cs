using System;
using System.Linq;
using System.Windows;
using System.Windows.Media;

namespace VerticalScenarioEditer;

public static class DocumentSettings
{
    public const string DefaultFontFamilyName = "游明朝";
    public const double DefaultFontSizePt = 10.5;
    public const double LineSpacing = 1.0;
    public const double RoleLabelHeightChars = 5.5;
    public const double RecordGapChars = 1.0;
    public const double PageGapDip = 24.0;

    public const double PageWidthMm = 297.0;
    public const double PageHeightMm = 210.0;
    public const double MarginLeftMm = 20.0;
    public const double MarginRightMm = 20.0;
    public const double MarginTopMm = 25.0;
    public const double MarginBottomMm = 25.0;

    public static double PageWidthDip => MmToDip(PageWidthMm);
    public static double PageHeightDip => MmToDip(PageHeightMm);
    public static double MarginLeftDip => MmToDip(MarginLeftMm);
    public static double MarginRightDip => MmToDip(MarginRightMm);
    public static double MarginTopDip => MmToDip(MarginTopMm);
    public static double MarginBottomDip => MmToDip(MarginBottomMm);
    public static double FontSizeDip => PointsToDip(DefaultFontSizePt);
    public static double ColumnAdvanceDip => FontSizeDip * LineSpacing;
    public static double RoleLabelHeightDip => ColumnAdvanceDip * RoleLabelHeightChars;
    public static double RecordGapDip => ColumnAdvanceDip * RecordGapChars;
    public static Thickness PageMargin => new Thickness(MarginLeftDip, MarginTopDip, MarginRightDip, MarginBottomDip);
    public static System.Windows.Media.FontFamily DefaultFontFamily { get; } = CreateDefaultFontFamily();

    private static double MmToDip(double mm)
    {
        return mm / 25.4 * 96.0;
    }

    private static double PointsToDip(double points)
    {
        return points / 72.0 * 96.0;
    }

    private static System.Windows.Media.FontFamily CreateDefaultFontFamily()
    {
        var preferred = Fonts.SystemFontFamilies.FirstOrDefault(font =>
            string.Equals(font.Source, DefaultFontFamilyName, StringComparison.OrdinalIgnoreCase) ||
            font.FamilyNames.Values.Any(name => string.Equals(name, DefaultFontFamilyName, StringComparison.OrdinalIgnoreCase)));

        return preferred ?? new System.Windows.Media.FontFamily("MS Mincho");
    }
}
