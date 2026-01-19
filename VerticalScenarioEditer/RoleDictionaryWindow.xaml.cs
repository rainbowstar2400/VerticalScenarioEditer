using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using VerticalScenarioEditer.Models;
using VerticalScenarioEditer.ViewModels;

namespace VerticalScenarioEditer;

public partial class RoleDictionaryWindow : Window
{
    private readonly DocumentState _document;
    public ObservableCollection<RoleColorEntry> Entries { get; } = new();
    public IReadOnlyList<string> ColorPresets { get; } = new[]
    {
        "黄色 (#FFFF99)",
        "橙 (#F2994A)",
        "赤 (#EB5757)",
        "紫 (#BB6BD9)",
        "青 (#2F80ED)",
        "水色 (#56CCF2)",
        "緑 (#6FCF97)",
        "深緑 (#219653)",
        "濃紫 (#9B51E0)",
        "黒 (#333333)",
    };

    public RoleDictionaryWindow(DocumentState document)
    {
        _document = document;
        InitializeComponent();
        DataContext = this;
        RefreshEntries();
    }

    private void RefreshEntries()
    {
        var colorByRole = new Dictionary<string, string>(_document.RoleDictionary);
        var rolesFromRecords = _document.Records
            .Select(record => record.RoleName?.Trim())
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Select(name => name!)
            .Distinct()
            .OrderBy(name => name);

        var allRoles = new List<string>();
        allRoles.AddRange(colorByRole.Keys);
        allRoles.AddRange(rolesFromRecords);

        Entries.Clear();
        foreach (var role in allRoles.Distinct().OrderBy(name => name))
        {
            Entries.Add(new RoleColorEntry
            {
                RoleName = role,
                Color = colorByRole.TryGetValue(role, out var color) ? color : string.Empty
            });
        }
    }

    private void OnRefreshClick(object sender, RoutedEventArgs e)
    {
        RefreshEntries();
    }

    private void OnSaveClick(object sender, RoutedEventArgs e)
    {
        RoleGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
        RoleGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);

        var nextDictionary = new Dictionary<string, string>();
        foreach (var entry in Entries)
        {
            var role = entry.RoleName?.Trim();
            var color = NormalizeColor(entry.Color);
            if (string.IsNullOrWhiteSpace(role) || string.IsNullOrWhiteSpace(color))
            {
                continue;
            }
            nextDictionary[role] = color;
        }

        _document.RoleDictionary = nextDictionary;
        DialogResult = true;
    }

    private static string NormalizeColor(string? input)
    {
        var color = input?.Trim();
        if (string.IsNullOrWhiteSpace(color))
        {
            return string.Empty;
        }

        var extracted = ExtractHexColor(color);
        if (!string.IsNullOrWhiteSpace(extracted))
        {
            return extracted;
        }

        if (color.StartsWith("#", StringComparison.Ordinal))
        {
            return color;
        }

        if (color.Length == 3 || color.Length == 6)
        {
            var isHex = color.All(ch =>
                (ch >= '0' && ch <= '9') ||
                (ch >= 'a' && ch <= 'f') ||
                (ch >= 'A' && ch <= 'F'));
            if (isHex)
            {
                return $"#{color}";
            }
        }

        return color;
    }

    private static string ExtractHexColor(string input)
    {
        var index = input.IndexOf('#');
        if (index < 0 || index + 7 > input.Length)
        {
            return string.Empty;
        }

        var candidate = input.Substring(index, 7);
        var isHex = candidate.Skip(1).All(ch =>
            (ch >= '0' && ch <= '9') ||
            (ch >= 'a' && ch <= 'f') ||
            (ch >= 'A' && ch <= 'F'));
        return isHex ? candidate : string.Empty;
    }

    private void OnPickColorClick(object sender, RoutedEventArgs e)
    {
        if (sender is not System.Windows.Controls.Button button)
        {
            return;
        }

        if (button.DataContext is not RoleColorEntry entry)
        {
            return;
        }

        using var dialog = new ColorDialog
        {
            FullOpen = true
        };

        if (TryParseHexColor(entry.Color, out var color))
        {
            dialog.Color = color;
        }

        if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK)
        {
            return;
        }

        entry.Color = $"#{dialog.Color.R:X2}{dialog.Color.G:X2}{dialog.Color.B:X2}";
        RoleGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Cell, true);
        RoleGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);
    }

    private static bool TryParseHexColor(string? input, out System.Drawing.Color color)
    {
        color = default;
        var value = NormalizeColor(input);
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (value.StartsWith("#", StringComparison.Ordinal))
        {
            value = value[1..];
        }

        if (value.Length != 6)
        {
            return false;
        }

        if (!int.TryParse(value, System.Globalization.NumberStyles.HexNumber, null, out var hex))
        {
            return false;
        }

        color = System.Drawing.Color.FromArgb(
            (hex >> 16) & 0xFF,
            (hex >> 8) & 0xFF,
            hex & 0xFF);
        return true;
    }
}
