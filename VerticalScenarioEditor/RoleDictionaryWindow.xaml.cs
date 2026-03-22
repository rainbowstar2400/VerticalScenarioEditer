using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using VerticalScenarioEditor.Models;
using VerticalScenarioEditor.RoleManagement;
using VerticalScenarioEditor.ViewModels;

namespace VerticalScenarioEditor;

public partial class RoleDictionaryWindow : Window
{
    private readonly DocumentState _document;
    public ObservableCollection<RoleColorEntry> Entries { get; } = new();
    public IReadOnlyList<string> ColorPresets { get; } = new[]
    {
        "黄 (#FFFF99)",
        "青 (#DEEAF6)",
        "緑 (#E2EFD9)",
        "桃 (#FFE1E1)",
        "橙 (#FBE4D5)",
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
        RoleDictionarySynchronizer.Synchronize(_document);

        Entries.Clear();
        foreach (var role in _document.RoleDictionary.Keys.OrderBy(name => name))
        {
            Entries.Add(new RoleColorEntry
            {
                RoleName = role,
                Color = _document.RoleDictionary.TryGetValue(role, out var color) ? color : string.Empty
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
            if (string.IsNullOrWhiteSpace(role))
            {
                continue;
            }
            nextDictionary[role] = color;
        }

        _document.RoleDictionary = nextDictionary;
        RoleDictionarySynchronizer.Synchronize(_document);
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

