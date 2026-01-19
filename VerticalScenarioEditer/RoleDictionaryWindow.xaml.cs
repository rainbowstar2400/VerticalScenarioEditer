using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using VerticalScenarioEditer.Models;
using VerticalScenarioEditer.ViewModels;

namespace VerticalScenarioEditer;

public partial class RoleDictionaryWindow : Window
{
    private readonly DocumentState _document;
    public ObservableCollection<RoleColorEntry> Entries { get; } = new();

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
}
