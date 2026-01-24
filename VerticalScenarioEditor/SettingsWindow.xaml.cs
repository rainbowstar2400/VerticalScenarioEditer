using System.Globalization;
using System.Windows;
using VerticalScenarioEditor.Settings;

namespace VerticalScenarioEditor;

public partial class SettingsWindow : Window
{
    public double RoleLabelHeightChars { get; private set; }

    public SettingsWindow(AppSettings settings)
    {
        InitializeComponent();
        RoleLabelHeightChars = settings.RoleLabelHeightChars;
        RoleHeightTextBox.Text = RoleLabelHeightChars.ToString(CultureInfo.CurrentCulture);
    }

    private void OnOkClick(object sender, RoutedEventArgs e)
    {
        if (!double.TryParse(RoleHeightTextBox.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out var value))
        {
            System.Windows.MessageBox.Show(this, "数値を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }
        if (value <= 0)
        {
            System.Windows.MessageBox.Show(this, "0より大きい数値を入力してください。", "入力エラー", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        RoleLabelHeightChars = value;
        DialogResult = true;
    }

    private void OnCancelClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}

