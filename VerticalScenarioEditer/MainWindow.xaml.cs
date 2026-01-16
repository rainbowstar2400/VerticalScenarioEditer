using System.Windows;

namespace VerticalScenarioEditer;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void OnFileOpenClick(object sender, RoutedEventArgs e)
    {
        ShowNotImplementedMessage("Open");
    }

    private void OnFileSaveClick(object sender, RoutedEventArgs e)
    {
        ShowNotImplementedMessage("Save");
    }

    private static void ShowNotImplementedMessage(string actionName)
    {
        MessageBox.Show($"{actionName} is not implemented in M1.", "Not Implemented", MessageBoxButton.OK, MessageBoxImage.Information);
    }
}
