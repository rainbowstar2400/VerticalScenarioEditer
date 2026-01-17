using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Text.Json;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using VerticalScenarioEditer.Models;
using VerticalScenarioEditer.Serialization;

namespace VerticalScenarioEditer;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private const string FileFilter = "縦書き脚本エディタ (*.vse)|*.vse|JSON (*.json)|*.json|すべてのファイル (*.*)|*.*";
    private DocumentState _document = DocumentState.CreateDefault();
    private string? _currentFilePath;
    private bool _isWebContentReady;

    public MainWindow()
    {
        InitializeComponent();
        UpdateTitle();
        Loaded += OnLoaded;
    }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        await InitializeWebViewAsync();
    }

    private async Task InitializeWebViewAsync()
    {
        try
        {
            await EditorWebView.EnsureCoreWebView2Async();
        }
        catch (WebView2RuntimeNotFoundException)
        {
            MessageBox.Show(
                this,
                "WebView2 Runtime が見つかりません。\nMicrosoft Edge WebView2 Runtime をインストールしてから再起動してください。",
                "WebView2 Runtime 未導入",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            return;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "WebView2 の初期化に失敗しました", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        EditorWebView.NavigationCompleted += OnWebViewNavigationCompleted;

        var htmlPath = Path.Combine(AppContext.BaseDirectory, "Web", "index.html");
        if (File.Exists(htmlPath))
        {
            EditorWebView.Source = new Uri(htmlPath);
        }
        else
        {
            EditorWebView.CoreWebView2.NavigateToString("<html lang='ja'><body><p>WebView2 の表示コンテンツが見つかりません。</p></body></html>");
        }
    }

    private void OnFileOpenClick(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Filter = FileFilter
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            _document = DocumentFileService.Load(dialog.FileName);
            _currentFilePath = dialog.FileName;
            UpdateTitle();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "読み込みに失敗しました", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        SendDocumentToWebView();
    }

    private void OnFileSaveClick(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(_currentFilePath))
        {
            SaveAs();
            return;
        }

        SaveToPath(_currentFilePath);
    }

    private void SaveAs()
    {
        var dialog = new SaveFileDialog
        {
            Filter = FileFilter,
            DefaultExt = ".vse"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        SaveToPath(dialog.FileName);
    }

    private void SaveToPath(string path)
    {
        try
        {
            DocumentFileService.Save(path, _document);
            _currentFilePath = path;
            UpdateTitle();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "保存に失敗しました", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void UpdateTitle()
    {
        var fileLabel = string.IsNullOrWhiteSpace(_currentFilePath) ? "無題" : _currentFilePath;
        Title = $"縦書き脚本エディタ - {fileLabel}";
    }

    private void OnWebViewNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (!e.IsSuccess)
        {
            MessageBox.Show(this, "WebView2 の読み込みに失敗しました。", "読み込みに失敗しました", MessageBoxButton.OK, MessageBoxImage.Error);
            return;
        }

        _isWebContentReady = true;
        SendDocumentToWebView();
    }

    private void SendDocumentToWebView()
    {
        if (EditorWebView.CoreWebView2 == null || !_isWebContentReady)
        {
            return;
        }

        var payload = new
        {
            type = "initDocument",
            document = _document,
            settings = new
            {
                pageWidthMm = DocumentSettings.PageWidthMm,
                pageHeightMm = DocumentSettings.PageHeightMm,
                marginLeftMm = DocumentSettings.MarginLeftMm,
                marginRightMm = DocumentSettings.MarginRightMm,
                marginTopMm = DocumentSettings.MarginTopMm,
                marginBottomMm = DocumentSettings.MarginBottomMm,
                fontFamily = DocumentSettings.DefaultFontFamilyName,
                fontSizePt = DocumentSettings.DefaultFontSizePt,
                lineSpacing = DocumentSettings.LineSpacing
            }
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        EditorWebView.CoreWebView2.PostWebMessageAsJson(json);
    }
}
