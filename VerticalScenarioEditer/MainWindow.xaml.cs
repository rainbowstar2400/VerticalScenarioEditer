using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Text.Json;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using VerticalScenarioEditer.Models;
using VerticalScenarioEditer.Serialization;
using VerticalScenarioEditer.Settings;

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
    private readonly AppSettings _appSettings;

    public MainWindow()
    {
        InitializeComponent();
        _appSettings = AppSettingsStore.Load();
        EnsureAtLeastOneRecord();
        UpdateTitle();
        Loaded += OnLoaded;
        Closing += OnClosing;
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

        EditorWebView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
        EditorWebView.NavigationCompleted += OnWebViewNavigationCompleted;

        var webRoot = Path.Combine(AppContext.BaseDirectory, "Web");
        var htmlPath = Path.Combine(webRoot, "index.html");
        if (!Directory.Exists(webRoot) || !File.Exists(htmlPath))
        {
            EditorWebView.CoreWebView2.NavigateToString("<html lang='ja'><body><p>WebView2 の表示コンテンツが見つかりません。</p></body></html>");
            return;
        }

        EditorWebView.CoreWebView2.SetVirtualHostNameToFolderMapping(
            "app",
            webRoot,
            CoreWebView2HostResourceAccessKind.Allow);
        EditorWebView.Source = new Uri("https://app/index.html");
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
            EnsureAtLeastOneRecord();
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

        EnsureAtLeastOneRecord();
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
                lineSpacing = DocumentSettings.LineSpacing,
                zoomScale = _appSettings.ZoomScale
            }
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        EditorWebView.CoreWebView2.PostWebMessageAsJson(json);
    }

    private void OnWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        try
        {
            using var document = JsonDocument.Parse(e.WebMessageAsJson);
            if (!document.RootElement.TryGetProperty("type", out var typeProperty))
            {
                return;
            }

            var type = typeProperty.GetString();
            switch (type)
            {
                case "documentReady":
                    _isWebContentReady = true;
                    SendDocumentToWebView();
                    break;
                case "inputPatch":
                    ApplyInputPatch(document.RootElement);
                    break;
                case "command":
                    ApplyCommand(document.RootElement);
                    break;
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "WebView2 メッセージの処理に失敗しました", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ApplyInputPatch(JsonElement root)
    {
        if (!root.TryGetProperty("recordIndex", out var indexProperty) || !indexProperty.TryGetInt32(out var recordIndex))
        {
            return;
        }

        if (recordIndex < 0 || recordIndex >= _document.Records.Count)
        {
            return;
        }

        if (!root.TryGetProperty("field", out var fieldProperty))
        {
            return;
        }

        var field = fieldProperty.GetString();
        var text = root.TryGetProperty("text", out var textProperty) ? textProperty.GetString() ?? string.Empty : string.Empty;

        var record = _document.Records[recordIndex];
        if (field == "roleName")
        {
            record.RoleName = text;
        }
        else if (field == "body")
        {
            record.Body = text;
        }
    }

    private void ApplyCommand(JsonElement root)
    {
        if (!root.TryGetProperty("name", out var nameProperty))
        {
            return;
        }

        var name = nameProperty.GetString();
        if (!root.TryGetProperty("recordIndex", out var indexProperty) || !indexProperty.TryGetInt32(out var recordIndex))
        {
            return;
        }

        if (recordIndex < 0 || recordIndex >= _document.Records.Count)
        {
            return;
        }

        switch (name)
        {
            case "insertAfter":
                _document.Records.Insert(recordIndex + 1, new ScriptRecord());
                SendDocumentToWebView();
                break;
            case "deleteRecord":
                _document.Records.RemoveAt(recordIndex);
                EnsureAtLeastOneRecord();
                SendDocumentToWebView();
                break;
        }
    }

    private void EnsureAtLeastOneRecord()
    {
        if (_document.Records.Count == 0)
        {
            _document.Records.Add(new ScriptRecord());
        }
    }

    private void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        try
        {
            AppSettingsStore.Save(_appSettings);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, ex.Message, "設定の保存に失敗しました", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
}
