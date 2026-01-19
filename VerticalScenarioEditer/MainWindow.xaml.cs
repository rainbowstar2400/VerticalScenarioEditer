using System;
using System.IO;
using System.Text;
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
    private bool _isUiReady;
    private int _currentPage = 1;
    private int _totalPages = 1;
    private bool _hasLayoutStatus;

    public MainWindow()
    {
        _appSettings = AppSettingsStore.Load();
        InitializeComponent();
        _isUiReady = true;
        EnsureAtLeastOneRecord();
        UpdateTitle();
        UpdateZoomUi(_appSettings.ZoomScale);
        UpdateStatusBar();
        Loaded += OnLoaded;
        Closing += OnClosing;
        AddHandler(System.Windows.UIElement.PreviewMouseWheelEvent, new System.Windows.Input.MouseWheelEventHandler(OnWindowPreviewMouseWheel), true);
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
            System.Windows.MessageBox.Show(
                this,
                "WebView2 Runtime が見つかりません。\nMicrosoft Edge WebView2 Runtime をインストールしてから再起動してください。",
                "WebView2 Runtime 未導入",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
            return;
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "WebView2 の初期化に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            return;
        }

        EditorWebView.CoreWebView2.WebMessageReceived += OnWebMessageReceived;
        EditorWebView.NavigationCompleted += OnWebViewNavigationCompleted;
        EditorWebView.CoreWebView2.Settings.IsZoomControlEnabled = true;
        EditorWebView.ZoomFactor = _appSettings.ZoomScale;

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
        var dialog = new Microsoft.Win32.OpenFileDialog
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
            UpdateStatusBar();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "読み込みに失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
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

    private void OnExportTextClick(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "テキスト (*.txt)|*.txt|すべてのファイル (*.*)|*.*",
            DefaultExt = ".txt"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        var text = BuildExportText();
        try
        {
            File.WriteAllText(dialog.FileName, text, new UTF8Encoding(true));
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "書き出しに失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    private void OnCopyTextClick(object sender, RoutedEventArgs e)
    {
        var text = BuildExportText();
        try
        {
            System.Windows.Clipboard.SetText(text);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "コピーに失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    private void OnRoleDictionaryClick(object sender, RoutedEventArgs e)
    {
        var window = new RoleDictionaryWindow(_document)
        {
            Owner = this
        };

        if (window.ShowDialog() == true)
        {
            SendDocumentToWebView();
            SendRoleDictionaryToWebView();
        }
    }

    private string BuildExportText()
    {
        var builder = new StringBuilder();
        foreach (var record in _document.Records)
        {
            var role = record.RoleName?.Trim() ?? string.Empty;
            var body = record.Body ?? string.Empty;
            body = body.Replace("\r", string.Empty).Replace("\n", string.Empty);

            string line;
            if (string.IsNullOrWhiteSpace(role))
            {
                line = body;
            }
            else
            {
                line = $"{role}「{body}」";
            }

            builder.AppendLine(line);
        }

        return builder.ToString().TrimEnd();
    }

    private void SaveAs()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
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
            System.Windows.MessageBox.Show(this, ex.Message, "保存に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
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
            System.Windows.MessageBox.Show(this, "WebView2 の読み込みに失敗しました。", "読み込みに失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
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
                pageGapPx = DocumentSettings.PageGapDip,
                zoomScale = _appSettings.ZoomScale
            }
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        EditorWebView.CoreWebView2.PostWebMessageAsJson(json);
    }

    private void SendRoleDictionaryToWebView()
    {
        if (EditorWebView.CoreWebView2 == null || !_isWebContentReady)
        {
            return;
        }

        var payload = new
        {
            type = "applyRoleDictionary",
            roleDictionary = _document.RoleDictionary
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
                case "zoomDelta":
                    ApplyZoomDelta(document.RootElement);
                    break;
                case "layoutStatus":
                    ApplyLayoutStatus(document.RootElement);
                    break;
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "WebView2 メッセージの処理に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
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

        UpdateStatusBar();
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
                UpdateStatusBar();
                break;
            case "deleteRecord":
                _document.Records.RemoveAt(recordIndex);
                EnsureAtLeastOneRecord();
                SendDocumentToWebView();
                UpdateStatusBar();
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

    private void UpdateStatusBar()
    {
        var totalChars = 0;
        foreach (var record in _document.Records)
        {
            if (record.RoleName != null)
            {
                totalChars += record.RoleName.Length;
            }
            if (record.Body != null)
            {
                totalChars += record.Body.Length;
            }
        }

        CharCountText.Text = $"文字数: {totalChars}";
        if (_hasLayoutStatus)
        {
            PageInfoText.Text = $"ページ: {_currentPage}/{_totalPages}";
        }
        else if (string.IsNullOrWhiteSpace(PageInfoText.Text))
        {
            PageInfoText.Text = "ページ: -";
        }
    }

    private void UpdateZoomUi(double zoomScale)
    {
        if (ZoomSlider == null || ZoomValueText == null)
        {
            return;
        }
        if (ZoomSlider.Value != zoomScale)
        {
            ZoomSlider.Value = zoomScale;
        }
        var percent = Math.Round(zoomScale * 100.0);
        ZoomValueText.Text = $"{percent}%";
    }

    private void OnZoomSliderValueChanged(object sender, System.Windows.RoutedPropertyChangedEventArgs<double> e)
    {
        if (!_isUiReady)
        {
            return;
        }
        var zoom = Math.Clamp(e.NewValue, 0.5, 2.0);
        _appSettings.ZoomScale = zoom;
        UpdateZoomUi(zoom);
        if (EditorWebView.CoreWebView2 != null)
        {
            EditorWebView.ZoomFactor = zoom;
        }
    }


    private void OnZoomSliderMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
    {
        if (sender is not System.Windows.Controls.Slider slider)
        {
            return;
        }

        var step = slider.SmallChange > 0 ? slider.SmallChange : 0.05;
        var direction = e.Delta > 0 ? 1 : -1;
        var next = slider.Value + (step * direction);
        slider.Value = Math.Clamp(next, slider.Minimum, slider.Maximum);
        e.Handled = true;
    }

    private void OnWindowPreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
    {
        if (!_isUiReady || ZoomSlider == null)
        {
            return;
        }

        if (System.Windows.Input.Keyboard.Modifiers != System.Windows.Input.ModifierKeys.Control)
        {
            return;
        }

        var step = ZoomSlider.SmallChange > 0 ? ZoomSlider.SmallChange : 0.05;
        var direction = e.Delta > 0 ? 1 : -1;
        var next = ZoomSlider.Value + (step * direction);
        ZoomSlider.Value = Math.Clamp(next, ZoomSlider.Minimum, ZoomSlider.Maximum);
        e.Handled = true;
    }

    private void ApplyZoomDelta(JsonElement root)
    {
        if (ZoomSlider == null || !_isUiReady)
        {
            return;
        }

        if (!root.TryGetProperty("direction", out var directionProperty))
        {
            return;
        }

        var direction = directionProperty.GetInt32();
        var step = ZoomSlider.SmallChange > 0 ? ZoomSlider.SmallChange : 0.05;
        var next = ZoomSlider.Value + (step * direction);
        ZoomSlider.Value = Math.Clamp(next, ZoomSlider.Minimum, ZoomSlider.Maximum);
    }

    private void ApplyLayoutStatus(JsonElement root)
    {
        if (root.TryGetProperty("totalPages", out var totalProperty) && totalProperty.TryGetInt32(out var totalPages))
        {
            _totalPages = Math.Max(1, totalPages);
        }

        if (root.TryGetProperty("currentPage", out var currentProperty) && currentProperty.TryGetInt32(out var currentPage))
        {
            _currentPage = Math.Clamp(currentPage, 1, _totalPages);
        }

        _hasLayoutStatus = true;
        UpdateStatusBar();
    }

    private void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        try
        {
            AppSettingsStore.Save(_appSettings);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "設定の保存に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }
}
