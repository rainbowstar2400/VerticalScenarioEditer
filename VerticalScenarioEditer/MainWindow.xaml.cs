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
    private bool _hasOverflow;
    private int? _lastFocusedRecordIndex;
    private readonly System.Collections.Generic.HashSet<int> _overflowWarnedOnceRecords = new();
    private readonly System.Collections.Generic.HashSet<int> _overflowWarnedAgainRecords = new();
    private readonly System.Collections.Generic.HashSet<int> _overflowAttentionRecords = new();
    private int? _selectionStartRecordIndex;
    private int? _selectionEndRecordIndex;
    private bool _isSelectionMode;
    private bool _isSummaryMode;
    private bool _isSimpleMode;
    private bool _isDirty;
    private bool _suppressToggleEvents;
    private TaskCompletionSource<bool>? _pdfReadyTcs;
    private bool _isApplyingHistory;
    private readonly System.Collections.Generic.Stack<DocumentState> _undoStack = new();
    private readonly System.Collections.Generic.Stack<DocumentState> _redoStack = new();

    public MainWindow()
    {
        _appSettings = AppSettingsStore.Load();
        InitializeComponent();
        _isUiReady = true;
        EnsureAtLeastOneRecord();
        UpdateTitle();
        UpdatePageNumberToggleState();
        UpdateGuideLineToggleState();
        UpdateSelectionModeToggleState();
        UpdateSummaryModeToggleState();
        UpdateSimpleModeToggleState();
        UpdateSelectionModeAvailability();
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
        if (!ConfirmSaveIfDirty())
        {
            return;
        }

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
            _isDirty = false;
            EnsureAtLeastOneRecord();
            UpdateTitle();
            UpdatePageNumberToggleState();
            UpdateGuideLineToggleState();
            UpdateSelectionModeToggleState();
            UpdateSummaryModeToggleState();
            UpdateSimpleModeToggleState();
            UpdateSelectionModeAvailability();
            ResetOverflowWarningState();
            ClearSelectionRange();
            ClearHistory();
            UpdateStatusBar();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "読み込みに失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }

        SendDocumentToWebView();
    }

    private void OnImportFormattedTextClick(object sender, RoutedEventArgs e)
    {
        if (!ConfirmSaveIfDirty())
        {
            return;
        }

        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "テキスト (*.txt)|*.txt|すべてのファイル (*.*)|*.*"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            var lines = File.ReadAllLines(dialog.FileName);
            _document = ParseFormattedText(lines);
            _currentFilePath = null;
            _isDirty = true;
            EnsureAtLeastOneRecord();
            UpdateTitle();
            UpdatePageNumberToggleState();
            UpdateGuideLineToggleState();
            UpdateSelectionModeToggleState();
            ResetOverflowWarningState();
            ClearSelectionRange();
            ClearHistory();
            UpdateStatusBar();
            SendDocumentToWebView();
            SendRoleDictionaryToWebView();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "読み込みに失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    private void OnFileNewClick(object sender, RoutedEventArgs e)
    {
        if (!ConfirmSaveIfDirty())
        {
            return;
        }

        _document = DocumentState.CreateDefault();
        _currentFilePath = null;
        _isDirty = false;
        EnsureAtLeastOneRecord();
        UpdateTitle();
        UpdatePageNumberToggleState();
        UpdateGuideLineToggleState();
        UpdateSelectionModeToggleState();
        ResetOverflowWarningState();
        ClearSelectionRange();
        ClearHistory();
        UpdateStatusBar();
        SendDocumentToWebView();
        SendRoleDictionaryToWebView();
    }

    private void OnFileCloseClick(object sender, RoutedEventArgs e)
    {
        Close();
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

    private void OnCopyDialogueTextClick(object sender, RoutedEventArgs e)
    {
        var text = BuildDialogueText();
        try
        {
            System.Windows.Clipboard.SetText(text);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "コピーに失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    private void OnUndoClick(object sender, RoutedEventArgs e)
    {
        Undo();
    }

    private void OnRedoClick(object sender, RoutedEventArgs e)
    {
        Redo();
    }

    private void OnRoleDictionaryClick(object sender, RoutedEventArgs e)
    {
        var snapshot = CloneDocument(_document);
        var window = new RoleDictionaryWindow(_document)
        {
            Owner = this
        };

        if (window.ShowDialog() == true)
        {
            PushUndoState(snapshot);
            MarkDirty();
            SendDocumentToWebView();
            SendRoleDictionaryToWebView();
        }
    }

    private void OnOpenSettingsClick(object sender, RoutedEventArgs e)
    {
        var window = new SettingsWindow(_appSettings)
        {
            Owner = this
        };

        if (window.ShowDialog() == true)
        {
            _appSettings.RoleLabelHeightChars = window.RoleLabelHeightChars;
            SendDocumentToWebView();
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

    private string BuildExportText()
    {
        var builder = new StringBuilder();
        var start = _selectionStartRecordIndex;
        var end = _selectionEndRecordIndex;
        var hasSelection = start.HasValue && end.HasValue;
        var selectionStart = hasSelection ? Math.Min(start!.Value, end!.Value) : 0;
        var selectionEnd = hasSelection ? Math.Max(start!.Value, end!.Value) : _document.Records.Count - 1;
        if (selectionStart < 0 || selectionEnd >= _document.Records.Count)
        {
            selectionStart = 0;
            selectionEnd = _document.Records.Count - 1;
        }

        for (var index = selectionStart; index <= selectionEnd && index < _document.Records.Count; index += 1)
        {
            var record = _document.Records[index];
            var role = record.RoleName?.Trim() ?? string.Empty;
            var body = record.Body ?? string.Empty;
            body = body.Replace("\r", string.Empty).Replace("\n", string.Empty);

            string line;
            if (string.IsNullOrWhiteSpace(role))
            {
                line = $"\t{body}";
            }
            else if (role.StartsWith("シーン", StringComparison.Ordinal))
            {
                line = $"【{role}：{body}】";
            }
            else if (body.StartsWith('（') && body.EndsWith('）'))
            {
                line = $"{role}{body}";
            }
            else
            {
                line = $"{role}「{body}」";
            }

            builder.AppendLine(line);
        }

        return builder.ToString().TrimEnd();
    }

    private string BuildDialogueText()
    {
        var builder = new StringBuilder();
        var start = _selectionStartRecordIndex;
        var end = _selectionEndRecordIndex;
        var hasSelection = start.HasValue && end.HasValue;
        var selectionStart = hasSelection ? Math.Min(start!.Value, end!.Value) : 0;
        var selectionEnd = hasSelection ? Math.Max(start!.Value, end!.Value) : _document.Records.Count - 1;
        if (selectionStart < 0 || selectionEnd >= _document.Records.Count)
        {
            selectionStart = 0;
            selectionEnd = _document.Records.Count - 1;
        }

        for (var index = selectionStart; index <= selectionEnd && index < _document.Records.Count; index += 1)
        {
            var record = _document.Records[index];
            var body = record.Body ?? string.Empty;
            body = body.Replace("\r", string.Empty).Replace("\n", string.Empty);
            if (string.IsNullOrWhiteSpace(body))
            {
                continue;
            }
            builder.AppendLine(body);
            builder.AppendLine();
        }

        return builder.ToString().TrimEnd();
    }

    private static DocumentState ParseFormattedText(string[] lines)
    {
        var document = DocumentState.CreateDefault();
        document.Records.Clear();

        foreach (var rawLine in lines)
        {
            var line = rawLine?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(line))
            {
                continue;
            }

            var record = new ScriptRecord();
            var trimmed = line.TrimStart('\t');
            if (trimmed != line)
            {
                line = trimmed;
            }
            var openIndex = line.IndexOf('「');
            var parenIndex = line.IndexOf('（');
            if (line.StartsWith('【') && line.EndsWith('】') && line.Contains('：') && line.StartsWith("【シーン", StringComparison.Ordinal))
            {
                var inner = line.Substring(1, line.Length - 2);
                var separatorIndex = inner.IndexOf('：');
                if (separatorIndex > 0)
                {
                    record.RoleName = inner.Substring(0, separatorIndex).Trim();
                    record.Body = inner.Substring(separatorIndex + 1);
                    document.Records.Add(record);
                    continue;
                }
            }
            var sceneSeparatorIndex = line.IndexOf('：');
            if (sceneSeparatorIndex > 0 && line.StartsWith("シーン", StringComparison.Ordinal))
            {
                record.RoleName = line.Substring(0, sceneSeparatorIndex).Trim();
                record.Body = line.Substring(sceneSeparatorIndex + 1);
            }
            else if (parenIndex > 0 && (openIndex < 0 || parenIndex < openIndex))
            {
                record.RoleName = line.Substring(0, parenIndex).Trim();
                record.Body = line.Substring(parenIndex);
            }
            else if (openIndex > 0 && line.EndsWith('」'))
            {
                record.RoleName = line.Substring(0, openIndex).Trim();
                record.Body = line.Substring(openIndex + 1, line.Length - openIndex - 2);
            }
            else if (openIndex > 0)
            {
                var closeIndex = line.LastIndexOf('」');
                if (closeIndex > openIndex)
                {
                    record.RoleName = line.Substring(0, openIndex).Trim();
                    var mainBody = line.Substring(openIndex + 1, closeIndex - openIndex - 1);
                    var tail = line.Substring(closeIndex + 1);
                    record.Body = $"{mainBody}{tail}";
                }
                else
                {
                    record.RoleName = string.Empty;
                    record.Body = line;
                }
            }
            else
            {
                record.RoleName = string.Empty;
                record.Body = line;
            }

            document.Records.Add(record);
        }

        return document;
    }

    private bool SaveAs()
    {
        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = FileFilter,
            DefaultExt = ".vse"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return false;
        }

        return SaveToPath(dialog.FileName);
    }

    private void SaveCopyAs()
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

        try
        {
            DocumentFileService.Save(dialog.FileName, _document);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "保存に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    private bool SaveToPath(string path)
    {
        try
        {
            DocumentFileService.Save(path, _document);
            _currentFilePath = path;
            _isDirty = false;
            UpdateTitle();
            return true;
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "保存に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            return false;
        }
    }

    private void UpdateTitle()
    {
        var fileLabel = string.IsNullOrWhiteSpace(_currentFilePath) ? "無題" : _currentFilePath;
        var dirtyMark = _isDirty ? " *" : string.Empty;
        Title = $"縦書き脚本エディタ - {fileLabel}{dirtyMark}";
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
                pageNumberEnabled = _document.PageNumberEnabled,
                showGuides = _document.ShowGuides,
                roleLabelHeightChars = _appSettings.RoleLabelHeightChars,
                zoomScale = _appSettings.ZoomScale,
                selectionMode = _isSelectionMode,
                summaryMode = _isSummaryMode,
                simpleMode = _isSimpleMode
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
                case "selectionChanged":
                    ApplySelectionChanged(document.RootElement);
                    break;
                case "pdfReady":
                    var hasOverflow = document.RootElement.TryGetProperty("hasOverflow", out var overflowProperty)
                        && overflowProperty.ValueKind == JsonValueKind.True;
                    _pdfReadyTcs?.TrySetResult(!hasOverflow);
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
            if (record.RoleName != text)
            {
                PushUndoState();
                MarkDirty();
            }
            record.RoleName = text;
        }
        else if (field == "body")
        {
            if (record.Body != text)
            {
                PushUndoState();
                MarkDirty();
            }
            record.Body = text;
        }
        else if (field == "summary")
        {
            if (_document.SummaryText != text)
            {
                PushUndoState();
                MarkDirty();
            }
            _document.SummaryText = text;
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
        if (name == "undo")
        {
            Undo();
            return;
        }
        if (name == "redo")
        {
            Redo();
            return;
        }

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
                PushUndoState();
                _document.Records.Insert(recordIndex + 1, new ScriptRecord());
                MarkDirty();
                SendDocumentToWebView();
                UpdateStatusBar();
                break;
            case "insertBefore":
                PushUndoState();
                _document.Records.Insert(recordIndex, new ScriptRecord());
                MarkDirty();
                SendDocumentToWebView();
                UpdateStatusBar();
                break;
            case "deleteRecord":
                PushUndoState();
                _document.Records.RemoveAt(recordIndex);
                MarkDirty();
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

    private void PushUndoState()
    {
        if (_isApplyingHistory)
        {
            return;
        }
        PushUndoState(CloneDocument(_document));
    }

    private void PushUndoState(DocumentState snapshot)
    {
        if (_isApplyingHistory)
        {
            return;
        }
        _undoStack.Push(snapshot);
        _redoStack.Clear();
    }

    private void ClearHistory()
    {
        _undoStack.Clear();
        _redoStack.Clear();
    }

    private DocumentState CloneDocument(DocumentState source)
    {
        var clone = DocumentState.CreateDefault();
        clone.PageNumberEnabled = source.PageNumberEnabled;
        clone.ShowGuides = source.ShowGuides;
        clone.SummaryText = source.SummaryText ?? string.Empty;
        clone.Records = new System.Collections.Generic.List<ScriptRecord>();
        foreach (var record in source.Records)
        {
            clone.Records.Add(new ScriptRecord
            {
                RoleName = record.RoleName,
                Body = record.Body
            });
        }

        clone.RoleDictionary = new System.Collections.Generic.Dictionary<string, string>(source.RoleDictionary);
        return clone;
    }

    private void ApplyDocumentState(DocumentState state)
    {
        _document = CloneDocument(state);
        EnsureAtLeastOneRecord();
        UpdatePageNumberToggleState();
        UpdateGuideLineToggleState();
        UpdateTitle();
        UpdateStatusBar();
        ResetOverflowWarningState();
        ClearSelectionRange();
        SendDocumentToWebView();
        SendRoleDictionaryToWebView();
    }

    private void Undo()
    {
        if (_undoStack.Count == 0)
        {
            return;
        }
        _isApplyingHistory = true;
        _redoStack.Push(CloneDocument(_document));
        var previous = _undoStack.Pop();
        _isDirty = true;
        ApplyDocumentState(previous);
        _isApplyingHistory = false;
    }

    private void Redo()
    {
        if (_redoStack.Count == 0)
        {
            return;
        }
        _isApplyingHistory = true;
        _undoStack.Push(CloneDocument(_document));
        var next = _redoStack.Pop();
        _isDirty = true;
        ApplyDocumentState(next);
        _isApplyingHistory = false;
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

        var modeLabel = _isSummaryMode ? "概要" : _isSimpleMode ? "簡易" : "本文";
        ModeStatusText.Text = $"モード: {modeLabel}";
    }

    private void UpdatePageNumberToggleState()
    {
        if (PageNumberToggleMenuItem == null)
        {
            return;
        }
        _suppressToggleEvents = true;
        PageNumberToggleMenuItem.IsChecked = _document.PageNumberEnabled;
        _suppressToggleEvents = false;
    }

    private void UpdateGuideLineToggleState()
    {
        if (GuideLineToggleMenuItem == null)
        {
            return;
        }
        _suppressToggleEvents = true;
        GuideLineToggleMenuItem.IsChecked = _document.ShowGuides;
        _suppressToggleEvents = false;
    }

    private void UpdateSelectionModeToggleState()
    {
        if (SelectionModeMenuItem == null)
        {
            return;
        }
        SelectionModeMenuItem.IsChecked = _isSelectionMode;
    }

    private void UpdateSummaryModeToggleState()
    {
        if (SummaryModeMenuItem == null)
        {
            return;
        }
        SummaryModeMenuItem.IsChecked = _isSummaryMode;
    }

    private void UpdateSimpleModeToggleState()
    {
        if (SimpleModeMenuItem == null)
        {
            return;
        }
        SimpleModeMenuItem.IsChecked = _isSimpleMode;
    }

    private void UpdateSelectionModeAvailability()
    {
        if (SelectionModeMenuItem == null)
        {
            return;
        }
        SelectionModeMenuItem.IsEnabled = !_isSummaryMode;
    }

    private void ResetOverflowWarningState()
    {
        _overflowWarnedOnceRecords.Clear();
        _overflowWarnedAgainRecords.Clear();
        _overflowAttentionRecords.Clear();
        _lastFocusedRecordIndex = null;
        _hasLayoutStatus = false;
        _hasOverflow = false;
        _currentPage = 1;
        _totalPages = 1;
        SendOverflowAttentionToWebView();
    }

    private void ClearSelectionRange()
    {
        _selectionStartRecordIndex = null;
        _selectionEndRecordIndex = null;
    }

    private void MarkDirty()
    {
        if (_isDirty)
        {
            return;
        }
        _isDirty = true;
        UpdateTitle();
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

    private async void OnExportPdfClick(object sender, RoutedEventArgs e)
    {
        if (_hasOverflow)
        {
            System.Windows.MessageBox.Show(this, "警告: 1ページに収まらないレコードがあるため、PDF出力を停止しました。", "PDF出力を停止しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            return;
        }

        if (EditorWebView.CoreWebView2 == null || !_isWebContentReady)
        {
            System.Windows.MessageBox.Show(this, "WebView2 が準備できていません。", "PDF出力を停止しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            return;
        }

        var dialog = new Microsoft.Win32.SaveFileDialog
        {
            Filter = "PDF (*.pdf)|*.pdf|すべてのファイル (*.*)|*.*",
            DefaultExt = ".pdf"
        };

        if (dialog.ShowDialog(this) != true)
        {
            return;
        }

        try
        {
            if (!string.IsNullOrWhiteSpace(_document.SummaryText))
            {
                if (!await EnterPdfCombinedModeAsync())
                {
                    System.Windows.MessageBox.Show(this, "警告: 概要ページまたは本文に収まらない内容があるため、PDF出力を停止しました。", "PDF出力を停止しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }
            }

            var settings = EditorWebView.CoreWebView2.Environment.CreatePrintSettings();
            settings.ShouldPrintBackgrounds = true;
            settings.Orientation = CoreWebView2PrintOrientation.Landscape;
            TrySetPdfPageSize(settings, 11.69, 8.27);
            var success = await EditorWebView.CoreWebView2.PrintToPdfAsync(dialog.FileName, settings);
            if (!success)
            {
                System.Windows.MessageBox.Show(this, "PDF出力に失敗しました。", "PDF出力に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "PDF出力に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
        finally
        {
            if (!string.IsNullOrWhiteSpace(_document.SummaryText))
            {
                ExitPdfCombinedMode();
            }
        }
    }

    private void ApplyLayoutStatus(JsonElement root)
    {
        var overflowRecords = GetOverflowRecords(root);
        var focusedRecordIndex = GetFocusedRecordIndex(root);
        _hasOverflow = overflowRecords.Count > 0;

        if (root.TryGetProperty("totalPages", out var totalProperty) && totalProperty.TryGetInt32(out var totalPages))
        {
            _totalPages = Math.Max(1, totalPages);
        }

        if (root.TryGetProperty("currentPage", out var currentProperty) && currentProperty.TryGetInt32(out var currentPage))
        {
            _currentPage = Math.Clamp(currentPage, 1, _totalPages);
        }

        if (root.TryGetProperty("overflowCount", out var overflowProperty) && overflowProperty.TryGetInt32(out var overflowCount))
        {
            _hasOverflow = overflowCount > 0;
        }

        _hasLayoutStatus = true;
        UpdateOverflowWarnings(overflowRecords, focusedRecordIndex);
        UpdateStatusBar();
    }

    private System.Collections.Generic.HashSet<int> GetOverflowRecords(JsonElement root)
    {
        var records = new System.Collections.Generic.HashSet<int>();
        if (!root.TryGetProperty("overflowRecords", out var overflowProperty) || overflowProperty.ValueKind != JsonValueKind.Array)
        {
            return records;
        }

        foreach (var item in overflowProperty.EnumerateArray())
        {
            if (item.TryGetInt32(out var index))
            {
                records.Add(index);
            }
        }

        return records;
    }

    private int? GetFocusedRecordIndex(JsonElement root)
    {
        if (!root.TryGetProperty("focusedRecordIndex", out var focusedProperty))
        {
            return null;
        }

        if (focusedProperty.ValueKind == JsonValueKind.Null)
        {
            return null;
        }

        if (focusedProperty.TryGetInt32(out var index))
        {
            return index;
        }

        return null;
    }

    private void ApplySelectionChanged(JsonElement root)
    {
        var start = GetNullableRecordIndex(root, "startRecordIndex");
        var end = GetNullableRecordIndex(root, "endRecordIndex");
        if (!start.HasValue || !end.HasValue)
        {
            ClearSelectionRange();
            return;
        }

        var normalizedStart = Math.Min(start.Value, end.Value);
        var normalizedEnd = Math.Max(start.Value, end.Value);
        if (normalizedStart < 0 || normalizedEnd >= _document.Records.Count)
        {
            ClearSelectionRange();
            return;
        }

        _selectionStartRecordIndex = normalizedStart;
        _selectionEndRecordIndex = normalizedEnd;
    }

    private static int? GetNullableRecordIndex(JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out var property))
        {
            return null;
        }
        if (property.ValueKind == JsonValueKind.Null)
        {
            return null;
        }
        if (property.TryGetInt32(out var index))
        {
            return index;
        }
        return null;
    }

    private void UpdateOverflowWarnings(System.Collections.Generic.HashSet<int> overflowRecords, int? focusedRecordIndex)
    {
        _overflowWarnedOnceRecords.RemoveWhere(index => !overflowRecords.Contains(index));
        _overflowWarnedAgainRecords.RemoveWhere(index => !overflowRecords.Contains(index));
        _overflowAttentionRecords.RemoveWhere(index => !overflowRecords.Contains(index));

        var hasNewOverflow = false;
        foreach (var index in overflowRecords)
        {
            if (_overflowWarnedOnceRecords.Add(index))
            {
                hasNewOverflow = true;
            }
        }

        if (hasNewOverflow)
        {
            System.Windows.MessageBox.Show(this, "警告: 1ページに収まらないレコードがあります。区切ってください。", "溢れ警告", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
        }

        if (focusedRecordIndex.HasValue && _lastFocusedRecordIndex.HasValue && focusedRecordIndex.Value != _lastFocusedRecordIndex.Value)
        {
            var previousIndex = _lastFocusedRecordIndex.Value;
            if (overflowRecords.Contains(previousIndex) && _overflowWarnedAgainRecords.Add(previousIndex))
            {
                _overflowAttentionRecords.Add(previousIndex);
                System.Windows.MessageBox.Show(this, "警告: 1ページに収まらないレコードが未解決のままです。該当レコードを確認してください。", "溢れ警告", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
            }
        }

        _lastFocusedRecordIndex = focusedRecordIndex;
        SendOverflowAttentionToWebView();
    }

    private void SendOverflowAttentionToWebView()
    {
        if (EditorWebView.CoreWebView2 == null || !_isWebContentReady)
        {
            return;
        }

        var payload = new
        {
            type = "applyOverflowAttention",
            overflowAttentionRecords = _overflowAttentionRecords
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        EditorWebView.CoreWebView2.PostWebMessageAsJson(json);
    }

    private void SendSelectionModeToWebView()
    {
        if (EditorWebView.CoreWebView2 == null || !_isWebContentReady)
        {
            return;
        }

        var payload = new
        {
            type = "applySelectionMode",
            selectionMode = _isSelectionMode
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        EditorWebView.CoreWebView2.PostWebMessageAsJson(json);
    }

    private static void TrySetPdfPageSize(CoreWebView2PrintSettings settings, double widthInches, double heightInches)
    {
        var type = settings.GetType();
        var widthProperty = type.GetProperty("PageWidth");
        var heightProperty = type.GetProperty("PageHeight");
        if (widthProperty != null && widthProperty.CanWrite)
        {
            widthProperty.SetValue(settings, widthInches);
        }
        if (heightProperty != null && heightProperty.CanWrite)
        {
            heightProperty.SetValue(settings, heightInches);
        }
    }

    private void OnClosing(object? sender, System.ComponentModel.CancelEventArgs e)
    {
        if (!ConfirmSaveIfDirty())
        {
            e.Cancel = true;
            return;
        }

        try
        {
            AppSettingsStore.Save(_appSettings);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(this, ex.Message, "設定の保存に失敗しました", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
        }
    }

    private void OnFileSaveAsClick(object sender, RoutedEventArgs e)
    {
        SaveAs();
    }

    private void OnPageNumberToggleChanged(object sender, RoutedEventArgs e)
    {
        if (PageNumberToggleMenuItem == null)
        {
            return;
        }
        if (_suppressToggleEvents)
        {
            return;
        }

        _document.PageNumberEnabled = PageNumberToggleMenuItem.IsChecked == true;
        MarkDirty();
        SendDocumentToWebView();
    }

    private void OnGuideLineToggleChanged(object sender, RoutedEventArgs e)
    {
        if (GuideLineToggleMenuItem == null)
        {
            return;
        }
        if (_suppressToggleEvents)
        {
            return;
        }

        _document.ShowGuides = GuideLineToggleMenuItem.IsChecked == true;
        MarkDirty();
        SendDocumentToWebView();
    }

    private async Task<bool> EnterPdfCombinedModeAsync()
    {
        if (EditorWebView.CoreWebView2 == null || !_isWebContentReady)
        {
            return false;
        }

        _pdfReadyTcs?.TrySetCanceled();
        _pdfReadyTcs = new TaskCompletionSource<bool>();

        var payload = new
        {
            type = "enterPdfMode",
            summaryText = _document.SummaryText ?? string.Empty
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        EditorWebView.CoreWebView2.PostWebMessageAsJson(json);
        var completed = await Task.WhenAny(_pdfReadyTcs.Task, Task.Delay(1500));
        return completed == _pdfReadyTcs.Task && _pdfReadyTcs.Task.Result;
    }

    private void ExitPdfCombinedMode()
    {
        if (EditorWebView.CoreWebView2 == null || !_isWebContentReady)
        {
            return;
        }

        var payload = new
        {
            type = "exitPdfMode"
        };

        var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        });

        EditorWebView.CoreWebView2.PostWebMessageAsJson(json);
    }

    private bool ConfirmSaveIfDirty()
    {
        if (!_isDirty)
        {
            return true;
        }

        var result = System.Windows.MessageBox.Show(
            this,
            "変更が保存されていません。保存しますか？",
            "変更の保存確認",
            System.Windows.MessageBoxButton.YesNoCancel,
            System.Windows.MessageBoxImage.Warning);
        if (result == System.Windows.MessageBoxResult.Cancel)
        {
            return false;
        }
        if (result == System.Windows.MessageBoxResult.Yes)
        {
            return string.IsNullOrWhiteSpace(_currentFilePath)
                ? SaveAs()
                : SaveToPath(_currentFilePath);
        }
        return true;
    }

    private void OnSelectionModeToggleChanged(object sender, RoutedEventArgs e)
    {
        if (SelectionModeMenuItem == null)
        {
            return;
        }

        _isSelectionMode = SelectionModeMenuItem.IsChecked == true;
        if (_isSelectionMode)
        {
            _isSummaryMode = false;
            UpdateSummaryModeToggleState();
        }
        ClearSelectionRange();
        SendSelectionModeToWebView();
        SendDocumentToWebView();
        UpdateSelectionModeAvailability();
    }

    private void OnSummaryModeToggleChanged(object sender, RoutedEventArgs e)
    {
        if (SummaryModeMenuItem == null)
        {
            return;
        }

        _isSummaryMode = SummaryModeMenuItem.IsChecked == true;
        if (_isSummaryMode)
        {
            _isSelectionMode = false;
            UpdateSelectionModeToggleState();
            _isSimpleMode = false;
            UpdateSimpleModeToggleState();
        }
        ClearSelectionRange();
        SendSelectionModeToWebView();
        SendDocumentToWebView();
        UpdateSelectionModeAvailability();
    }

    private void OnSimpleModeToggleChanged(object sender, RoutedEventArgs e)
    {
        if (SimpleModeMenuItem == null)
        {
            return;
        }

        _isSimpleMode = SimpleModeMenuItem.IsChecked == true;
        if (_isSimpleMode)
        {
            _isSummaryMode = false;
            UpdateSummaryModeToggleState();
        }
        SendDocumentToWebView();
    }
}
