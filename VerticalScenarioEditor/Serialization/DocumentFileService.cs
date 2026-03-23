using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using VerticalScenarioEditor.Models;

namespace VerticalScenarioEditor.Serialization;

public static class DocumentFileService
{
    private const string InvalidFileMessage = "ファイルの内容が不正です。";
    private const string UnsupportedVersionMessagePrefix = "このファイルのバージョンには対応していません。";

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    public static DocumentState Load(string path)
    {
        var json = File.ReadAllText(path);
        DocumentFile? file;
        try
        {
            file = JsonSerializer.Deserialize<DocumentFile>(json, JsonOptions);
        }
        catch (JsonException ex)
        {
            throw new InvalidDataException(InvalidFileMessage, ex);
        }
        catch (NotSupportedException ex)
        {
            throw new InvalidDataException(InvalidFileMessage, ex);
        }

        if (file == null)
        {
            throw new InvalidDataException(InvalidFileMessage);
        }

        if (file.Version != DocumentFile.CurrentVersion)
        {
            throw new InvalidDataException($"{UnsupportedVersionMessagePrefix} (version: {file.Version})");
        }

        return NormalizeDocument(file.Document);
    }

    public static void Save(string path, DocumentState document)
    {
        if (document == null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        var file = new DocumentFile
        {
            Version = DocumentFile.CurrentVersion,
            Document = NormalizeDocument(document)
        };

        var json = JsonSerializer.Serialize(file, JsonOptions);
        File.WriteAllText(path, json);
    }

    private static DocumentState NormalizeDocument(DocumentState? source)
    {
        var sourceDocument = source ?? DocumentState.CreateDefault();
        var normalized = new DocumentState
        {
            PageNumberEnabled = sourceDocument.PageNumberEnabled,
            ShowGuides = sourceDocument.ShowGuides,
            SummaryText = sourceDocument.SummaryText ?? string.Empty,
            Records = new List<ScriptRecord>(),
            RoleDictionary = new Dictionary<string, string>()
        };

        if (sourceDocument.Records != null)
        {
            foreach (var record in sourceDocument.Records)
            {
                normalized.Records.Add(new ScriptRecord
                {
                    RoleName = record?.RoleName ?? string.Empty,
                    Body = record?.Body ?? string.Empty,
                    PageBreakBefore = record?.PageBreakBefore == true
                });
            }
        }

        if (sourceDocument.RoleDictionary != null)
        {
            foreach (var pair in sourceDocument.RoleDictionary)
            {
                if (string.IsNullOrWhiteSpace(pair.Key))
                {
                    continue;
                }

                normalized.RoleDictionary[pair.Key] = pair.Value ?? string.Empty;
            }
        }

        if (normalized.Records.Count > 0)
        {
            normalized.Records[0].PageBreakBefore = false;
        }

        return normalized;
    }
}

