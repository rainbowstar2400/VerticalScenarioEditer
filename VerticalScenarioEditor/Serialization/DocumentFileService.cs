using System;
using System.IO;
using System.Text.Json;
using VerticalScenarioEditor.Models;

namespace VerticalScenarioEditor.Serialization;

public static class DocumentFileService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true
    };

    public static DocumentState Load(string path)
    {
        var json = File.ReadAllText(path);
        var file = JsonSerializer.Deserialize<DocumentFile>(json, JsonOptions);
        if (file == null)
        {
            throw new InvalidDataException("Invalid file contents.");
        }

        if (file.Version != DocumentFile.CurrentVersion)
        {
            throw new InvalidDataException($"Unsupported file version: {file.Version}.");
        }

        return file.Document ?? DocumentState.CreateDefault();
    }

    public static void Save(string path, DocumentState document)
    {
        var file = new DocumentFile
        {
            Version = DocumentFile.CurrentVersion,
            Document = document
        };

        var json = JsonSerializer.Serialize(file, JsonOptions);
        File.WriteAllText(path, json);
    }
}

