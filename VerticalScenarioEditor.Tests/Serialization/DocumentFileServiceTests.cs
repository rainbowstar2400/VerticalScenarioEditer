using System;
using System.IO;
using System.Linq;
using VerticalScenarioEditor.Models;
using VerticalScenarioEditor.Serialization;
using Xunit;

namespace VerticalScenarioEditor.Tests.Serialization;

public sealed class DocumentFileServiceTests : IDisposable
{
    private readonly string _tempDirectory;

    public DocumentFileServiceTests()
    {
        _tempDirectory = Path.Combine(
            Path.GetTempPath(),
            "VerticalScenarioEditor.Tests",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_tempDirectory);
    }

    [Fact]
    public void SaveAndLoad_RoundTrip_ShouldPreserveDocument()
    {
        var path = Path.Combine(_tempDirectory, "roundtrip.vse");
        var document = new DocumentState
        {
            PageNumberEnabled = false,
            ShowGuides = false,
            SummaryText = "概要",
            Records =
            {
                new ScriptRecord
                {
                    RoleName = "太郎",
                    Body = "台詞"
                }
            },
            RoleDictionary =
            {
                ["太郎"] = "#111111"
            }
        };

        DocumentFileService.Save(path, document);
        var loaded = DocumentFileService.Load(path);

        Assert.False(loaded.PageNumberEnabled);
        Assert.False(loaded.ShowGuides);
        Assert.Equal("概要", loaded.SummaryText);
        Assert.Single(loaded.Records);
        Assert.Equal("太郎", loaded.Records.Single().RoleName);
        Assert.Equal("台詞", loaded.Records.Single().Body);
        Assert.Equal("#111111", loaded.RoleDictionary["太郎"]);
    }

    [Fact]
    public void Load_WhenDocumentIsNull_ShouldReturnDefaultDocument()
    {
        var path = Path.Combine(_tempDirectory, "null-document.vse");
        var json = """
                   {
                     "version": 1,
                     "document": null
                   }
                   """;
        File.WriteAllText(path, json);

        var loaded = DocumentFileService.Load(path);

        Assert.NotNull(loaded);
        Assert.NotNull(loaded.Records);
        Assert.NotNull(loaded.RoleDictionary);
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDirectory))
            {
                Directory.Delete(_tempDirectory, recursive: true);
            }
        }
        catch
        {
            // テストの後始末失敗は本体検証に影響させない
        }
    }
}
