using System;
using System.IO;
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
                },
                new ScriptRecord
                {
                    RoleName = "花子",
                    Body = "返答",
                    PageBreakBefore = true
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
        Assert.Equal(2, loaded.Records.Count);
        Assert.Equal("太郎", loaded.Records[0].RoleName);
        Assert.Equal("台詞", loaded.Records[0].Body);
        Assert.False(loaded.Records[0].PageBreakBefore);
        Assert.Equal("花子", loaded.Records[1].RoleName);
        Assert.Equal("返答", loaded.Records[1].Body);
        Assert.True(loaded.Records[1].PageBreakBefore);
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

    [Fact]
    public void Load_WhenStructureInvalid_ShouldThrowJapaneseError()
    {
        var path = Path.Combine(_tempDirectory, "invalid-structure.vse");
        File.WriteAllText(path, "{");

        var exception = Assert.Throws<InvalidDataException>(() => DocumentFileService.Load(path));

        Assert.Contains("ファイルの内容が不正です。", exception.Message);
    }

    [Fact]
    public void Load_WhenVersionUnsupported_ShouldThrowJapaneseError()
    {
        var path = Path.Combine(_tempDirectory, "unsupported-version.vse");
        var json = """
                   {
                     "version": 999,
                     "document": {}
                   }
                   """;
        File.WriteAllText(path, json);

        var exception = Assert.Throws<InvalidDataException>(() => DocumentFileService.Load(path));

        Assert.Contains("このファイルのバージョンには対応していません。", exception.Message);
    }

    [Fact]
    public void Load_WhenDocumentHasNullMembers_ShouldNormalizeValues()
    {
        var path = Path.Combine(_tempDirectory, "normalize.vse");
        var json = """
                   {
                     "version": 1,
                     "document": {
                        "summaryText": null,
                        "records": [{ "roleName": null, "body": null, "pageBreakBefore": true }, { "roleName": null, "body": null, "pageBreakBefore": true }],
                        "roleDictionary": {
                          "太郎": null,
                          "": "#ff0000"
                        }
                     }
                   }
                   """;
        File.WriteAllText(path, json);

        var loaded = DocumentFileService.Load(path);

        Assert.Equal(string.Empty, loaded.SummaryText);
        Assert.Equal(2, loaded.Records.Count);
        Assert.All(loaded.Records, record => Assert.NotNull(record));
        Assert.Equal(string.Empty, loaded.Records[0].RoleName);
        Assert.Equal(string.Empty, loaded.Records[0].Body);
        Assert.False(loaded.Records[0].PageBreakBefore);
        Assert.Equal(string.Empty, loaded.Records[1].RoleName);
        Assert.Equal(string.Empty, loaded.Records[1].Body);
        Assert.True(loaded.Records[1].PageBreakBefore);
        Assert.Equal(string.Empty, loaded.RoleDictionary["太郎"]);
        Assert.DoesNotContain(string.Empty, loaded.RoleDictionary.Keys);
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
