using System.Linq;
using VerticalScenarioEditor.Models;
using VerticalScenarioEditor.RoleManagement;
using Xunit;

namespace VerticalScenarioEditor.Tests.RoleManagement;

public sealed class RoleDictionarySynchronizerTests
{
    [Fact]
    public void ExtractRoleNames_ShouldSplitByDelimiterAndExcludeScene()
    {
        var roles = RoleDictionarySynchronizer.ExtractRoleNames("太郎／ シーン1 ／花子／ ").ToArray();

        Assert.Equal(["太郎", "花子"], roles);
    }

    [Fact]
    public void Synchronize_ShouldRemoveUnusedRolesAndPreserveColor()
    {
        var document = new DocumentState
        {
            Records =
            {
                new ScriptRecord { RoleName = "太郎", Body = "本文" }
            },
            RoleDictionary =
            {
                ["太郎"] = "#111111",
                ["花子"] = "#222222"
            }
        };

        RoleDictionarySynchronizer.Synchronize(document);

        Assert.True(document.RoleDictionary.ContainsKey("太郎"));
        Assert.Equal("#111111", document.RoleDictionary["太郎"]);
        Assert.False(document.RoleDictionary.ContainsKey("花子"));
    }

    [Fact]
    public void Synchronize_ShouldAddUsedRoleEvenWhenColorIsMissing()
    {
        var document = new DocumentState
        {
            Records =
            {
                new ScriptRecord { RoleName = "太郎／花子", Body = "本文" }
            },
            RoleDictionary =
            {
                ["太郎"] = "#111111"
            }
        };

        RoleDictionarySynchronizer.Synchronize(document);

        Assert.Equal("#111111", document.RoleDictionary["太郎"]);
        Assert.True(document.RoleDictionary.ContainsKey("花子"));
        Assert.Equal(string.Empty, document.RoleDictionary["花子"]);
    }
}
