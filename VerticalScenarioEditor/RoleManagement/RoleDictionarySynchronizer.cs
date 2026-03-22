using System;
using System.Collections.Generic;
using System.Linq;
using VerticalScenarioEditor.Models;

namespace VerticalScenarioEditor.RoleManagement;

public static class RoleDictionarySynchronizer
{
    public static IEnumerable<string> ExtractRoleNames(string? roleName)
    {
        if (string.IsNullOrWhiteSpace(roleName))
        {
            yield break;
        }

        foreach (var part in roleName.Split('／', StringSplitOptions.RemoveEmptyEntries))
        {
            var trimmed = part.Trim();
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                continue;
            }

            if (trimmed.StartsWith("シーン", StringComparison.Ordinal))
            {
                continue;
            }

            yield return trimmed;
        }
    }

    public static void Synchronize(DocumentState document)
    {
        if (document == null)
        {
            return;
        }

        document.Records ??= new List<ScriptRecord>();
        document.RoleDictionary ??= new Dictionary<string, string>();

        var usedRoles = document.Records
            .SelectMany(record => ExtractRoleNames(record?.RoleName))
            .ToHashSet(StringComparer.Ordinal);

        var existingColors = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var entry in document.RoleDictionary)
        {
            foreach (var role in ExtractRoleNames(entry.Key))
            {
                if (!existingColors.ContainsKey(role))
                {
                    existingColors[role] = entry.Value ?? string.Empty;
                }
            }
        }

        var nextDictionary = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var role in usedRoles.OrderBy(name => name, StringComparer.Ordinal))
        {
            nextDictionary[role] = existingColors.TryGetValue(role, out var color)
                ? color
                : string.Empty;
        }

        document.RoleDictionary = nextDictionary;
    }
}
