using System;
using System.Collections.Generic;

namespace VerticalScenarioEditor.Models;

public static class DocumentStateCloner
{
    public static DocumentState Clone(DocumentState source)
    {
        if (source == null)
        {
            return DocumentState.CreateDefault();
        }

        var clone = new DocumentState
        {
            PageNumberEnabled = source.PageNumberEnabled,
            ShowGuides = source.ShowGuides,
            SummaryText = source.SummaryText ?? string.Empty,
            Records = new List<ScriptRecord>(),
            RoleDictionary = new Dictionary<string, string>()
        };

        if (source.Records != null)
        {
            foreach (var record in source.Records)
            {
                clone.Records.Add(new ScriptRecord
                {
                    RoleName = record?.RoleName ?? string.Empty,
                    Body = record?.Body ?? string.Empty
                });
            }
        }

        if (source.RoleDictionary != null)
        {
            foreach (var pair in source.RoleDictionary)
            {
                if (string.IsNullOrWhiteSpace(pair.Key))
                {
                    continue;
                }

                clone.RoleDictionary[pair.Key] = pair.Value ?? string.Empty;
            }
        }

        return clone;
    }
}
