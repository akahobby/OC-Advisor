using System;
using System.Collections.Generic;

namespace OcAdvisor;

internal static class RecommendationDedupe
{
    public static List<Recommendation> Dedupe(IEnumerable<Recommendation> recs)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var outList = new List<Recommendation>();

        foreach (var r in recs)
        {
            if (r == null) continue;

            // Key is intentionally strict to avoid deleting legitimately different rows.
            var key = $"{r.Section}||{r.Setting}||{r.Value}||{r.Path}||{r.Reason}";
            if (seen.Add(key))
                outList.Add(r);
        }

        return outList;
    }
}
