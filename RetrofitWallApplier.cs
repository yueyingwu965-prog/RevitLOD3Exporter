using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitLOD3Exporter
{
    public static class RetrofitWallApplier
    {
        private const string PARAM_BUILDING_ID = "building_id";
        private const int BATCH_SIZE = 50;

        // ✅ Strategy param written back already (baseline2 from your table, baseline3 from API)
        private const string PARAM_EA_WALL_STRATEGY = "ea_wall_envelope_strategy"; // baseline|insulated_upgrade|high_performance

        public static void Apply(Document doc, string retrofitMode)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            // ✅ Strict 1-to-1 mapping by building_id (thickness-safe)
            // This mapping is used ONLY when strategy requests upgrade (insulated_upgrade / high_performance)
            var mapByBuilding_InsulatedL2 = BuildInsulatedTypeMapping();

            // Index all WallTypes by normalized name (Basic only)
            var wallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .Where(wt => wt.Kind == WallKind.Basic)
                .GroupBy(wt => NormalizeKey(wt.Name), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            // Collect all wall instances
            var walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            int totalWalls = walls.Count;

            int skippedNotBasic = 0;
            int skippedNoBuildingId = 0;
            int skippedBuildingNotMapped = 0;

            int skippedNoUpgradeRequested = 0;      // ✅ NEW: strategy says baseline/none/empty
            int skippedHighPerformanceNoMap = 0;    // ✅ NEW: high_performance requested but no L3 map implemented

            int skippedTypeNotInMapping = 0;
            int skippedTargetTypeMissing = 0;
            int alreadyTarget = 0;

            // diagnostics
            int strategyReadFromInstance = 0;
            int strategyReadFromType = 0;

            var samples = new List<string>();
            var toChange = new List<(ElementId wallId, ElementId targetTypeId, string fromName, string toName, string strategy)>();

            foreach (var wall in walls)
            {
                var curType = wall.WallType;
                if (curType == null || curType.Kind != WallKind.Basic)
                {
                    skippedNotBasic++;
                    continue;
                }

                // building_id
                string buildingId = wall.LookupParameter(PARAM_BUILDING_ID)?.AsString();
                if (string.IsNullOrWhiteSpace(buildingId))
                {
                    skippedNoBuildingId++;
                    continue;
                }
                buildingId = SanitizeId(buildingId);

                // read strategy (instance first, then type) — THIS is the critical fix
                string strategyRaw = GetParamAsText(wall, PARAM_EA_WALL_STRATEGY);
                bool gotFromInstance = !string.IsNullOrWhiteSpace(strategyRaw);

                if (string.IsNullOrWhiteSpace(strategyRaw))
                    strategyRaw = GetParamAsText(curType, PARAM_EA_WALL_STRATEGY);

                if (gotFromInstance) strategyReadFromInstance++;
                else if (!string.IsNullOrWhiteSpace(strategyRaw)) strategyReadFromType++;

                string strategy = NormalizeWallStrategy(strategyRaw);

                // ✅ HARD RULE: baseline/none/empty => DO NOT UPGRADE (keep existing wall)
                if (strategy == "BASELINE" || strategy == "NONE")
                {
                    skippedNoUpgradeRequested++;
                    continue;
                }

                // building mapping must exist
                if (!mapByBuilding_InsulatedL2.TryGetValue(buildingId, out var typeMapL2))
                {
                    skippedBuildingNotMapped++;
                    continue;
                }

                string curTypeNameKey = NormalizeKey(curType.Name);

                // Decide which mapping to use
                // - insulated_upgrade => use L2 insulated mapping (existing -> insulated)
                // - high_performance  => currently no mapping provided in your code; skip safely
                Dictionary<string, string> activeMap = null;

                if (strategy == "INSULATED_UPGRADE")
                {
                    activeMap = typeMapL2;
                }
                else if (strategy == "HIGH_PERFORMANCE")
                {
                    // You haven't provided L3 wall type names/mapping in this file.
                    // To avoid accidental upgrades, we skip and report.
                    skippedHighPerformanceNoMap++;
                    if (samples.Count < 6)
                        samples.Add($"Wall {wall.Id}: strategy=high_performance but no L3 wall mapping implemented -> SKIPPED");
                    continue;
                }
                else
                {
                    // unknown strategy -> treat as baseline
                    skippedNoUpgradeRequested++;
                    continue;
                }

                // Map current type -> target type name
                if (!activeMap.TryGetValue(curTypeNameKey, out string targetTypeName) || string.IsNullOrWhiteSpace(targetTypeName))
                {
                    skippedTypeNotInMapping++;
                    if (samples.Count < 6)
                        samples.Add($"Wall {wall.Id}: strategy={strategyRaw} but current type '{curType.Name}' not in mapping");
                    continue;
                }

                string targetTypeKey = NormalizeKey(targetTypeName);

                if (!wallTypes.TryGetValue(targetTypeKey, out WallType targetType))
                {
                    skippedTargetTypeMissing++;
                    if (samples.Count < 6) samples.Add($"Wall {wall.Id}: target type '{targetTypeKey}' NOT found");
                    continue;
                }

                if (wall.GetTypeId() == targetType.Id)
                {
                    alreadyTarget++;
                    continue;
                }

                toChange.Add((wall.Id, targetType.Id, curType.Name, targetType.Name, strategyRaw));
            }

            int changed = 0;
            int idx = 0;

            while (idx < toChange.Count)
            {
                int take = Math.Min(BATCH_SIZE, toChange.Count - idx);

                using (Transaction t = new Transaction(doc, "Retrofit - Walls (by strategy param)"))
                {
                    t.Start();

                    for (int i = 0; i < take; i++)
                    {
                        var item = toChange[idx + i];
                        var wall = doc.GetElement(item.wallId) as Wall;
                        if (wall == null) continue;

                        if (wall.GetTypeId() == item.targetTypeId)
                        {
                            alreadyTarget++;
                            continue;
                        }

                        wall.ChangeTypeId(item.targetTypeId);
                        changed++;

                        if (samples.Count < 6)
                            samples.Add($"Wall {wall.Id}: [{item.strategy}] '{item.fromName}' -> '{item.toName}'");
                    }

                    t.Commit();
                }

                idx += take;
            }

            string report =
                $"Mode: {retrofitMode}\n" +
                $"Total walls: {totalWalls}\n" +
                $"Changed: {changed}\n" +
                $"Skipped (not Basic): {skippedNotBasic}\n" +
                $"Skipped (no building_id): {skippedNoBuildingId}\n" +
                $"Skipped (building_id not mapped): {skippedBuildingNotMapped}\n" +
                $"Skipped (strategy baseline/none/empty): {skippedNoUpgradeRequested}\n" +
                $"Skipped (high_performance but no L3 mapping): {skippedHighPerformanceNoMap}\n" +
                $"Skipped (type not in mapping): {skippedTypeNotInMapping}\n" +
                $"Skipped (target type missing): {skippedTargetTypeMissing}\n" +
                $"Already target: {alreadyTarget}\n" +
                $"Strategy read from instance: {strategyReadFromInstance}\n" +
                $"Strategy read from type: {strategyReadFromType}\n\n" +
                $"Samples:\n- " + (samples.Count > 0 ? string.Join("\n- ", samples) : " ");

            TaskDialog.Show("Retrofit - Walls", report);
        }

        /// <summary>
        /// ✅ EXACT 1-to-1 mapping: Existing -> Insulated (L2)
        /// Keys/Values must match Revit WallType.Name (we normalize spaces)
        /// Used ONLY when strategy == insulated_upgrade
        /// </summary>
        private static Dictionary<string, Dictionary<string, string>> BuildInsulatedTypeMapping()
        {
            return new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "building_0001",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        { NormalizeKey("Ziegel+WD hart 240+160"),                 NormalizeKey("Ziegel+WD hart 240+160 Insulated") },
                        { NormalizeKey("Ziegel+WD hart 240+160 mit Fliesen"),     NormalizeKey("Ziegel+WD hart 240+160 mit Fliesen Insulated") },
                        { NormalizeKey("Ziegel 200 beidseitig verputzt"),         NormalizeKey("Ziegel 200 beidseitig verputzt Insulated") },
                        { NormalizeKey("Ziegel 200 gefliest und verputzt"),       NormalizeKey("Ziegel 200 gefliest und verputzt Insulated") },
                        { NormalizeKey("Ziegel 200 gefliest und verputzt 2"),     NormalizeKey("Ziegel 200 gefliest und verputzt Insulated") },
                    }
                },
                {
                    "building_0002",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        { NormalizeKey("Ziegel 200"),                             NormalizeKey("Ziegel 200 Insulated") },
                        { NormalizeKey("Ziegel 240"),                             NormalizeKey("Ziegel 240 Insulated") },
                        { NormalizeKey("Ziegel+WD hart 200+40"),                  NormalizeKey("Ziegel+WD hart 200+40 Insulated") },
                        { NormalizeKey("Ziegel+WD hart 310+40"),                  NormalizeKey("Ziegel+WD hart 310+40 Insulated") },
                    }
                }
            };
        }

        // =====================================================
        // Helpers
        // =====================================================

        private static string NormalizeWallStrategy(string raw)
        {
            raw = (raw ?? "").Trim().ToLowerInvariant();

            // empty => treat as baseline (safe)
            if (string.IsNullOrWhiteSpace(raw)) return "BASELINE";

            // normalize baseline/none
            if (raw == "baseline" || raw == "0" || raw == "false" || raw == "no") return "BASELINE";
            if (raw == "none") return "NONE";

            // normalize insulated
            if (raw.Contains("insulated")) return "INSULATED_UPGRADE";

            // normalize high performance
            if (raw.Contains("high")) return "HIGH_PERFORMANCE";
            if (raw.Contains("performance")) return "HIGH_PERFORMANCE";

            // unknown => baseline (safe)
            return "BASELINE";
        }

        private static string GetParamAsText(Element e, string paramName)
        {
            try
            {
                var p = e?.LookupParameter(paramName);
                if (p == null) return "";

                switch (p.StorageType)
                {
                    case StorageType.String:
                        return p.AsString() ?? "";
                    case StorageType.Integer:
                        return p.AsInteger().ToString();
                    case StorageType.Double:
                        return p.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        var id = p.AsElementId();
                        return (id != null && id != ElementId.InvalidElementId) ? id.Value.ToString() : "";
                    default:
                        return p.AsValueString() ?? "";
                }
            }
            catch { return ""; }
        }

        private static string SanitizeId(string s)
        {
            return (s ?? "").Trim().Trim('\'').Trim('"');
        }

        private static string NormalizeKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim();
            s = s.Replace('\u00A0', ' ').Replace('\u2007', ' ').Replace('\u202F', ' ');
            while (s.IndexOf("  ", StringComparison.Ordinal) >= 0)
                s = s.Replace("  ", " ");
            return s;
        }
    }
}
