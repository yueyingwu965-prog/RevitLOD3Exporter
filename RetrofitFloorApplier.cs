using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitLOD3Exporter
{
    public static class RetrofitFloorApplier
    {
        private const string PARAM_BUILDING_ID = "building_id";
        private const int BATCH_SIZE = 30;

        // ✅ Strategy params (written back already)
        private const string PARAM_EA_FLOOR_THERMAL_STRATEGY = "ea_floor_thermal_strategy";              // insulated|baseline
        private const string PARAM_WE_FLOOR_PERMEABILITY_STRATEGY = "we_floor_permeability_strategy";   // permeable|semi_permeable|impervious

        // ✅ Ground detection tolerance (internal units: feet)
        // 0.5m ≈ 1.64042 ft, 1.0m ≈ 3.28084 ft
        private const double GROUND_LEVEL_TOL_FT = 3.28084; // 1.0 m

        public static void Apply(Document doc, string retrofitMode)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            // 0) Cache levels and find "ground" reference (lowest level elevation)
            var levelsById = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToDictionary(lv => lv.Id, lv => lv, new ElementIdComparer());

            double groundElevFt = double.PositiveInfinity;
            foreach (var lv in levelsById.Values)
                groundElevFt = Math.Min(groundElevFt, lv.Elevation);

            // 1) Index FloorTypes by normalized name
            var floorTypesByKey = new FilteredElementCollector(doc)
                .OfClass(typeof(FloorType))
                .Cast<FloorType>()
                .GroupBy(ft => NormalizeKey(ft.Name), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            // 2) Collect floor instances
            var floors = new FilteredElementCollector(doc)
                .OfClass(typeof(Floor))
                .WhereElementIsNotElementType()
                .Cast<Floor>()
                .ToList();

            // 3) Build strict 1:1 mappings
            var mapInsulated = BuildInsulatedMappingByBuilding();
            var mapPermeable = BuildPermeableMappingByBuilding();

            int total = floors.Count;

            int planned = 0, changed = 0, alreadyTarget = 0;

            int skippedNoBuildingId = 0;
            int skippedBuildingNotMapped = 0;
            int skippedNoUpgradeRequested = 0;
            int skippedTypeNotInMapping = 0;
            int skippedTargetMissing = 0;
            int skippedFamilyMismatch = 0;
            int skippedNotChangeable = 0;

            // diagnostics
            int ignoredPermeabilityBecauseNotGroundExterior = 0;
            int ignoredPermeabilityBecauseNoMap = 0;

            var samples = new List<string>();
            var debugReads = new List<string>();

            var plan = new List<FloorChangeItem>();

            foreach (var fl in floors)
            {
                // building_id
                string buildingId = fl.LookupParameter(PARAM_BUILDING_ID)?.AsString();
                if (string.IsNullOrWhiteSpace(buildingId))
                {
                    skippedNoBuildingId++;
                    continue;
                }
                buildingId = SanitizeId(buildingId);

                if (!mapInsulated.ContainsKey(buildingId))
                {
                    skippedBuildingNotMapped++;
                    if (samples.Count < 6) samples.Add($"Floor {fl.Id}: building_id='{buildingId}' not mapped");
                    continue;
                }

                var curType = doc.GetElement(fl.GetTypeId()) as FloorType;
                if (curType == null) continue;

                string curTypeName = (curType.Name ?? "").Trim();
                string curTypeKey = NormalizeKey(curTypeName);
                string curFamily = curType.FamilyName ?? "";

                // Read strategies (instance-first then type)
                var typeElem = doc.GetElement(fl.GetTypeId());
                string thermalRaw = GetParamAsTextInstanceOrType(fl, typeElem, PARAM_EA_FLOOR_THERMAL_STRATEGY);
                string permRaw = GetParamAsTextInstanceOrType(fl, typeElem, PARAM_WE_FLOOR_PERMEABILITY_STRATEGY);

                string thermal = NormalizeThermal(thermalRaw);        // INSULATED / BASELINE
                string perm = NormalizePermeability(permRaw);         // PERMEABLE / OTHER (only PERMEABLE triggers switch)

                if (debugReads.Count < 8)
                {
                    debugReads.Add(
                        $"Floor {fl.Id} | Type='{curTypeName}' | ea_floor_thermal_strategy='{thermalRaw}' -> {thermal} | we_floor_permeability_strategy='{permRaw}' -> {perm}"
                    );
                }

                bool wantInsulated = (thermal == "INSULATED");
                bool wantPermeable = (perm == "PERMEABLE");

                // ✅ Critical rule (FIXED):
                // Permeability ONLY applies to GROUND-CONNECTED EXTERIOR floors (not interior floors, not roof decks).
                bool allowPermeableSwitch = IsGroundConnectedExteriorFloor(fl, curTypeName, levelsById, groundElevFt);

                if (wantPermeable && !allowPermeableSwitch)
                {
                    // Roof decks + interior + upper terraces: do NOT treat as permeable infiltration.
                    wantPermeable = false;
                    ignoredPermeabilityBecauseNotGroundExterior++;

                    // Optional sample log
                    if (samples.Count < 6)
                        samples.Add($"Floor {fl.Id}: permeability ignored (not ground-connected exterior). Type='{curTypeName}'");
                }

                // If neither upgrade requested => skip
                if (!wantInsulated && !wantPermeable)
                {
                    skippedNoUpgradeRequested++;
                    continue;
                }

                // Decide target type (strict 1:1)
                // Priority:
                // - If permeability switch allowed and requested => PERMEABLE
                // - Else if insulated requested => INSULATED
                string targetTypeKey = null;
                string targetReason = "";

                if (wantPermeable)
                {
                    if (!mapPermeable.TryGetValue(buildingId, out var permMap))
                    {
                        ignoredPermeabilityBecauseNoMap++;
                        wantPermeable = false; // fallback to thermal
                    }
                    else
                    {
                        // If current type already contains "Permeable" and strategy says permeable => keep as-is
                        if (ContainsIgnoreCase(curTypeName, "Permeable"))
                        {
                            if (!wantInsulated)
                            {
                                skippedNoUpgradeRequested++;
                                continue;
                            }
                            // If you later create "Permeable Insulated", extend here.
                        }

                        if (!permMap.TryGetValue(curTypeKey, out string targetPermName) || string.IsNullOrWhiteSpace(targetPermName))
                        {
                            // Second chance: Dachterrasse/Dachterasse typo normalization (kept for robustness)
                            string altKey = NormalizeKey(curTypeName.Replace("Dachterrasse", "Dachterasse"));
                            if (!permMap.TryGetValue(altKey, out targetPermName) || string.IsNullOrWhiteSpace(targetPermName))
                            {
                                skippedTypeNotInMapping++;
                                if (samples.Count < 6) samples.Add($"Floor {fl.Id}: permeable requested but current type '{curTypeName}' not in permeable mapping");
                                continue;
                            }
                        }

                        targetTypeKey = NormalizeKey(targetPermName);
                        targetReason = "PERMEABLE";
                    }
                }

                if (targetTypeKey == null && wantInsulated)
                {
                    var insMap = mapInsulated[buildingId];

                    if (!insMap.TryGetValue(curTypeKey, out string targetInsName) || string.IsNullOrWhiteSpace(targetInsName))
                    {
                        // Second chance for Dachterrasse/Dachterasse variant
                        string altKey = NormalizeKey(curTypeName.Replace("Dachterrasse", "Dachterasse"));
                        if (!insMap.TryGetValue(altKey, out targetInsName) || string.IsNullOrWhiteSpace(targetInsName))
                        {
                            skippedTypeNotInMapping++;
                            if (samples.Count < 6) samples.Add($"Floor {fl.Id}: insulated requested but current type '{curTypeName}' not in insulated mapping");
                            continue;
                        }
                    }

                    targetTypeKey = NormalizeKey(targetInsName);
                    targetReason = "INSULATED";
                }

                if (string.IsNullOrWhiteSpace(targetTypeKey))
                {
                    skippedNoUpgradeRequested++;
                    continue;
                }

                if (!floorTypesByKey.TryGetValue(targetTypeKey, out FloorType targetType))
                {
                    skippedTargetMissing++;
                    if (samples.Count < 6) samples.Add($"Floor {fl.Id}: target type '{targetTypeKey}' NOT found in model");
                    continue;
                }

                // ✅ Keep same Family (do NOT jump to Bodenplatte etc.)
                string targetFamily = targetType.FamilyName ?? "";
                if (!string.Equals(curFamily, targetFamily, StringComparison.OrdinalIgnoreCase))
                {
                    skippedFamilyMismatch++;
                    if (samples.Count < 6) samples.Add($"Floor {fl.Id}: family mismatch '{curFamily}' -> '{targetFamily}'");
                    continue;
                }

                if (fl.GetTypeId() == targetType.Id)
                {
                    alreadyTarget++;
                    continue;
                }

                // snapshot to keep offset / room-bounding stable
                var snap = CaptureInstancePositionSnapshot(fl);

                plan.Add(new FloorChangeItem
                {
                    FloorId = fl.Id,
                    TargetTypeId = targetType.Id,
                    FromTypeName = curTypeName,
                    ToTypeName = targetType.Name,
                    Reason = targetReason,
                    Snapshot = snap
                });

                planned++;
            }

            // Apply in batches
            int idx = 0;
            while (idx < plan.Count)
            {
                int take = Math.Min(BATCH_SIZE, plan.Count - idx);

                using (Transaction t = new Transaction(doc, "Retrofit - Floors (by strategies)"))
                {
                    t.Start();

                    for (int i = 0; i < take; i++)
                    {
                        var item = plan[idx + i];
                        var fl = doc.GetElement(item.FloorId) as Floor;
                        if (fl == null) continue;

                        try
                        {
                            fl.ChangeTypeId(item.TargetTypeId);
                            RestoreInstancePositionSnapshot(fl, item.Snapshot);

                            changed++;

                            if (samples.Count < 8)
                                samples.Add($"Floor {fl.Id}: [{item.Reason}] '{item.FromTypeName}' -> '{item.ToTypeName}'");
                        }
                        catch
                        {
                            skippedNotChangeable++;
                            if (samples.Count < 6) samples.Add($"Floor {fl.Id}: ChangeTypeId failed");
                        }
                    }

                    t.Commit();
                }

                idx += take;
            }

            string report =
                $"Total floors: {total}\n" +
                $"Changed: {changed}\n" +
                $"Ignored permeability (not ground-connected exterior): {ignoredPermeabilityBecauseNotGroundExterior}\n" +
                $"Skipped (building not mapped): {skippedBuildingNotMapped}\n" +
                $"Skipped (no upgrade requested): {skippedNoUpgradeRequested}\n" +
                $"Skipped (type not in mapping): {skippedTypeNotInMapping}\n" +
                $"Skipped (target type missing): {skippedTargetMissing}\n" +
                $"Skipped (family mismatch): {skippedFamilyMismatch}\n" +
                $"Skipped (not changeable): {skippedNotChangeable}\n\n" +
                $"Samples:\n- " + (samples.Count > 0 ? string.Join("\n- ", samples) : " ");

            TaskDialog.Show("Retrofit - Floors", report);
        }

        // =====================================================
        // ✅ NEW: Ground-connected exterior floor detection
        // =====================================================
        private static bool IsGroundConnectedExteriorFloor(
            Floor fl,
            string typeName,
            Dictionary<ElementId, Level> levelsById,
            double groundElevFt)
        {
            if (fl == null) return false;

            // 1) Exclude roof decks explicitly by name (these are NOT permeable-infiltration)
            //    Dachterrasse / Dachterasse = roof terrace => should be retention/detention, not infiltration
            if (ContainsIgnoreCase(typeName, "Dachterrasse")) return false;
            if (ContainsIgnoreCase(typeName, "Dachterasse")) return false;

            // 2) Basic exterior hint by type name
            //    (You said you want name-based detection; we keep it deterministic)
            bool looksExterior =
                ContainsIgnoreCase(typeName, "Terrasse") ||
                ContainsIgnoreCase(typeName, "Außen") ||
                ContainsIgnoreCase(typeName, "Aussen") ||
                ContainsIgnoreCase(typeName, "Outdoor");

            if (!looksExterior) return false;

            // 3) Must be on/near lowest level (ground-connected)
            Level lv = null;
            if (levelsById != null && levelsById.TryGetValue(fl.LevelId, out lv) && lv != null)
            {
                double elev = lv.Elevation;
                if (Math.Abs(elev - groundElevFt) <= GROUND_LEVEL_TOL_FT)
                    return true;

                return false;
            }

            // 4) Fallback: if level missing, be conservative (do NOT treat as permeable)
            return false;
        }

        private static bool ContainsIgnoreCase(string text, string needle)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(needle)) return false;
            return text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        // =====================================================
        // Strategy normalization
        // =====================================================
        private static string NormalizeThermal(string raw)
        {
            raw = (raw ?? "").Trim().ToLowerInvariant();
            if (raw.Contains("insulated")) return "INSULATED";
            if (raw == "1" || raw == "true" || raw == "yes") return "INSULATED";
            return "BASELINE";
        }

        private static string NormalizePermeability(string raw)
        {
            raw = (raw ?? "").Trim().ToLowerInvariant();

            // ONLY fully permeable triggers type switch
            if (raw.Contains("permeable") && !raw.Contains("semi") && !raw.Contains("imper"))
                return "PERMEABLE";

            if (raw == "1" || raw == "true" || raw == "yes") return "PERMEABLE";
            return "OTHER";
        }

        // =====================================================
        // ✅ Strict 1:1 mappings (your latest corrected names)
        // =====================================================
        private static Dictionary<string, Dictionary<string, string>> BuildInsulatedMappingByBuilding()
        {
            return new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "building_0001",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        // Existing -> Insulated
                        { NormalizeKey("STB 200"),                                "STB 200 Insulated" },
                        { NormalizeKey("STB 200 (WU Beton) + Abdichtung + Perimeterdämmung 150 (tot. 365)"),   "STB 200 (WU Beton) + Abdichtung + Perimeterdämmung" },

                        { NormalizeKey("FB 150 Vinyl"),                           "FB 150 Vinyl Insulated" },
                        { NormalizeKey("FB 150 Fliese Grau 300 x 300"),           "FB 150 Fliese Grau 300 x 300 Insulated" },
                        { NormalizeKey("FB 150 Terrasse Holz"),                   "FB 150 Terrasse Holz Insulated" },

                        { NormalizeKey("FB 240 Dachterasse Holz (Dominic)"),      "FB 240 Dachterasse Holz Insulated" },
                        { NormalizeKey("FB 240 Dachterrasse Holz (Dominic)"),     "FB 240 Dachterasse Holz Insulated" },

                        { NormalizeKey("FB 320 Dachterasse begrünt (Dominic)"),   "FB 320 Dachterasse begrünt Insulated" },
                    }
                },
                {
                    "building_0002",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        { NormalizeKey("FB 150 leer"),                            "FB 150 leer Insulated" },
                        { NormalizeKey("FB 150 Terrasse Holz"),                   "FB 150 Terrasse Holz Insulated" },
                        { NormalizeKey("FB 150 Terrasse Stein 500 x 500"),        "FB 150 Terrasse Stein 500 x 500 Insulated" },
                    }
                }
            };
        }

        private static Dictionary<string, Dictionary<string, string>> BuildPermeableMappingByBuilding()
        {
            return new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "building_0001",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        // NOTE: these mappings may still exist for legacy reasons,
                        // but will ONLY be applied if IsGroundConnectedExteriorFloor() returns true.
                        { NormalizeKey("FB 150 Terrasse Holz"),                   "FB 150 Terrasse Holz Permeable" },
                        { NormalizeKey("FB 240 Dachterasse Holz (Dominic)"),      "FB 240 Dachterasse Holz Permeable" },
                        { NormalizeKey("FB 240 Dachterrasse Holz (Dominic)"),     "FB 240 Dachterasse Holz Permeable" },
                        { NormalizeKey("FB 320 Dachterasse begrünt (Dominic)"),   "FB 320 Dachterasse begrünt Permeable" },
                    }
                },
                {
                    "building_0002",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        { NormalizeKey("FB 150 Terrasse Holz"),                   "FB 150 Terrasse Holz Permeable" },
                        { NormalizeKey("FB 150 Terrasse Stein 500 x 500"),        "FB 150 Terrasse Stein 500 x 500 Permeable" },
                    }
                }
            };
        }

        // =====================================================
        // Snapshot: keep placement stable (Offset + Room Bounding)
        // =====================================================
        private class FloorInstanceSnapshot
        {
            public double? LevelOffsetInternal { get; set; } = null; // internal units
            public int? RoomBoundingInt { get; set; } = null;        // 0/1
        }

        private class FloorChangeItem
        {
            public ElementId FloorId { get; set; }
            public ElementId TargetTypeId { get; set; }
            public string FromTypeName { get; set; }
            public string ToTypeName { get; set; }
            public string Reason { get; set; }
            public FloorInstanceSnapshot Snapshot { get; set; }
        }

        private static FloorInstanceSnapshot CaptureInstancePositionSnapshot(Floor fl)
        {
            var s = new FloorInstanceSnapshot();

            try
            {
                var pOff = fl.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                if (pOff != null && pOff.StorageType == StorageType.Double)
                    s.LevelOffsetInternal = pOff.AsDouble();
            }
            catch { }

            try
            {
                Parameter pRb =
                    fl.LookupParameter("Raumbegrenzung") ??
                    fl.LookupParameter("Room Bounding");

                if (pRb != null && pRb.StorageType == StorageType.Integer)
                    s.RoomBoundingInt = pRb.AsInteger();
            }
            catch { }

            return s;
        }

        private static void RestoreInstancePositionSnapshot(Floor fl, FloorInstanceSnapshot s)
        {
            if (fl == null || s == null) return;

            try
            {
                if (s.LevelOffsetInternal.HasValue)
                {
                    var pOff = fl.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                    if (pOff != null && !pOff.IsReadOnly && pOff.StorageType == StorageType.Double)
                        pOff.Set(s.LevelOffsetInternal.Value);
                }
            }
            catch { }

            try
            {
                if (s.RoomBoundingInt.HasValue)
                {
                    Parameter pRb =
                        fl.LookupParameter("Raumbegrenzung") ??
                        fl.LookupParameter("Room Bounding");

                    if (pRb != null && !pRb.IsReadOnly && pRb.StorageType == StorageType.Integer)
                        pRb.Set(s.RoomBoundingInt.Value);
                }
            }
            catch { }
        }

        // =====================================================
        // Param reading helpers (robust)
        // =====================================================
        private static string GetParamAsTextInstanceOrType(Element inst, Element type, string paramName)
        {
            string v = GetParamAsText(inst, paramName);
            if (string.IsNullOrWhiteSpace(v) && type != null)
                v = GetParamAsText(type, paramName);
            return v ?? "";
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

        // =====================================================
        // Utils
        // =====================================================
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

        private sealed class ElementIdComparer : IEqualityComparer<ElementId>
        {
            public bool Equals(ElementId x, ElementId y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (x is null || y is null) return false;
                return x.Value == y.Value;
            }

            public int GetHashCode(ElementId obj)
            {
                return obj?.Value.GetHashCode() ?? 0;
            }
        }
    }
}
