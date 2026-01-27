using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitLOD3Exporter
{
    public static class RetrofitRoofApplier
    {
        // =========================
        // Parameters (Shared Parameters, EXACT NAMES)
        // =========================
        private const string PARAM_BUILDING_ID = "building_id";

        private const string PARAM_EA_PV_STRATEGY = "ea_roof_pv_strategy";                 // none|limited|full
        private const string PARAM_EA_PV_SYSTEM_TYPE = "ea_roof_pv_system_type";           // attached|mounted|bipv (text)
        private const string PARAM_EA_STRUCT_FLAG = "ea_roof_structural_capacity_flag";    // bool/int
        private const string PARAM_EA_PV_PANEL_COUNT = "ea_pv_panel_count";                // int
        private const string PARAM_EA_PV_ADDED_AREA_M2 = "ea_pv_added_area_m2";            // area
        private const string PARAM_EA_PV_LAYOUT_STATUS = "ea_pv_layout_status";            // text

        private const string PARAM_WE_ROOF_RAIN_STRATEGY = "we_roof_rainwater_strategy";   // none|basic|enhanced
        private const string PARAM_WE_TANK_CAP_M3 = "we_rainwater_storage_capacity_m3";    // volume

        private const string PARAM_SS_ROOF_HEAT = "ss_roof_heat_strategy";                 // none|green_roof|cool_roof

        // =========================
        // Performance
        // =========================
        private const int BATCH_SIZE = 30; // reduce lag; 20-50 reasonable

        public static void Apply(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            // 1) Index all RoofTypes by Name (case-insensitive)
            var roofTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .Cast<RoofType>()
                .GroupBy(rt => rt.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            // 2) Collect all roof instances
            var roofs = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Roofs)
                .WhereElementIsNotElementType()
                .ToElements();

            int total = roofs.Count;

            // Counters
            int skippedNoBuildingId = 0;
            int skippedBuildingNotMapped = 0;
            int skippedNoTarget = 0;
            int skippedTargetMissing = 0;
            int skippedNotChangeable = 0;

            int alreadyTarget = 0;
            int planned = 0;
            int changed = 0;

            var samples = new List<string>();

            // 3) Build mapping: building_id -> names of RoofTypes in Revit
            var map = BuildRoofTypeMapping();

            // 4) Build plan outside of transaction (fast)
            var plan = new List<RooftopChange>();

            foreach (var e in roofs)
            {
                Element roof = e;
                if (roof == null) continue;

                string buildingId = GetStringParam(roof, PARAM_BUILDING_ID);
                if (string.IsNullOrWhiteSpace(buildingId))
                {
                    skippedNoBuildingId++;
                    continue;
                }
                buildingId = SanitizeId(buildingId);

                if (!map.TryGetValue(buildingId, out RoofTypeNames names))
                {
                    skippedBuildingNotMapped++;
                    if (samples.Count < 6) samples.Add($"Roof {roof.Id}: building_id='{buildingId}' not mapped");
                    continue;
                }

                // Determine target roof type name based on rules:
                // (1) base retrofit level / heat strategy -> EXISTING / RETROFIT_L1 / GREEN_L2 / COOL_L3
                // (2) PV overlay: if pv_strategy != none AND structural_flag true -> ROOF_PV
                string targetTypeName = DecideTargetRoofTypeName(roof, names);

                if (string.IsNullOrWhiteSpace(targetTypeName))
                {
                    skippedNoTarget++;
                    if (samples.Count < 6) samples.Add($"Roof {roof.Id}: no target decided");
                    continue;
                }

                if (!roofTypes.TryGetValue(targetTypeName, out RoofType targetType))
                {
                    skippedTargetMissing++;
                    if (samples.Count < 6) samples.Add($"Roof {roof.Id}: target type '{targetTypeName}' NOT found");
                    continue;
                }

                // Some roofs may be special types; ChangeTypeId generally works for Roofs, but we guard anyway.
                if (roof.GetTypeId() == ElementId.InvalidElementId)
                {
                    skippedNotChangeable++;
                    if (samples.Count < 6) samples.Add($"Roof {roof.Id}: invalid TypeId");
                    continue;
                }

                if (roof.GetTypeId() == targetType.Id)
                {
                    alreadyTarget++;
                    continue;
                }

                // Snapshot to keep position stable
                var snap = CaptureRoofSnapshot(roof);

                plan.Add(new RooftopChange
                {
                    RoofId = roof.Id,
                    TargetTypeId = targetType.Id,
                    TargetTypeName = targetTypeName,
                    FromTypeName = (doc.GetElement(roof.GetTypeId()) as ElementType)?.Name ?? "",
                    Snapshot = snap
                });

                planned++;
            }

            // 5) Apply changes in batches (less lag)
            int idx = 0;
            while (idx < plan.Count)
            {
                int take = Math.Min(BATCH_SIZE, plan.Count - idx);

                using (Transaction tx = new Transaction(doc, "EA Retrofit - Roofs"))
                {
                    tx.Start();

                    for (int i = 0; i < take; i++)
                    {
                        var item = plan[idx + i];
                        Element roof = doc.GetElement(item.RoofId);
                        if (roof == null) continue;

                        try
                        {
                            roof.ChangeTypeId(item.TargetTypeId);

                            // Restore critical instance placement params (offset + room bounding)
                            RestoreRoofSnapshot(roof, item.Snapshot);

                            changed++;

                            if (samples.Count < 6)
                                samples.Add($"Roof {roof.Id}: '{item.FromTypeName}' -> '{item.TargetTypeName}'");
                        }
                        catch
                        {
                            skippedNotChangeable++;
                            if (samples.Count < 6) samples.Add($"Roof {roof.Id}: ChangeTypeId failed");
                        }
                    }

                    tx.Commit();
                }

                idx += take;
            }

            string report =
                $"Total roofs: {total}\n" +
                $"Planned changes: {planned}\n" +
                $"Changed: {changed}\n" +
                $"Already target: {alreadyTarget}\n\n" +
                $"Skipped (no building_id): {skippedNoBuildingId}\n" +
                $"Skipped (building_id not mapped): {skippedBuildingNotMapped}\n" +
                $"Skipped (no target decided): {skippedNoTarget}\n" +
                $"Skipped (target type missing): {skippedTargetMissing}\n" +
                $"Skipped (not changeable): {skippedNotChangeable}\n\n" +
                $"Samples:\n- " + (samples.Count > 0 ? string.Join("\n- ", samples) : " ");

            TaskDialog.Show("Retrofit - Roofs", report);
        }

        // =====================================================
        // Decision rules
        // =====================================================
        private static string DecideTargetRoofTypeName(Element roof, RoofTypeNames names)
        {
            // ---- Base choice: by SS roof heat strategy + (optional) retrofit level
            // Your current roof family set:
            // ROOF_EXISTING, ROOF_RETROFIT_L1, ROOF_GREEN_L2, ROOF_COOL_L3, ROOF_PV

            string heat = NormalizeHeat(GetStringParam(roof, PARAM_SS_ROOF_HEAT)); // GREEN / COOL / NONE

            // If you want to use retrofit L1 as default when no SS strategy is set,
            // change base = names.ROOF_RETROFIT_L1 instead.
            string baseType =
                heat == "GREEN" ? names.ROOF_GREEN_L2 :
                heat == "COOL" ? names.ROOF_COOL_L3 :
                names.ROOF_EXISTING;

            // ---- PV overlay: ONLY switch to ROOF_PV if want PV AND structural flag is true
            string pvStrategy = NormalizePvStrategy(GetStringParam(roof, PARAM_EA_PV_STRATEGY)); // NONE / PV
            bool structOk = GetBoolLike(roof, PARAM_EA_STRUCT_FLAG);

            // Extra guard: if pv numbers are 0 and strategy empty, treat as NONE
            int panelCount = GetIntLike(roof, PARAM_EA_PV_PANEL_COUNT);
            double pvAreaM2 = GetDoubleLike(roof, PARAM_EA_PV_ADDED_AREA_M2);

            bool wantPv = (pvStrategy == "PV") || panelCount > 0 || pvAreaM2 > 0.0001;

            if (wantPv && structOk && !string.IsNullOrWhiteSpace(names.ROOF_PV))
                return names.ROOF_PV;

            return baseType;
        }

        private static string NormalizeHeat(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "NONE";
            raw = raw.Trim().ToLowerInvariant();

            if (raw.Contains("green")) return "GREEN"; // green_roof
            if (raw.Contains("cool")) return "COOL";  // cool_roof
            return "NONE";
        }

        private static string NormalizePvStrategy(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "NONE";
            raw = raw.Trim().ToLowerInvariant();

            if (raw == "none" || raw == "0") return "NONE";
            // limited / full / pv_ready etc.
            return "PV";
        }

        // =====================================================
        // Mapping: fill with YOUR Revit RoofType names
        // =====================================================
        private class RoofTypeNames
        {
            public string ROOF_EXISTING;
            public string ROOF_RETROFIT_L1;
            public string ROOF_GREEN_L2;
            public string ROOF_COOL_L3;
            public string ROOF_PV;
        }

        private static Dictionary<string, RoofTypeNames> BuildRoofTypeMapping()
        {
            // ✅ building_id are exactly: building_0001 / building_0002
            // ✅ Values must match RoofType.Name in Revit 100%
            return new Dictionary<string, RoofTypeNames>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "building_0001",
                    new RoofTypeNames
                    {
                        ROOF_EXISTING    = "Flachdachaufbau begrünt 320",
                        ROOF_RETROFIT_L1 = "Flachdachaufbau begrünt 320 Insulated",
                        ROOF_GREEN_L2    = "Flachdachaufbau begrünt 320 Green",
                        ROOF_COOL_L3     = "Flachdachaufbau 320 Cool",
                        ROOF_PV          = "Flachdachaufbau begrünt 320 PV"
                    }
                },
                {
                    "building_0002",
                    new RoofTypeNames
                    {
                        ROOF_EXISTING    = "KLH 200",
                        ROOF_RETROFIT_L1 = "KLH 200 Insulated",
                        ROOF_GREEN_L2    = "KLH 200 Green",
                        ROOF_COOL_L3     = "KLH 200 Cool",
                        ROOF_PV          = "KLH 200 PV"
                    }
                }
            };
        }

        // =====================================================
        // Snapshot: keep position stable (Offset + Room Bounding)
        // =====================================================
        private class RoofSnapshot
        {
            public ElementId LevelId = ElementId.InvalidElementId; // for checking only
            public double? OffsetInternal = null;                 // internal feet
            public int? RoomBoundingInt = null;                   // 0/1
        }

        private class RooftopChange
        {
            public ElementId RoofId;
            public ElementId TargetTypeId;
            public string TargetTypeName;
            public string FromTypeName;
            public RoofSnapshot Snapshot;
        }

        private static RoofSnapshot CaptureRoofSnapshot(Element roof)
        {
            var s = new RoofSnapshot();

            // Level (check only)
            try
            {
                var pLvl = roof.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                if (pLvl != null && pLvl.StorageType == StorageType.ElementId)
                    s.LevelId = pLvl.AsElementId();
            }
            catch { }

            // Offset (Roof level offset)
            // Note: for roofs this is commonly ROOF_LEVEL_OFFSET_PARAM
            try
            {
                var pOff = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
                if (pOff != null && pOff.StorageType == StorageType.Double)
                    s.OffsetInternal = pOff.AsDouble();
            }
            catch { }

            // Room Bounding - cross-language by parameter name
            try
            {
                var pRb =
                    roof.LookupParameter("Raumbegrenzung") ??
                    roof.LookupParameter("Room Bounding");
                if (pRb != null && pRb.StorageType == StorageType.Integer)
                    s.RoomBoundingInt = pRb.AsInteger();
            }
            catch { }

            return s;
        }

        private static void RestoreRoofSnapshot(Element roof, RoofSnapshot s)
        {
            if (roof == null || s == null) return;

            // Offset restore
            try
            {
                if (s.OffsetInternal.HasValue)
                {
                    var pOff = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
                    if (pOff != null && !pOff.IsReadOnly && pOff.StorageType == StorageType.Double)
                        pOff.Set(s.OffsetInternal.Value);
                }
            }
            catch { }

            // Room Bounding restore by name
            try
            {
                if (s.RoomBoundingInt.HasValue)
                {
                    var pRb =
                        roof.LookupParameter("Raumbegrenzung") ??
                        roof.LookupParameter("Room Bounding");
                    if (pRb != null && !pRb.IsReadOnly && pRb.StorageType == StorageType.Integer)
                        pRb.Set(s.RoomBoundingInt.Value);
                }
            }
            catch { }
        }

        // =====================================================
        // Helpers
        // =====================================================
        private static string SanitizeId(string s)
        {
            s = (s ?? "").Trim();
            s = s.Trim('\'').Trim('"');
            return s;
        }

        private static string GetStringParam(Element e, string name)
        {
            try
            {
                var p = e.LookupParameter(name);
                return p != null ? p.AsString() : null;
            }
            catch { return null; }
        }

        private static bool GetBoolLike(Element e, string name)
        {
            try
            {
                var p = e.LookupParameter(name);
                if (p == null) return false;

                if (p.StorageType == StorageType.Integer)
                    return p.AsInteger() != 0;

                if (p.StorageType == StorageType.String)
                {
                    string s = (p.AsString() ?? "").Trim();
                    if (bool.TryParse(s, out bool b)) return b;
                    if (s == "1") return true;
                    if (s == "0") return false;
                    if (s.Equals("yes", StringComparison.OrdinalIgnoreCase)) return true;
                    if (s.Equals("no", StringComparison.OrdinalIgnoreCase)) return false;
                }
            }
            catch { }
            return false;
        }

        private static int GetIntLike(Element e, string name)
        {
            try
            {
                var p = e.LookupParameter(name);
                if (p == null) return 0;

                if (p.StorageType == StorageType.Integer) return p.AsInteger();

                if (p.StorageType == StorageType.String)
                {
                    var s = (p.AsString() ?? "").Trim().Replace(",", ".");
                    if (int.TryParse(s, out int iv)) return iv;
                    if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dv))
                        return (int)Math.Round(dv);
                }

                if (p.StorageType == StorageType.Double)
                {
                    // might be unitless; treat as int
                    return (int)Math.Round(p.AsDouble());
                }
            }
            catch { }
            return 0;
        }

        private static double GetDoubleLike(Element e, string name)
        {
            try
            {
                var p = e.LookupParameter(name);
                if (p == null) return 0.0;

                if (p.StorageType == StorageType.Double) return p.AsDouble(); // NOTE: might be internal units
                if (p.StorageType == StorageType.Integer) return p.AsInteger();

                if (p.StorageType == StorageType.String)
                {
                    var s = (p.AsString() ?? "").Trim().Replace(",", ".");
                    if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double dv))
                        return dv;
                }
            }
            catch { }
            return 0.0;
        }
    }
}
