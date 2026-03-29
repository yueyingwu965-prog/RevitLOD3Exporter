using System;
using System.Collections.Generic;
using System.Globalization;
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
        private const string PARAM_EA_PV_ADDED_AREA_M2 = "ea_pv_added_area_m2";            // area (m2 in your logic)
        private const string PARAM_EA_PV_LAYOUT_STATUS = "ea_pv_layout_status";            // text

        private const string PARAM_WE_ROOF_RAIN_STRATEGY = "we_roof_rainwater_strategy";   // none|basic|enhanced
        private const string PARAM_WE_TANK_CAP_M3 = "we_rainwater_storage_capacity_m3";    // volume

        private const string PARAM_SS_ROOF_HEAT = "ss_roof_heat_strategy";                 // none|green_roof|cool_roof

        // =========================
        // Performance
        // =========================
        private const int BATCH_SIZE = 30;

        // =========================
        // Base choice default
        // =========================
        // ✅ You said you previously "统一写给所有屋顶" and want strategy-driven.
        // If ss_roof_heat_strategy is NONE/empty:
        // - set to true => default to ROOF_RETROFIT_L1
        // - set to false => default to ROOF_EXISTING
        private const bool DEFAULT_TO_RETROFIT_L1_WHEN_NO_HEAT_STRATEGY = true;

        public static void Apply(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            // 1) Index RoofTypes by normalized name (robust)
            var roofTypesByKey = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .Cast<RoofType>()
                .GroupBy(rt => NormalizeKey(rt.Name), StringComparer.OrdinalIgnoreCase)
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

            // diagnostics
            int pvWantedCount = 0;
            int pvAppliedCount = 0;
            int pvBlockedByStructFlag = 0;

            var samples = new List<string>();

            // 3) Mapping: building_id -> names of RoofTypes in Revit
            var map = BuildRoofTypeMapping();

            // 4) Build plan outside transaction
            var plan = new List<RooftopChange>();

            foreach (var e in roofs)
            {
                Element roof = e;
                if (roof == null) continue;

                string buildingId = GetParamAsText(roof, PARAM_BUILDING_ID);
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

                // Decide target roof type name based on instance strategy params
                var decision = DecideTargetRoofTypeName(doc, roof, names);

                if (decision.WantPv) pvWantedCount++;
                if (decision.PvBlockedByStructFlag) pvBlockedByStructFlag++;
                if (decision.TargetIsPv) pvAppliedCount++;

                string targetTypeName = decision.TargetTypeName;
                if (string.IsNullOrWhiteSpace(targetTypeName))
                {
                    skippedNoTarget++;
                    if (samples.Count < 6) samples.Add($"Roof {roof.Id}: no target decided");
                    continue;
                }

                string targetKey = NormalizeKey(targetTypeName);
                if (!roofTypesByKey.TryGetValue(targetKey, out RoofType targetType))
                {
                    skippedTargetMissing++;
                    if (samples.Count < 6) samples.Add($"Roof {roof.Id}: target type '{targetTypeName}' NOT found");
                    continue;
                }

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

            // 5) Apply in batches
            int idx = 0;
            while (idx < plan.Count)
            {
                int take = Math.Min(BATCH_SIZE, plan.Count - idx);

                using (Transaction tx = new Transaction(doc, "EA Retrofit - Roofs (by strategies)"))
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
                $"Planned: {planned}\n" +
                $"Changed: {changed}\n" +
                $"Already target: {alreadyTarget}\n" +
                $"PV wanted: {pvWantedCount}\n" +
                $"PV applied (roof type switched to PV): {pvAppliedCount}\n" +
                $"PV blocked by structural flag: {pvBlockedByStructFlag}\n" +
                $"Skipped (no building_id): {skippedNoBuildingId}\n" +
                $"Skipped (building_id not mapped): {skippedBuildingNotMapped}\n" +
                $"Skipped (no target decided): {skippedNoTarget}\n" +
                $"Skipped (target type missing): {skippedTargetMissing}\n" +
                $"Skipped (not changeable): {skippedNotChangeable}\n\n" +
                $"Samples:\n- " + (samples.Count > 0 ? string.Join("\n- ", samples) : " ");

            TaskDialog.Show("Retrofit - Roofs", report);
        }

        // =====================================================
        // Decision rules (instance-first; robust param reading)
        // =====================================================

        private struct RoofDecision
        {
            public string TargetTypeName;
            public bool WantPv;
            public bool TargetIsPv;
            public bool PvBlockedByStructFlag;
        }

        private static RoofDecision DecideTargetRoofTypeName(Document doc, Element roof, RoofTypeNames names)
        {
            // ---- Base choice by SS heat strategy
            string heatRaw = GetParamAsText(roof, PARAM_SS_ROOF_HEAT);
            string heat = NormalizeHeat(heatRaw); // GREEN / COOL / NONE

            string baseType =
                heat == "GREEN" ? names.ROOF_GREEN_L2 :
                heat == "COOL" ? names.ROOF_COOL_L3 :
                (DEFAULT_TO_RETROFIT_L1_WHEN_NO_HEAT_STRATEGY ? names.ROOF_RETROFIT_L1 : names.ROOF_EXISTING);

            // ---- PV overlay (only if want PV AND structural flag true)
            string pvRaw = GetParamAsText(roof, PARAM_EA_PV_STRATEGY);
            string pvStrategy = NormalizePvStrategy(pvRaw); // NONE / PV

            bool structOk = GetBoolLike(roof, PARAM_EA_STRUCT_FLAG);

            // panelCount often stored as int (or string)
            int panelCount = GetIntLike(roof, PARAM_EA_PV_PANEL_COUNT);

            // pv added area stored as Area (internal units ft^2). We convert to m^2 if it's a Double area.
            double pvAreaM2 = GetAreaM2Like(doc, roof, PARAM_EA_PV_ADDED_AREA_M2);

            bool wantPv = (pvStrategy == "PV") || panelCount > 0 || pvAreaM2 > 1e-6;

            // If want PV but struct flag false => keep base
            if (wantPv && !structOk)
            {
                return new RoofDecision
                {
                    TargetTypeName = baseType,
                    WantPv = true,
                    TargetIsPv = false,
                    PvBlockedByStructFlag = true
                };
            }

            // If want PV and struct ok => PV type
            if (wantPv && structOk && !string.IsNullOrWhiteSpace(names.ROOF_PV))
            {
                return new RoofDecision
                {
                    TargetTypeName = names.ROOF_PV,
                    WantPv = true,
                    TargetIsPv = true,
                    PvBlockedByStructFlag = false
                };
            }

            // default base
            return new RoofDecision
            {
                TargetTypeName = baseType,
                WantPv = wantPv,
                TargetIsPv = false,
                PvBlockedByStructFlag = false
            };
        }

        private static string NormalizeHeat(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "NONE";
            raw = raw.Trim().ToLowerInvariant();

            if (raw.Contains("green")) return "GREEN"; // green_roof
            if (raw.Contains("cool")) return "COOL";   // cool_roof
            if (raw == "none" || raw == "0" || raw == "false" || raw == "no") return "NONE";

            return "NONE";
        }

        private static string NormalizePvStrategy(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "NONE";
            raw = raw.Trim().ToLowerInvariant();

            if (raw == "none" || raw == "0" || raw == "false" || raw == "no") return "NONE";
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
                        ROOF_EXISTING    = "WD 200",
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

            try
            {
                var pLvl = roof.get_Parameter(BuiltInParameter.LEVEL_PARAM);
                if (pLvl != null && pLvl.StorageType == StorageType.ElementId)
                    s.LevelId = pLvl.AsElementId();
            }
            catch { }

            try
            {
                var pOff = roof.get_Parameter(BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM);
                if (pOff != null && pOff.StorageType == StorageType.Double)
                    s.OffsetInternal = pOff.AsDouble();
            }
            catch { }

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
        // Robust param reading helpers (StorageType-safe)
        // =====================================================

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
                        return p.AsInteger().ToString(CultureInfo.InvariantCulture);
                    case StorageType.Double:
                        return p.AsDouble().ToString(CultureInfo.InvariantCulture);
                    case StorageType.ElementId:
                        var id = p.AsElementId();
                        return (id != null && id != ElementId.InvalidElementId) ? id.Value.ToString(CultureInfo.InvariantCulture) : "";
                    default:
                        return p.AsValueString() ?? "";
                }
            }
            catch { return ""; }
        }

        private static bool GetBoolLike(Element e, string name)
        {
            try
            {
                var p = e?.LookupParameter(name);
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
                    if (s.Equals("true", StringComparison.OrdinalIgnoreCase)) return true;
                    if (s.Equals("false", StringComparison.OrdinalIgnoreCase)) return false;
                }
            }
            catch { }
            return false;
        }

        private static int GetIntLike(Element e, string name)
        {
            try
            {
                var p = e?.LookupParameter(name);
                if (p == null) return 0;

                if (p.StorageType == StorageType.Integer) return p.AsInteger();

                if (p.StorageType == StorageType.Double)
                    return (int)Math.Round(p.AsDouble());

                if (p.StorageType == StorageType.String)
                {
                    var s = (p.AsString() ?? "").Trim().Replace(",", ".");
                    if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int iv)) return iv;
                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double dv))
                        return (int)Math.Round(dv);
                }
            }
            catch { }
            return 0;
        }

        private static double GetAreaM2Like(Document doc, Element e, string name)
        {
            try
            {
                var p = e?.LookupParameter(name);
                if (p == null) return 0.0;

                if (p.StorageType == StorageType.Double)
                {
                    double internalVal = p.AsDouble();

                    // Try convert if units API available (Revit 2021+)
                    try
                    {
                        return UnitUtils.ConvertFromInternalUnits(internalVal, UnitTypeId.SquareMeters);
                    }
                    catch
                    {
                        // If conversion fails, at least return the raw (still OK for >0 checks)
                        return internalVal;
                    }
                }

                if (p.StorageType == StorageType.Integer) return p.AsInteger();

                if (p.StorageType == StorageType.String)
                {
                    var s = (p.AsString() ?? "").Trim().Replace(",", ".");
                    if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double dv))
                        return dv;
                }
            }
            catch { }
            return 0.0;
        }

        // =====================================================
        // Utils
        // =====================================================

        private static string SanitizeId(string s)
        {
            s = (s ?? "").Trim();
            s = s.Trim('\'').Trim('"');
            return s;
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