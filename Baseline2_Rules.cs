using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace RevitLOD3Exporter
{
    public class Baseline2_Rules
    {
        // Roof priority: PV > cool/green
        private const bool ROOF_PV_PRIORITY = true;

        // Keep disabled unless you really want type swaps
        private static readonly bool ENABLE_TYPE_CHANGE = false;

        // PV constants (report derivation only; deterministic PV calc is in ImportRetrofitCSV)
        private const double DEFAULT_PV_COVERAGE = 0.60;
        private const double DEFAULT_PANEL_W_M = 1.494;
        private const double DEFAULT_PANEL_L_M = 1.219;

        public class ResultInfo
        {
            public Dictionary<string, string> Actuation = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // per-element computed results (for debugging + report)
            public Dictionary<ElementId, string> WallStrategyById = new Dictionary<ElementId, string>();
            public Dictionary<ElementId, double> WallUById = new Dictionary<ElementId, double>();

            public Dictionary<ElementId, string> FloorThermalById = new Dictionary<ElementId, string>();
            public Dictionary<ElementId, double> FloorUById = new Dictionary<ElementId, double>();
            public Dictionary<ElementId, string> FloorPermeabilityById = new Dictionary<ElementId, string>();
            public Dictionary<ElementId, bool> FloorOutdoorById = new Dictionary<ElementId, bool>();

            public Dictionary<ElementId, string> RoofPvStrategyById = new Dictionary<ElementId, string>();
            public Dictionary<ElementId, double> RoofPvUsableById = new Dictionary<ElementId, double>();
            public Dictionary<ElementId, string> RoofHeatById = new Dictionary<ElementId, string>();
            public Dictionary<ElementId, string> RoofRainwaterById = new Dictionary<ElementId, string>();

            public int WriteCount = 0;
            public int TypeChangeCount = 0;

            public List<string> Notes = new List<string>();
        }

        private class RevitTypeNameMap
        {
            public Dictionary<string, string> Roof = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> Wall = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> Floor = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        // ============================================================
        // Mapping (only used when ENABLE_TYPE_CHANGE=true)
        // ============================================================
        private RevitTypeNameMap GetMapForBuilding(string buildingId)
        {
            string bid = (buildingId ?? "").Trim().ToLowerInvariant();

            // ---------- Building1 ----------
            if (bid.Contains("0001") || bid.Contains("building_0001"))
            {
                var m = new RevitTypeNameMap();

                m.Roof["ROOF_EXISTING"] = "Flachdachaufbau begrünt 320";
                m.Roof["ROOF_RETROFIT_L1"] = "Flachdachaufbau begrünt 320 Insulated";
                m.Roof["ROOF_GREEN_L2"] = "Flachdachaufbau 320 Green";
                m.Roof["ROOF_COOL_L3"] = "Flachdachaufbau 320 Cool";
                m.Roof["ROOF_PV"] = "Flachdachaufbau begrünt 320 PV";

                m.Wall["WALL_RETROFIT_L1_SUFFIX"] = " Insulated";

                m.Floor["FLOOR_INSULATED_SUFFIX"] = " Insulated";
                m.Floor["FLOOR_PERMEABLE"] = "FB 240 Dachterasse Holz Permeable";

                return m;
            }

            // ---------- Building2 ----------
            if (bid.Contains("0002") || bid.Contains("building_0002"))
            {
                var m = new RevitTypeNameMap();

                m.Roof["ROOF_EXISTING"] = "KLH 200";
                m.Roof["ROOF_RETROFIT_L1"] = "KLH 200 Insulated";
                m.Roof["ROOF_GREEN_L2"] = "KLH 200 Green";
                m.Roof["ROOF_COOL_L3"] = "KLH 200 Cool";
                m.Roof["ROOF_PV"] = "KLH 200 PV";

                m.Wall["WALL_RETROFIT_L1_SUFFIX"] = " Insulated";

                m.Floor["FLOOR_INSULATED_SUFFIX"] = " Insulated";
                m.Floor["FLOOR_PERMEABLE"] = "FB 150 Terrasse Holz Permeable";

                return m;
            }

            return new RevitTypeNameMap();
        }

        // ============================================================
        // Backward-compatible Apply (OLD signature: no doors)
        // ============================================================
        public ResultInfo Apply(
            Document doc,
            Dictionary<string, string> row,
            IList<Element> roofs,
            IList<Element> windows,
            IList<Element> walls,
            IList<Element> floors)
        {
            return Apply(doc, row, roofs, windows, new List<Element>(), walls, floors);
        }

        // ============================================================
        // Apply: includes doors
        // ============================================================
        public ResultInfo Apply(
            Document doc,
            Dictionary<string, string> row,
            IList<Element> roofs,
            IList<Element> windows,
            IList<Element> doors,
            IList<Element> walls,
            IList<Element> floors)
        {
            var res = new ResultInfo();

            // 1) Compute GLOBAL actuation (things that are truly building/site-level)
            res.Actuation = ComputeGlobalActuation(row);

            // 2) Write global + per-element actuation
            res.WriteCount += WriteActuation(doc, roofs, windows, doors, walls, floors, row, res);

            // 3) Optional: type changes
            if (ENABLE_TYPE_CHANGE)
            {
                string buildingId = GetOrEmpty(row, "building_id");
                var typeMap = GetMapForBuilding(buildingId);
                res.TypeChangeCount += ApplyTriggers(doc, roofs, walls, floors, row, res, typeMap);
            }
            else
            {
                res.Notes.Add("Type changes disabled: Baseline2 writes parameters, computes strategies, and exports report only.");
            }

            if (doors == null || doors.Count == 0)
                res.Notes.Add("Doors list is empty: ea_door_upgrade_strategy not written (pass doors collector to Apply to enable).");

            return res;
        }

        // ============================================================
        // Report
        // ============================================================
        public string BuildReport(Dictionary<string, string> row, ResultInfo info)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Baseline2 Retrofit Report (Rules)");
            sb.AppendLine("========================================");
            sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine();

            var pvPairs = new List<KeyValuePair<string, string>>();
            var nonPvPairs = new List<KeyValuePair<string, string>>();

            foreach (var kv in row ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase))
            {
                if (IsPvKey(kv.Key)) pvPairs.Add(kv);
                else nonPvPairs.Add(kv);
            }

            sb.AppendLine("Input Metrics (CSV selected row)");
            sb.AppendLine("----------------------------------------");
            foreach (var kv in nonPvPairs) sb.AppendLine($"{kv.Key}: {kv.Value}");
            sb.AppendLine();

            sb.AppendLine("Baseline2 Actuation (GLOBAL, computed by rules)");
            sb.AppendLine("----------------------------------------");
            if (info?.Actuation != null)
            {
                foreach (var kv in info.Actuation) sb.AppendLine($"{kv.Key}: {kv.Value}");
            }
            sb.AppendLine();

            AppendRoofSummary(sb, info);
            AppendFloorSummary(sb, info);
            AppendWallSummary(sb, info);

            sb.AppendLine($"Writes (actuation params): {info?.WriteCount ?? 0}");
            sb.AppendLine($"Type changes (triggers):   {info?.TypeChangeCount ?? 0}");
            sb.AppendLine();

            if (info?.Notes != null && info.Notes.Count > 0)
            {
                sb.AppendLine("Notes");
                sb.AppendLine("----------------------------------------");
                foreach (var n in info.Notes) sb.AppendLine($"- {n}");
                sb.AppendLine();
            }

            sb.AppendLine("PV Metrics (from CSV + deterministic calc)");
            sb.AppendLine("----------------------------------------");
            if (pvPairs.Count > 0)
            {
                foreach (var kv in pvPairs) sb.AppendLine($"{kv.Key}: {kv.Value}");
            }
            else sb.AppendLine("(No PV keys found in row)");
            sb.AppendLine();

            AppendPvDerivation(sb, row);

            return sb.ToString();
        }

        private static void AppendRoofSummary(StringBuilder sb, ResultInfo info)
        {
            sb.AppendLine("Roof Strategy (per-roof)");
            sb.AppendLine("----------------------------------------");
            if (info == null || info.RoofPvStrategyById == null || info.RoofPvStrategyById.Count == 0)
            {
                sb.AppendLine("(No per-roof strategy computed. Check: roof list passed to Apply.)");
                sb.AppendLine();
                return;
            }

            int pvNone = 0, pvLimited = 0, pvFull = 0;
            foreach (var kv in info.RoofPvStrategyById.Values)
            {
                string s = (kv ?? "").Trim().ToLowerInvariant();
                if (s == "none") pvNone++;
                else if (s == "limited") pvLimited++;
                else if (s == "full") pvFull++;
            }

            sb.AppendLine($"Total roofs evaluated: {info.RoofPvStrategyById.Count}");
            sb.AppendLine($"- pv none:    {pvNone}");
            sb.AppendLine($"- pv limited: {pvLimited}");
            sb.AppendLine($"- pv full:    {pvFull}");
            sb.AppendLine();

            int n = 0;
            foreach (var id in info.RoofPvStrategyById.Keys)
            {
                if (n >= 6) break;
                string pv = info.RoofPvStrategyById[id];
                string heat = info.RoofHeatById.ContainsKey(id) ? info.RoofHeatById[id] : "";
                string rw = info.RoofRainwaterById.ContainsKey(id) ? info.RoofRainwaterById[id] : "";
                double usable = info.RoofPvUsableById.ContainsKey(id) ? info.RoofPvUsableById[id] : double.NaN;
                sb.AppendLine($"Sample Roof {id.Value}: pv_usable≈{Fmt(usable)} m² -> pv={pv}, heat={heat}, rainwater={rw}");
                n++;
            }
            sb.AppendLine();
        }

        private static void AppendFloorSummary(StringBuilder sb, ResultInfo info)
        {
            sb.AppendLine("Floor Strategy (per-floor)");
            sb.AppendLine("----------------------------------------");
            if (info == null || info.FloorThermalById == null || info.FloorThermalById.Count == 0)
            {
                sb.AppendLine("(No per-floor strategy computed. Check: floor list passed to Apply.)");
                sb.AppendLine();
                return;
            }

            int thermIns = 0, thermBase = 0, thermOther = 0;
            foreach (var v in info.FloorThermalById.Values)
            {
                string s = (v ?? "").Trim().ToLowerInvariant();
                if (s == "insulated") thermIns++;
                else if (s == "baseline") thermBase++;
                else thermOther++;
            }

            int permNone = 0, permPer = 0, permSemi = 0, permImp = 0, permOther = 0;
            foreach (var v in info.FloorPermeabilityById.Values)
            {
                string s = (v ?? "").Trim().ToLowerInvariant();
                if (s == "none") permNone++;
                else if (s == "permeable") permPer++;
                else if (s == "semi_permeable") permSemi++;
                else if (s == "impervious") permImp++;
                else permOther++;
            }

            sb.AppendLine($"Total floors evaluated: {info.FloorThermalById.Count}");
            sb.AppendLine($"- thermal insulated: {thermIns}");
            sb.AppendLine($"- thermal baseline:  {thermBase}");
            if (thermOther > 0) sb.AppendLine($"- thermal other:     {thermOther}");
            sb.AppendLine($"- permeability none(indoor): {permNone}");
            sb.AppendLine($"- permeability permeable:    {permPer}");
            sb.AppendLine($"- permeability semi:         {permSemi}");
            sb.AppendLine($"- permeability impervious:   {permImp}");
            if (permOther > 0) sb.AppendLine($"- permeability other:        {permOther}");
            sb.AppendLine();

            int n = 0;
            foreach (var id in info.FloorThermalById.Keys)
            {
                if (n >= 6) break;
                string therm = info.FloorThermalById[id];
                double u = info.FloorUById.ContainsKey(id) ? info.FloorUById[id] : double.NaN;
                string perm = info.FloorPermeabilityById.ContainsKey(id) ? info.FloorPermeabilityById[id] : "";
                bool outdoor = info.FloorOutdoorById.ContainsKey(id) && info.FloorOutdoorById[id];
                sb.AppendLine($"Sample Floor {id.Value}: outdoor={outdoor}, U={Fmt(u)} -> thermal={therm}, permeability={perm}");
                n++;
            }
            sb.AppendLine();
        }

        private static void AppendWallSummary(StringBuilder sb, ResultInfo info)
        {
            sb.AppendLine("Wall Strategy (per-wall, from each wall's ea_wall_u_value)");
            sb.AppendLine("----------------------------------------");

            if (info == null || info.WallStrategyById == null || info.WallStrategyById.Count == 0)
            {
                sb.AppendLine("(No per-wall strategy computed. Check: wall list passed to Apply, and ea_wall_u_value exists.)");
                sb.AppendLine();
                return;
            }

            int hp = 0, bl = 0, ins = 0, unk = 0;
            foreach (var kv in info.WallStrategyById.Values)
            {
                string s = (kv ?? "").Trim().ToLowerInvariant();
                if (s == "high_performance") hp++;
                else if (s == "baseline") bl++;
                else if (s == "insulated_upgrade") ins++;
                else unk++;
            }

            sb.AppendLine($"Total walls evaluated: {info.WallStrategyById.Count}");
            sb.AppendLine($"- high_performance (U ≤ 0.6): {hp}");
            sb.AppendLine($"- baseline (0.6 < U ≤ 0.9):   {bl}");
            sb.AppendLine($"- insulated_upgrade (U > 0.9): {ins}");
            if (unk > 0) sb.AppendLine($"- unknown: {unk}");
            sb.AppendLine();

            int n = 0;
            foreach (var id in info.WallStrategyById.Keys)
            {
                if (n >= 6) break;
                string strat = info.WallStrategyById[id];
                double u = info.WallUById.ContainsKey(id) ? info.WallUById[id] : double.NaN;
                sb.AppendLine($"Sample Wall {id.Value}: U={Fmt(u)} -> {strat}");
                n++;
            }
            sb.AppendLine();
        }

        private static bool IsPvKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return false;
            string k = key.Trim().ToLowerInvariant();

            return k.StartsWith("pv_") ||
                   k.StartsWith("ea_pv_") ||
                   k.Contains("panel_unit") ||
                   k.Contains("coverage_factor") ||
                   k.Contains("layout_status");
        }

        private static void AppendPvDerivation(StringBuilder sb, Dictionary<string, string> row)
        {
            double pvUsable = D(row, "pv_usable_roof_area_m2");
            double pvAdded = D(row, "ea_pv_added_area_m2");
            int panels = I(row, "ea_pv_panel_count");

            double cov = D(row, "coverage_factor_used");
            if (cov <= 1e-9)
            {
                if (pvUsable > 1e-9 && pvAdded > 1e-9) cov = pvAdded / pvUsable;
                else cov = DEFAULT_PV_COVERAGE;
            }

            double pw = D(row, "pv_panel_unit_width_m");
            double pl = D(row, "pv_panel_unit_length_m");
            if (pw <= 1e-9) pw = DEFAULT_PANEL_W_M;
            if (pl <= 1e-9) pl = DEFAULT_PANEL_L_M;

            double pA = D(row, "pv_panel_unit_area_m2");
            if (pA <= 1e-9) pA = pw * pl;

            int expectedPanels = 0;
            if (pvAdded > 1e-9 && pA > 1e-9) expectedPanels = (int)Math.Floor(pvAdded / pA);

            sb.AppendLine("PV Derivation (deterministic)");
            sb.AppendLine("----------------------------------------");
            sb.AppendLine($"coverage_factor_used: {Fmt(cov)}  (default = {DEFAULT_PV_COVERAGE})");
            sb.AppendLine("ea_pv_added_area_m2 = pv_usable_roof_area_m2 × coverage_factor_used");
            sb.AppendLine($"- pv_usable_roof_area_m2: {Fmt(pvUsable)}");
            sb.AppendLine($"- ea_pv_added_area_m2:    {Fmt(pvAdded)}");
            sb.AppendLine();
            sb.AppendLine("panel_area_m2 = panel_width_m × panel_length_m");
            sb.AppendLine($"- panel_width_m:  {Fmt(pw)}");
            sb.AppendLine($"- panel_length_m: {Fmt(pl)}");
            sb.AppendLine($"- panel_area_m2:  {Fmt(pA)}");
            sb.AppendLine();
            sb.AppendLine("ea_pv_panel_count = floor(ea_pv_added_area_m2 ÷ panel_area_m2)");
            sb.AppendLine($"- ea_pv_panel_count: {panels}");
            if (expectedPanels > 0 && panels != expectedPanels)
                sb.AppendLine($"- check: expected {expectedPanels} from formula (verify rounding/inputs in ImportRetrofitCSV).");
            sb.AppendLine();
        }

        // ============================================================
        // Safe gets + parse
        // ============================================================
        private static string GetOrEmpty(Dictionary<string, string> d, string key)
        {
            if (d == null || string.IsNullOrWhiteSpace(key)) return "";
            if (d.TryGetValue(key, out var v) && v != null) return v;
            return "";
        }

        private static string GetOr(Dictionary<string, string> d, string key, string fallback)
        {
            var v = GetOrEmpty(d, key);
            return string.IsNullOrWhiteSpace(v) ? (fallback ?? "") : v;
        }

        private static double D(Dictionary<string, string> row, string key)
        {
            string s = RetrofitShared.CleanRaw(GetOrEmpty(row, key));
            return RetrofitShared.ParseDouble(s);
        }

        private static int I(Dictionary<string, string> row, string key)
        {
            string s = RetrofitShared.CleanRaw(GetOrEmpty(row, key));
            if (string.IsNullOrWhiteSpace(s)) return 0;

            double v = RetrofitShared.ParseDouble(s);
            if (double.IsNaN(v) || double.IsInfinity(v)) return 0;
            return (int)Math.Round(v);
        }

        private static string Fmt(double v)
        {
            if (double.IsNaN(v) || double.IsInfinity(v)) return "NaN";
            return v.ToString("0.###", CultureInfo.InvariantCulture);
        }

        // ============================================================
        // GLOBAL actuation (building/site-level only)
        // ============================================================
        private Dictionary<string, string> ComputeGlobalActuation(Dictionary<string, string> row)
        {
            var a = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // WE water efficiency level (global)
            double wri = D(row, "baseline_water_stress_mean");
            if (wri >= 1.5) a["we_water_efficiency_level"] = "high";
            else if (wri >= 1.0) a["we_water_efficiency_level"] = "medium";
            else a["we_water_efficiency_level"] = "low";

            // Fixture class (derived, global)
            string weLevel = GetOr(a, "we_water_efficiency_level", "low");
            if (string.Equals(weLevel, "high", StringComparison.OrdinalIgnoreCase)) a["we_fixture_efficiency_class"] = "A";
            else if (string.Equals(weLevel, "medium", StringComparison.OrdinalIgnoreCase)) a["we_fixture_efficiency_class"] = "B";
            else a["we_fixture_efficiency_class"] = "C";

            // Window upgrade (global input value)
            double winU = D(row, "ea_window_u_value");
            if (winU >= 2.0) a["ea_window_upgrade_strategy"] = "low_e_double";
            else if (winU >= 1.4) a["ea_window_upgrade_strategy"] = "frame_upgrade";
            else a["ea_window_upgrade_strategy"] = "baseline";

            // Window shading (global inputs)
            double dlp = D(row, "lt_daylight_potential");
            double ori = D(row, "ea_orientation_deg");
            bool southish = (ori >= 90 && ori <= 270);
            if (dlp > 0.40 && southish) a["ea_window_shading_strategy"] = "fixed";
            else a["ea_window_shading_strategy"] = "none";

            // Door upgrade (global input u)
            double doorU = D(row, "ea_door_u_value");
            if (doorU >= 2.0) a["ea_door_upgrade_strategy"] = "door_replacement_high_perf";
            else if (doorU >= 1.4) a["ea_door_upgrade_strategy"] = "frame_upgrade";
            else if (doorU >= 1.0) a["ea_door_upgrade_strategy"] = "air_seal_only";
            else a["ea_door_upgrade_strategy"] = "baseline";

            // LT constraint flag (global)
            double dist = D(row, "distance_to_nearest_transit_m");
            double amen = D(row, "amenity_count_400m");
            if (dist > 800 && amen < 10) a["lt_constraint_flag"] = "true";
            else a["lt_constraint_flag"] = "false";

            return a;
        }

        // ============================================================
        // Per-element rules
        // ============================================================

        // Wall: your final logic
        // U ≤ 0.6 => high_performance (no upgrade)
        // 0.6 < U ≤ 0.9 => baseline (no upgrade)
        // U > 0.9 => insulated_upgrade (upgrade)
        private static string ComputeWallStrategyFromU(double u)
        {
            if (double.IsNaN(u) || double.IsInfinity(u) || u <= 0) return "baseline";
            if (u <= 0.6) return "high_performance";
            if (u <= 0.9) return "baseline";
            return "insulated_upgrade";
        }

        // Floor thermal
        private static string ComputeFloorThermalFromU(double uSlab)
        {
            if (double.IsNaN(uSlab) || double.IsInfinity(uSlab) || uSlab < 0) return "baseline";
            return (uSlab >= 0.60) ? "insulated" : "baseline";
        }

        // Floor permeability (only meaningful outdoor; indoor => none)
        private static string ComputeFloorPermeabilityFromImpervious(double imperviousMeanRatio)
        {
            if (double.IsNaN(imperviousMeanRatio) || double.IsInfinity(imperviousMeanRatio) || imperviousMeanRatio < 0)
                return "impervious";

            if (imperviousMeanRatio >= 0.70) return "impervious";
            if (imperviousMeanRatio >= 0.40) return "semi_permeable";
            return "permeable";
        }

        // Roof heat base (global inputs, but applied per roof with PV priority)
        private static string ComputeRoofHeatBase(double ss_heat_island, double lst_mean_celsius)
        {
            if (ss_heat_island >= 0.66 || lst_mean_celsius >= 35) return "cool_roof";
            if (ss_heat_island >= 0.33) return "green_roof";
            return "none";
        }

        // Roof PV strategy per roof (thresholds from your table, using per-roof usable area)
        private static string ComputeRoofPvStrategy(double pvUsableRoofM2, double shadowRatioMean)
        {
            if (pvUsableRoofM2 <= 10 || shadowRatioMean >= 0.60) return "none";
            if (pvUsableRoofM2 >= 80) return "full";
            return "limited";
        }

        private static string ComputeRoofPvSystemType(string pvStrategy)
        {
            return string.Equals((pvStrategy ?? "").Trim(), "none", StringComparison.OrdinalIgnoreCase) ? "none" : "mounted";
        }

        private static string ComputeRoofStructuralFlag(string pvStrategy)
        {
            return string.Equals((pvStrategy ?? "").Trim(), "full", StringComparison.OrdinalIgnoreCase) ? "true" : "false";
        }

        private static string ComputeRoofRainwater(double baselineWaterStressMean, double pvUsableRoofM2)
        {
            if (baselineWaterStressMean >= 1.0 && pvUsableRoofM2 >= 20) return "basic";
            return "none";
        }

        // ============================================================
        // Write actuation params
        // ============================================================
        private int WriteActuation(
            Document doc,
            IList<Element> roofs,
            IList<Element> windows,
            IList<Element> doors,
            IList<Element> walls,
            IList<Element> floors,
            Dictionary<string, string> row,
            ResultInfo info)
        {
            int c = 0;
            var act = info?.Actuation ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // ---------- Roofs (PER ROOF) ----------
            if (roofs != null)
            {
                // Inputs (global)
                double pvUsableTotal = D(row, "pv_usable_roof_area_m2");
                double roofTotal = D(row, "roof_total_area_m2"); // if present
                double shadow = D(row, "shadow_ratio_mean");
                double wri = D(row, "baseline_water_stress_mean");
                double ss = D(row, "ss_heat_island");
                double lst = D(row, "lst_mean_celsius");

                // usable ratio (for distributing usable area across multiple roofs)
                double usableRatio = 1.0;
                if (roofTotal > 1e-9 && pvUsableTotal >= 0)
                    usableRatio = Math.Max(0.0, Math.Min(1.0, pvUsableTotal / roofTotal));

                string heatBase = ComputeRoofHeatBase(ss, lst);

                foreach (var r in roofs)
                {
                    double roofAreaM2 = TryGetElementAreaM2(r);
                    double pvUsableThis = (roofAreaM2 > 1e-9) ? roofAreaM2 * usableRatio : pvUsableTotal; // fallback

                    string pvStrat = ComputeRoofPvStrategy(pvUsableThis, shadow);

                    string heat = heatBase;
                    if (ROOF_PV_PRIORITY && !string.Equals(pvStrat, "none", StringComparison.OrdinalIgnoreCase))
                        heat = "none";

                    string rw = ComputeRoofRainwater(wri, pvUsableThis);
                    string pvType = ComputeRoofPvSystemType(pvStrat);
                    string pvFlag = ComputeRoofStructuralFlag(pvStrat);

                    // cache for report/debug
                    if (info != null)
                    {
                        info.RoofPvUsableById[r.Id] = pvUsableThis;
                        info.RoofPvStrategyById[r.Id] = pvStrat;
                        info.RoofHeatById[r.Id] = heat;
                        info.RoofRainwaterById[r.Id] = rw;
                    }

                    if (RetrofitShared.EnsureSetStringInstanceOrType(doc, r, "ea_roof_pv_strategy", pvStrat)) c++;
                    if (RetrofitShared.EnsureSetStringInstanceOrType(doc, r, "ea_roof_pv_system_type", pvType)) c++;
                    if (RetrofitShared.EnsureSetStringInstanceOrType(doc, r, "ea_roof_structural_capacity_flag", pvFlag)) c++;
                    if (RetrofitShared.EnsureSetStringInstanceOrType(doc, r, "ss_roof_heat_strategy", heat)) c++;
                    if (RetrofitShared.EnsureSetStringInstanceOrType(doc, r, "we_roof_rainwater_strategy", rw)) c++;
                }
            }

            // ---------- Windows (GLOBAL strategies are OK) ----------
            if (windows != null)
            {
                foreach (var w in windows)
                {
                    c += SetS(doc, w, "ea_window_upgrade_strategy", act);
                    c += SetS(doc, w, "ea_window_shading_strategy", act);
                }
            }

            // ---------- Doors (GLOBAL strategy) ----------
            if (doors != null)
            {
                foreach (var d in doors)
                    c += SetS(doc, d, "ea_door_upgrade_strategy", act);
            }

            // ---------- Walls (PER WALL) ----------
            if (walls != null)
            {
                foreach (var wa in walls)
                {
                    double u = GetDoubleOnElementOrType(doc, wa, "ea_wall_u_value");
                    string strat = ComputeWallStrategyFromU(u);

                    if (info != null)
                    {
                        info.WallUById[wa.Id] = u;
                        info.WallStrategyById[wa.Id] = strat;
                    }

                    if (RetrofitShared.EnsureSetStringInstanceOrType(doc, wa, "ea_wall_envelope_strategy", strat))
                        c++;
                }
            }

            // ---------- Floors (PER FLOOR) ----------
            if (floors != null)
            {
                double imperv = D(row, "impervious_mean_ratio");
                string outdoorPerm = ComputeFloorPermeabilityFromImpervious(imperv);

                foreach (var f in floors)
                {
                    // thermal per slab U
                    double uSlab = GetDoubleOnElementOrType(doc, f, "ea_slab_u_value");
                    string therm = ComputeFloorThermalFromU(uSlab);

                    // permeability: only outdoor floors; indoor => none
                    bool isOutdoor = IsOutdoorFloor(doc, f);
                    string perm = isOutdoor ? outdoorPerm : "none";

                    if (info != null)
                    {
                        info.FloorUById[f.Id] = uSlab;
                        info.FloorThermalById[f.Id] = therm;
                        info.FloorOutdoorById[f.Id] = isOutdoor;
                        info.FloorPermeabilityById[f.Id] = perm;
                    }

                    if (RetrofitShared.EnsureSetStringInstanceOrType(doc, f, "ea_floor_thermal_strategy", therm)) c++;

                    // IMPORTANT: do not contaminate indoor floors with permeability strategy
                    if (RetrofitShared.EnsureSetStringInstanceOrType(doc, f, "we_floor_permeability_strategy", perm)) c++;
                }
            }

            // ---------- ProjectInfo (GLOBAL) ----------
            var proj = doc?.ProjectInformation;
            if (proj != null)
            {
                c += SetS(doc, proj, "lt_constraint_flag", act);
                c += SetS(doc, proj, "we_water_efficiency_level", act);
                c += SetS(doc, proj, "we_fixture_efficiency_class", act);
            }

            return c;
        }

        private int SetS(Document doc, Element e, string key, Dictionary<string, string> act)
        {
            if (act == null || !act.TryGetValue(key, out var v)) return 0;
            if (string.IsNullOrWhiteSpace(v)) return 0;
            return RetrofitShared.EnsureSetStringInstanceOrType(doc, e, key, v) ? 1 : 0;
        }

        // ============================================================
        // Optional: type changes (only if ENABLE_TYPE_CHANGE=true)
        // ============================================================
        private int ApplyTriggers(
            Document doc,
            IList<Element> roofs,
            IList<Element> walls,
            IList<Element> floors,
            Dictionary<string, string> row,
            ResultInfo info,
            RevitTypeNameMap map)
        {
            int changed = 0;

            // ---- Roof: choose target type per roof based on per-roof computed strategies ----
            foreach (var r in roofs ?? new List<Element>())
            {
                string pv = (info != null && info.RoofPvStrategyById.TryGetValue(r.Id, out var pv0)) ? pv0 : "none";
                string heat = (info != null && info.RoofHeatById.TryGetValue(r.Id, out var h0)) ? h0 : "none";

                string logical;
                if (ROOF_PV_PRIORITY && !string.Equals(pv, "none", StringComparison.OrdinalIgnoreCase))
                    logical = "ROOF_PV";
                else if (string.Equals(heat, "cool_roof", StringComparison.OrdinalIgnoreCase))
                    logical = "ROOF_COOL_L3";
                else if (string.Equals(heat, "green_roof", StringComparison.OrdinalIgnoreCase))
                    logical = "ROOF_GREEN_L2";
                else
                    logical = "ROOF_EXISTING";

                if (map.Roof.TryGetValue(logical, out var targetTypeName) && !string.IsNullOrWhiteSpace(targetTypeName))
                {
                    if (TryChangeElementTypeByName(doc, r, targetTypeName, out _))
                        changed++;
                }
            }

            // ---- Walls: upgrade only when per-wall strategy == insulated_upgrade ----
            foreach (var w in walls ?? new List<Element>())
            {
                string strat = (info != null && info.WallStrategyById.TryGetValue(w.Id, out var s0)) ? s0 : "baseline";
                bool wantUpgrade = string.Equals((strat ?? "").Trim(), "insulated_upgrade", StringComparison.OrdinalIgnoreCase);
                if (!wantUpgrade) continue;

                string suffix = map.Wall.TryGetValue("WALL_RETROFIT_L1_SUFFIX", out var sfx) ? sfx : " Insulated";
                string currentTypeName = GetCurrentTypeName(doc, w);
                if (string.IsNullOrWhiteSpace(currentTypeName)) continue;

                string targetTypeName = currentTypeName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                    ? currentTypeName
                    : (currentTypeName + suffix);

                if (TryChangeElementTypeByName(doc, w, targetTypeName, out _))
                    changed++;
            }

            // ---- Floors: implement your note logic (Outdoor + permeable + baseline => FLOOR_PERMEABLE; insulated => FLOOR_INSULATED) ----
            foreach (var f in floors ?? new List<Element>())
            {
                bool isOutdoor = IsOutdoorFloor(doc, f);

                string therm = (info != null && info.FloorThermalById.TryGetValue(f.Id, out var t0)) ? t0 : "baseline";
                string perm = (info != null && info.FloorPermeabilityById.TryGetValue(f.Id, out var p0)) ? p0 : "none";

                bool wantInsFloor = string.Equals(therm, "insulated", StringComparison.OrdinalIgnoreCase);
                bool wantPermeable = string.Equals(perm, "permeable", StringComparison.OrdinalIgnoreCase);

                if (isOutdoor && wantPermeable && !wantInsFloor)
                {
                    if (map.Floor.TryGetValue("FLOOR_PERMEABLE", out var targetTypeName) && !string.IsNullOrWhiteSpace(targetTypeName))
                    {
                        if (TryChangeElementTypeByName(doc, f, targetTypeName, out _))
                            changed++;
                    }
                    continue;
                }

                if (wantInsFloor)
                {
                    string suffix = map.Floor.TryGetValue("FLOOR_INSULATED_SUFFIX", out var sfx) ? sfx : " Insulated";
                    string currentTypeName = GetCurrentTypeName(doc, f);
                    if (string.IsNullOrWhiteSpace(currentTypeName)) continue;

                    string targetTypeName = currentTypeName.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)
                        ? currentTypeName
                        : (currentTypeName + suffix);

                    if (TryChangeElementTypeByName(doc, f, targetTypeName, out _))
                        changed++;
                }
            }

            return changed;
        }

        // ============================================================
        // Helpers: element reading + geometry
        // ============================================================
        private static double GetDoubleOnElementOrType(Document doc, Element e, string paramName)
        {
            if (doc == null || e == null || string.IsNullOrWhiteSpace(paramName)) return double.NaN;

            double v = TryReadDouble(e, paramName);
            if (!double.IsNaN(v) && !double.IsInfinity(v)) return v;

            try
            {
                var tid = e.GetTypeId();
                if (tid == ElementId.InvalidElementId) return double.NaN;
                var t = doc.GetElement(tid) as ElementType;
                if (t == null) return double.NaN;

                v = TryReadDouble(t, paramName);
                return v;
            }
            catch { return double.NaN; }
        }

        private static double TryReadDouble(Element e, string paramName)
        {
            try
            {
                var p = e?.LookupParameter(paramName);
                if (p == null) return double.NaN;

                if (p.StorageType == StorageType.Double)
                    return p.AsDouble();

                string s = "";
                try { s = p.AsValueString() ?? ""; } catch { s = ""; }
                if (string.IsNullOrWhiteSpace(s))
                {
                    try { s = p.AsString() ?? ""; } catch { s = ""; }
                }

                s = RetrofitShared.CleanRaw(s);
                if (string.IsNullOrWhiteSpace(s)) return double.NaN;

                return RetrofitShared.ParseDouble(s);
            }
            catch
            {
                return double.NaN;
            }
        }

        // Try to read element area (internal ft² -> m²)
        private static double TryGetElementAreaM2(Element e)
        {
            if (e == null) return 0.0;

            try
            {
                // Many host elements expose HOST_AREA_COMPUTED
                Parameter p = e.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    double aInternal = p.AsDouble(); // ft²
                    if (aInternal > 1e-9)
                        return UnitUtils.ConvertFromInternalUnits(aInternal, UnitTypeId.SquareMeters);
                }
            }
            catch { }

            // fallback: try common "Area" parameter
            try
            {
                var p2 = e.LookupParameter("Area") ?? e.LookupParameter("Fläche");
                if (p2 != null && p2.StorageType == StorageType.Double)
                {
                    double aInternal = p2.AsDouble();
                    if (aInternal > 1e-9)
                        return UnitUtils.ConvertFromInternalUnits(aInternal, UnitTypeId.SquareMeters);
                }
            }
            catch { }

            return 0.0;
        }

        private static string GetCurrentTypeName(Document doc, Element e)
        {
            try
            {
                if (doc == null || e == null) return "";
                var tid = e.GetTypeId();
                if (tid == ElementId.InvalidElementId) return "";
                var t = doc.GetElement(tid) as ElementType;
                return t?.Name ?? "";
            }
            catch { return ""; }
        }

        private static bool IsOutdoorFloor(Document doc, Element floor)
        {
            try
            {
                if (doc == null || floor == null) return false;

                string tn = GetCurrentTypeName(doc, floor);
                if (!string.IsNullOrWhiteSpace(tn))
                {
                    string t = tn.ToLowerInvariant();
                    if (t.Contains("terrasse") ||
                        t.Contains("dachterrasse") ||
                        t.Contains("balkon") ||
                        t.Contains("balcon") ||
                        t.Contains("loggia") ||
                        t.Contains("außen") ||
                        t.Contains("aussen"))
                    {
                        return true;
                    }
                }

                Parameter p =
                    floor.LookupParameter("Room Bounding") ??
                    floor.LookupParameter("Raumbegrenzung");

                if (p != null && p.StorageType == StorageType.Integer)
                {
                    int rb = p.AsInteger(); // 1=true, 0=false
                    if (rb == 0) return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool TryChangeElementTypeByName(Document doc, Element e, string targetTypeName, out string msg)
        {
            msg = "";
            if (doc == null || e == null) { msg = "doc/element null"; return false; }
            if (string.IsNullOrWhiteSpace(targetTypeName)) { msg = "targetTypeName empty"; return false; }

            try
            {
                ElementId tid = e.GetTypeId();
                if (tid == ElementId.InvalidElementId)
                {
                    msg = $"Element has no type: {e.Id.Value}";
                    return false;
                }

                var catId = e.Category?.Id;
                if (catId == null)
                {
                    msg = $"No category: {e.Id.Value}";
                    return false;
                }

                var types = new FilteredElementCollector(doc)
                    .WhereElementIsElementType()
                    .Where(x => x?.Category != null && x.Category.Id.Value == catId.Value)
                    .Cast<ElementType>()
                    .ToList();

                var target = types.FirstOrDefault(t => string.Equals(t.Name, targetTypeName, StringComparison.OrdinalIgnoreCase));
                if (target == null)
                {
                    msg = $"Type not found in category [{e.Category.Name}]: {targetTypeName}";
                    return false;
                }

                if (target.Id == tid)
                {
                    msg = $"Already that type: {targetTypeName}";
                    return false;
                }

                e.ChangeTypeId(target.Id);
                return true;
            }
            catch (Exception ex)
            {
                msg = ex.Message;
                return false;
            }
        }
    }
}