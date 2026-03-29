using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace RevitLOD3Exporter
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ImportRetrofitCSV : IExternalCommand
    {
        private const string REPORT_PARAM_NAME = "leed_retrofit_report";
        private const bool EXPORT_REPORT_TXT = true;

        // ✅ Updated: include fields required by your Baseline2 rules
        private static readonly string[] GLOBAL_PARAMS = new[]
        {
            "building_id",
            "retrofit_score_total",
            "ea_score",
            "ss_score",
            "we_score",
            "lt_score",

            "lt_daylight_potential",
            "ss_heat_island",
            "lst_mean_celsius",
            "shadow_ratio_mean",

            "baseline_water_stress_mean",
            "impervious_mean_ratio",
            "pv_usable_roof_area_m2",

            "amenity_count_400m",
            "distance_to_nearest_transit_m",

            // optional / legacy (kept harmless)
            "rainwater_harvest_m3_year",
        };

        private static readonly string[] WINDOW_CSV_FORCE_PARAMS = new[]
        {
            "ea_window_u_value",
            "ea_window_shgc",
            "climate_zone",
            "heating_degree_days",
            "cooling_degree_days",
        };

        private enum BaselineMode
        {
            Baseline2_Rules,
            Baseline3_AI
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;

                string csvPath = RetrofitShared.SelectCSVFile();
                if (string.IsNullOrWhiteSpace(csvPath))
                {
                    TaskDialog.Show("Cancelled", "CSV import cancelled.");
                    return Result.Cancelled;
                }

                string reportOutDir = null;
                if (EXPORT_REPORT_TXT)
                    reportOutDir = RetrofitShared.SelectReportOutputFolder();

                BaselineMode mode;
                try { mode = PromptBaselineMode(); }
                catch
                {
                    TaskDialog.Show("Cancelled", "Baseline selection cancelled.");
                    return Result.Cancelled;
                }

                string forcedBuildingId = RetrofitShared.GetForcedBuildingIdFromFileName(csvPath);

                var rowDict = RetrofitShared.ReadMatchingRowAsDict(csvPath, forcedBuildingId);
                if (rowDict == null || rowDict.Count == 0)
                {
                    TaskDialog.Show("CSV Error", "CSV is empty/invalid, or matching building_id row not found.");
                    return Result.Failed;
                }

                if (!string.IsNullOrWhiteSpace(forcedBuildingId))
                    rowDict["building_id"] = forcedBuildingId;

                var roofs = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Roofs)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

                var windows = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Windows)
                    .WhereElementIsNotElementType()
                    .OfType<FamilyInstance>()
                    .Cast<Element>()
                    .ToList();

                var walls = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .OfType<Wall>()
                    .Cast<Element>()
                    .ToList();

                var floors = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

                var doors = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Doors)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

                var views = new FilteredElementCollector(doc)
                    .OfClass(typeof(Autodesk.Revit.DB.View))
                    .Cast<Autodesk.Revit.DB.View>()
                    .Where(v => v != null && !v.IsTemplate)
                    .ToList();

                // ✅ Deterministic PV calculation written into rowDict (Baseline2/3 share it)
                ApplyPvDeterministicToRow(rowDict);

                bool apiOk = false;
                string aiRawPlan = "";
                string aiRawAdvice = "";
                Baseline3_DeepSeek.LeedPlan plan = null;
                Baseline3_DeepSeek.LeedAdvice advice = null;

                if (mode == BaselineMode.Baseline3_AI)
                {
                    try
                    {
                        string sriNow = RetrofitShared.ComputeSriTargetFromHeatIsland(rowDict);
                        if (!string.IsNullOrWhiteSpace(sriNow))
                            rowDict["ss_roof_sri_target"] = sriNow;

                        plan = Baseline3_DeepSeek.CallDeepSeekLeedPlan(doc, roofs, rowDict, out aiRawPlan);
                        apiOk = (plan != null);

                        if (apiOk)
                        {
                            LockPlanPvTargetsToRow(plan, rowDict);
                            advice = Baseline3_DeepSeek.CallDeepSeekLeedAdvice(rowDict, plan, out aiRawAdvice);
                        }

                        RetrofitShared.ShowTimedPopup(
                            "DeepSeek API Status",
                            apiOk ? "✅ DeepSeek API call successful: PLAN + DETAILED ADVICE generated."
                                  : "❌ DeepSeek API call failed: PLAN not generated.",
                            5
                        );
                    }
                    catch (Exception ex)
                    {
                        apiOk = false;
                        aiRawPlan = ex.ToString();
                        plan = null;
                        advice = null;

                        RetrofitShared.ShowTimedPopup(
                            "DeepSeek API Status",
                            "❌ DeepSeek API call failed. See report for error text.",
                            5
                        );
                    }
                }

                string reportText = "";

                using (Transaction tx = new Transaction(doc, "Retrofit Import (Baseline)"))
                {
                    tx.Start();

                    // 1) Write ALL CSV params to elements (if parameter exists on those elements/types)
                    foreach (var r in roofs) RetrofitShared.WriteAllCsvParams(doc, r, rowDict);
                    foreach (var w in windows)
                    {
                        RetrofitShared.WriteAllCsvParams(doc, w, rowDict);
                        RetrofitShared.WriteWindowCsvForced(doc, w, rowDict);
                    }
                    foreach (var wa in walls) RetrofitShared.WriteAllCsvParams(doc, wa, rowDict);
                    foreach (var f in floors) RetrofitShared.WriteAllCsvParams(doc, f, rowDict);
                    foreach (var d in doors) RetrofitShared.WriteAllCsvParams(doc, d, rowDict);

                    // 2) Write global fields into ProjectInfo + Views
                    RetrofitShared.WriteCsvParamsByName(doc, doc.ProjectInformation, rowDict, GLOBAL_PARAMS);
                    foreach (var v in views)
                        RetrofitShared.WriteCsvParamsByName(doc, v, rowDict, GLOBAL_PARAMS);

                    // 3) Ensure exact write for GLOBAL_PARAMS if present in CSV row
                    foreach (var pName in GLOBAL_PARAMS)
                        RetrofitShared.EnsureWriteCsvExact(doc, doc.ProjectInformation, rowDict, pName);
                    foreach (var v in views)
                    {
                        foreach (var pName in GLOBAL_PARAMS)
                            RetrofitShared.EnsureWriteCsvExact(doc, v, rowDict, pName);
                    }

                    // 4) Extract U-values from element types and write to instance/type params
                    foreach (var e in walls)
                    {
                        var wall = e as Wall;
                        if (wall == null) continue;

                        double uWall = 0.0;
                        if (RetrofitShared.TryGetUValueFromElementType(doc, wall, out double foundU) && !double.IsNaN(foundU) && foundU >= 0)
                            uWall = foundU;

                        RetrofitShared.EnsureSetDoubleInstanceOrType(doc, wall, "ea_wall_u_value", uWall);
                    }

                    foreach (var r in roofs)
                    {
                        double uRoof = 0.0;
                        if (RetrofitShared.TryGetUValueFromElementType(doc, r, out double foundU) && !double.IsNaN(foundU) && foundU >= 0)
                            uRoof = foundU;

                        RetrofitShared.EnsureSetDoubleInstanceOrType(doc, r, "ea_roof_u_value", uRoof);
                    }

                    foreach (var f in floors)
                    {
                        double uSlab = 0.0;
                        if (RetrofitShared.TryGetUValueFromElementType(doc, f, out double foundU) && !double.IsNaN(foundU) && foundU >= 0)
                            uSlab = foundU;

                        RetrofitShared.EnsureSetDoubleInstanceOrType(doc, f, "ea_slab_u_value", uSlab);
                    }

                    foreach (var d in doors)
                    {
                        double uDoor = 0.0;
                        if (RetrofitShared.TryGetUValueFromElementType(doc, d, out double foundU) && !double.IsNaN(foundU) && foundU >= 0)
                            uDoor = foundU;

                        RetrofitShared.EnsureSetDoubleInstanceOrType(doc, d, "ea_door_u_value", uDoor);
                    }

                    // 5) Orientation: write ea_orientation_deg to exterior walls + their hosted windows
                    var wallWinMap = RetrofitShared.BuildWallToWindowsMap(doc, windows);
                    foreach (var e in walls)
                    {
                        var wall = e as Wall;
                        if (wall == null) continue;
                        if (!RetrofitShared.IsExteriorWall(wall)) continue;
                        if (!RetrofitShared.TryGetWallAzimuthDeg(wall, out double azDeg)) continue;

                        RetrofitShared.EnsureSetDoubleInstanceOrType(doc, wall, "ea_orientation_deg", azDeg);

                        if (wallWinMap.TryGetValue(wall.Id, out var winsOnWall) && winsOnWall != null)
                        {
                            foreach (var w in winsOnWall)
                                RetrofitShared.EnsureSetDoubleInstanceOrType(doc, w, "ea_orientation_deg", azDeg);
                        }
                    }

                    // 6) SRI target from heat island (optional but kept for Baseline3/compat)
                    string sri = RetrofitShared.ComputeSriTargetFromHeatIsland(rowDict);
                    if (!string.IsNullOrWhiteSpace(sri))
                    {
                        RetrofitShared.EnsureSetStringInstanceOrType(doc, doc.ProjectInformation, "ss_roof_sri_target", sri);
                        foreach (var v in views)
                            RetrofitShared.EnsureSetStringInstanceOrType(doc, v, "ss_roof_sri_target", sri);

                        foreach (var r in roofs)
                            RetrofitShared.EnsureSetStringInstanceOrType(doc, r, "ss_roof_sri_target", sri);

                        rowDict["ss_roof_sri_target"] = sri;
                    }

                    // 7) WWR ratio (optional)
                    foreach (var e in walls)
                    {
                        var wall = e as Wall;
                        if (wall == null) continue;
                        if (!RetrofitShared.IsExteriorWall(wall)) continue;

                        if (!wallWinMap.TryGetValue(wall.Id, out var winsOnWall) || winsOnWall.Count == 0)
                            continue;

                        double wallGrossAreaM2 = RetrofitShared.TryGetWallGrossAreaM2(wall);
                        if (wallGrossAreaM2 <= 1e-9) continue;

                        double winAreaM2 = 0.0;
                        foreach (var w in winsOnWall)
                            winAreaM2 += RetrofitShared.TryGetWindowAreaM2(w);

                        if (winAreaM2 <= 1e-9) continue;

                        double ratio = winAreaM2 / wallGrossAreaM2;
                        ratio = Math.Max(0, Math.Min(1, ratio));

                        RetrofitShared.EnsureSetDoubleInstanceOrType(doc, wall, "ea_wwr_ratio", ratio);
                        foreach (var w in winsOnWall)
                            RetrofitShared.EnsureSetDoubleInstanceOrType(doc, w, "ea_wwr_ratio", ratio);
                    }

                    // ✅ write PV deterministic values to ProjectInfo + Views + Roofs
                    WritePvDeterministicToRevit(doc, views, roofs, rowDict);

                    // 8) Baseline logic
                    if (mode == BaselineMode.Baseline2_Rules)
                    {
                        var b2 = new Baseline2_Rules();
                        var b2Result = b2.Apply(doc, rowDict, roofs, windows, doors, walls, floors);
                        reportText = b2.BuildReport(rowDict, b2Result);
                    }
                    else
                    {
                        var b3 = new Baseline3_DeepSeek();
                        var b3Result = b3.Apply(doc, rowDict, roofs, windows, doors, walls, floors, plan);

                        reportText = b3.BuildReport(rowDict, plan, advice, apiOk, b3Result);

                        if (!apiOk && !string.IsNullOrWhiteSpace(aiRawPlan))
                        {
                            reportText += "\n\n[AI ERROR]\n----------------------------------------\n";
                            reportText += aiRawPlan;
                            reportText += "\n";
                        }
                    }

                    RetrofitShared.WriteReportToProjectInformation(doc, REPORT_PARAM_NAME, reportText);

                    tx.Commit();
                }

                string txtOutPath = null;
                if (EXPORT_REPORT_TXT && !string.IsNullOrWhiteSpace(reportOutDir))
                {
                    try
                    {
                        txtOutPath = Path.Combine(
                            reportOutDir,
                            $"LEED_Retrofit_Report_{Path.GetFileNameWithoutExtension(csvPath)}_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
                        );
                        File.WriteAllText(txtOutPath, reportText, Encoding.UTF8);
                    }
                    catch
                    {
                        txtOutPath = null;
                    }
                }

                TaskDialog.Show(
                    "Retrofit Import Completed",
                    $"Mode: {mode}\n\n" +
                    $"CSV file: {Path.GetFileName(csvPath)}\n" +
                    $"Forced building_id: {(forcedBuildingId ?? "(none)")}\n\n" +
                    $"Report param: {REPORT_PARAM_NAME}\n" +
                    (txtOutPath != null ? $"TXT exported:\n{txtOutPath}\n" : "TXT export: skipped/failed.\n")
                );

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Import Error", ex.ToString());
                return Result.Failed;
            }
        }

        private BaselineMode PromptBaselineMode()
        {
            var td = new TaskDialog("Select Baseline Mode");
            td.MainInstruction = "Choose which baseline to run";
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                "Baseline 2 (Predefined Rules)",
                "Run deterministic if-then rules.");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                "Baseline 3 (AI)",
                "Call AI: PLAN + detailed advice, write parameters and report.");
            td.CommonButtons = TaskDialogCommonButtons.Cancel;

            var r = td.Show();
            if (r == TaskDialogResult.CommandLink1) return BaselineMode.Baseline2_Rules;
            if (r == TaskDialogResult.CommandLink2) return BaselineMode.Baseline3_AI;
            throw new OperationCanceledException();
        }

        // ✅ Updated to match your Baseline2 deterministic PV spec
        private static void ApplyPvDeterministicToRow(Dictionary<string, string> rowDict)
        {
            if (rowDict == null) return;

            double pvUsableM2 = 0.0;
            if (rowDict.TryGetValue("pv_usable_roof_area_m2", out var rawUsable))
                pvUsableM2 = RetrofitShared.ParseDouble(rawUsable);
            pvUsableM2 = Math.Max(0, pvUsableM2);

            // ✅ fixed, per your rule table
            double coverage = 0.60;

            // ✅ fixed, per your table
            const double panelW = 1.494;
            const double panelL = 1.219;
            double panelAreaM2 = panelW * panelL;

            double pvAddedM2 = pvUsableM2 * coverage;

            int panelCount = 0;
            if (pvAddedM2 > 1e-9 && panelAreaM2 > 1e-9)
                panelCount = (int)Math.Floor(pvAddedM2 / panelAreaM2);
            if (panelCount < 0) panelCount = 0;

            rowDict["coverage_factor_used"] = coverage.ToString(CultureInfo.InvariantCulture);
            rowDict["pv_panel_unit_width_m"] = panelW.ToString(CultureInfo.InvariantCulture);
            rowDict["pv_panel_unit_length_m"] = panelL.ToString(CultureInfo.InvariantCulture);
            rowDict["pv_panel_unit_area_m2"] = panelAreaM2.ToString(CultureInfo.InvariantCulture);

            rowDict["ea_pv_added_area_m2"] = pvAddedM2.ToString(CultureInfo.InvariantCulture);
            rowDict["ea_pv_panel_count"] = panelCount.ToString(CultureInfo.InvariantCulture);
            rowDict["ea_pv_layout_status"] = $"usable_x_coverage_{coverage.ToString("0.##", CultureInfo.InvariantCulture)}";
        }

        private static void WritePvDeterministicToRevit(Document doc, List<View> views, List<Element> roofs, Dictionary<string, string> rowDict)
        {
            if (doc == null || rowDict == null) return;

            if (!rowDict.ContainsKey("ea_pv_added_area_m2")) rowDict["ea_pv_added_area_m2"] = "0";
            if (!rowDict.ContainsKey("ea_pv_panel_count")) rowDict["ea_pv_panel_count"] = "0";
            if (!rowDict.ContainsKey("ea_pv_layout_status")) rowDict["ea_pv_layout_status"] = "csv_deterministic";

            RetrofitShared.EnsureSetByRawInstanceOrType(doc, doc.ProjectInformation, "ea_pv_added_area_m2", rowDict["ea_pv_added_area_m2"]);
            RetrofitShared.EnsureSetByRawInstanceOrType(doc, doc.ProjectInformation, "ea_pv_panel_count", rowDict["ea_pv_panel_count"]);
            RetrofitShared.EnsureSetStringInstanceOrType(doc, doc.ProjectInformation, "ea_pv_layout_status", rowDict["ea_pv_layout_status"]);

            if (views != null)
            {
                foreach (var v in views)
                {
                    RetrofitShared.EnsureSetByRawInstanceOrType(doc, v, "ea_pv_added_area_m2", rowDict["ea_pv_added_area_m2"]);
                    RetrofitShared.EnsureSetByRawInstanceOrType(doc, v, "ea_pv_panel_count", rowDict["ea_pv_panel_count"]);
                    RetrofitShared.EnsureSetStringInstanceOrType(doc, v, "ea_pv_layout_status", rowDict["ea_pv_layout_status"]);
                }
            }

            if (roofs != null)
            {
                foreach (var r in roofs)
                {
                    RetrofitShared.EnsureSetByRawInstanceOrType(doc, r, "ea_pv_added_area_m2", rowDict["ea_pv_added_area_m2"]);
                    RetrofitShared.EnsureSetByRawInstanceOrType(doc, r, "ea_pv_panel_count", rowDict["ea_pv_panel_count"]);
                    RetrofitShared.EnsureSetStringInstanceOrType(doc, r, "ea_pv_layout_status", rowDict["ea_pv_layout_status"]);
                }
            }
        }

        private static void LockPlanPvTargetsToRow(Baseline3_DeepSeek.LeedPlan plan, Dictionary<string, string> rowDict)
        {
            if (plan == null || rowDict == null) return;

            if (plan.ea == null) plan.ea = new Baseline3_DeepSeek.EaSection();
            if (plan.ea.targets == null) plan.ea.targets = new Baseline3_DeepSeek.EaTargets();

            double pvAdded = RetrofitShared.ParseDouble(rowDict["ea_pv_added_area_m2"]);
            int pvCount = (int)Math.Round(RetrofitShared.ParseDouble(rowDict["ea_pv_panel_count"]));
            string status = rowDict.ContainsKey("ea_pv_layout_status") ? rowDict["ea_pv_layout_status"] : "csv_deterministic";

            plan.ea.targets.pv_added_area_m2 = Math.Max(0, pvAdded);
            plan.ea.targets.pv_panel_count = Math.Max(0, pvCount);
            plan.ea.targets.pv_layout_status = status;
        }
    }

    internal static class RetrofitShared
    {
        public static string GetForcedBuildingIdFromFileName(string csvPath)
        {
            string name = (Path.GetFileNameWithoutExtension(csvPath) ?? "").ToLowerInvariant();
            if (name.Contains("building1") || name.Contains("0001")) return "building_0001";
            if (name.Contains("building2") || name.Contains("0002")) return "building_0002";
            return null;
        }

        public static string SelectCSVFile()
        {
            var dlg = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "CSV Files (*.csv)|*.csv",
                Title = "Select QGIS CSV Export"
            };
            return dlg.ShowDialog() == true ? dlg.FileName : null;
        }

        public static string SelectReportOutputFolder()
        {
            try
            {
                using (var dlg = new System.Windows.Forms.FolderBrowserDialog())
                {
                    dlg.Description = "Select folder to export report TXT";
                    dlg.ShowNewFolderButton = true;
                    var result = dlg.ShowDialog();
                    return result == System.Windows.Forms.DialogResult.OK ? dlg.SelectedPath : null;
                }
            }
            catch { return null; }
        }

        public static void ShowTimedPopup(string title, string text, int seconds)
        {
            try
            {
                using (var f = new System.Windows.Forms.Form())
                {
                    f.Text = title;
                    f.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
                    f.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
                    f.MaximizeBox = false;
                    f.MinimizeBox = false;
                    f.TopMost = true;
                    f.ShowInTaskbar = false;
                    f.Width = 560;
                    f.Height = 170;

                    var lbl = new System.Windows.Forms.Label();
                    lbl.Text = text;
                    lbl.AutoSize = false;
                    lbl.Dock = System.Windows.Forms.DockStyle.Fill;
                    lbl.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                    lbl.Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Regular);
                    f.Controls.Add(lbl);

                    var timer = new System.Windows.Forms.Timer();
                    timer.Interval = Math.Max(1, seconds) * 1000;
                    timer.Tick += (s, e) =>
                    {
                        timer.Stop();
                        try { f.Close(); } catch { }
                    };
                    timer.Start();

                    f.ShowDialog();
                }
            }
            catch { }
        }

        public static Dictionary<string, string> ReadMatchingRowAsDict(string path, string forcedBuildingId)
        {
            var lines = File.ReadAllLines(path);
            if (lines.Length < 2) return null;

            char d = lines[0].Contains(";") ? ';' : ',';

            var headerRaw = lines[0].Split(d).ToList();
            var header = headerRaw.Select(NormalizeHeaderName).ToList();

            int idxBuilding = header.FindIndex(h => string.Equals(h, "building_id", StringComparison.OrdinalIgnoreCase));
            Dictionary<string, string> firstNonEmpty = null;

            for (int row = 1; row < lines.Length; row++)
            {
                if (string.IsNullOrWhiteSpace(lines[row])) continue;

                var values = lines[row].Split(d);
                var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                bool any = false;
                for (int i = 0; i < header.Count; i++)
                {
                    string key = header[i];
                    if (string.IsNullOrWhiteSpace(key)) continue;

                    string val = i < values.Length ? (values[i] ?? "").Trim() : "";
                    dict[key] = val;
                    if (!string.IsNullOrWhiteSpace(val)) any = true;
                }

                if (!any) continue;
                if (firstNonEmpty == null) firstNonEmpty = dict;

                if (!string.IsNullOrWhiteSpace(forcedBuildingId) && idxBuilding >= 0)
                {
                    if (dict.TryGetValue("building_id", out var bid) &&
                        string.Equals((bid ?? "").Trim(), forcedBuildingId, StringComparison.OrdinalIgnoreCase))
                    {
                        return dict;
                    }
                }
            }

            return firstNonEmpty;
        }

        private static string NormalizeHeaderName(string s)
        {
            if (s == null) return "";
            s = s.Trim();
            if (s.Length > 0 && s[0] == '\uFEFF')
                s = s.Substring(1);
            while (s.Contains("  "))
                s = s.Replace("  ", " ");
            return s.Trim();
        }

        public static void EnsureWriteCsvExact(Document doc, Element e, Dictionary<string, string> row, string exactName)
        {
            try
            {
                if (doc == null || e == null || row == null || string.IsNullOrWhiteSpace(exactName)) return;
                if (!row.TryGetValue(exactName, out string raw)) return;
                if (string.IsNullOrWhiteSpace(raw)) return;

                EnsureSetByRawInstanceOrType(doc, e, exactName, raw);
            }
            catch { }
        }

        public static int WriteWindowCsvForced(Document doc, Element window, Dictionary<string, string> row)
        {
            if (doc == null || window == null || row == null) return 0;
            if (window.Category == null || window.Category.Id.Value != (int)BuiltInCategory.OST_Windows) return 0;

            int c = 0;

            if (row.TryGetValue("ea_window_u_value", out string rawU))
            {
                if (EnsureSetByRawInstanceOrType(doc, window, "ea_window_u_value", rawU))
                    c++;
            }

            if (row.TryGetValue("ea_window_shgc", out string rawShgc))
            {
                if (EnsureSetByRawInstanceOrType(doc, window, "ea_window_shgc", rawShgc))
                    c++;
            }

            return c;
        }

        public static int WriteAllCsvParams(Document doc, Element e, Dictionary<string, string> row)
        {
            int count = 0;
            foreach (var kv in row)
            {
                var ps = GetAllParamsInsensitive(doc, e, kv.Key);
                if (ps == null || ps.Count == 0) continue;

                foreach (var p in ps.Where(x => x != null && !x.IsReadOnly))
                {
                    if (SetParamValue(doc, p, kv.Key, kv.Value))
                    {
                        count++;
                        break;
                    }
                }
            }
            return count;
        }

        public static int WriteCsvParamsByName(Document doc, Element e, Dictionary<string, string> row, IEnumerable<string> names)
        {
            if (doc == null || e == null || row == null || names == null) return 0;

            int count = 0;
            foreach (var name in names)
            {
                if (!row.TryGetValue(name, out string raw)) continue;

                Parameter p = LookupParamInsensitive(e, name);
                if (p == null || p.IsReadOnly) continue;

                if (SetParamValue(doc, p, name, raw))
                    count++;
            }
            return count;
        }

        public static Parameter LookupParamInsensitive(Element e, string name)
        {
            if (e == null || string.IsNullOrWhiteSpace(name)) return null;

            try
            {
                var p0 = e.LookupParameter(name);
                if (p0 != null) return p0;
            }
            catch { }

            try
            {
                foreach (Parameter p in e.Parameters)
                {
                    if (p?.Definition?.Name == null) continue;
                    if (string.Equals(p.Definition.Name, name, StringComparison.OrdinalIgnoreCase))
                        return p;
                }
            }
            catch { }

            return null;
        }

        private static List<Parameter> GetAllParamsInsensitive(Document doc, Element e, string name)
        {
            var list = new List<Parameter>();
            if (doc == null || e == null || string.IsNullOrWhiteSpace(name)) return list;

            try
            {
                var p0 = e.LookupParameter(name);
                if (p0 != null) list.Add(p0);
            }
            catch { }

            try
            {
                foreach (Parameter p in e.Parameters)
                {
                    if (p?.Definition?.Name == null) continue;
                    if (string.Equals(p.Definition.Name, name, StringComparison.OrdinalIgnoreCase))
                        list.Add(p);
                }
            }
            catch { }

            try
            {
                ElementId tid = e.GetTypeId();
                if (tid != ElementId.InvalidElementId)
                {
                    var t = doc.GetElement(tid) as ElementType;
                    if (t != null)
                    {
                        var tp0 = t.LookupParameter(name);
                        if (tp0 != null) list.Add(tp0);

                        foreach (Parameter p in t.Parameters)
                        {
                            if (p?.Definition?.Name == null) continue;
                            if (string.Equals(p.Definition.Name, name, StringComparison.OrdinalIgnoreCase))
                                list.Add(p);
                        }
                    }
                }
            }
            catch { }

            return list.Distinct().ToList();
        }

        private static bool SetParamValue(Document doc, Parameter p, string paramName, string raw)
        {
            if (p == null) return false;
            if (string.IsNullOrWhiteSpace(raw)) return false;

            raw = raw.Trim();

            try
            {
                switch (p.StorageType)
                {
                    case StorageType.Double:
                        {
                            if (string.Equals(paramName, "ea_window_shgc", StringComparison.OrdinalIgnoreCase))
                            {
                                double shgc = NormalizeShgc(raw);
                                if (!double.IsNaN(shgc))
                                {
                                    try
                                    {
                                        if (p.SetValueString(shgc.ToString(CultureInfo.InvariantCulture)))
                                            return true;
                                    }
                                    catch { }

                                    p.Set(shgc);
                                    return true;
                                }
                                return false;
                            }

                            try
                            {
                                if (p.SetValueString(raw))
                                    return true;
                            }
                            catch { }

                            double v = ParseDouble(raw);
                            double internalV = ConvertToInternalIfNeeded(paramName, v);
                            p.Set(internalV);
                            return true;
                        }

                    case StorageType.Integer:
                        {
                            if (bool.TryParse(raw, out bool bb))
                            {
                                p.Set(bb ? 1 : 0);
                                return true;
                            }
                            p.Set((int)Math.Round(ParseDouble(raw)));
                            return true;
                        }
                    case StorageType.String:
                        {
                            p.Set(raw);
                            return true;
                        }
                }
            }
            catch { }

            return false;
        }

        public static double ParseDouble(string s)
        {
            s = (s ?? "").Trim().Replace(",", ".");
            double.TryParse(
                s,
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out double v
            );
            return v;
        }

        private static double NormalizeShgc(string raw)
        {
            string s = CleanRaw(raw);
            if (string.IsNullOrWhiteSpace(s)) return double.NaN;

            s = s.Replace("%", "").Trim();

            double v = ParseDouble(s);
            if (double.IsNaN(v)) return double.NaN;

            if (v > 1.0) v = v / 100.0;

            v = Math.Max(0.0, Math.Min(1.0, v));
            return v;
        }

        public static string CleanRaw(string raw)
        {
            if (raw == null) return "";

            string s = raw.Trim();

            if (s.Length >= 2 && s.StartsWith("\"") && s.EndsWith("\""))
                s = s.Substring(1, s.Length - 2).Trim();

            s = s.Trim().TrimEnd('\"');
            s = s.Replace(",", ".");

            return s;
        }

        public static double ConvertToInternalIfNeeded(string paramName, double metricValue)
        {
            if (string.IsNullOrWhiteSpace(paramName)) return metricValue;
            string n = paramName.Trim().ToLowerInvariant();

            if (n.EndsWith("_m2"))
                return UnitUtils.ConvertToInternalUnits(metricValue, UnitTypeId.SquareMeters);
            if (n.EndsWith("_m3"))
                return UnitUtils.ConvertToInternalUnits(metricValue, UnitTypeId.CubicMeters);
            if (n.EndsWith("_mm"))
                return UnitUtils.ConvertToInternalUnits(metricValue / 1000.0, UnitTypeId.Meters);
            if (n.EndsWith("_m"))
                return UnitUtils.ConvertToInternalUnits(metricValue, UnitTypeId.Meters);
            if (n.EndsWith("_deg"))
                return UnitUtils.ConvertToInternalUnits(metricValue, UnitTypeId.Degrees);

            return metricValue;
        }

        public static bool EnsureSetStringInstanceOrType(Document doc, Element e, string exactName, string value)
        {
            if (doc == null || e == null || string.IsNullOrWhiteSpace(exactName)) return false;

            if (TrySetStringExact(e, exactName, value)) return true;

            ElementType t = GetElementType(doc, e);
            if (t != null && TrySetStringExact(t, exactName, value)) return true;

            return false;
        }

        public static bool EnsureSetDoubleInstanceOrType(Document doc, Element e, string exactName, double valueMetric)
        {
            if (doc == null || e == null || string.IsNullOrWhiteSpace(exactName)) return false;

            if (TrySetDoubleExact(e, exactName, valueMetric)) return true;

            ElementType t = GetElementType(doc, e);
            if (t != null && TrySetDoubleExact(t, exactName, valueMetric)) return true;

            return false;
        }

        public static bool EnsureSetByRawInstanceOrType(Document doc, Element e, string exactName, string raw)
        {
            if (doc == null || e == null || string.IsNullOrWhiteSpace(exactName) || string.IsNullOrWhiteSpace(raw))
                return false;

            if (TrySetByRawExact(doc, e, exactName, raw)) return true;

            ElementType t = GetElementType(doc, e);
            if (t != null && TrySetByRawExact(doc, t, exactName, raw)) return true;

            return false;
        }

        private static ElementType GetElementType(Document doc, Element e)
        {
            try
            {
                ElementId tid = e.GetTypeId();
                if (tid == ElementId.InvalidElementId) return null;
                return doc.GetElement(tid) as ElementType;
            }
            catch { return null; }
        }

        private static bool TrySetStringExact(Element e, string exactName, string value)
        {
            try
            {
                var p = e.LookupParameter(exactName);
                if (p == null || p.IsReadOnly) return false;

                try { if (p.SetValueString(value ?? "")) return true; } catch { }

                if (p.StorageType == StorageType.String)
                {
                    p.Set(value ?? "");
                    return true;
                }
                return false;
            }
            catch { return false; }
        }

        private static bool TrySetDoubleExact(Element e, string exactName, double valueMetric)
        {
            try
            {
                var p = e.LookupParameter(exactName);
                if (p == null || p.IsReadOnly) return false;

                string s = valueMetric.ToString(CultureInfo.InvariantCulture);

                try { if (p.SetValueString(s)) return true; } catch { }

                if (p.StorageType == StorageType.Double)
                {
                    double internalV = ConvertToInternalIfNeeded(exactName, valueMetric);
                    p.Set(internalV);
                    return true;
                }

                if (p.StorageType == StorageType.String)
                {
                    p.Set(s);
                    return true;
                }

                return false;
            }
            catch { return false; }
        }

        private static bool TrySetByRawExact(Document doc, Element e, string exactName, string raw)
        {
            try
            {
                var p = e.LookupParameter(exactName);
                if (p == null || p.IsReadOnly) return false;

                string s = (raw ?? "").Trim();
                if (s.Length >= 2 && s.StartsWith("\"") && s.EndsWith("\""))
                    s = s.Substring(1, s.Length - 2).Trim();
                if (s.EndsWith("\""))
                    s = s.TrimEnd('\"').Trim();
                s = s.Replace(",", ".");

                try { if (p.SetValueString(s)) return true; } catch { }

                if (p.StorageType == StorageType.String)
                {
                    p.Set(s);
                    return true;
                }

                if (string.Equals(exactName, "ea_window_shgc", StringComparison.OrdinalIgnoreCase))
                {
                    double shgc = NormalizeShgc(raw);
                    if (!double.IsNaN(shgc))
                    {
                        try
                        {
                            if (p.SetValueString(shgc.ToString(CultureInfo.InvariantCulture)))
                                return true;
                        }
                        catch { }

                        if (p.StorageType == StorageType.Double)
                        {
                            p.Set(shgc);
                            return true;
                        }

                        if (p.StorageType == StorageType.String)
                        {
                            p.Set(shgc.ToString(CultureInfo.InvariantCulture));
                            return true;
                        }
                    }
                }

                if (p.StorageType == StorageType.Double)
                {
                    double v = ParseDouble(s);
                    double internalV = ConvertToInternalIfNeeded(exactName, v);
                    p.Set(internalV);
                    return true;
                }

                if (p.StorageType == StorageType.Integer)
                {
                    if (bool.TryParse(s, out bool bb))
                    {
                        p.Set(bb ? 1 : 0);
                        return true;
                    }
                    p.Set((int)Math.Round(ParseDouble(s)));
                    return true;
                }

                return false;
            }
            catch { return false; }
        }

        public static void WriteReportToProjectInformation(Document doc, string paramName, string reportText)
        {
            if (doc == null || string.IsNullOrWhiteSpace(paramName)) return;

            Element projInfo = doc.ProjectInformation;
            if (projInfo == null) return;

            Parameter p = LookupParamInsensitive(projInfo, paramName);
            if (p == null || p.IsReadOnly) return;

            try
            {
                if (p.StorageType == StorageType.String)
                    p.Set(reportText ?? "");
            }
            catch { }
        }

        public static bool TryGetUValueFromElementType(Document doc, Element e, out double uValue)
        {
            uValue = double.NaN;
            if (doc == null || e == null) return false;

            ElementType t = GetElementType(doc, e);
            if (t == null) return false;

            string[] candidates = new[]
            {
                "Wärmedurchgangskoeffizient (U)",
                "Waermedurchgangskoeffizient (U)",
                "Wärmedurchgangskoeffizient",
                "Waermedurchgangskoeffizient",
                "U-Wert",
                "U Wert",
                "U-Value",
                "U Value",
                "Thermal Transmittance"
            };

            foreach (var name in candidates)
            {
                Parameter p = null;
                try { p = t.LookupParameter(name) ?? e.LookupParameter(name); } catch { p = null; }
                if (p == null) continue;

                try
                {
                    string vs = "";
                    try { vs = p.AsValueString(); } catch { vs = ""; }

                    if (!string.IsNullOrWhiteSpace(vs))
                    {
                        double v = ParseDouble(StripUnits(vs));
                        if (!double.IsNaN(v) && v >= 0)
                        {
                            uValue = v;
                            return true;
                        }
                    }

                    if (p.StorageType == StorageType.Double)
                    {
                        double internalVal = p.AsDouble();
                        if (!double.IsNaN(internalVal) && internalVal >= 0)
                        {
                            uValue = internalVal;
                            return true;
                        }
                    }

                    if (p.StorageType == StorageType.String)
                    {
                        string s = (p.AsString() ?? "").Trim();
                        double v = ParseDouble(StripUnits(s));
                        if (!double.IsNaN(v) && v >= 0)
                        {
                            uValue = v;
                            return true;
                        }
                    }
                }
                catch { }
            }

            return false;
        }

        private static string StripUnits(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var sb = new StringBuilder();
            foreach (char c in s)
            {
                if (char.IsDigit(c) || c == '.' || c == ',' || c == '-' || c == '+')
                    sb.Append(c);
            }
            return sb.ToString();
        }

        public static Dictionary<ElementId, List<FamilyInstance>> BuildWallToWindowsMap(Document doc, List<Element> allWindows)
        {
            var map = new Dictionary<ElementId, List<FamilyInstance>>();
            foreach (var e in allWindows)
            {
                var fi = e as FamilyInstance;
                if (fi == null) continue;

                Element host = null;
                try { host = fi.Host; } catch { host = null; }
                var wall = host as Wall;
                if (wall == null) continue;

                if (!map.TryGetValue(wall.Id, out var list))
                {
                    list = new List<FamilyInstance>();
                    map[wall.Id] = list;
                }
                list.Add(fi);
            }
            return map;
        }

        public static bool IsExteriorWall(Wall wall)
        {
            try
            {
                var wt = wall?.WallType;
                if (wt == null) return false;
                return wt.Function == WallFunction.Exterior;
            }
            catch { return false; }
        }

        public static bool TryGetWallAzimuthDeg(Wall wall, out double azDeg)
        {
            azDeg = double.NaN;
            if (wall == null) return false;

            try
            {
                var lc = wall.Location as LocationCurve;
                if (lc?.Curve == null) return false;

                XYZ dir = (lc.Curve.GetEndPoint(1) - lc.Curve.GetEndPoint(0));
                dir = new XYZ(dir.X, dir.Y, 0);
                if (dir.GetLength() <= 1e-9) return false;
                dir = dir.Normalize();

                double angle = Math.Atan2(dir.X, dir.Y); // atan2(East, North)
                double deg = angle * 180.0 / Math.PI;
                if (deg < 0) deg += 360.0;

                azDeg = deg;
                return true;
            }
            catch { return false; }
        }

        public static double TryGetWallGrossAreaM2(Wall wall)
        {
            try
            {
                var lc = wall.Location as LocationCurve;
                if (lc?.Curve == null) return 0;

                double lenFt = lc.Curve.Length;

                double hFt = 0;
                var ph = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (ph != null && ph.StorageType == StorageType.Double)
                    hFt = ph.AsDouble();

                if (hFt <= 1e-9)
                {
                    var ph2 = wall.LookupParameter("Unconnected Height") ?? wall.LookupParameter("Nicht verknüpfte Höhe");
                    if (ph2 != null && ph2.StorageType == StorageType.Double)
                        hFt = ph2.AsDouble();
                }

                if (lenFt <= 1e-9 || hFt <= 1e-9) return 0;

                double lenM = UnitUtils.ConvertFromInternalUnits(lenFt, UnitTypeId.Meters);
                double hM = UnitUtils.ConvertFromInternalUnits(hFt, UnitTypeId.Meters);

                return lenM * hM;
            }
            catch { return 0; }
        }

        public static double TryGetWindowAreaM2(FamilyInstance fi)
        {
            if (fi == null) return 0.0;
            double best = 0.0;

            try
            {
                double wFt = 0, hFt = 0;
                var pw = fi.get_Parameter(BuiltInParameter.FAMILY_ROUGH_WIDTH_PARAM);
                var ph = fi.get_Parameter(BuiltInParameter.FAMILY_ROUGH_HEIGHT_PARAM);

                if (pw != null && pw.StorageType == StorageType.Double) wFt = pw.AsDouble();
                if (ph != null && ph.StorageType == StorageType.Double) hFt = ph.AsDouble();

                if (wFt > 1e-9 && hFt > 1e-9)
                {
                    double aFt2 = wFt * hFt;
                    double aM2 = UnitUtils.ConvertFromInternalUnits(aFt2, UnitTypeId.SquareMeters);
                    best = Math.Max(best, aM2);
                }
            }
            catch { }

            try
            {
                var p = fi.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                if (p != null && p.StorageType == StorageType.Double)
                {
                    double internalVal = p.AsDouble();
                    double aM2 = UnitUtils.ConvertFromInternalUnits(internalVal, UnitTypeId.SquareMeters);
                    best = Math.Max(best, aM2);
                }
            }
            catch { }

            return best;
        }

        public static string ComputeSriTargetFromHeatIsland(Dictionary<string, string> row)
        {
            if (row == null) return "";
            if (!row.TryGetValue("ss_heat_island", out string raw)) return "";
            if (string.IsNullOrWhiteSpace(raw)) return "";

            double v = ParseDouble(raw);
            if (double.IsNaN(v)) return "";

            if (v < 0.33) return "Low";
            if (v <= 0.66) return "Medium";
            return "High";
        }
    }
}