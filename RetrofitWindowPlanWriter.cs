using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;

namespace RevitLOD3Exporter
{
    internal static class RetrofitWindowPlanWriter
    {
        // =========================
        // ✅ PERF: static HttpClient
        // =========================
        private static readonly HttpClient _http = CreateHttpClient();
        private static string _lastBearerKey = null;

        private static HttpClient CreateHttpClient()
        {
            var http = new HttpClient();
            http.Timeout = TimeSpan.FromSeconds(60);
            return http;
        }

        private static void EnsureAuthHeader(string apiKey)
        {
            if (!string.Equals(_lastBearerKey, apiKey, StringComparison.Ordinal))
            {
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                _lastBearerKey = apiKey;
            }
        }

        public static void GenerateAndWritePlans(Document doc)
        {
            ExportOpeningsReportTxt(doc); // ✅ now includes WINDOWS + DOORS

        }
        // ✅ Backward-compatible entry point for older callers
        public static void ExportWindowsReportTxt(Document doc)
        {
            GenerateAndWritePlans(doc);
        }

        // ====== Key params for WINDOWS ======
        private static readonly string[] WindowKeyParamNames = new[]
        {
            "ea_score","lt_score","we_score","ss_score","retrofit_score_total",
            "ea_window_u_value","lst_mean_celsius",
            "ea_shading_count","ea_shading_depth_m","ea_shading_status","shadow_ratio_mean",
            "lt_vlt","lt_glare_risk_class",
            "ea_window_upgrade_strategy","ea_window_shading_strategy"
        };

        // ====== Key params for DOORS (you can expand later) ======
        private static readonly string[] DoorKeyParamNames = new[]
        {
            "ea_score","lt_score","we_score","ss_score","retrofit_score_total",
            "ea_door_u_value","lst_mean_celsius",
            // strategy params (AI will output these; your other code can write them back if you want)
            "ea_door_upgrade_strategy",
            "ea_door_air_tightness_strategy",
            "ea_door_glazing_strategy"
        };

        // Still collect all prefixed params (future use)
        private static readonly string[] ParamPrefixes = new[] { "ea_", "we_", "ss_", "lt_" };

        // ============================================================
        // ✅ ONE REPORT: WINDOWS + DOORS
        // ============================================================
        public static void ExportOpeningsReportTxt(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var windows = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Windows)
                .WhereElementIsNotElementType()
                .ToElements();

            var doors = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Doors)
                .WhereElementIsNotElementType()
                .ToElements();

            if ((windows == null || windows.Count == 0) && (doors == null || doors.Count == 0))
            {
                TaskDialog.Show("Openings", "No window/door instances found.");
                return;
            }

            // ✅ default output path
            string defaultOut = GetDefaultOutputPath(doc);

            string outputPath;
            if (!PickOutputPath(defaultOut, out outputPath))
                return;

            string externalContext = "";

            // ✅ Type summary: Family | Type -> Count
            var windowTypeSummary = (windows ?? new List<Element>())
                .Select(e =>
                {
                    var symbol = doc.GetElement(e.GetTypeId()) as FamilySymbol;
                    string fam = symbol?.FamilyName ?? "UnknownFamily";
                    string type = symbol?.Name ?? "UnknownType";
                    return $"{fam} | {type}";
                })
                .GroupBy(x => x)
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            var doorTypeSummary = (doors ?? new List<Element>())
                .Select(e =>
                {
                    var symbol = doc.GetElement(e.GetTypeId()) as FamilySymbol;
                    string fam = symbol?.FamilyName ?? "UnknownFamily";
                    string type = symbol?.Name ?? "UnknownType";
                    return $"{fam} | {type}";
                })
                .GroupBy(x => x)
                .OrderByDescending(g => g.Count())
                .ThenBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            int okW = 0, failW = 0;
            int okD = 0, failD = 0;

            // ============================================================
            // ✅ PERF: advice caches (avoid calling API for each instance)
            // ============================================================
            var windowAdviceCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var doorAdviceCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var sb = new StringBuilder();
            sb.AppendLine("OPENINGS RETROFIT REPORT (TXT EXPORT)");
            sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"Model: {doc.Title}");
            sb.AppendLine($"Window instance count: {(windows?.Count ?? 0)}");
            sb.AppendLine($"Door instance count: {(doors?.Count ?? 0)}");
            sb.AppendLine();

            // ======================
            // WINDOWS TYPE SUMMARY
            // ======================
            sb.AppendLine("WINDOW TYPE SUMMARY (Family | Type -> Count)");
            if (windowTypeSummary.Count == 0) sb.AppendLine("- (none)");
            foreach (var g in windowTypeSummary)
                sb.AppendLine($"- {g.Key} : {g.Count()}");
            sb.AppendLine();

            // ======================
            // DOORS TYPE SUMMARY
            // ======================
            sb.AppendLine("DOOR TYPE SUMMARY (Family | Type -> Count)");
            if (doorTypeSummary.Count == 0) sb.AppendLine("- (none)");
            foreach (var g in doorTypeSummary)
                sb.AppendLine($"- {g.Key} : {g.Count()}");
            sb.AppendLine();

            sb.AppendLine(new string('=', 90));
            sb.AppendLine();

            // ✅ Print "framework" ONCE
            AppendGlobalWindowFramework(sb);
            sb.AppendLine(new string('-', 90));
            AppendGlobalDoorFramework(sb);

            sb.AppendLine(new string('=', 90));
            sb.AppendLine();

            // ======================
            // WINDOWS SECTION
            // ======================
            sb.AppendLine("SECTION 1) WINDOWS");
            sb.AppendLine(new string('=', 90));
            sb.AppendLine();

            foreach (var e in (windows ?? new List<Element>()))
            {
                try
                {
                    var info = ExtractWindowFullInfo(doc, e);

                    string key = BuildWindowAdviceKey(info);

                    string advice;
                    if (!windowAdviceCache.TryGetValue(key, out advice))
                    {
                        string prompt = BuildWindowPrompt(info, externalContext);
                        advice = CallRetrofitApi(prompt);
                        windowAdviceCache[key] = advice;
                    }

                    AppendWindowBlock(sb, info, advice);
                    okW++;
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"[Window] ElementId: {SafeElementId(e)} | UniqueId: {SafeUniqueId(e)}");
                    sb.AppendLine("AI Retrofit Advice:");
                    sb.AppendLine("(failed to generate advice) " + ex.Message);
                    sb.AppendLine(new string('-', 90));
                    sb.AppendLine();

                    failW++;
                }
            }

            // ======================
            // DOORS SECTION
            // ======================
            sb.AppendLine();
            sb.AppendLine("SECTION 2) DOORS");
            sb.AppendLine(new string('=', 90));
            sb.AppendLine();

            foreach (var e in (doors ?? new List<Element>()))
            {
                try
                {
                    var info = ExtractDoorFullInfo(doc, e);

                    string key = BuildDoorAdviceKey(info);

                    string advice;
                    if (!doorAdviceCache.TryGetValue(key, out advice))
                    {
                        string prompt = BuildDoorPrompt(info, externalContext);
                        advice = CallRetrofitApi(prompt);
                        doorAdviceCache[key] = advice;
                    }

                    AppendDoorBlock(sb, info, advice);
                    okD++;
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"[Door] ElementId: {SafeElementId(e)} | UniqueId: {SafeUniqueId(e)}");
                    sb.AppendLine("AI Retrofit Advice:");
                    sb.AppendLine("(failed to generate advice) " + ex.Message);
                    sb.AppendLine(new string('-', 90));
                    sb.AppendLine();

                    failD++;
                }
            }

            // ======================
            // SUMMARY
            // ======================
            sb.AppendLine();
            sb.AppendLine(new string('=', 90));
            sb.AppendLine("SUMMARY");
            sb.AppendLine($"Windows: Success {okW}, Failed {failW}, Unique advice cached: {windowAdviceCache.Count}");
            sb.AppendLine($"Doors:   Success {okD}, Failed {failD}, Unique advice cached: {doorAdviceCache.Count}");
            sb.AppendLine(new string('=', 90));

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);

            TaskDialog.Show(
                "Openings",
                $"TXT exported.\n\n" +
                $"Windows -> Success: {okW}, Failed: {failW}, Cached: {windowAdviceCache.Count}\n" +
                $"Doors   -> Success: {okD}, Failed: {failD}, Cached: {doorAdviceCache.Count}\n\n" +
                $"Path:\n{outputPath}"
            );
        }

        // ✅ Print once: parameter framework + LEED linkage + assumptions
        private static void AppendGlobalWindowFramework(StringBuilder sb)
        {
            sb.AppendLine("A) WINDOW ANALYSIS FRAMEWORK");
            sb.AppendLine();
            sb.AppendLine("LEED Linkage (Window Retrofit Strategies):");
            sb.AppendLine("    - EA: glazing U-value, SHGC control, shading -> lower heating/cooling loads");
            sb.AppendLine("    - LT: daylight / glare control (depends on VLT + shading/blinds)");
            sb.AppendLine("    - Climate zone / weather file not provided (affects U/SHGC targets and savings).");
            sb.AppendLine("    - SHGC (g-value) not provided (critical for solar gain control).");
            sb.AppendLine("    - Airtightness condition assumed from model; on-site verification may differ.");
            sb.AppendLine();
        }

        private static void AppendGlobalDoorFramework(StringBuilder sb)
        {
            sb.AppendLine("B) DOOR ANALYSIS FRAMEWORK");
            sb.AppendLine();
            sb.AppendLine("LEED Linkage (Door Retrofit Strategies):");
            sb.AppendLine("    - EA: door leaf U-value + frame thermal bridge + air leakage sealing -> heating/cooling reduction");
            sb.AppendLine("    - LT: only relevant for glazed doors / vision panels (glare + light) — otherwise minor");
            sb.AppendLine("    - Door air leakage class not provided; sealing effectiveness depends on installation quality.");
            sb.AppendLine("    - Threshold / weatherstrip condition not modeled reliably in BIM.");
            sb.AppendLine();
        }

        // =========================
        // Output path helpers
        // =========================
        private static bool PickOutputPath(string defaultOutputPath, out string outputPath)
        {
            outputPath = defaultOutputPath;

            try
            {
                var sfd = new FileSaveDialog("Text file (*.txt)|*.txt");
                sfd.Title = "Select output TXT path";
                sfd.InitialFileName = Path.GetFileName(defaultOutputPath);

                if (sfd.Show() != ItemSelectionDialogResult.Confirmed)
                    return false;

                var mp = sfd.GetSelectedModelPath();
                outputPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp);

                if (!outputPath.EndsWith(".txt", StringComparison.OrdinalIgnoreCase))
                    outputPath += ".txt";

                return true;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Export Error", ex.ToString());
                return false;
            }
        }

        private static string GetDefaultOutputPath(Document doc)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string ts = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string safeTitle = SanitizeFileName(doc?.Title ?? "RevitModel");
            return Path.Combine(desktop, $"Openings_RetrofitReport_{safeTitle}_{ts}.txt");
        }

        private static string SanitizeFileName(string s)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            return s;
        }

        // ============================================================
        // WINDOW info model + extract + key + prompt + block
        // ============================================================
        private class WindowFullInfo
        {
            public long ElementId { get; set; }
            public string UniqueId { get; set; }

            public string FamilyName { get; set; }
            public string TypeName { get; set; }

            public string Width { get; set; }
            public string Height { get; set; }

            public string LevelName { get; set; }
            public string HostName { get; set; }

            public string LocationPointMeters { get; set; }
            public string FacingOrientation { get; set; }
            public string HandOrientation { get; set; }

            public string SillHeight { get; set; }
            public string HeadHeight { get; set; }

            public string FromRoom { get; set; }
            public string ToRoom { get; set; }

            public List<string> Materials { get; set; } = new List<string>();
            public Dictionary<string, string> Params { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private static WindowFullInfo ExtractWindowFullInfo(Document doc, Element windowInstance)
        {
            var info = new WindowFullInfo
            {
                ElementId = windowInstance?.Id?.Value ?? -1L,
                UniqueId = windowInstance?.UniqueId ?? ""
            };

            var symbol = doc.GetElement(windowInstance.GetTypeId()) as FamilySymbol;
            info.FamilyName = symbol?.FamilyName ?? "";
            info.TypeName = symbol?.Name ?? "";

            info.Width = GetAsString(windowInstance, BuiltInParameter.WINDOW_WIDTH);
            info.Height = GetAsString(windowInstance, BuiltInParameter.WINDOW_HEIGHT);

            info.LevelName = GetLevelName(doc, windowInstance);
            info.HostName = GetHostName(windowInstance);
            info.LocationPointMeters = GetLocationStringMeters(windowInstance);

            var fi = windowInstance as FamilyInstance;
            info.FacingOrientation = ToXYZString(fi?.FacingOrientation);
            info.HandOrientation = ToXYZString(fi?.HandOrientation);

            info.SillHeight = GetAsString(windowInstance, BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);

            info.HeadHeight = GetParamFromInstanceOrType(windowInstance, symbol, "Head Height");
            if (string.IsNullOrWhiteSpace(info.HeadHeight))
                info.HeadHeight = GetParamFromInstanceOrType(windowInstance, symbol, "Sturzhöhe");

            info.FromRoom = SafeRoomToString(fi?.FromRoom);
            info.ToRoom = SafeRoomToString(fi?.ToRoom);

            info.Materials = GetMaterialNames(doc, windowInstance, symbol);

            foreach (var pn in WindowKeyParamNames)
            {
                var v = GetParamFromInstanceOrType(windowInstance, symbol, pn);
                if (!string.IsNullOrWhiteSpace(v))
                    info.Params[pn] = v.Trim();
            }

            var fuzzyUpgrade = FindParamByPrefix(windowInstance, symbol, "ea_window_upgrade");
            if (fuzzyUpgrade.HasValue && !string.IsNullOrWhiteSpace(fuzzyUpgrade.Value.Value))
                info.Params[fuzzyUpgrade.Value.Key] = fuzzyUpgrade.Value.Value.Trim();

            CollectByPrefixes(info.Params, windowInstance, symbol, ParamPrefixes);

            return info;
        }

        private static string BuildWindowAdviceKey(WindowFullInfo w)
        {
            string GetP(string name)
            {
                if (w?.Params == null) return "";
                if (!w.Params.TryGetValue(name, out var v)) return "";
                return (v ?? "").Trim();
            }

            return "WINDOW|" + string.Join("|", new[]
            {
                (w.FamilyName ?? "").Trim(),
                (w.TypeName ?? "").Trim(),
                GetP("ea_window_u_value"),
                GetP("ea_window_upgrade_strategy"),
                GetP("ea_window_shading_strategy"),
                GetP("ea_shading_status"),
                GetP("ea_shading_depth_m"),
                GetP("ea_shading_count"),
                GetP("shadow_ratio_mean"),
                GetP("lt_glare_risk_class"),
                GetP("lt_vlt"),
                GetP("ea_score"),
                GetP("lt_score"),
                GetP("we_score"),
                GetP("ss_score")
            });
        }

        private static void AppendWindowBlock(StringBuilder sb, WindowFullInfo w, string advice)
        {
            sb.AppendLine($"[Window] ElementId: {w.ElementId} | UniqueId: {w.UniqueId}");
            sb.AppendLine($"Family: {w.FamilyName} | Type: {w.TypeName}");
            sb.AppendLine($"Level: {w.LevelName} | Host: {w.HostName}");
            sb.AppendLine($"Width: {w.Width} | Height: {w.Height}");
            if (!string.IsNullOrWhiteSpace(w.SillHeight)) sb.AppendLine($"SillHeight: {w.SillHeight}");
            if (!string.IsNullOrWhiteSpace(w.HeadHeight)) sb.AppendLine($"HeadHeight: {w.HeadHeight}");
            if (!string.IsNullOrWhiteSpace(w.FromRoom)) sb.AppendLine($"FromRoom: {w.FromRoom}");
            if (!string.IsNullOrWhiteSpace(w.ToRoom)) sb.AppendLine($"ToRoom: {w.ToRoom}");
            sb.AppendLine($"Location: {(string.IsNullOrWhiteSpace(w.LocationPointMeters) ? "(unknown)" : w.LocationPointMeters)}");
            sb.AppendLine($"FacingOrientation: {w.FacingOrientation}");
            if (!string.IsNullOrWhiteSpace(w.HandOrientation)) sb.AppendLine($"HandOrientation: {w.HandOrientation}");
            sb.AppendLine($"Materials: {(w.Materials.Count > 0 ? string.Join(", ", w.Materials) : "Unknown")}");

            sb.AppendLine();
            sb.AppendLine("Key Metrics (compact):");
            AppendMetricLine(sb, w.Params, "retrofit_score_total", "retrofit_score_total");
            AppendMetricLine(sb, w.Params, "ea_score", "ea_score");
            AppendMetricLine(sb, w.Params, "lt_score", "lt_score");
            AppendMetricLine(sb, w.Params, "ss_score", "ss_score");
            AppendMetricLine(sb, w.Params, "we_score", "we_score");
            AppendMetricLine(sb, w.Params, "ea_window_u_value", "ea_window_u_value");
            AppendMetricLine(sb, w.Params, "ea_window_upgrade_strategy", "ea_window_upgrade_strategy");
            AppendMetricLine(sb, w.Params, "ea_window_shading_strategy", "ea_window_shading_strategy");
            AppendMetricLine(sb, w.Params, "shadow_ratio_mean", "shadow_ratio_mean");
            AppendMetricLine(sb, w.Params, "lt_vlt", "lt_vlt");
            AppendMetricLine(sb, w.Params, "lt_glare_risk_class", "lt_glare_risk_class");
            AppendMetricLine(sb, w.Params, "lst_mean_celsius", "lst_mean_celsius");

            sb.AppendLine();
            sb.AppendLine("AI Retrofit Advice (no repeated LEED/assumptions):");
            sb.AppendLine(string.IsNullOrWhiteSpace(advice) ? "(no advice)" : advice.Trim());
            sb.AppendLine(new string('-', 90));
            sb.AppendLine();
        }

        private static string BuildWindowPrompt(WindowFullInfo w, string externalContext)
        {
            string mats = (w.Materials != null && w.Materials.Count > 0) ? string.Join(", ", w.Materials) : "Unknown";

            string GetP(string name)
            {
                if (w?.Params == null) return "";
                if (!w.Params.TryGetValue(name, out var v)) return "";
                return (v ?? "").Trim();
            }

            string compactMetrics =
$@"- retrofit_score_total: {GetP("retrofit_score_total")}
- ea_score: {GetP("ea_score")}
- lt_score: {GetP("lt_score")}
- ss_score: {GetP("ss_score")}
- we_score: {GetP("we_score")}
- ea_window_u_value: {GetP("ea_window_u_value")}
- ea_window_upgrade_strategy: {GetP("ea_window_upgrade_strategy")}
- ea_window_shading_strategy: {GetP("ea_window_shading_strategy")}
- shadow_ratio_mean: {GetP("shadow_ratio_mean")}
- lt_vlt: {GetP("lt_vlt")}
- lt_glare_risk_class: {GetP("lt_glare_risk_class")}
- lst_mean_celsius: {GetP("lst_mean_celsius")}";

            string ext = string.IsNullOrWhiteSpace(externalContext) ? "" : ("\nEXTERNAL CONTEXT:\n" + externalContext + "\n");

            return
$@"You are a LEED-oriented retrofit assistant.
Task: Provide retrofit advice for ONE WINDOW instance using window info + compact metrics.
Rules:
- TEXT ONLY.
- Do NOT output generic sections like 'LEED Linkage' or 'Assumptions/Missing Data' (they are reported globally).
- Do NOT instruct direct Revit geometry/type edits; focus on practical retrofit actions/spec changes.
- Keep output concise and actionable.
{ext}

WINDOW OWN INFO:
- ElementId: {w.ElementId}
- Family: {w.FamilyName}
- Type: {w.TypeName}
- Width: {w.Width}
- Height: {w.Height}
- Level: {w.LevelName}
- Host: {w.HostName}
- SillHeight: {w.SillHeight}
- HeadHeight: {w.HeadHeight}
- Location: {w.LocationPointMeters}
- FacingOrientation: {w.FacingOrientation}
- Materials: {mats}

COMPACT METRICS:
{compactMetrics}

OUTPUT FORMAT (ONLY):
1) Diagnosis (1-2 lines)
2) Retrofit measures (5-8 bullets) – prioritize based on U-value, shading, glare risk and scores
3) Instance-specific notes (2-4 bullets) – only what depends on this window’s context/orientation/room
";
        }

        // ============================================================
        // DOOR info model + extract + key + prompt + block
        // ============================================================
        private class DoorFullInfo
        {
            public long ElementId { get; set; }
            public string UniqueId { get; set; }

            public string FamilyName { get; set; }
            public string TypeName { get; set; }

            public string Width { get; set; }
            public string Height { get; set; }

            public string LevelName { get; set; }
            public string HostName { get; set; }

            public string LocationPointMeters { get; set; }
            public string FacingOrientation { get; set; }
            public string HandOrientation { get; set; }

            public string FromRoom { get; set; }
            public string ToRoom { get; set; }

            public List<string> Materials { get; set; } = new List<string>();
            public Dictionary<string, string> Params { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private static DoorFullInfo ExtractDoorFullInfo(Document doc, Element doorInstance)
        {
            var info = new DoorFullInfo
            {
                ElementId = doorInstance?.Id?.Value ?? -1L,
                UniqueId = doorInstance?.UniqueId ?? ""
            };

            var symbol = doc.GetElement(doorInstance.GetTypeId()) as FamilySymbol;
            info.FamilyName = symbol?.FamilyName ?? "";
            info.TypeName = symbol?.Name ?? "";

            // doors: width/height built-in
            info.Width = GetAsString(doorInstance, BuiltInParameter.DOOR_WIDTH);
            info.Height = GetAsString(doorInstance, BuiltInParameter.DOOR_HEIGHT);

            info.LevelName = GetLevelName(doc, doorInstance);
            info.HostName = GetHostName(doorInstance);
            info.LocationPointMeters = GetLocationStringMeters(doorInstance);

            var fi = doorInstance as FamilyInstance;
            info.FacingOrientation = ToXYZString(fi?.FacingOrientation);
            info.HandOrientation = ToXYZString(fi?.HandOrientation);

            info.FromRoom = SafeRoomToString(fi?.FromRoom);
            info.ToRoom = SafeRoomToString(fi?.ToRoom);

            info.Materials = GetMaterialNames(doc, doorInstance, symbol);

            foreach (var pn in DoorKeyParamNames)
            {
                var v = GetParamFromInstanceOrType(doorInstance, symbol, pn);
                if (!string.IsNullOrWhiteSpace(v))
                    info.Params[pn] = v.Trim();
            }

            // fuzzy pick-ups if your params are named with prefix
            var fuzzyUpgrade = FindParamByPrefix(doorInstance, symbol, "ea_door_upgrade");
            if (fuzzyUpgrade.HasValue && !string.IsNullOrWhiteSpace(fuzzyUpgrade.Value.Value))
                info.Params[fuzzyUpgrade.Value.Key] = fuzzyUpgrade.Value.Value.Trim();

            var fuzzyTight = FindParamByPrefix(doorInstance, symbol, "ea_door_air");
            if (fuzzyTight.HasValue && !string.IsNullOrWhiteSpace(fuzzyTight.Value.Value))
                info.Params[fuzzyTight.Value.Key] = fuzzyTight.Value.Value.Trim();

            CollectByPrefixes(info.Params, doorInstance, symbol, ParamPrefixes);

            return info;
        }

        private static string BuildDoorAdviceKey(DoorFullInfo d)
        {
            string GetP(string name)
            {
                if (d?.Params == null) return "";
                if (!d.Params.TryGetValue(name, out var v)) return "";
                return (v ?? "").Trim();
            }

            return "DOOR|" + string.Join("|", new[]
            {
                (d.FamilyName ?? "").Trim(),
                (d.TypeName ?? "").Trim(),
                GetP("ea_door_u_value"),
                GetP("ea_door_upgrade_strategy"),
                GetP("ea_door_air_tightness_strategy"),
                GetP("ea_score"),
                GetP("lt_score"),
                GetP("we_score"),
                GetP("ss_score"),
                GetP("lst_mean_celsius")
            });
        }

        private static void AppendDoorBlock(StringBuilder sb, DoorFullInfo d, string advice)
        {
            sb.AppendLine($"[Door] ElementId: {d.ElementId} | UniqueId: {d.UniqueId}");
            sb.AppendLine($"Family: {d.FamilyName} | Type: {d.TypeName}");
            sb.AppendLine($"Level: {d.LevelName} | Host: {d.HostName}");
            sb.AppendLine($"Width: {d.Width} | Height: {d.Height}");
            if (!string.IsNullOrWhiteSpace(d.FromRoom)) sb.AppendLine($"FromRoom: {d.FromRoom}");
            if (!string.IsNullOrWhiteSpace(d.ToRoom)) sb.AppendLine($"ToRoom: {d.ToRoom}");
            sb.AppendLine($"Location: {(string.IsNullOrWhiteSpace(d.LocationPointMeters) ? "(unknown)" : d.LocationPointMeters)}");
            sb.AppendLine($"FacingOrientation: {d.FacingOrientation}");
            if (!string.IsNullOrWhiteSpace(d.HandOrientation)) sb.AppendLine($"HandOrientation: {d.HandOrientation}");
            sb.AppendLine($"Materials: {(d.Materials.Count > 0 ? string.Join(", ", d.Materials) : "Unknown")}");

            sb.AppendLine();
            sb.AppendLine("Key Metrics (compact):");
            AppendMetricLine(sb, d.Params, "retrofit_score_total", "retrofit_score_total");
            AppendMetricLine(sb, d.Params, "ea_score", "ea_score");
            AppendMetricLine(sb, d.Params, "lt_score", "lt_score");
            AppendMetricLine(sb, d.Params, "ss_score", "ss_score");
            AppendMetricLine(sb, d.Params, "we_score", "we_score");
            AppendMetricLine(sb, d.Params, "ea_door_u_value", "ea_door_u_value");
            AppendMetricLine(sb, d.Params, "ea_door_upgrade_strategy", "ea_door_upgrade_strategy");
            AppendMetricLine(sb, d.Params, "ea_door_air_tightness_strategy", "ea_door_air_tightness_strategy");
            AppendMetricLine(sb, d.Params, "ea_door_glazing_strategy", "ea_door_glazing_strategy");
            AppendMetricLine(sb, d.Params, "lst_mean_celsius", "lst_mean_celsius");

            sb.AppendLine();
            sb.AppendLine("AI Retrofit Advice (no repeated LEED/assumptions):");
            sb.AppendLine(string.IsNullOrWhiteSpace(advice) ? "(no advice)" : advice.Trim());
            sb.AppendLine(new string('-', 90));
            sb.AppendLine();
        }

        private static string BuildDoorPrompt(DoorFullInfo d, string externalContext)
        {
            string mats = (d.Materials != null && d.Materials.Count > 0) ? string.Join(", ", d.Materials) : "Unknown";

            string GetP(string name)
            {
                if (d?.Params == null) return "";
                if (!d.Params.TryGetValue(name, out var v)) return "";
                return (v ?? "").Trim();
            }

            string compactMetrics =
$@"- retrofit_score_total: {GetP("retrofit_score_total")}
- ea_score: {GetP("ea_score")}
- lt_score: {GetP("lt_score")}
- ss_score: {GetP("ss_score")}
- we_score: {GetP("we_score")}
- ea_door_u_value: {GetP("ea_door_u_value")}
- lst_mean_celsius: {GetP("lst_mean_celsius")}";

            string ext = string.IsNullOrWhiteSpace(externalContext) ? "" : ("\nEXTERNAL CONTEXT:\n" + externalContext + "\n");

            return
$@"You are a LEED-oriented retrofit assistant.
Task: Provide retrofit advice for ONE DOOR instance using door info + compact metrics.
Rules:
- TEXT ONLY.
- Do NOT output generic sections like 'LEED Linkage' or 'Assumptions/Missing Data' (they are reported globally).
- Do NOT instruct direct Revit geometry/type edits; focus on practical retrofit actions/spec changes.
- Keep output concise and actionable.
{ext}

DOOR OWN INFO:
- ElementId: {d.ElementId}
- Family: {d.FamilyName}
- Type: {d.TypeName}
- Width: {d.Width}
- Height: {d.Height}
- Level: {d.LevelName}
- Host: {d.HostName}
- Location: {d.LocationPointMeters}
- FacingOrientation: {d.FacingOrientation}
- Materials: {mats}
- FromRoom: {d.FromRoom}
- ToRoom: {d.ToRoom}

COMPACT METRICS:
{compactMetrics}

OUTPUT FORMAT (ONLY):
1) Diagnosis (1-2 lines)
2) Retrofit measures (5-8 bullets) – prioritize U-value + air leakage sealing + frame/threshold improvements
3) Instance-specific notes (2-4 bullets) – room-to-room usage, orientation, exterior exposure, safety/operation constraints
4) Strategy fields (3-5 lines) – output these keys with concise values:
   - ea_door_upgrade_strategy:
   - ea_door_air_tightness_strategy:
   - ea_door_glazing_strategy:
";
        }

        // ============================================================
        // ✅ SINGLE API CALL FUNCTION (for both windows & doors)
        // ============================================================
        private static string CallRetrofitApi(string prompt)
        {
            string apiKey =
                Environment.GetEnvironmentVariable("DEEPSEEK_API_KEY") ??
                Environment.GetEnvironmentVariable("OPENAI_API_KEY");

            if (string.IsNullOrWhiteSpace(apiKey))
                return "(API key missing: set env var DEEPSEEK_API_KEY or OPENAI_API_KEY)";

            string url =
                Environment.GetEnvironmentVariable("RETROFIT_API_URL") ??
                "https://api.deepseek.com/v1/chat/completions";

            string model =
                Environment.GetEnvironmentVariable("RETROFIT_API_MODEL") ??
                "deepseek-chat";

            try
            {
                EnsureAuthHeader(apiKey);

                var payload = new ChatCompletionsRequest
                {
                    model = model,
                    temperature = 0.2,
                    messages = new List<ChatMessage>
                    {
                        new ChatMessage { role = "system", content = "You are a LEED-oriented retrofit assistant." },
                        new ChatMessage { role = "user", content = prompt }
                    }
                };

                string json = JsonConvert.SerializeObject(payload);
                using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
                {
                    var resp = _http.PostAsync(url, content).GetAwaiter().GetResult();
                    var respText = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

                    if (!resp.IsSuccessStatusCode)
                        return $"(API error {(int)resp.StatusCode}): {TrimTo(respText, 1000)}";

                    var parsed = JsonConvert.DeserializeObject<ChatCompletionsResponse>(respText);
                    var advice = parsed?.choices?.FirstOrDefault()?.message?.content;

                    return string.IsNullOrWhiteSpace(advice) ? "(API returned empty advice)" : advice.Trim();
                }
            }
            catch (Exception ex)
            {
                return "(API call exception): " + ex.Message;
            }
        }

        private class ChatCompletionsRequest
        {
            public string model { get; set; }
            public double temperature { get; set; }
            public List<ChatMessage> messages { get; set; }
        }

        private class ChatMessage
        {
            public string role { get; set; }     // system/user/assistant
            public string content { get; set; }
        }

        private class ChatCompletionsResponse
        {
            public List<Choice> choices { get; set; }
        }

        private class Choice
        {
            public ChatMessage message { get; set; }
        }

        private static string TrimTo(string s, int maxLen)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            s = s.Trim();
            return s.Length <= maxLen ? s : s.Substring(0, maxLen) + "...";
        }

        // ============================================================
        // Shared extract helpers (unchanged)
        // ============================================================
        private static string SafeRoomToString(Room r)
        {
            if (r == null) return "";
            try
            {
                string name = r.Name ?? "";
                string num = r.Number ?? "";
                if (!string.IsNullOrWhiteSpace(num) && !string.IsNullOrWhiteSpace(name))
                    return $"{num} - {name}";
                return !string.IsNullOrWhiteSpace(name) ? name : num;
            }
            catch { return ""; }
        }

        private static void CollectByPrefixes(Dictionary<string, string> dict, Element inst, Element type, string[] prefixes)
        {
            if (dict == null) return;

            foreach (Parameter p in inst.Parameters)
            {
                string name = p.Definition?.Name ?? "";
                if (string.IsNullOrWhiteSpace(name)) continue;

                if (prefixes.Any(pre => name.StartsWith(pre, StringComparison.OrdinalIgnoreCase)))
                {
                    string val = GetParameterValueString(p);
                    if (!string.IsNullOrWhiteSpace(val) && !dict.ContainsKey(name))
                        dict[name] = val.Trim();
                }
            }

            if (type != null)
            {
                foreach (Parameter p in type.Parameters)
                {
                    string name = p.Definition?.Name ?? "";
                    if (string.IsNullOrWhiteSpace(name)) continue;

                    if (prefixes.Any(pre => name.StartsWith(pre, StringComparison.OrdinalIgnoreCase)))
                    {
                        string val = GetParameterValueString(p);
                        if (!string.IsNullOrWhiteSpace(val) && !dict.ContainsKey(name))
                            dict[name] = val.Trim();
                    }
                }
            }
        }

        private static KeyValuePair<string, string>? FindParamByPrefix(Element inst, Element type, string prefix)
        {
            foreach (Parameter p in inst.Parameters)
            {
                string name = p.Definition?.Name ?? "";
                if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return new KeyValuePair<string, string>(name, GetParameterValueString(p));
            }

            if (type != null)
            {
                foreach (Parameter p in type.Parameters)
                {
                    string name = p.Definition?.Name ?? "";
                    if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                        return new KeyValuePair<string, string>(name, GetParameterValueString(p));
                }
            }
            return null;
        }

        private static string GetParamFromInstanceOrType(Element inst, Element type, string paramName)
        {
            string v = GetParamAsString(inst, paramName);
            if (string.IsNullOrWhiteSpace(v) && type != null)
                v = GetParamAsString(type, paramName);
            return v;
        }

        private static string GetParamAsString(Element e, string paramName)
        {
            try
            {
                var p = e?.LookupParameter(paramName);
                if (p == null) return "";
                return GetParameterValueString(p);
            }
            catch { return ""; }
        }

        private static string GetParameterValueString(Parameter p)
        {
            try
            {
                if (p == null) return "";
                if (p.StorageType == StorageType.String) return p.AsString() ?? "";
                return p.AsValueString() ?? "";
            }
            catch { return ""; }
        }

        private static string GetAsString(Element e, BuiltInParameter bip)
        {
            try
            {
                var p = e.get_Parameter(bip);
                if (p == null) return "";
                if (p.StorageType == StorageType.Double) return p.AsValueString() ?? "";
                return p.AsString() ?? p.AsValueString() ?? "";
            }
            catch { return ""; }
        }

        private static string GetLevelName(Document doc, Element e)
        {
            try
            {
                var id = e.LevelId;
                if (id == ElementId.InvalidElementId) return "";
                return (doc.GetElement(id) as Level)?.Name ?? "";
            }
            catch { return ""; }
        }

        private static string GetHostName(Element e)
        {
            try
            {
                if (e is FamilyInstance fi && fi.Host != null)
                    return fi.Host.Name ?? "";
                return "";
            }
            catch { return ""; }
        }

        private static string GetLocationStringMeters(Element e)
        {
            try
            {
                if (e.Location is LocationPoint lp)
                    return FormatPointMeters(lp.Point);

                if (e.Location is LocationCurve lc)
                {
                    var p0 = lc.Curve.GetEndPoint(0);
                    var p1 = lc.Curve.GetEndPoint(1);
                    return $"{FormatPointMeters(p0)} -> {FormatPointMeters(p1)}";
                }
                return "";
            }
            catch { return ""; }
        }

        private static string FormatPointMeters(XYZ p)
        {
            double x = p.X * 0.3048;
            double y = p.Y * 0.3048;
            double z = p.Z * 0.3048;

            return $"{x.ToString("0.###", CultureInfo.InvariantCulture)}," +
                   $"{y.ToString("0.###", CultureInfo.InvariantCulture)}," +
                   $"{z.ToString("0.###", CultureInfo.InvariantCulture)} (m)";
        }

        private static string ToXYZString(XYZ v)
        {
            if (v == null) return "";
            return $"{v.X.ToString("0.###", CultureInfo.InvariantCulture)}," +
                   $"{v.Y.ToString("0.###", CultureInfo.InvariantCulture)}," +
                   $"{v.Z.ToString("0.###", CultureInfo.InvariantCulture)}";
        }

        private static List<string> GetMaterialNames(Document doc, Element inst, FamilySymbol type)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                if (type != null)
                {
                    foreach (var mid in type.GetMaterialIds(false))
                    {
                        var mat = doc.GetElement(mid) as Autodesk.Revit.DB.Material;
                        if (mat != null && !string.IsNullOrWhiteSpace(mat.Name))
                            names.Add(mat.Name.Trim());
                    }
                }

                foreach (var mid in inst.GetMaterialIds(false))
                {
                    var mat = doc.GetElement(mid) as Autodesk.Revit.DB.Material;
                    if (mat != null && !string.IsNullOrWhiteSpace(mat.Name))
                        names.Add(mat.Name.Trim());
                }
            }
            catch { }

            return names.ToList();
        }

        private static void AppendMetricLine(StringBuilder sb, Dictionary<string, string> dict, string key, string label)
        {
            if (dict == null) return;
            if (dict.TryGetValue(key, out var v) && !string.IsNullOrWhiteSpace(v))
                sb.AppendLine($"- {label}: {v.Trim()}");
        }

        private static string SafeUniqueId(Element e)
        {
            try { return e?.UniqueId ?? ""; } catch { return ""; }
        }

        private static long SafeElementId(Element e)
        {
            try { return e?.Id?.Value ?? -1L; } catch { return -1L; }
        }
    }
}
