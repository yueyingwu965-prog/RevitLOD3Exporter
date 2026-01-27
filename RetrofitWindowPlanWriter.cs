using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace RevitLOD3Exporter
{
    internal static class RetrofitWindowPlanWriter
    {
        // 兼容旧调用名（如果 RetrofitCommand 还在调用 GenerateAndWritePlans）
        public static void GenerateAndWritePlans(Document doc)
        {
            ExportWindowsReportTxt(doc);
        }

        // 你截图里重点关心的参数名（严格匹配）
        private static readonly string[] KeyParamNames = new[]
        {
            "ea_score","lt_score","we_score","ss_score","retrofit_score_total",
            "ea_window_u_value","lst_mean_celsius",
            "ea_shading_count","ea_shading_depth_m","ea_shading_status","shadow_ratio_mean"
        };

        // 额外：收集所有这些前缀的参数（实例+类型）
        private static readonly string[] ParamPrefixes = new[] { "ea_", "we_", "ss_", "lt_" };

        public static void ExportWindowsReportTxt(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var windows = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Windows)
                .WhereElementIsNotElementType()
                .ToElements();

            if (windows.Count == 0)
            {
                TaskDialog.Show("Windows", "No window instances found.");
                return;
            }

            // ✅ FIX 1: defaultOut 未定义 -> 这里定义
            string defaultOut = GetDefaultOutputPath(doc);

            // ✅ 只弹窗选择输出路径（不再有 inputPath）
            string outputPath;
            if (!PickOutputPath(defaultOut, out outputPath))
                return;

            // 不再使用任何 externalContext
            string externalContext = "";

            // ✅ Window type summary: Family | Type -> Count
            var windowTypeSummary = windows
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

            int ok = 0, fail = 0;

            var sb = new StringBuilder();
            sb.AppendLine("WINDOWS RETROFIT REPORT (TXT EXPORT)");
            sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"Model: {doc.Title}");
            sb.AppendLine($"Window count: {windows.Count}");
            sb.AppendLine();

            sb.AppendLine("WINDOW TYPE SUMMARY (Family | Type -> Count)");
            foreach (var g in windowTypeSummary)
                sb.AppendLine($"- {g.Key} : {g.Count()}");

            sb.AppendLine();
            sb.AppendLine(new string('=', 90));
            sb.AppendLine();

            foreach (var e in windows)
            {
                try
                {
                    var info = ExtractWindowFullInfo(doc, e);

                    string prompt = BuildPrompt(info, externalContext);

                    // API建议（把这里替换成你真实的调用）
                    string advice = CallRetrofitApiForWindowPlan(prompt);

                    AppendWindowBlock(sb, info, advice);

                    ok++;
                }
                catch
                {
                    fail++;
                }
            }

            File.WriteAllText(outputPath, sb.ToString(), Encoding.UTF8);

            TaskDialog.Show(
                "Windows",
                $"TXT exported.\nSuccess: {ok}\nFailed: {fail}\n\nPath:\n{outputPath}"
            );
        }

        // ✅ 只选择输出路径（Revit原生 FileSaveDialog）
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
            return Path.Combine(desktop, $"Windows_RetrofitReport_{safeTitle}_{ts}.txt");
        }

        private static string SanitizeFileName(string s)
        {
            foreach (var c in Path.GetInvalidFileNameChars())
                s = s.Replace(c, '_');
            return s;
        }

        // -------------------------
        // Window info model
        // -------------------------
        private class WindowFullInfo
        {
            public int ElementId { get; set; }
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

        // -------------------------
        // Extract: window own info + parameters
        // -------------------------
        private static WindowFullInfo ExtractWindowFullInfo(Document doc, Element windowInstance)
        {
            // ⚠️ 你原来文件里应该已经有完整实现。
            // 这里给你一个“可编译版本”的实现（不会写回Revit，只读取信息）

            var info = new WindowFullInfo
            {
                ElementId = windowInstance.Id.IntegerValue,
                UniqueId = windowInstance.UniqueId
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

            foreach (var pn in KeyParamNames)
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
            sb.AppendLine("Collected Parameters (key + ea/we/ss/lt-prefixed):");
            foreach (var kv in w.Params.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase))
                sb.AppendLine($"- {kv.Key}: {kv.Value}");

            sb.AppendLine();
            sb.AppendLine("AI Retrofit Advice:");
            sb.AppendLine(string.IsNullOrWhiteSpace(advice) ? "(no advice)" : advice.Trim());
            sb.AppendLine(new string('-', 90));
            sb.AppendLine();
        }

        private static string BuildPrompt(WindowFullInfo w, string externalContext)
        {
            string mats = (w.Materials != null && w.Materials.Count > 0) ? string.Join(", ", w.Materials) : "Unknown";
            string paramLines = (w.Params != null && w.Params.Count > 0)
                ? string.Join("\n", w.Params.Select(kv => $"- {kv.Key}: {kv.Value}"))
                : "- (none)";

            return
$@"You are a LEED-oriented retrofit assistant.
Task: Provide retrofit advice for a WINDOW, using both window own info and extracted Revit parameters.
Rules:
- TEXT ONLY.
- Prioritize based on EA/WE/SS/LT scores and ea_* metrics (e.g., shading, U-value).
- Do NOT instruct direct Revit geometry/type edits; focus on practical retrofit actions/spec changes.

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

EXTRACTED PARAMETERS:
{paramLines}

OUTPUT FORMAT:
1) Diagnosis (1-2 lines)
2) Retrofit measures (5-9 bullets)
3) LEED linkage (EA/WE/SS/LT) (short bullets)
4) Assumptions / missing data (if any)
";
        }

        private static string CallRetrofitApiForWindowPlan(string prompt)
        {
            // TODO: 替换为你的真实 API 调用
            return "API advice placeholder.";
        }
    }
}
