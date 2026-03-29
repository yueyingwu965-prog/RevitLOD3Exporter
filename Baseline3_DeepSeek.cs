using Autodesk.Revit.DB;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace RevitLOD3Exporter
{
    public class Baseline3_DeepSeek
    {
        private const string DEFAULT_MODEL = "deepseek-chat";
        private const string DEFAULT_URL = "https://api.deepseek.com/v1/chat/completions";
        private const double DEFAULT_PANEL_AREA_M2 = 1.90;   // typical ~400-450W module footprint (proxy only)
        private const double DEFAULT_MODULE_EFF = 0.20;      // 20%
        private const double DEFAULT_PERFORMANCE_RATIO = 0.80; // losses
        private const double DEFAULT_AZIMUTH_DEG = 180.0;    // south (Europe)
        private const double DEFAULT_TILT_MIN = 10.0;
        private const double DEFAULT_TILT_MAX = 40.0;

        private static readonly HttpClient _http = CreateHttpClient();
        private static string _lastBearerKey = null;

        private static HttpClient CreateHttpClient()
        {
            var handler = new HttpClientHandler
            {
                AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate
            };

            var http = new HttpClient(handler);
            http.Timeout = System.Threading.Timeout.InfiniteTimeSpan;
            http.DefaultRequestHeaders.Accept.Clear();
            http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            http.DefaultRequestHeaders.UserAgent.ParseAdd("RevitLOD3Exporter/1.0");
            return http;
        }

        private static string PostJsonWithRetry(string url, string jsonBody, int timeoutSeconds, int maxAttempts)
        {
            Exception last = null;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    using (var req = new HttpRequestMessage(HttpMethod.Post, url))
                    {
                        req.Content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

                        using (var cts = new System.Threading.CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds)))
                        {
                            var resp = _http.SendAsync(req, HttpCompletionOption.ResponseHeadersRead, cts.Token)
                                            .GetAwaiter().GetResult();

                            var raw = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult();

                            if (!resp.IsSuccessStatusCode)
                                throw new Exception($"HTTP {(int)resp.StatusCode} {resp.ReasonPhrase}\n{raw}");

                            return raw;
                        }
                    }
                }
                catch (Exception ex)
                {
                    last = ex;

                    bool retryable =
                        ex is HttpRequestException ||
                        ex is TaskCanceledException ||
                        (ex.InnerException is TimeoutException);

                    if (!retryable || attempt == maxAttempts) break;

                    int delayMs = 1000 * attempt * attempt;
                    System.Threading.Thread.Sleep(delayMs);
                }
            }

            throw new Exception($"DeepSeek POST failed after {maxAttempts} attempts.\nLast error: {last}");
        }

        private static void EnsureAuthHeader(string apiKey)
        {
            if (!string.Equals(_lastBearerKey, apiKey, StringComparison.Ordinal))
            {
                _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);
                _lastBearerKey = apiKey;
            }
        }

        private static string GetApiUrl()
        {
            var url = Environment.GetEnvironmentVariable("RETROFIT_API_URL");
            if (string.IsNullOrWhiteSpace(url)) url = DEFAULT_URL;
            return url.Trim();
        }

        private static string GetModel()
        {
            var m = Environment.GetEnvironmentVariable("RETROFIT_API_MODEL");
            if (string.IsNullOrWhiteSpace(m)) m = DEFAULT_MODEL;
            return m.Trim();
        }

        public class ResultInfo
        {
            public int ApiWrites = 0;
            public List<string> Notes = new List<string>();
        }

        public class LeedPlan
        {
            public EaSection ea { get; set; }
            public SsSection ss { get; set; }
            public WeSection we { get; set; }
            public LtSection lt { get; set; }
            public List<string> assumptions { get; set; }
        }

        public class EaSection
        {
            public EaTargets targets { get; set; }
            public EaStrategy strategy { get; set; }
        }

        public class EaTargets
        {
            public int pv_panel_count { get; set; }
            public double pv_added_area_m2 { get; set; }
            public string pv_layout_status { get; set; }
            public double shading_depth_m { get; set; }
            public int shading_count { get; set; }
            public string shading_status { get; set; }
        }

        public class EaStrategy
        {
            public string ea_roof_pv_strategy { get; set; }
            public string ea_roof_pv_system_type { get; set; }
            public bool ea_roof_structural_capacity_flag { get; set; }
            public string ea_window_shading_strategy { get; set; }
            public string ea_window_upgrade_strategy { get; set; }

            public string ea_door_upgrade_strategy { get; set; }

            public string ea_wall_envelope_strategy { get; set; }
            public string ea_floor_thermal_strategy { get; set; }
        }

        public class SsSection
        {
            public SsStrategies strategies { get; set; }
            public string summary { get; set; }
            public List<string> recommendations { get; set; }
        }

        public class SsStrategies
        {
            public string ss_roof_heat_strategy { get; set; }
        }

        public class WeSection
        {
            public WeTargets targets { get; set; }
            public WeStrategies strategies { get; set; }
            public string summary { get; set; }
            public List<string> recommendations { get; set; }
        }

        public class WeTargets
        {
            public double we_rainwater_storage_capacity_m3 { get; set; }
            public string we_water_efficiency_level { get; set; }
        }

        public class WeStrategies
        {
            public string we_roof_rainwater_strategy { get; set; }
            public string we_floor_permeability_strategy { get; set; }
        }

        public class LtSection
        {
            public LtTargets targets { get; set; }
            public string summary { get; set; }
            public List<string> recommendations { get; set; }
        }

        public class LtTargets
        {
            public bool lt_constraint_flag { get; set; }
            public double lt_retrofit_intensity_factor { get; set; }
            public double lt_vlt { get; set; }
            public string lt_glare_risk_class { get; set; }
        }

        public class LeedAdvice
        {
            public AdviceSection ea { get; set; }
            public AdviceSection ss { get; set; }
            public AdviceSection we { get; set; }
            public AdviceSection lt { get; set; }

            public ComponentAdvice components { get; set; }

            public List<string> priorities { get; set; }
            public List<string> risks { get; set; }
        }

        public class AdviceSection
        {
            public string diagnosis { get; set; }
            public List<string> actions { get; set; }
            public List<ParamAdvice> parameters { get; set; }
        }

        public class ParamAdvice
        {
            public string parameter { get; set; }
            public string target { get; set; }
            public string note { get; set; }
        }

        public class ComponentAdvice
        {
            public List<ComponentItem> roof { get; set; }
            public List<ComponentItem> window { get; set; }
            public List<ComponentItem> door { get; set; }
            public List<ComponentItem> wall { get; set; }
            public List<ComponentItem> floor { get; set; }
        }

        public class ComponentItem
        {
            public string measure { get; set; }
            public string where { get; set; }
            public string why { get; set; }
            public List<string> steps { get; set; }
            public List<string> qa_check { get; set; }
        }

        public ResultInfo Apply(
            Document doc,
            Dictionary<string, string> row,
            IList<Element> roofs,
            IList<Element> windows,
            IList<Element> doors,
            IList<Element> walls,
            IList<Element> floors,
            LeedPlan plan)
        {
            var res = new ResultInfo();

            if (plan == null)
            {
                res.Notes.Add("AI plan is null (API failed or not called).");
                return res;
            }

            try
            {
                if (plan.ea == null) plan.ea = new EaSection();
                if (plan.ea.targets == null) plan.ea.targets = new EaTargets();

                double pvAdded = TryGetMetricDouble(row, "ea_pv_added_area_m2");

                int pvCount = 0;
                if (row != null && row.TryGetValue("ea_pv_panel_count", out var cntRaw))
                    pvCount = (int)Math.Round(RetrofitShared.ParseDouble(cntRaw));
                if (pvCount < 0) pvCount = 0;

                string status = "csv_deterministic";
                if (row != null && row.TryGetValue("ea_pv_layout_status", out var stRaw) && !string.IsNullOrWhiteSpace(stRaw))
                    status = stRaw.Trim();

                if (double.IsNaN(pvAdded) || pvAdded < 0) pvAdded = 0;

                plan.ea.targets.pv_added_area_m2 = pvAdded;
                plan.ea.targets.pv_panel_count = pvCount;
                plan.ea.targets.pv_layout_status = status;

                res.Notes.Add($"PV targets locked: added={pvAdded:0.###} m2, count={pvCount}, status={status}");
            }
            catch (Exception ex)
            {
                res.Notes.Add("PV target lock failed: " + ex.Message);
            }

            foreach (var r in roofs ?? new List<Element>()) res.ApiWrites += WriteApiToRoof(doc, r, plan);
            foreach (var w in windows ?? new List<Element>()) res.ApiWrites += WriteApiToWindow(doc, w, plan);
            foreach (var d in doors ?? new List<Element>()) res.ApiWrites += WriteApiToDoor(doc, d, plan);
            foreach (var wa in walls ?? new List<Element>()) res.ApiWrites += WriteApiToWall(doc, wa, plan);
            foreach (var f in floors ?? new List<Element>()) res.ApiWrites += WriteApiToFloor(doc, f, plan);

            res.ApiWrites += WriteApiToProjectInfo(doc, plan);

            if (doors == null || doors.Count == 0)
                res.Notes.Add("Doors list is empty: ea_door_upgrade_strategy not written (pass doors collector to Apply to enable).");

            return res;
        }

        // =========================
        // Report (metrics once + deterministic PV formula + plan + detailed advice)
        // =========================
        public string BuildReport(
            Dictionary<string, string> metrics,
            LeedPlan plan,
            LeedAdvice advice,
            bool apiOk,
            ResultInfo info)
        {
            var sb = new StringBuilder();

            sb.AppendLine("LEED Retrofit Report (Baseline3 + AI)");
            sb.AppendLine("========================================");
            sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"AI status: {(apiOk ? "OK" : "FAILED")}");
            sb.AppendLine();

            sb.AppendLine("Metrics (CSV + Revit-derived + computed)");
            sb.AppendLine("----------------------------------------");
            if (metrics != null)
            {
                foreach (var kv in metrics)
                    sb.AppendLine($"- {CleanMetricKey(kv.Key)}: {CleanMetricValue(kv.Value)}");
            }
            sb.AppendLine();

            sb.AppendLine("PV Deterministic Calculation (Baseline2 logic)");
            sb.AppendLine("----------------------------------------");
            try
            {
                double pvUsable = TryGetMetricDouble(metrics, "pv_usable_roof_area_m2");
                double coverage = TryGetMetricDouble(metrics, "coverage_factor_used");
                double pw = TryGetMetricDouble(metrics, "pv_panel_unit_width_m");
                double pl = TryGetMetricDouble(metrics, "pv_panel_unit_length_m");
                double pArea = TryGetMetricDouble(metrics, "pv_panel_unit_area_m2");
                double pvAdded = TryGetMetricDouble(metrics, "ea_pv_added_area_m2");

                int pvCount = 0;
                if (metrics != null && metrics.TryGetValue("ea_pv_panel_count", out var pcRaw))
                    pvCount = (int)Math.Round(RetrofitShared.ParseDouble(pcRaw));

                sb.AppendLine($"pv_usable_roof_area_m2 (CSV): {pvUsable.ToString("0.###", CultureInfo.InvariantCulture)}");
                sb.AppendLine($"coverage_factor_used: {coverage.ToString("0.###", CultureInfo.InvariantCulture)}");
                sb.AppendLine($"ea_pv_added_area_m2 = pv_usable_roof_area_m2 × coverage_factor_used = {pvAdded.ToString("0.###", CultureInfo.InvariantCulture)}");
                sb.AppendLine($"panel_width_m: {pw.ToString("0.###", CultureInfo.InvariantCulture)}");
                sb.AppendLine($"panel_length_m: {pl.ToString("0.###", CultureInfo.InvariantCulture)}");
                sb.AppendLine($"panel_area_m2 = {pArea.ToString("0.###", CultureInfo.InvariantCulture)}");
                sb.AppendLine($"ea_pv_panel_count = floor(ea_pv_added_area_m2 ÷ panel_area_m2) = {pvCount}");
            }
            catch
            {
                sb.AppendLine("- (PV formula block unavailable)");
            }
            sb.AppendLine();

            sb.AppendLine("AI Plan (strategies + targets)");
            sb.AppendLine("----------------------------------------");
            if (plan == null)
            {
                sb.AppendLine("(plan is null)");
            }
            else
            {
                sb.AppendLine("EA");
                if (plan.ea?.strategy != null)
                {
                    sb.AppendLine($"ea_roof_pv_strategy: {plan.ea.strategy.ea_roof_pv_strategy}");
                    sb.AppendLine($"ea_roof_pv_system_type: {plan.ea.strategy.ea_roof_pv_system_type}");
                    sb.AppendLine($"ea_roof_structural_capacity_flag: {plan.ea.strategy.ea_roof_structural_capacity_flag}");
                    sb.AppendLine($"ea_window_shading_strategy: {plan.ea.strategy.ea_window_shading_strategy}");
                    sb.AppendLine($"ea_window_upgrade_strategy: {plan.ea.strategy.ea_window_upgrade_strategy}");
                    sb.AppendLine($"ea_door_upgrade_strategy: {plan.ea.strategy.ea_door_upgrade_strategy}");
                    sb.AppendLine($"ea_wall_envelope_strategy: {plan.ea.strategy.ea_wall_envelope_strategy}");
                    sb.AppendLine($"ea_floor_thermal_strategy: {plan.ea.strategy.ea_floor_thermal_strategy}");
                }
                if (plan.ea?.targets != null)
                {
                    sb.AppendLine($"pv_panel_count: {plan.ea.targets.pv_panel_count}");
                    sb.AppendLine($"pv_added_area_m2: {plan.ea.targets.pv_added_area_m2.ToString("0.###", CultureInfo.InvariantCulture)}");
                    sb.AppendLine($"pv_layout_status: {plan.ea.targets.pv_layout_status}");
                    sb.AppendLine($"shading_depth_m: {plan.ea.targets.shading_depth_m.ToString("0.###", CultureInfo.InvariantCulture)}");
                    sb.AppendLine($"shading_count: {plan.ea.targets.shading_count}");
                    sb.AppendLine($"shading_status: {plan.ea.targets.shading_status}");
                }
                sb.AppendLine();

                sb.AppendLine("SS");
                if (plan.ss?.strategies != null)
                    sb.AppendLine($"ss_roof_heat_strategy: {plan.ss.strategies.ss_roof_heat_strategy}");
                if (!string.IsNullOrWhiteSpace(plan.ss?.summary))
                    sb.AppendLine($"summary: {plan.ss.summary}");
                if (plan.ss?.recommendations != null && plan.ss.recommendations.Count > 0)
                {
                    sb.AppendLine("recommendations:");
                    foreach (var r in plan.ss.recommendations) sb.AppendLine($"- {r}");
                }
                sb.AppendLine();

                sb.AppendLine("WE");
                if (plan.we?.targets != null)
                {
                    sb.AppendLine($"we_rainwater_storage_capacity_m3: {plan.we.targets.we_rainwater_storage_capacity_m3.ToString("0.###", CultureInfo.InvariantCulture)}");
                    sb.AppendLine($"we_water_efficiency_level: {plan.we.targets.we_water_efficiency_level}");
                }
                if (plan.we?.strategies != null)
                {
                    sb.AppendLine($"we_roof_rainwater_strategy: {plan.we.strategies.we_roof_rainwater_strategy}");
                    sb.AppendLine($"we_floor_permeability_strategy: {plan.we.strategies.we_floor_permeability_strategy}");
                }
                if (!string.IsNullOrWhiteSpace(plan.we?.summary))
                    sb.AppendLine($"summary: {plan.we.summary}");
                if (plan.we?.recommendations != null && plan.we.recommendations.Count > 0)
                {
                    sb.AppendLine("recommendations:");
                    foreach (var r in plan.we.recommendations) sb.AppendLine($"- {r}");
                }
                sb.AppendLine();

                sb.AppendLine("LT");
                if (plan.lt?.targets != null)
                {
                    sb.AppendLine($"lt_constraint_flag: {plan.lt.targets.lt_constraint_flag}");
                    sb.AppendLine($"lt_retrofit_intensity_factor: {plan.lt.targets.lt_retrofit_intensity_factor.ToString("0.###", CultureInfo.InvariantCulture)}");
                    sb.AppendLine($"lt_vlt: {plan.lt.targets.lt_vlt.ToString("0.###", CultureInfo.InvariantCulture)}");
                    sb.AppendLine($"lt_glare_risk_class: {plan.lt.targets.lt_glare_risk_class}");
                }
                if (!string.IsNullOrWhiteSpace(plan.lt?.summary))
                    sb.AppendLine($"summary: {plan.lt.summary}");
                if (plan.lt?.recommendations != null && plan.lt.recommendations.Count > 0)
                {
                    sb.AppendLine("recommendations:");
                    foreach (var r in plan.lt.recommendations) sb.AppendLine($"- {r}");
                }
                sb.AppendLine();
            }

            sb.AppendLine("Detailed Retrofit Advice (EA / SS / WE / LT)");
            sb.AppendLine("----------------------------------------");
            AppendAdviceSection(sb, "EA", advice?.ea);
            AppendAdviceSection(sb, "SS", advice?.ss);
            AppendAdviceSection(sb, "WE", advice?.we);
            AppendAdviceSection(sb, "LT", advice?.lt);

            sb.AppendLine("Component-level Recommendations");
            sb.AppendLine("----------------------------------------");
            AppendComponentAdvice(sb, "ROOF", advice?.components?.roof);
            AppendComponentAdvice(sb, "WINDOW", advice?.components?.window);
            AppendComponentAdvice(sb, "DOOR", advice?.components?.door);
            AppendComponentAdvice(sb, "WALL", advice?.components?.wall);
            AppendComponentAdvice(sb, "FLOOR", advice?.components?.floor);

            sb.AppendLine("Priorities");
            sb.AppendLine("----------------------------------------");
            if (advice?.priorities != null && advice.priorities.Count > 0)
            {
                foreach (var p in advice.priorities) sb.AppendLine($"- {p}");
            }
            else sb.AppendLine("- (none)");
            sb.AppendLine();

            sb.AppendLine("Risks / Assumptions (Advice)");
            sb.AppendLine("----------------------------------------");
            if (advice?.risks != null && advice.risks.Count > 0)
            {
                foreach (var r in advice.risks) sb.AppendLine($"- {r}");
            }
            else sb.AppendLine("- (none)");
            sb.AppendLine();

            sb.AppendLine($"AI writes applied: {info?.ApiWrites ?? 0}");
            if (info?.Notes != null && info.Notes.Count > 0)
            {
                sb.AppendLine("Notes");
                sb.AppendLine("----------------------------------------");
                foreach (var n in info.Notes) sb.AppendLine($"- {n}");
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static void AppendAdviceSection(StringBuilder sb, string title, AdviceSection s)
        {
            sb.AppendLine(title);
            if (s == null)
            {
                sb.AppendLine("- (no advice)");
                sb.AppendLine();
                return;
            }

            if (!string.IsNullOrWhiteSpace(s.diagnosis))
                sb.AppendLine($"diagnosis: {s.diagnosis}");

            if (s.actions != null && s.actions.Count > 0)
            {
                sb.AppendLine("actions:");
                foreach (var a in s.actions) sb.AppendLine($"- {a}");
            }

            if (s.parameters != null && s.parameters.Count > 0)
            {
                sb.AppendLine("parameters:");
                foreach (var p in s.parameters)
                {
                    string note = string.IsNullOrWhiteSpace(p.note) ? "" : $" ({p.note})";
                    sb.AppendLine($"- {p.parameter}: {p.target}{note}");
                }
            }

            sb.AppendLine();
        }

        private static void AppendComponentAdvice(StringBuilder sb, string title, List<ComponentItem> items)
        {
            sb.AppendLine(title);
            if (items == null || items.Count == 0)
            {
                sb.AppendLine("- (no items)");
                sb.AppendLine();
                return;
            }

            int idx = 1;
            foreach (var it in items)
            {
                sb.AppendLine($"{idx}. {it.measure}");
                if (!string.IsNullOrWhiteSpace(it.where)) sb.AppendLine($"   where: {it.where}");
                if (!string.IsNullOrWhiteSpace(it.why)) sb.AppendLine($"   why: {it.why}");

                if (it.steps != null && it.steps.Count > 0)
                {
                    sb.AppendLine("   steps:");
                    foreach (var s in it.steps) sb.AppendLine($"   - {s}");
                }

                if (it.qa_check != null && it.qa_check.Count > 0)
                {
                    sb.AppendLine("   qa_check:");
                    foreach (var q in it.qa_check) sb.AppendLine($"   - {q}");
                }

                idx++;
            }
            sb.AppendLine();
        }

        public static LeedPlan CallDeepSeekLeedPlan(Dictionary<string, string> metrics, out string raw)
        {
            return CallDeepSeekLeedPlan(doc: null, roofs: null, metrics: metrics, raw: out raw);
        }

        public static LeedPlan CallDeepSeekLeedPlan(Document doc, IList<Element> roofs, Dictionary<string, string> metrics, out string raw)
        {
            string key = GetApiKey();
            EnsureAuthHeader(key);

            // Enrich metrics with:
            // - roof_total_area_m2 (reference only)
            // - pv_installable_area_m2 = CSV pv_usable_roof_area_m2 (fallback roof total if missing)
            // - site_latitude / site_longitude
            // - simple weather/irradiance proxy (reproducible)
            EnsurePvMetricsFromRevit(doc, roofs, metrics);

            string systemPrompt =
@"You are a LEED building retrofit consultant.
Use the given building metrics to propose retrofit strategies and parameter values.

IMPORTANT PV RULES (must follow):
- The metric 'pv_installable_area_m2' represents the MAX PV area available FROM CSV pv_usable_roof_area_m2 (fallback to roof total if missing).
- Set pv_added_area_m2 <= pv_installable_area_m2.
- Keep pv_panel_count consistent with pv_added_area_m2 using metric 'pv_panel_area_m2':
  pv_panel_count ≈ round(pv_added_area_m2 / pv_panel_area_m2).
- If any PV input is missing, state it in assumptions and keep PV conservative.

Return ONLY valid JSON with this structure (no markdown, no extra text):

{
  ""ea"": {
    ""targets"": {
      ""pv_panel_count"": int,
      ""pv_added_area_m2"": number,
      ""pv_layout_status"": string,
      ""shading_depth_m"": number,
      ""shading_count"": int,
      ""shading_status"": string
    },
    ""strategy"": {
      ""ea_roof_pv_strategy"": ""none|limited|full"",
      ""ea_roof_pv_system_type"": ""attached|mounted|bipv"",
      ""ea_roof_structural_capacity_flag"": true|false,
      ""ea_window_shading_strategy"": ""none|fixed|adjustable"",
      ""ea_window_upgrade_strategy"": ""baseline|low_e_double|triple|frame_upgrade"",
      ""ea_door_upgrade_strategy"": ""baseline|seal_threshold|replace_insulated_set"",
      ""ea_wall_envelope_strategy"": ""baseline|insulated_upgrade|high_performance"",
      ""ea_floor_thermal_strategy"": ""baseline|insulated""
    }
  },
  ""ss"": {
    ""strategies"": {
      ""ss_roof_heat_strategy"": ""none|green_roof|cool_roof""
    },
    ""summary"": ""..."",
    ""recommendations"": [""...""] 
  },
  ""we"": {
    ""targets"": {
      ""we_rainwater_storage_capacity_m3"": number,
      ""we_water_efficiency_level"": ""low|medium|high""
    },
    ""strategies"": {
      ""we_roof_rainwater_strategy"": ""none|basic|enhanced"",
      ""we_floor_permeability_strategy"": ""impervious|semi_permeable|permeable""
    },
    ""summary"": ""..."",
    ""recommendations"": [""...""] 
  },
  ""lt"": {
    ""targets"": {
      ""lt_constraint_flag"": true|false,
      ""lt_retrofit_intensity_factor"": number,
      ""lt_vlt"": number,
      ""lt_glare_risk_class"": ""low|medium|high""
    },
    ""summary"": ""..."",
    ""recommendations"": [""...""] 
  },
  ""assumptions"": [""...""] 
}

Rules:
- Fill ALL strategy fields with one of the allowed enum values.
- lt_vlt should be between 0 and 1.
- lt_glare_risk_class must be one of: low, medium, high.
- Keep summaries short; recommendations as an engineering checklist.
- Do NOT include any fields outside this JSON.";

            var sb = new StringBuilder("Building metrics:\n");
            if (metrics != null)
            {
                foreach (var kv in metrics)
                    sb.AppendLine($"- {CleanMetricKey(kv.Key)}: {CleanMetricValue(kv.Value)}");
            }

            var body = new
            {
                model = GetModel(),
                messages = new[]
                {
                    new { role = "system", content = systemPrompt },
                    new { role = "user", content = sb.ToString() }
                },
                temperature = 0.0
            };

            string json = JsonConvert.SerializeObject(body);

            raw = PostJsonWithRetry(GetApiUrl(), json, timeoutSeconds: 180, maxAttempts: 3);

            dynamic root = JsonConvert.DeserializeObject<dynamic>(raw);
            string aiText = root.choices[0].message.content.ToString();

            string jsonOnly = ExtractJsonBlock(aiText);
            if (string.IsNullOrEmpty(jsonOnly))
                throw new Exception("AI did not return valid JSON (plan).");

            var plan = JsonConvert.DeserializeObject<LeedPlan>(jsonOnly);

            if (plan.assumptions == null) plan.assumptions = new List<string>();
            return plan;
        }

        // =========================
        // DeepSeek Call 2: DETAILED ADVICE (JSON)
        // =========================
        public static LeedAdvice CallDeepSeekLeedAdvice(
            Dictionary<string, string> metrics,
            LeedPlan plan,
            out string rawAdvice)
        {
            string key = GetApiKey();
            EnsureAuthHeader(key);

            string systemPrompt =
@"You are a LEED building retrofit consultant.
Given (1) building metrics and (2) the PLAN JSON already generated,
produce DETAILED retrofit advice grounded in both.

Return ONLY valid JSON with this structure (no markdown, no extra text):

{
  ""ea"": { ""diagnosis"": ""..."", ""actions"": [""...""], ""parameters"": [ {""parameter"": ""..."", ""target"": ""..."", ""note"": ""...""} ] },
  ""ss"": { ""diagnosis"": ""..."", ""actions"": [""...""], ""parameters"": [ ... ] },
  ""we"": { ""diagnosis"": ""..."", ""actions"": [""...""], ""parameters"": [ ... ] },
  ""lt"": { ""diagnosis"": ""..."", ""actions"": [""...""], ""parameters"": [ ... ] },
  ""components"": {
    ""roof"":   [ { ""measure"": ""..."", ""where"": ""..."", ""why"": ""..."", ""steps"": [""...""], ""qa_check"": [""...""] } ],
    ""window"": [ ... ],
    ""door"":   [ ... ],
    ""wall"":   [ ... ],
    ""floor"":  [ ... ]
  },
  ""priorities"": [""quick wins..."", ""mid term..."", ""major works...""],
  ""risks"": [""assumption/constraint..."", ""coordination risks...""]
}

Rules:
- Advice must be explicitly grounded in provided metrics and the plan.
- Actions should be practical and implementable, including component-level measures.
- Diagnosis short; actions detailed.
- If key metrics are missing or contradictory, state it in risks.";

            var user = new StringBuilder();
            user.AppendLine("METRICS:");
            if (metrics != null)
            {
                foreach (var kv in metrics)
                    user.AppendLine($"- {CleanMetricKey(kv.Key)}: {CleanMetricValue(kv.Value)}");
            }
            user.AppendLine();
            user.AppendLine("PLAN_JSON:");
            user.AppendLine(JsonConvert.SerializeObject(plan));

            var body = new
            {
                model = GetModel(),
                messages = new[]
                {
                    new { role = "system", content = systemPrompt },
                    new { role = "user", content = user.ToString() }
                },
                temperature = 0.2
            };

            string json = JsonConvert.SerializeObject(body);

            rawAdvice = PostJsonWithRetry(GetApiUrl(), json, timeoutSeconds: 240, maxAttempts: 3);

            dynamic root = JsonConvert.DeserializeObject<dynamic>(rawAdvice);
            string aiText = root.choices[0].message.content.ToString();

            string jsonOnly = ExtractJsonBlock(aiText);
            if (string.IsNullOrEmpty(jsonOnly))
                throw new Exception("AI did not return valid JSON (advice).");

            return JsonConvert.DeserializeObject<LeedAdvice>(jsonOnly);
        }

        private static string CleanMetricKey(string k)
        {
            if (k == null) return "";
            return k.Trim().Trim('"').Replace("\r", "").Replace("\n", "");
        }

        private static string CleanMetricValue(string v)
        {
            if (v == null) return "";
            var s = v.Trim();

            // strip wrapping quotes
            s = s.Trim('"');

            // collapse whitespace/newlines
            s = s.Replace("\r", " ").Replace("\n", " ");
            while (s.Contains("  ")) s = s.Replace("  ", " ");

            return s;
        }

        // ✅ robust JSON extraction
        private static string ExtractJsonBlock(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;

            int start = text.IndexOf('{');
            if (start < 0) return null;

            int depth = 0;
            for (int i = start; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch == '{') depth++;
                else if (ch == '}')
                {
                    depth--;
                    if (depth == 0)
                        return text.Substring(start, i - start + 1);
                }
            }

            return null;
        }

        private static string GetApiKey()
        {
            var key =
                Environment.GetEnvironmentVariable("DEEPSEEK_API_KEY") ??
                Environment.GetEnvironmentVariable("OPENAI_API_KEY");

            if (string.IsNullOrWhiteSpace(key))
                throw new Exception("Missing environment variable: DEEPSEEK_API_KEY (or OPENAI_API_KEY fallback).");

            return key.Trim().Replace("\r", "").Replace("\n", "");
        }

        // =========================
        // Write PLAN to Elements
        // =========================
        private int WriteApiToRoof(Document doc, Element r, LeedPlan p)
        {
            int c = 0;

            if (p?.ea?.targets != null)
            {
                c += Set(doc, r, "ea_pv_panel_count", p.ea.targets.pv_panel_count);
                c += Set(doc, r, "ea_pv_added_area_m2", p.ea.targets.pv_added_area_m2);
                c += Set(doc, r, "ea_pv_layout_status", p.ea.targets.pv_layout_status);
            }

            if (p?.ea?.strategy != null)
            {
                c += Set(doc, r, "ea_roof_pv_strategy", p.ea.strategy.ea_roof_pv_strategy);
                c += Set(doc, r, "ea_roof_pv_system_type", p.ea.strategy.ea_roof_pv_system_type);
                c += Set(doc, r, "ea_roof_structural_capacity_flag", p.ea.strategy.ea_roof_structural_capacity_flag);
            }

            if (p?.ss?.strategies != null)
                c += Set(doc, r, "ss_roof_heat_strategy", p.ss.strategies.ss_roof_heat_strategy);

            if (p?.we?.strategies != null)
                c += Set(doc, r, "we_roof_rainwater_strategy", p.we.strategies.we_roof_rainwater_strategy);

            if (p?.we?.targets != null)
            {
                c += Set(doc, r, "we_rainwater_storage_capacity_m3", p.we.targets.we_rainwater_storage_capacity_m3);
                c += Set(doc, r, "we_water_efficiency_level", p.we.targets.we_water_efficiency_level);
            }

            if (p?.lt?.targets != null)
            {
                c += Set(doc, r, "lt_constraint_flag", p.lt.targets.lt_constraint_flag);
                c += Set(doc, r, "lt_retrofit_intensity_factor", p.lt.targets.lt_retrofit_intensity_factor);
            }

            return c;
        }

        private int WriteApiToWindow(Document doc, Element w, LeedPlan p)
        {
            int c = 0;

            if (p?.ea?.targets != null)
            {
                c += Set(doc, w, "ea_shading_depth_m", p.ea.targets.shading_depth_m);
                c += Set(doc, w, "ea_shading_count", p.ea.targets.shading_count);
                c += Set(doc, w, "ea_shading_status", p.ea.targets.shading_status);
            }

            if (p?.ea?.strategy != null)
            {
                c += Set(doc, w, "ea_window_shading_strategy", p.ea.strategy.ea_window_shading_strategy);
                c += Set(doc, w, "ea_window_upgrade_strategy", p.ea.strategy.ea_window_upgrade_strategy);
            }

            if (p?.lt?.targets != null)
            {
                c += Set(doc, w, "lt_vlt", p.lt.targets.lt_vlt);
                c += Set(doc, w, "lt_glare_risk_class", p.lt.targets.lt_glare_risk_class);
            }

            return c;
        }

        private int WriteApiToDoor(Document doc, Element d, LeedPlan p)
        {
            int c = 0;
            if (p?.ea?.strategy != null)
                c += Set(doc, d, "ea_door_upgrade_strategy", p.ea.strategy.ea_door_upgrade_strategy);
            return c;
        }

        private int WriteApiToWall(Document doc, Element wa, LeedPlan p)
        {
            int c = 0;
            if (p?.ea?.strategy != null)
                c += Set(doc, wa, "ea_wall_envelope_strategy", p.ea.strategy.ea_wall_envelope_strategy);
            return c;
        }

        private int WriteApiToFloor(Document doc, Element f, LeedPlan p)
        {
            int c = 0;

            if (p?.we?.strategies != null)
                c += Set(doc, f, "we_floor_permeability_strategy", p.we.strategies.we_floor_permeability_strategy);

            if (p?.ea?.strategy != null)
                c += Set(doc, f, "ea_floor_thermal_strategy", p.ea.strategy.ea_floor_thermal_strategy);

            return c;
        }

        private int WriteApiToProjectInfo(Document doc, LeedPlan p)
        {
            int c = 0;
            Element projInfo = doc?.ProjectInformation;
            if (projInfo == null || p == null) return 0;

            if (p.lt?.targets != null)
            {
                c += Set(doc, projInfo, "lt_constraint_flag", p.lt.targets.lt_constraint_flag);
                c += Set(doc, projInfo, "lt_retrofit_intensity_factor", p.lt.targets.lt_retrofit_intensity_factor);
            }

            if (p.we?.targets != null)
                c += Set(doc, projInfo, "we_water_efficiency_level", p.we.targets.we_water_efficiency_level);

            return c;
        }

        private int Set(Document doc, Element e, string name, object val)
        {
            if (doc == null || e == null || string.IsNullOrWhiteSpace(name) || val == null) return 0;

            if (TrySetOnElement(doc, e, name, val)) return 1;

            try
            {
                var tid = e.GetTypeId();
                if (tid != ElementId.InvalidElementId)
                {
                    var t = doc.GetElement(tid) as Element;
                    if (t != null && TrySetOnElement(doc, t, name, val)) return 1;
                }
            }
            catch { }

            return 0;
        }

        private bool TrySetOnElement(Document doc, Element e, string name, object val)
        {
            Parameter p = RetrofitShared.LookupParamInsensitive(e, name);
            if (p == null || p.IsReadOnly) return false;

            try
            {
                if (p.StorageType == StorageType.Integer)
                {
                    if (val is bool b) { p.Set(b ? 1 : 0); return true; }
                    if (val is string sBool && bool.TryParse(sBool.Trim(), out bool bb)) { p.Set(bb ? 1 : 0); return true; }

                    int iv;
                    if (val is string si) iv = (int)Math.Round(RetrofitShared.ParseDouble(si));
                    else iv = Convert.ToInt32(val);

                    p.Set(iv);
                    return true;
                }

                if (p.StorageType == StorageType.String)
                {
                    p.Set(val.ToString());
                    return true;
                }

                if (p.StorageType == StorageType.Double)
                {
                    double v = (val is string sd) ? RetrofitShared.ParseDouble(sd) : Convert.ToDouble(val);
                    double internalV = RetrofitShared.ConvertToInternalIfNeeded(name, v);

                    try
                    {
                        if (p.SetValueString(v.ToString(CultureInfo.InvariantCulture)))
                            return true;
                    }
                    catch { }

                    p.Set(internalV);
                    return true;
                }

                try { return p.SetValueString(val.ToString()); } catch { }
                return false;
            }
            catch { return false; }
        }

        // =========================
        // PV metrics enrichment (Revit coords + roof total ref + PV proxy)
        // PV area max uses CSV pv_usable_roof_area_m2 (fallback roof total)
        // =========================
        private static void EnsurePvMetricsFromRevit(Document doc, IList<Element> roofs, Dictionary<string, string> metrics)
        {
            if (metrics == null) metrics = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Roof total area (m2) - reference only
            double roofTotalAreaM2 = GetTotalRoofAreaM2(doc, roofs);
            if (roofTotalAreaM2 > 0)
            {
                metrics["roof_total_area_m2"] = roofTotalAreaM2.ToString("0.##", CultureInfo.InvariantCulture);
            }

            // PV installable area preference: CSV pv_usable_roof_area_m2, else roof total
            double pvUsable = TryGetMetricDouble(metrics, "pv_usable_roof_area_m2");
            if (!double.IsNaN(pvUsable) && pvUsable > 0)
                metrics["pv_installable_area_m2"] = pvUsable.ToString("0.##", CultureInfo.InvariantCulture);
            else if (roofTotalAreaM2 > 0)
                metrics["pv_installable_area_m2"] = roofTotalAreaM2.ToString("0.##", CultureInfo.InvariantCulture);

            // Revit site coordinates (degrees)
            var (lat, lon) = GetLatLonDegrees(doc);
            if (!double.IsNaN(lat) && !double.IsNaN(lon))
            {
                metrics["site_latitude_deg"] = lat.ToString("0.######", CultureInfo.InvariantCulture);
                metrics["site_longitude_deg"] = lon.ToString("0.######", CultureInfo.InvariantCulture);
            }

            // PV assumptions/constants exposed to AI (proxy only)
            double panelAreaM2 = GetDoubleFromEnv("PV_PANEL_AREA_M2", DEFAULT_PANEL_AREA_M2);
            double eff = GetDoubleFromEnv("PV_MODULE_EFF", DEFAULT_MODULE_EFF);
            double pr = GetDoubleFromEnv("PV_PERFORMANCE_RATIO", DEFAULT_PERFORMANCE_RATIO);

            metrics["pv_panel_area_m2"] = panelAreaM2.ToString("0.###", CultureInfo.InvariantCulture);
            metrics["pv_module_efficiency"] = eff.ToString("0.###", CultureInfo.InvariantCulture);
            metrics["pv_performance_ratio"] = pr.ToString("0.###", CultureInfo.InvariantCulture);

            // Tilt suggestion
            double tilt = EstimateTiltDeg(lat);
            metrics["pv_tilt_deg_proxy"] = tilt.ToString("0.#", CultureInfo.InvariantCulture);
            metrics["pv_azimuth_deg_proxy"] = DEFAULT_AZIMUTH_DEG.ToString("0.#", CultureInfo.InvariantCulture);

            double hdd = TryGetMetricDouble(metrics, "heating_degree_days");
            double cdd = TryGetMetricDouble(metrics, "cooling_degree_days");

            double irr = EstimateAnnualIrradianceKwhM2YrProxy(lat, hdd, cdd);
            metrics["pv_irradiance_kwh_m2_yr_proxy"] = irr.ToString("0", CultureInfo.InvariantCulture);

            double pvAreaM2 = TryGetMetricDouble(metrics, "pv_installable_area_m2");
            if (double.IsNaN(pvAreaM2) || pvAreaM2 < 0) pvAreaM2 = 0;

            double kwp = (pvAreaM2 > 0) ? Math.Max(0, pvAreaM2 * eff) : 0;
            metrics["pv_peak_power_kwp_proxy"] = kwp.ToString("0.##", CultureInfo.InvariantCulture);

            double specific = Math.Max(0, irr * pr);
            metrics["pv_specific_yield_kwh_per_kwp_proxy"] = specific.ToString("0", CultureInfo.InvariantCulture);

            double annualKwh = Math.Max(0, pvAreaM2 * irr * eff * pr);
            metrics["pv_annual_kwh_proxy"] = annualKwh.ToString("0", CultureInfo.InvariantCulture);

            metrics["pv_weather_method_note"] =
                "Weather is approximated by a latitude-based irradiance proxy (no external API); pv_annual_kwh_proxy is a rough estimate.";
        }

        private static (double latDeg, double lonDeg) GetLatLonDegrees(Document doc)
        {
            try
            {
                var loc = doc?.SiteLocation;
                if (loc == null) return (double.NaN, double.NaN);

                double lat = loc.Latitude;
                double lon = loc.Longitude;
                return (lat, lon);
            }
            catch
            {
                return (double.NaN, double.NaN);
            }
        }

        private static double EstimateTiltDeg(double latDeg)
        {
            if (double.IsNaN(latDeg)) return 30.0;

            double tilt = latDeg - 10.0;
            if (tilt < DEFAULT_TILT_MIN) tilt = DEFAULT_TILT_MIN;
            if (tilt > DEFAULT_TILT_MAX) tilt = DEFAULT_TILT_MAX;
            return tilt;
        }

        private static double EstimateAnnualIrradianceKwhM2YrProxy(double latDeg, double hdd, double cdd)
        {
            double lat = double.IsNaN(latDeg) ? 45.0 : Math.Abs(latDeg);

            double irr;
            if (lat <= 35) irr = 1300;
            else if (lat >= 55) irr = 900;
            else
                irr = 1300 + (lat - 35) * (900 - 1300) / (55 - 35);

            double adj = 0.0;
            if (!double.IsNaN(hdd) && hdd > 0) adj -= Math.Min(0.08, (hdd / 4000.0) * 0.08);
            if (!double.IsNaN(cdd) && cdd > 0) adj += Math.Min(0.08, (cdd / 1200.0) * 0.08);

            irr *= (1.0 + adj);

            if (irr < 800) irr = 800;
            if (irr > 1500) irr = 1500;
            return irr;
        }

        private static double TryGetMetricDouble(Dictionary<string, string> metrics, string key)
        {
            if (metrics == null || string.IsNullOrWhiteSpace(key)) return double.NaN;
            if (!metrics.TryGetValue(key, out var s) || string.IsNullOrWhiteSpace(s)) return double.NaN;
            try
            {
                return RetrofitShared.ParseDouble(s);
            }
            catch
            {
                if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
                return double.NaN;
            }
        }

        private static double GetDoubleFromEnv(string name, double def)
        {
            try
            {
                var s = Environment.GetEnvironmentVariable(name);
                if (string.IsNullOrWhiteSpace(s)) return def;
                if (double.TryParse(s.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var v)) return v;
                return def;
            }
            catch { return def; }
        }

        private static double GetTotalRoofAreaM2(Document doc, IList<Element> roofs)
        {
            if (doc == null || roofs == null || roofs.Count == 0) return 0;

            double totalFt2 = 0.0;
            foreach (var r in roofs)
            {
                try
                {
                    var p = r.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED);
                    if (p != null && p.StorageType == StorageType.Double)
                    {
                        double a = p.AsDouble();
                        if (a > 0) totalFt2 += a;
                    }
                }
                catch { }
            }

            try
            {
                return UnitUtils.ConvertFromInternalUnits(totalFt2, UnitTypeId.SquareMeters);
            }
            catch
            {
                return 0;
            }
        }
    }
}
