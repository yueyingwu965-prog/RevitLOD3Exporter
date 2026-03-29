using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace RevitLOD3Exporter
{
    /// <summary>
    /// AI 返回的配置结果：types + attributes
    /// </summary>
    public class AiConfigResult
    {
        [JsonProperty("types")]
        public List<string> Types { get; set; } = new List<string>();

        [JsonProperty("attributes")]
        public List<string> Attributes { get; set; } = new List<string>();
    }

    public enum AiProvider
    {
        OpenAI,
        DeepSeek,
        Gemini
    }

    public class AiExportHelper : IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;
        private readonly string _model;
        private readonly AiProvider _provider;

        private const string OpenAiDefaultModel = "gpt-4.1-mini";
        private const string DeepSeekDefaultModel = "deepseek-chat";
        private const string GeminiDefaultModel = "gemini-1.5-flash";

        public AiExportHelper(AiProvider provider, string apiKey, string model = null)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new ArgumentException("API key is required.", nameof(apiKey));

            _provider = provider;
            _apiKey = apiKey;

            switch (_provider)
            {
                case AiProvider.OpenAI:
                    _model = string.IsNullOrWhiteSpace(model) ? OpenAiDefaultModel : model;
                    break;
                case AiProvider.DeepSeek:
                    _model = string.IsNullOrWhiteSpace(model) ? DeepSeekDefaultModel : model;
                    break;
                case AiProvider.Gemini:
                    _model = string.IsNullOrWhiteSpace(model) ? GeminiDefaultModel : model;
                    break;
            }

            _httpClient = new HttpClient();

            if (_provider == AiProvider.OpenAI)
            {
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}");
            }
            else if (_provider == AiProvider.DeepSeek)
            {
                _httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {_apiKey}");
            }
        }


        public async Task<AiConfigResult> ParseUserTextAsync(string userText)
        {
            if (string.IsNullOrWhiteSpace(userText))
                return new AiConfigResult();

            // 共用的 system prompt：要求模型输出 JSON
            string systemPrompt = @"
You are a parser that converts the user's natural language into a JSON configuration for a CityJSON LOD3 export tool.

The user can mention:
- which geometry types they want to keep (e.g. walls, roofs, openings, windows, doors, ceilings, floors, buildings...)
- which attributes they want to keep (e.g. class, category, revitElementId, element_id, originalName, name, function...)

The output MUST be valid JSON with this exact shape:

{
  ""types"": [ ""WallSurface"", ""RoofSurface"" ],
  ""attributes"": [ ""class"", ""category"", ""revitElementId"" ]
}

Rules:
- Use these canonical CityJSON type names when possible:
  - walls → ""WallSurface""
  - wall → ""WallSurface""
  - roofs / roof → ""RoofSurface""
  - ground / floor / terrain → ""GroundSurface""
  - openings / doors / windows / door / window → ""Opening""
  - building / buildings → ""Building""
  - ceiling / ceilings → ""CeilingSurface""

- For attributes, use these canonical attribute names and mappings:

  • Element ID (Revit element identifier)
    When the user refers to the element ID in any way, such as:
      - ""element id"", ""element-id"", ""elementid""
      - ""id of element"", ""id of the element""
      - ""revit id"", ""revit element id"", ""ID del elemento"", ""元素ID""
    you MUST include **both** of these attribute names in the JSON:
      - ""revitElementId""
      - ""element_id""

  • Original name:
      - ""originalName"", ""original name"", ""name in revit"", ""object name""
      → use ""originalName""

  • Class:
      - ""class"", ""usage class"", ""use class""
      → use ""class""

  • Category:
      - ""category"", ""revit category"", ""family category""
      → use ""category""

- If the user explicitly writes an attribute name that already looks like a key
  (e.g. ""revitElementId"", ""element_id"", ""originalName"", ""class"", ""category""),
  reuse it as-is in the ""attributes"" array.

- If the user says ""all types"" or doesn't specify types, return an empty list for ""types"": [].
- If the user says ""all attributes"" or doesn't specify attributes, return an empty list for ""attributes"": [].

- Do NOT invent new attribute names that do not appear in the user's request or in the canonical list above.
- Do NOT add any other fields.
- Do NOT write any explanation.
- Output JSON only.
";

            try
            {
                string messageContent;

                switch (_provider)
                {
                    case AiProvider.OpenAI:
                        messageContent = await CallOpenAiAsync(systemPrompt, userText);
                        break;

                    case AiProvider.DeepSeek:
                        messageContent = await CallDeepSeekAsync(systemPrompt, userText);
                        break;

                    case AiProvider.Gemini:
                        messageContent = await CallGeminiAsync(systemPrompt, userText);
                        break;

                    default:
                        return new AiConfigResult();
                }

                AiConfigResult parsed;
                try
                {
                    parsed = JsonConvert.DeserializeObject<AiConfigResult>(messageContent);
                }
                catch
                {

                    return new AiConfigResult();
                }

                if (parsed == null)
                    parsed = new AiConfigResult();

                parsed.Types = new List<string>(
                    new HashSet<string>(parsed.Types ?? new List<string>(), StringComparer.OrdinalIgnoreCase));
                parsed.Attributes = new List<string>(
                    new HashSet<string>(parsed.Attributes ?? new List<string>(), StringComparer.OrdinalIgnoreCase));

                return parsed;
            }
            catch (Exception ex)
            {

                throw new Exception($"AI provider '{_provider}' failed: {ex.Message}", ex);
            }
        }


        private async Task<string> CallOpenAiAsync(string systemPrompt, string userText)
        {
            var requestBody = new
            {
                model = _model,
                messages = new[]
                {
                    new { role = "system", content = systemPrompt },
                    new { role = "user",   content = userText   }
                },
                temperature = 0.0
            };

            string jsonBody = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("https://api.openai.com/v1/chat/completions", content);
            string responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"OpenAI API error {response.StatusCode}: {responseJson}");
            }

            dynamic result = JsonConvert.DeserializeObject(responseJson);
            string messageContent = result.choices[0].message.content.ToString().Trim();
            return messageContent;
        }

        private async Task<string> CallDeepSeekAsync(string systemPrompt, string userText)
        {
            var requestBody = new
            {
                model = _model, // deepseek-chat
                messages = new[]
                {
                    new { role = "system", content = systemPrompt },
                    new { role = "user",   content = userText   }
                },
                temperature = 0.0
            };

            string jsonBody = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            // DeepSeek API endpoint（OpenAI 风格）
            var response = await _httpClient.PostAsync("https://api.deepseek.com/chat/completions", content);
            string responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"DeepSeek API error {response.StatusCode}: {responseJson}");
            }

            dynamic result = JsonConvert.DeserializeObject(responseJson);
            string messageContent = result.choices[0].message.content.ToString().Trim();
            return messageContent;
        }

        private async Task<string> CallGeminiAsync(string systemPrompt, string userText)
        {
            // Gemini 的 prompt 结构稍微不一样，我们把 system + user 合并成一段文本
            var requestBody = new
            {
                contents = new[]
                {
                    new
                    {
                        parts = new[]
                        {
                            new { text = systemPrompt + "\n\nUser input:\n" + userText }
                        }
                    }
                }
            };

            string jsonBody = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");

            // Gemini 用 URL query 传 key
            string url =
                $"https://generativelanguage.googleapis.com/v1beta/models/{_model}:generateContent?key={_apiKey}";

            var response = await _httpClient.PostAsync(url, content);
            string responseJson = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"Gemini API error {response.StatusCode}: {responseJson}");
            }

            dynamic result = JsonConvert.DeserializeObject(responseJson);
            // 典型结构：candidates[0].content.parts[0].text
            string messageContent = result.candidates[0].content.parts[0].text.ToString().Trim();
            return messageContent;
        }

        public void Dispose()
        {
            _httpClient?.Dispose();
        }
    }
}
