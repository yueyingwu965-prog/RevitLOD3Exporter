using System;
using System.Collections.Generic;
using System.Linq;

namespace RevitLOD3Exporter
{
    /// <summary>
    /// Parse the user-input sentence into an ExportConfig
    /// Example Input:
    ///   "walls, roofs; attrs: class, revitElementId"
    /// </summary>
    public static class ExportConfigParser
    {
        // --- 1) 类型关键词映射：用户说什么 → CityJSON type ---
        private static readonly Dictionary<string, string> TypeKeywordMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "wall", "WallSurface" },
            { "walls", "WallSurface" },

            { "roof", "RoofSurface" },
            { "roofs", "RoofSurface" },

            { "ground", "GroundSurface" },

            { "opening", "Opening" },
            { "openings", "Opening" },
            { "door", "Opening" },
            { "doors", "Opening" },
            { "window", "Opening" },
            { "windows", "Opening" },
        };

        // --- 2) 属性关键词映射：用户说什么 → attributes 里的 key ---
        private static readonly Dictionary<string, string> AttrKeywordMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "id", "revitElementId" },
            { "elementid", "revitElementId" },
            { "revitid", "revitElementId" },

            { "class", "class" },
            { "category", "category" },

            { "originalname", "originalName" },
            { "name", "originalName" },
        };

        /// <summary>
        /// 主解析函数：输入一行用户文本，输出 ExportConfig。
        /// 推荐用户输入格式：
        ///   "walls, roofs; attrs: class, revitElementId"
        /// </summary>
        public static ExportConfig ParseUserInputToConfig(string userInput)
        {
            var config = new ExportConfig();

            if (string.IsNullOrWhiteSpace(userInput))
                return config; // 空配置 = 不做过滤

            // 按 ; 把“类型部分”和“属性部分”拆开
            string[] parts = userInput.Split(
                new[] { ';' },
                StringSplitOptions.RemoveEmptyEntries);

            foreach (var rawPart in parts)
            {
                string part = rawPart.Trim();

                // ---------- 属性部分 ----------
                if (part.IndexOf("attr", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    part.Contains("属性"))
                {
                    // 找冒号，把 "attrs:" 去掉
                    int colonIndex = part.IndexOf(':');
                    string attrList = colonIndex >= 0
                        ? part.Substring(colonIndex + 1)
                        : part;

                    string[] tokens = attrList.Split(
                        new[] { ',', '，', ' ' },
                        StringSplitOptions.RemoveEmptyEntries);

                    foreach (var tk in tokens)
                    {
                        string key = tk.Trim();
                        string attrName;

                        if (AttrKeywordMap.TryGetValue(key, out attrName))
                        {
                            // 直接往 SelectedAttributes 里加
                            config.SelectedAttributes.Add(attrName);
                        }
                        else
                        {
                            // 用户可能直接写了 "revitElementId" 这样的原始 key
                            config.SelectedAttributes.Add(key);
                        }
                    }
                }
                // ---------- 类型部分 ----------
                else
                {
                    string[] tokens = part.Split(
                        new[] { ',', '，', ' ' },
                        StringSplitOptions.RemoveEmptyEntries);

                    foreach (var tk in tokens)
                    {
                        string key = tk.Trim();
                        if (TypeKeywordMap.TryGetValue(key, out string typeName))
                        {
                            config.AllowedTypes.Add(typeName);
                        }
                        else
                        {
                            // 用户直接写 CityJSON type，例如 "WallSurface"
                            config.AllowedTypes.Add(key);
                        }
                    }
                }
            }

            // 对属性做一次去重（忽略大小写），并对 element id 做归一化
            if (config.SelectedAttributes != null && config.SelectedAttributes.Count > 0)
            {
                var attrSet = new HashSet<string>(config.SelectedAttributes, StringComparer.OrdinalIgnoreCase);

                // 如果用户选择了任意一种“元素ID字段”，就强制同时保留这两个：
                //  - revitElementId
                //  - element_id
                if (attrSet.Contains("revitElementId") || attrSet.Contains("element_id"))
                {
                    attrSet.Add("revitElementId");
                    attrSet.Add("element_id");
                }

                config.SelectedAttributes = attrSet.ToList();
            }

            return config;
        }
    }
}
