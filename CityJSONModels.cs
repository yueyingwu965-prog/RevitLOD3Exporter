using Newtonsoft.Json;
using System.Collections.Generic;

namespace RevitLOD3Exporter
{
    public class CityJSONObject
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("vertices")]
        public List<List<double>> Vertices { get; set; }

        [JsonProperty("geometry")]
        public List<CityJSONGeometry> Geometry { get; set; }

        [JsonProperty("attributes")]
        public Dictionary<string, object> Attributes { get; set; }
    }

    public class CityJSONGeometry
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("lod")]
        public string Lod { get; set; }

        [JsonProperty("boundaries")]
        public List<object> Boundaries { get; set; }

        [JsonProperty("semantics")]
        public CityJSONSemantics Semantics { get; set; }
    }

    public class CityJSONSemantics
    {
        [JsonProperty("values")]
        public List<int> Values { get; set; }

        [JsonProperty("surfaces")]
        public List<CityJSONSurface> Surfaces { get; set; }
    }

    public class CityJSONSurface
    {
        [JsonProperty("type")]
        public string Type { get; set; }
    }

    public class CityJSONRoot
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("CityObjects")]
        public Dictionary<string, CityJSONObject> CityObjects { get; set; }

        [JsonProperty("vertices")]
        public List<List<double>> Vertices { get; set; }

        [JsonProperty("transform")]
        public CityJSONTransform Transform { get; set; }
    }

    public class CityJSONTransform
    {
        [JsonProperty("scale")]
        public List<double> Scale { get; set; } = new List<double> { 1.0, 1.0, 1.0 };

        [JsonProperty("translate")]
        public List<double> Translate { get; set; } = new List<double> { 0.0, 0.0, 0.0 };
    }
}