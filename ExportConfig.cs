using System;
using System.Collections.Generic;

namespace RevitLOD3Exporter
{
    /// <summary>
    /// Configuration used to store “what the user chooses to export”:
    /// - Which CityJSON types need to be retained (WallSurface / RoofSurface / Opening, etc.)
    /// - Which attributes need to be retained (class / category / revitElementId, etc.)
    /// </summary>
    public class ExportConfig
    {
        /// <summary>
        /// CityObject types permitted for export
        /// e.g.: WallSurface, RoofSurface, Opening ...
        /// Empty set = No type filtering (all types exported)
        /// </summary>
        public HashSet<string> AllowedTypes { get; set; } = new HashSet<string>();

        /// <summary>
        /// Permitted attribute names (keys within the attributes dictionary).
        /// e.g.：class, category, revitElementId ...
        /// Empty set = No type filtering (all types exported)
        /// </summary>
        public List<string> SelectedAttributes { get; set; } = new List<string>();

    }
}
