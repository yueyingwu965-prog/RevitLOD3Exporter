using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitLOD3Exporter
{
    public static class RetrofitFloorApplier
    {
        private const string PARAM_BUILDING_ID = "building_id";
        private const int BATCH_SIZE = 30;

        public static void Apply(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            var mapByBuilding = BuildInsulatedFloorTypeMapping();

            // FloorType
            var floorTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(FloorType))
                .Cast<FloorType>()
                .GroupBy(ft => ft.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            // Floors
            var floors = new FilteredElementCollector(doc)
                .OfClass(typeof(Floor))
                .WhereElementIsNotElementType()
                .Cast<Floor>()
                .ToList();

            int totalFloors = floors.Count;
            int skippedNoBuildingId = 0;
            int skippedBuildingNotMapped = 0;
            int skippedTypeNotInPlan = 0;
            int skippedTargetTypeMissing = 0;
            int alreadyTarget = 0;

            var samples = new List<string>();

            var toChange = new List<FloorChangeItem>();

            foreach (var fl in floors)
            {
                string buildingId = fl.LookupParameter(PARAM_BUILDING_ID)?.AsString();
                if (string.IsNullOrWhiteSpace(buildingId))
                {
                    skippedNoBuildingId++;
                    if (samples.Count < 6) samples.Add($"Floor {fl.Id}: building_id empty");
                    continue;
                }
                buildingId = buildingId.Trim().Trim('\'').Trim('"');

                if (!mapByBuilding.TryGetValue(buildingId, out var typeMap))
                {
                    skippedBuildingNotMapped++;
                    if (samples.Count < 6) samples.Add($"Floor {fl.Id}: building_id='{buildingId}' not mapped");
                    continue;
                }

                string curTypeName = (doc.GetElement(fl.GetTypeId()) as ElementType)?.Name ?? "";
                curTypeName = curTypeName.Trim();

                if (!typeMap.TryGetValue(curTypeName, out string targetTypeName) || string.IsNullOrWhiteSpace(targetTypeName))
                {
                    skippedTypeNotInPlan++;
                    continue;
                }
                targetTypeName = targetTypeName.Trim();

                if (!floorTypes.TryGetValue(targetTypeName, out FloorType targetType))
                {
                    skippedTargetTypeMissing++;
                    if (samples.Count < 6) samples.Add($"Floor {fl.Id}: target type '{targetTypeName}' NOT found");
                    continue;
                }

                if (fl.GetTypeId() == targetType.Id)
                {
                    alreadyTarget++;
                    continue;
                }

                var snap = CaptureInstancePositionSnapshot(fl);

                toChange.Add(new FloorChangeItem
                {
                    FloorId = fl.Id,
                    TargetTypeId = targetType.Id,
                    FromTypeName = curTypeName,
                    ToTypeName = targetTypeName,
                    Snapshot = snap
                });
            }
            
            int changed = 0;
            int idx = 0;

            while (idx < toChange.Count)
            {
                int take = Math.Min(BATCH_SIZE, toChange.Count - idx);

                using (Transaction t = new Transaction(doc, "EA Retrofit - Floors (Insulated)"))
                {
                    t.Start();

                    for (int i = 0; i < take; i++)
                    {
                        var item = toChange[idx + i];
                        var fl = doc.GetElement(item.FloorId) as Floor;
                        if (fl == null) continue;

                        if (fl.GetTypeId() == item.TargetTypeId)
                        {
                            alreadyTarget++;
                            continue;
                        }

                        fl.ChangeTypeId(item.TargetTypeId);

                        RestoreInstancePositionSnapshot(fl, item.Snapshot);

                        changed++;

                        if (samples.Count < 6)
                            samples.Add($"Floor {fl.Id}: '{item.FromTypeName}' -> '{item.ToTypeName}'");
                    }

                    t.Commit();
                }

                idx += take;
            }

            // Report
            string report =
                $"Total floors: {totalFloors}\n" +
                $"Planned changes: {toChange.Count}\n" +
                $"Changed: {changed}\n" +
                $"Already target: {alreadyTarget}\n\n" +
                $"Skipped (no building_id): {skippedNoBuildingId}\n" +
                $"Skipped (building_id not mapped): {skippedBuildingNotMapped}\n" +
                $"Skipped (type not in update plan): {skippedTypeNotInPlan}\n" +
                $"Skipped (target type missing): {skippedTargetTypeMissing}\n\n" +
                $"Samples:\n- " + string.Join("\n- ", samples);

            TaskDialog.Show("Retrofit - Floors", report);
        }

        // =========================
        // ✅ 位置快照：Ebene / Offset / Room Bounding
        // =========================
        private class FloorInstanceSnapshot
        {
            public ElementId LevelId { get; set; } = ElementId.InvalidElementId;
            public double? LevelOffsetInternal { get; set; } = null; // internal units
            public int? RoomBoundingInt { get; set; } = null;        // 0/1
        }

        private class FloorChangeItem
        {
            public ElementId FloorId { get; set; }
            public ElementId TargetTypeId { get; set; }
            public string FromTypeName { get; set; }
            public string ToTypeName { get; set; }
            public FloorInstanceSnapshot Snapshot { get; set; }
        }

        private static FloorInstanceSnapshot CaptureInstancePositionSnapshot(Floor fl)
        {
            var s = new FloorInstanceSnapshot();

            try
            {
                // Ebene (Level)
                if (fl.LevelId != null && fl.LevelId != ElementId.InvalidElementId)
                    s.LevelId = fl.LevelId;
            }
            catch { }

            try
            {
                // Höhenversatz von Ebene (Offset)
                var pOff = fl.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                if (pOff != null && pOff.StorageType == StorageType.Double)
                    s.LevelOffsetInternal = pOff.AsDouble();
            }
            catch { }

            try
            {
                // ✅ Raumbegrenzung（Floor 没有稳定 BuiltInParameter，只能按名字）
                Parameter pRb =
                    fl.LookupParameter("Raumbegrenzung") ??
                    fl.LookupParameter("Room Bounding");

                if (pRb != null && pRb.StorageType == StorageType.Integer)
                    s.RoomBoundingInt = pRb.AsInteger();
            }
            catch { }

            return s;
        }


        private static void RestoreInstancePositionSnapshot(Floor fl, FloorInstanceSnapshot s)
        {
            if (fl == null || s == null) return;

            try
            {
                // Offset 写回
                if (s.LevelOffsetInternal.HasValue)
                {
                    var pOff = fl.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                    if (pOff != null && !pOff.IsReadOnly && pOff.StorageType == StorageType.Double)
                        pOff.Set(s.LevelOffsetInternal.Value);
                }
            }
            catch { }

            try
            {
                // ✅ Raumbegrenzung 写回（按名字）
                if (s.RoomBoundingInt.HasValue)
                {
                    Parameter pRb =
                        fl.LookupParameter("Raumbegrenzung") ??
                        fl.LookupParameter("Room Bounding");

                    if (pRb != null && !pRb.IsReadOnly && pRb.StorageType == StorageType.Integer)
                        pRb.Set(s.RoomBoundingInt.Value);
                }
            }
            catch { }
        }



        // =========================
        // ✅ 按你表格：Existing -> Insulated（Building1）
        // =========================
        private static Dictionary<string, Dictionary<string, string>> BuildInsulatedFloorTypeMapping()
        {
            return new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "building_0001",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        { "STB 200",                                   "STB 200 Insulated" },
                        { "STB 200 (WU Beton) + Abdichtung + Perimeter", "STB 200 (WU Beton) + Abdichtung + Perimeter Insulated" },
                        { "FB 150 Vinyl",                               "FB 150 Vinyl Insulated" },
                        { "FB 150 Fliese Grau 300 x 300",               "FB 150 Fliese Grau 300 x 300 Insulated" },
                        { "FB 240 Dachterrasse Holz (Dominic)",         "FB 240 Dachterrasse Holz Insulated" },                        
                    }
                },
                {
                    "building_0002",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        { "FB 150 leer",                                 "FB 150 leer Insulated" },
                        { "FB 150 Terrasse Holz",                        "FB 150 Terrasse Holz Insulated" }
                    }
                }
            };
        }
    }
}
