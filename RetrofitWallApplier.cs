using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitLOD3Exporter
{
    public static class RetrofitWallApplier
    {
        private const string PARAM_BUILDING_ID = "building_id";
        private const int BATCH_SIZE = 50;

        public static void Apply(Document doc)
        {
            if (doc == null) throw new ArgumentNullException(nameof(doc));

            // 1) 映射：当前类型名 -> 对应 Insulated 类型名（按 building_id 分开）
            var mapByBuilding = BuildInsulatedTypeMapping();

            // 2) 索引所有 Basic WallType（Type.Name -> WallType）
            var wallTypes = new FilteredElementCollector(doc)
                .OfClass(typeof(WallType))
                .Cast<WallType>()
                .Where(wt => wt.Kind == WallKind.Basic)
                .GroupBy(wt => wt.Name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            // 3) 收集墙
            var walls = new FilteredElementCollector(doc)
                .OfClass(typeof(Wall))
                .WhereElementIsNotElementType()
                .Cast<Wall>()
                .ToList();

            int totalWalls = walls.Count;
            int skippedNotBasic = 0;
            int skippedNoBuildingId = 0;
            int skippedBuildingNotMapped = 0;
            int skippedTypeNotInPlan = 0;
            int skippedTargetTypeMissing = 0;
            int alreadyTarget = 0;

            var samples = new List<string>();

            // 4) 事务外先算出要改的列表（减少事务内耗时）
            var toChange = new List<(ElementId wallId, ElementId targetTypeId, string fromName, string toName)>();

            foreach (var wall in walls)
            {
                WallType curType = wall.WallType;
                if (curType == null || curType.Kind != WallKind.Basic)
                {
                    skippedNotBasic++;
                    continue;
                }

                // building_id
                string buildingId = wall.LookupParameter(PARAM_BUILDING_ID)?.AsString();
                if (string.IsNullOrWhiteSpace(buildingId))
                {
                    skippedNoBuildingId++;
                    if (samples.Count < 6) samples.Add($"Wall {wall.Id}: building_id empty");
                    continue;
                }
                buildingId = buildingId.Trim().Trim('\'').Trim('"');

                if (!mapByBuilding.TryGetValue(buildingId, out var typeMap))
                {
                    skippedBuildingNotMapped++;
                    if (samples.Count < 6) samples.Add($"Wall {wall.Id}: building_id='{buildingId}' not mapped");
                    continue;
                }

                // 当前类型名（Type.Name，不含 Basiswand）
                string curTypeName = (curType.Name ?? "").Trim();

                // 只改你计划内的外墙类型
                if (!typeMap.TryGetValue(curTypeName, out string targetTypeName) || string.IsNullOrWhiteSpace(targetTypeName))
                {
                    skippedTypeNotInPlan++;
                    continue;
                }

                targetTypeName = targetTypeName.Trim();

                // 找目标类型
                if (!wallTypes.TryGetValue(targetTypeName, out WallType targetType))
                {
                    skippedTargetTypeMissing++;
                    if (samples.Count < 6) samples.Add($"Wall {wall.Id}: target type '{targetTypeName}' NOT found");
                    continue;
                }

                if (wall.GetTypeId() == targetType.Id)
                {
                    alreadyTarget++;
                    continue;
                }

                toChange.Add((wall.Id, targetType.Id, curTypeName, targetTypeName));
            }

            // 5) 分批修改（减少卡顿）
            int changed = 0;

            int idx = 0;
            while (idx < toChange.Count)
            {
                int take = Math.Min(BATCH_SIZE, toChange.Count - idx);

                using (Transaction t = new Transaction(doc, "EA Retrofit - Walls (Insulated)"))
                {
                    t.Start();

                    for (int i = 0; i < take; i++)
                    {
                        var item = toChange[idx + i];
                        var wall = doc.GetElement(item.wallId) as Wall;
                        if (wall == null) continue;

                        if (wall.GetTypeId() == item.targetTypeId)
                        {
                            alreadyTarget++;
                            continue;
                        }

                        wall.ChangeTypeId(item.targetTypeId);
                        changed++;

                        if (samples.Count < 6)
                            samples.Add($"Wall {wall.Id}: '{item.fromName}' -> '{item.toName}'");
                    }

                    t.Commit();
                }

                idx += take;
            }

            // 6) Report
            string report =
                $"Total walls: {totalWalls}\n" +
                $"Planned changes: {toChange.Count}\n" +
                $"Changed: {changed}\n" +
                $"Already target: {alreadyTarget}\n\n" +
                $"Skipped (not Basic): {skippedNotBasic}\n" +
                $"Skipped (no building_id): {skippedNoBuildingId}\n" +
                $"Skipped (building_id not mapped): {skippedBuildingNotMapped}\n" +
                $"Skipped (type not in update plan): {skippedTypeNotInPlan}\n" +
                $"Skipped (target type missing): {skippedTargetTypeMissing}\n\n" +
                $"Samples:\n- " + string.Join("\n- ", samples);

            TaskDialog.Show("Retrofit - Walls", report);
        }

        /// <summary>
        /// ✅ 一对一：原外墙类型 -> 对应 Insulated 类型（厚度不变）
        /// 注意：这里写的是 Type.Name（不带“Basiswand”族名）
        /// </summary>
        private static Dictionary<string, Dictionary<string, string>> BuildInsulatedTypeMapping()
        {
            return new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
            {
                {
                    "building_0001",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        // Building1 EXISTING -> Insulated（你表里：每一种墙都有一个 Insulated 版本）
                        { "Ziegel+WD hart 240+160",                 "Ziegel+WD hart 240+160 Insulated" },
                        { "Ziegel+WD hart 240+160 mit Fliesen",     "Ziegel+WD hart 240+160 mit Fliesen Insulated" },
                        { "Ziegel 200 beidseitig verputzt",         "Ziegel 200 beidseitig verputzt Insulated" },
                        { "Ziegel 200 gefliest und verputzt",       "Ziegel 200 gefliest und verputzt Insulated" }
                    }
                },
                {
                    "building_0002",
                    new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        // Building2 EXISTING -> Insulated
                        { "Ziegel 200",                             "Ziegel 200 Insulated" },
                        { "Ziegel 240",                             "Ziegel 240 Insulated" },
                        { "Ziegel+WD hart 200+40",                  "Ziegel+WD hart 200+40 Insulated" },
                        { "Ziegel+WD hart 310+40",                  "Ziegel+WD hart 310+40 Insulated" }
                    }
                }
            };
        }
    }
}
