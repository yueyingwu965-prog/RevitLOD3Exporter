using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using System.Drawing;

namespace RevitLOD3Exporter
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class ExportLOD3 : IExternalCommand
    {
        private string defaultLod2FilePath = @"D:\A. Master thesis\FME sample\building_lod2.city.json";
        private string defaultLod3OutputPath = @"D:\A. Master thesis\FME sample\building_lod3_revit.city.json";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                string lod2FilePath = SelectInputPath();
                if (string.IsNullOrEmpty(lod2FilePath))
                {
                    TaskDialog.Show("Cancelled", "LOD3 export operation was cancelled by user!");
                    return Result.Cancelled;
                }

                string lod3OutputPath = SelectOutputPath();
                if (string.IsNullOrEmpty(lod3OutputPath))
                {
                    TaskDialog.Show("Cancelled", "LOD3 export operation was cancelled by user!");
                    return Result.Cancelled;
                }

                string introMessage =
                    "✔ Plugin Loaded Successfully!\n" +
                    "✔ Input File: " + Path.GetFileName(lod2FilePath) + "\n" +
                    "✔ Output File: " + Path.GetFileName(lod3OutputPath) + "\n\n" +
                    "Starting LOD2 to LOD3 conversion process...";

                using (var infoForm = new AutoCloseInfoForm(introMessage, 3000))
                {
                    infoForm.ShowDialog();
                }

                CityJSONData lod2Data = ReadLOD2File(lod2FilePath);
                if (lod2Data == null)
                {
                    TaskDialog.Show("Error", "Failed to read LOD2 file: " + lod2FilePath);
                    return Result.Failed;
                }

                ExportConfig config;
                using (var chatForm = new ChatbotForm(lod2Data))
                {
                    DialogResult dr = chatForm.ShowDialog();
                    if (dr != DialogResult.OK)
                    {
                        TaskDialog.Show("Cancelled", "LOD3 export cancelled by user in chatbot window.");
                        return Result.Cancelled;
                    }

                    config = chatForm.SelectedConfig ?? new ExportConfig();
                }

                CityJSONData lod3Data = ExtractLOD3Data(doc, lod2Data, config);

                SaveLOD3File(lod3Data, lod3OutputPath);

                TaskDialog.Show("Success",
                    "✅ LOD3 Export Completed Successfully!\n\n" +
                    $"Input File: {Path.GetFileName(lod2FilePath)}\n" +
                    $"Output File: {Path.GetFileName(lod3OutputPath)}\n" +
                    $"Objects Processed: {lod3Data.CityObjects?.Count ?? 0}\n" +
                    $"Total Vertices: {lod3Data.Vertices?.Count ?? 0}");

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = $"Error: {ex.Message}\n\nStack Trace:\n{ex.StackTrace}";
                TaskDialog.Show("Error", $"Plugin execution failed: {ex.Message}\n\n{ex.StackTrace}");
                return Result.Failed;
            }
        }

        private CityJSONData ExtractLOD3Data(Document doc, CityJSONData lod2Data, ExportConfig config)
        {
            CityJSONData lod3Data = ExtractLOD3Data(doc, lod2Data);

            if (config == null)
                return lod3Data;

            bool filterTypes = config.AllowedTypes != null && config.AllowedTypes.Count > 0;

            var attrFilterSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (config.SelectedAttributes != null)
                attrFilterSet.UnionWith(config.SelectedAttributes);
            bool filterAttrs = attrFilterSet.Count > 0;

            if (!filterTypes && !filterAttrs)
                return lod3Data;

            if (lod3Data.CityObjects == null)
                return lod3Data;

            if (filterTypes)
            {
                var keysToRemove = new List<string>();

                foreach (var kv in lod3Data.CityObjects)
                {
                    string id = kv.Key;
                    CityJSONObj obj = kv.Value;
                    string type = obj.Type ?? string.Empty;

                    if (!config.AllowedTypes.Contains(type))
                    {
                        keysToRemove.Add(id);
                    }
                }

                foreach (var id in keysToRemove)
                {
                    lod3Data.CityObjects.Remove(id);
                }
            }

            if (filterAttrs)
            {
                var alwaysKeep = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "lod",
                    "cityjsonObjectId",
                    "element_id",
                    "revitBoundingHeight",
                    "revitCategory",
                    "revitFamily",
                };

                foreach (var kv in lod3Data.CityObjects)
                {
                    CityJSONObj obj = kv.Value;
                    if (obj.Attributes == null)
                        continue;

                    var newAttrs = new Dictionary<string, object>();

                    foreach (var attr in obj.Attributes)
                    {
                        string key = attr.Key;
                        if (alwaysKeep.Contains(key) || attrFilterSet.Contains(key))
                        {
                            newAttrs[key] = attr.Value;
                        }
                    }

                    obj.Attributes = newAttrs;
                }
            }

            return lod3Data;
        }


        private string SelectInputPath()
        {
            TaskDialog inputDialog = new TaskDialog("LOD3 Export - Input File Selection");
            inputDialog.MainInstruction = "Select LOD2 input file";
            inputDialog.MainContent = $"Default input path:\n{defaultLod2FilePath}";

            inputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                "✅ Use Default Input Path",
                $"Use: {Path.GetFileName(defaultLod2FilePath)}");
            inputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                "📁 Select Custom Input File",
                "Choose a different LOD2 CityJSON file");
            inputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                "❌ Cancel",
                "Cancel the export operation");

            inputDialog.CommonButtons = TaskDialogCommonButtons.Close;
            inputDialog.DefaultButton = TaskDialogResult.CommandLink1;

            TaskDialogResult inputResult = inputDialog.Show();

            switch (inputResult)
            {
                case TaskDialogResult.CommandLink1:
                    return defaultLod2FilePath;

                case TaskDialogResult.CommandLink2:
                    string inputPath = SelectFileDialog("Select LOD2 CityJSON file", "CityJSON files|*.city.json;*.json|All files|*.*");
                    return !string.IsNullOrEmpty(inputPath) ? inputPath : null;

                case TaskDialogResult.CommandLink3:
                case TaskDialogResult.Close:
                case TaskDialogResult.Cancel:
                    return null;

                default:
                    return null;
            }
        }

        private string SelectOutputPath()
        {
            TaskDialog outputDialog = new TaskDialog("LOD3 Export - Output File Selection");
            outputDialog.MainInstruction = "Select LOD3 output file";
            outputDialog.MainContent = $"Default output path:\n{defaultLod3OutputPath}";

            outputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                "✅ Use Default Output Path",
                $"Use: {Path.GetFileName(defaultLod3OutputPath)}");
            outputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2,
                "💾 Select Custom Output File",
                "Choose a different location for LOD3 CityJSON file");
            outputDialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink3,
                "❌ Cancel",
                "Cancel the export operation");

            outputDialog.CommonButtons = TaskDialogCommonButtons.Close;
            outputDialog.DefaultButton = TaskDialogResult.CommandLink1;

            TaskDialogResult outputResult = outputDialog.Show();

            switch (outputResult)
            {
                case TaskDialogResult.CommandLink1:
                    return defaultLod3OutputPath;

                case TaskDialogResult.CommandLink2:
                    string outputPath = SaveFileDialog("Save LOD3 CityJSON file", "CityJSON files|*.city.json;*.json|All files|*.*");
                    return !string.IsNullOrEmpty(outputPath) ? outputPath : null;

                case TaskDialogResult.CommandLink3:
                case TaskDialogResult.Close:
                case TaskDialogResult.Cancel:
                    return null;

                default:
                    return null;
            }
        }

        private string SelectFileDialog(string title, string filter)
        {
            try
            {
                var openDialog = new System.Windows.Forms.OpenFileDialog();
                openDialog.Title = title;
                openDialog.Filter = filter;
                openDialog.DefaultExt = "city.json";
                openDialog.CheckFileExists = true;

                if (openDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return openDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("File Selection Error",
                    $"Could not open file dialog: {ex.Message}\n\nUsing default input path.");
                return defaultLod2FilePath;
            }

            return null;
        }

        private string SaveFileDialog(string title, string filter)
        {
            try
            {
                var saveDialog = new System.Windows.Forms.SaveFileDialog();
                saveDialog.Title = title;
                saveDialog.Filter = filter;
                saveDialog.DefaultExt = "city.json";
                saveDialog.FileName = "building_lod3_revit.city.json";
                saveDialog.OverwritePrompt = true;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    return saveDialog.FileName;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("File Selection Error",
                    $"Could not open save dialog: {ex.Message}\n\nUsing default output path.");
                return defaultLod3OutputPath;
            }

            return null;
        }

        private CityJSONData ReadLOD2File(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    TaskDialog.Show("Error", $"File not found: {filePath}");
                    return null;
                }

                string jsonContent = File.ReadAllText(filePath);
                return JsonConvert.DeserializeObject<CityJSONData>(jsonContent);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Read Error", $"Failed to read LOD2 file: {ex.Message}");
                return null;
            }
        }

        private void SaveLOD3File(CityJSONData lod3Data, string outputPath)
        {
            try
            {
                string directory = Path.GetDirectoryName(outputPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (lod3Data?.CityObjects != null)
                {
                    foreach (var cityObject in lod3Data.CityObjects.Values)
                    {
                        if (cityObject.Attributes != null)
                        {
                            var attributesToRemove = new List<string>
                            {
                                "revitType",
                                "category",
                                "revitPluginVersion",
                                "revitDocument",
                                "extractionDate",
                                "dataSource",
                                "software",
                                "building_id",
                                "building",
                                "sem_type"
                            };

                            foreach (var attr in attributesToRemove)
                            {
                                if (cityObject.Attributes.ContainsKey(attr))
                                {
                                    cityObject.Attributes.Remove(attr);
                                }
                            }
                        }
                    }
                }

                JsonSerializerSettings settings = new JsonSerializerSettings
                {
                    Formatting = Formatting.Indented,
                    NullValueHandling = NullValueHandling.Ignore,
                    DefaultValueHandling = DefaultValueHandling.Ignore
                };

                string jsonString = JsonConvert.SerializeObject(lod3Data, settings);
                File.WriteAllText(outputPath, jsonString);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to save LOD3 file: {ex.Message}");
            }
        }


        private CityJSONData ExtractLOD3Data(Document doc, CityJSONData lod2Data)
        {
            if (lod2Data.Vertices == null || lod2Data.Vertices.Count == 0)
            {
                TaskDialog.Show("Warning", "LOD2 file contains no vertices. Creating sample geometry.");
                lod2Data.Vertices = CreateSampleVertices();
            }

            if (lod2Data.CityObjects == null || lod2Data.CityObjects.Count == 0)
            {
                TaskDialog.Show("Warning", "LOD2 file contains no CityObjects. Creating sample building.");
                lod2Data.CityObjects = CreateSampleCityObjects();
            }

            CityJSONMeta lod3Meta;
            if (lod2Data.Metadata != null)
            {
                lod3Meta = lod2Data.Metadata;
                if (lod3Meta.GeographicalExtent == null ||
                    lod3Meta.GeographicalExtent.Count != 6)
                {
                    lod3Meta.GeographicalExtent = CalculateExtent(lod2Data.Vertices);
                }
            }
            else
            {
                lod3Meta = new CityJSONMeta
                {
                    ReferenceSystem = "EPSG::4326",
                    GeographicalExtent = CalculateExtent(lod2Data.Vertices)
                };
            }

            if (lod3Meta.RevitInfo == null)
                lod3Meta.RevitInfo = new Dictionary<string, object>();

            lod3Meta.RevitInfo["lod"] = "3";
            lod3Meta.RevitInfo["lod3ExportVersion"] = "1.0";

            CityJSONData lod3Data = new CityJSONData
            {
                Type = lod2Data.Type ?? "CityJSON",
                Version = lod2Data.Version ?? "2.0",
                CityObjects = new Dictionary<string, CityJSONObj>(),
                Vertices = new List<List<double>>(lod2Data.Vertices ?? new List<List<double>>()),
                Transform = lod2Data.Transform ?? new CityJSONTrans(),
                Metadata = lod3Meta
            };

            int processedCount = 0;

            foreach (var entry in lod2Data.CityObjects)
            {
                string objId = entry.Key;
                CityJSONObj lod2Obj = entry.Value;

                CityJSONObj lod3Obj = new CityJSONObj
                {
                    Type = lod2Obj.Type,
                    Attributes = new Dictionary<string, object>(lod2Obj.Attributes ?? new Dictionary<string, object>()),
                    Geometry = new List<CityJSONGeom>(),
                    Children = lod2Obj.Children != null ? new List<string>(lod2Obj.Children) : null,
                    Parents = null
                };

                lod3Obj.Attributes["cityjsonObjectId"] = objId;

                ConvertToLOD3Geometry(lod2Obj, lod3Obj, objId);

                AddRevitAttributes(doc, lod3Obj, objId);

                if (string.Equals(lod3Obj.Type, "Building", StringComparison.OrdinalIgnoreCase))
                {
                    EnhanceBuildingAttributes(lod3Obj);
                }
                else
                {
                    EnhanceGenericAttributes(lod3Obj);
                }

                if (lod3Obj.Geometry != null && lod3Obj.Geometry.Count > 0)
                {
                    lod3Data.CityObjects[objId] = lod3Obj;
                    processedCount++;
                }
            }

            foreach (var entry in lod2Data.CityObjects)
            {
                string parentId = entry.Key;
                CityJSONObj parentObj = entry.Value;

                if (parentObj.Children == null || parentObj.Children.Count == 0)
                    continue;

                foreach (string childId in parentObj.Children)
                {
                    if (lod3Data.CityObjects.TryGetValue(childId, out CityJSONObj childObj))
                    {
                        if (childObj.Parents == null)
                            childObj.Parents = new List<string>();

                        if (!childObj.Parents.Contains(parentId))
                            childObj.Parents.Add(parentId);
                    }
                }
            }

            if (processedCount == 0)
            {
                TaskDialog.Show("Warning", "No valid geometry found. Creating sample building.");
                CreateSampleBuilding(lod3Data);
            }

            AddAnalysisAttributes(lod3Data);

            return lod3Data;
        }

        private void AddBuildingIdAttributes(CityJSONData data)
        {
            if (data?.CityObjects == null || data.CityObjects.Count == 0)
                return;

            var buildingIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var kv in data.CityObjects)
            {
                if (string.Equals(kv.Value.Type, "Building", StringComparison.OrdinalIgnoreCase))
                {
                    buildingIds.Add(kv.Key);
                }
            }

            foreach (var kv in data.CityObjects)
            {
                string objId = kv.Key;
                CityJSONObj obj = kv.Value;

                if (obj.Attributes == null)
                    obj.Attributes = new Dictionary<string, object>();

                if (string.Equals(obj.Type, "Building", StringComparison.OrdinalIgnoreCase))
                {
                    obj.Attributes["building_id"] = objId;
                }
                else if (obj.Parents != null && obj.Parents.Count > 0)
                {
                    string parentBuildingId = null;
                    foreach (string parentId in obj.Parents)
                    {
                        if (buildingIds.Contains(parentId))
                        {
                            parentBuildingId = parentId;
                            break;
                        }
                    }

                    obj.Attributes["building_id"] = parentBuildingId ?? objId;
                }
                else
                {
                    obj.Attributes["building_id"] = objId;
                }
            }
        }

        private List<List<double>> CreateSampleVertices()
        {
            return new List<List<double>>
            {
                new List<double> {0, 0, 0},
                new List<double> {10, 0, 0},
                new List<double> {10, 10, 0},
                new List<double> {0, 10, 0},
                new List<double> {0, 0, 10},
                new List<double> {10, 0, 10},
                new List<double> {10, 10, 10},
                new List<double> {0, 10, 10}
            };
        }

        private Dictionary<string, CityJSONObj> CreateSampleCityObjects()
        {
            var cityObjects = new Dictionary<string, CityJSONObj>();

            var building = new CityJSONObj
            {
                Type = "Building",
                Attributes = new Dictionary<string, object>
                {
                    {"name", "Sample Building"},
                    {"function", "residential"},
                    {"lod", "2"}
                },
                Geometry = new List<CityJSONGeom>
                {
                    new CityJSONGeom
                    {
                        Type = "Solid",
                        Lod = "2",
                        Boundaries = new List<object>
                        {
                            new List<List<List<int>>>
                            {
                                new List<List<int>>
                                {
                                    new List<int> {0, 1, 2, 3}
                                },
                                new List<List<int>>
                                {
                                    new List<int> {4, 5, 6, 7}
                                },
                                new List<List<int>>
                                {
                                    new List<int> {0, 1, 5, 4}
                                },
                                new List<List<int>>
                                {
                                    new List<int> {1, 2, 6, 5}
                                },
                                new List<List<int>>
                                {
                                    new List<int> {2, 3, 7, 6}
                                },
                                new List<List<int>>
                                {
                                    new List<int> {3, 0, 4, 7}
                                }
                            }
                        }
                    }
                }
            };

            cityObjects["sample_building_1"] = building;
            return cityObjects;
        }

        private void CreateSampleBuilding(CityJSONData lod3Data)
        {
            if (lod3Data.Vertices == null || lod3Data.Vertices.Count < 8)
            {
                lod3Data.Vertices = CreateSampleVertices();
            }

            var building = new CityJSONObj
            {
                Type = "Building",
                Attributes = new Dictionary<string, object>
                {
                    {"name", "Revit Sample Building"},
                    {"function", "residential"},
                    {"lod", "2"},
                    {"dataSource", "Revit Sample"},
                    {"measuredHeight", 10.0}
                },
                Geometry = new List<CityJSONGeom>()
            };

            EnhanceBuildingAttributes(building);

            var boundaries = new List<object>
            {
                new List<List<List<int>>>
                {
                    new List<List<int>> { new List<int> {0, 1, 2, 3} },
                    new List<List<int>> { new List<int> {4, 5, 6, 7} },
                    new List<List<int>> { new List<int> {0, 1, 5, 4} },
                    new List<List<int>> { new List<int> {1, 2, 6, 5} },
                    new List<List<int>> { new List<int> {2, 3, 7, 6} },
                    new List<List<int>> { new List<int> {3, 0, 4, 7} }
                }
            };

            boundaries = CleanBoundaries(boundaries);

            building.Geometry.Add(new CityJSONGeom
            {
                Type = "MultiSurface",
                Lod = "3",
                Boundaries = boundaries,
                Semantics = new CityJSONSemantic
                {
                    Values = new List<int> { 0, 1, 2, 3, 4, 5 },
                    Surfaces = new List<CityJSONSurf>
                    {
                        new CityJSONSurf { Type = "GroundSurface" },
                        new CityJSONSurf { Type = "RoofSurface" },
                        new CityJSONSurf { Type = "WallSurface" },
                        new CityJSONSurf { Type = "WallSurface" },
                        new CityJSONSurf { Type = "WallSurface" },
                        new CityJSONSurf { Type = "WallSurface" }
                    }
                }
            });

            lod3Data.CityObjects["sample_building_1"] = building;
        }

        private List<double> CalculateExtent(List<List<double>> vertices)
        {
            if (vertices == null || vertices.Count == 0)
                return new List<double> { 0, 0, 0, 100, 100, 50 };

            double minX = double.MaxValue, minY = double.MaxValue, minZ = double.MaxValue;
            double maxX = double.MinValue, maxY = double.MinValue, maxZ = double.MinValue;

            foreach (var vertex in vertices)
            {
                if (vertex.Count >= 3)
                {
                    minX = Math.Min(minX, vertex[0]);
                    minY = Math.Min(minY, vertex[1]);
                    minZ = Math.Min(minZ, vertex[2]);
                    maxX = Math.Max(maxX, vertex[0]);
                    maxY = Math.Max(maxY, vertex[1]);
                    maxZ = Math.Max(maxZ, vertex[2]);
                }
            }

            if (minX == double.MaxValue) return new List<double> { 0, 0, 0, 100, 100, 50 };

            double buffer = Math.Max((maxX - minX) * 0.1, 10);
            return new List<double>
            {
                minX - buffer, minY - buffer, minZ - buffer,
                maxX + buffer, maxY + buffer, maxZ + buffer
            };
        }

        private CityJSONObj CreateEnhancedLOD3Building(Document doc, CityJSONObj lod2Building, string buildingId)
        {
            CityJSONObj lod3Building = new CityJSONObj
            {
                Type = lod2Building.Type,
                Attributes = new Dictionary<string, object>(lod2Building.Attributes ?? new Dictionary<string, object>()),
                Geometry = new List<CityJSONGeom>(),
                Children = lod2Building.Children != null ? new List<string>(lod2Building.Children) : null,
                Parents = null
            };

            lod3Building.Attributes["cityjsonObjectId"] = buildingId;

            ConvertToLOD3Geometry(lod2Building, lod3Building, buildingId);
            AddRevitAttributes(doc, lod3Building, buildingId);
            EnhanceBuildingAttributes(lod3Building);

            if (lod3Building.Geometry.Count == 0)
            {
                CreateDefaultBuildingGeometry(lod3Building);
            }

            return lod3Building;
        }

        private void ConvertToLOD3Geometry(CityJSONObj sourceBuilding, CityJSONObj targetBuilding, string buildingId)
        {
            try
            {
                if (sourceBuilding.Geometry != null)
                {
                    foreach (var sourceGeom in sourceBuilding.Geometry)
                    {
                        if (sourceGeom.Lod == "2" || sourceGeom.Lod == "2.0" ||
                            sourceGeom.Lod == "2.1" || sourceGeom.Lod == "2.2")
                        {
                            List<object> cleanedBoundaries = sourceGeom.Boundaries != null
                                ? CleanBoundaries(sourceGeom.Boundaries)
                                : new List<object>();

                            CityJSONGeom lod3Geom = new CityJSONGeom
                            {
                                Type = "MultiSurface",
                                Lod = "3",
                                Boundaries = cleanedBoundaries,
                                Semantics = sourceGeom.Semantics != null ?
                                    new CityJSONSemantic
                                    {
                                        Values = new List<int>(sourceGeom.Semantics.Values),
                                        Surfaces = CreateCompatibleSurfaces(sourceGeom.Semantics.Surfaces)
                                    } :
                                    null
                            };

                            targetBuilding.Geometry.Add(lod3Geom);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Geometry conversion error for {buildingId}: {ex.Message}");
            }
        }

        private List<object> CleanBoundaries(List<object> original)
        {
            var cleaned = new List<object>();
            if (original == null)
                return cleaned;

            foreach (var faceObj in original)
            {
                var ringCollection = faceObj as System.Collections.IEnumerable;
                if (ringCollection != null && !(faceObj is string))
                {
                    var validRings = new List<List<int>>();

                    foreach (var ringObj in ringCollection)
                    {
                        var ringEnum = ringObj as System.Collections.IEnumerable;
                        if (ringEnum == null || ringObj is string)
                            continue;

                        var ring = new List<int>();
                        foreach (var v in ringEnum)
                        {
                            try
                            {
                                ring.Add(Convert.ToInt32(v));
                            }
                            catch
                            {
                            }
                        }

                        if (ring.Count < 3)
                            continue;

                        for (int i = ring.Count - 1; i > 0; i--)
                        {
                            if (ring[i] == ring[i - 1])
                                ring.RemoveAt(i);
                        }
                        if (ring.Count < 3)
                            continue;

                        if (ring[0] == ring[ring.Count - 1])
                        {
                            ring.RemoveAt(ring.Count - 1);
                            if (ring.Count < 3)
                                continue;
                        }

                        var distinctCount = new HashSet<int>(ring).Count;
                        if (distinctCount < 3)
                            continue;

                        validRings.Add(ring);
                    }

                    if (validRings.Count > 0)
                        cleaned.Add(validRings);
                }
                else
                {
                    cleaned.Add(faceObj);
                }
            }

            return cleaned;
        }

        private List<CityJSONSurf> CreateCompatibleSurfaces(List<CityJSONSurf> originalSurfaces)
        {
            var compatibleSurfaces = new List<CityJSONSurf>();

            if (originalSurfaces != null)
            {
                foreach (var surface in originalSurfaces)
                {
                    compatibleSurfaces.Add(new CityJSONSurf
                    {
                        Type = GetCompatibleSurfaceType(surface.Type)
                    });
                }
            }

            return compatibleSurfaces;
        }

        private string GetCompatibleSurfaceType(string originalType)
        {
            var compatibleTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"WallSurface", "WallSurface"},
                {"RoofSurface", "RoofSurface"},
                {"GroundSurface", "GroundSurface"},
                {"FloorSurface", "FloorSurface"},
                {"ClosureSurface", "ClosureSurface"},
                {"OuterCeilingSurface", "OuterCeilingSurface"},
                {"OuterFloorSurface", "OuterFloorSurface"},
                {"CeilingSurface", "CeilingSurface"},
                {"InteriorWall", "WallSurface"},
                {"ExteriorWall", "WallSurface"}
            };

            if (compatibleTypes.ContainsKey(originalType))
                return compatibleTypes[originalType];

            return "WallSurface";
        }

        private void CreateDefaultBuildingGeometry(CityJSONObj building)
        {
            var boundaries = new List<object>
            {
                new List<List<List<int>>>
                {
                    new List<List<int>> { new List<int> {0, 1, 2, 3} },
                    new List<List<int>> { new List<int> {4, 5, 6, 7} },
                    new List<List<int>> { new List<int> {0, 1, 5, 4} },
                    new List<List<int>> { new List<int> {1, 2, 6, 5} },
                    new List<List<int>> { new List<int> {2, 3, 7, 6} },
                    new List<List<int>> { new List<int> {3, 0, 4, 7} }
                }
            };

            boundaries = CleanBoundaries(boundaries);

            building.Geometry.Add(new CityJSONGeom
            {
                Type = "MultiSurface",
                Lod = "3",
                Boundaries = boundaries,
                Semantics = new CityJSONSemantic
                {
                    Values = new List<int> { 0, 1, 2, 3, 4, 5 },
                    Surfaces = new List<CityJSONSurf>
                    {
                        new CityJSONSurf { Type = "GroundSurface" },
                        new CityJSONSurf { Type = "RoofSurface" },
                        new CityJSONSurf { Type = "WallSurface" },
                        new CityJSONSurf { Type = "WallSurface" },
                        new CityJSONSurf { Type = "WallSurface" },
                        new CityJSONSurf { Type = "WallSurface" }
                    }
                }
            });
        }

        private void AddRevitAttributes(Document doc, CityJSONObj obj, string objId)
        {
            try
            {
                if (obj.Attributes == null)
                    obj.Attributes = new Dictionary<string, object>();

                obj.Attributes["revitDocument"] = doc.Title;
                obj.Attributes["revitExportDate"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                obj.Attributes["revitPluginVersion"] = "1.0";
                obj.Attributes["geometrySource"] = "Revit LOD3 Export";

                if (obj.Attributes.TryGetValue("element_id", out var eidObj))
                {
                    int intId = -1;

                    if (eidObj is int i)
                        intId = i;
                    else if (eidObj is long l)
                        intId = (int)l;
                    else if (eidObj is string s && int.TryParse(s, out int parsed))
                        intId = parsed;

                    if (intId > 0)
                    {
                        ElementId rid = new ElementId((long)intId);
                        Element elem = doc.GetElement(rid);
                        if (elem != null)
                        {

                            if (elem.Category != null)
                                obj.Attributes["revitCategory"] = elem.Category.Name;

                            if (elem is FamilyInstance fi)
                            {
                                if (fi.Symbol != null)
                                {
                                    obj.Attributes["revitFamily"] = fi.Symbol.Family?.Name;
                                    obj.Attributes["revitTypeName"] = fi.Symbol.Name;
                                }
                            }
                            else
                            {
                                obj.Attributes["revitTypeName"] = elem.Name;
                            }

                            BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                            if (bb != null)
                            {
                                double hFeet = bb.Max.Z - bb.Min.Z;
                                double hMeters = hFeet * 0.3048;
                                obj.Attributes["revitBoundingHeight"] = Math.Round(hMeters, 3);
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Revit attribute error: {ex.Message}");
            }
        }

        private void EnhanceBuildingAttributes(CityJSONObj building)
        {
            if (building.Attributes == null)
                building.Attributes = new Dictionary<string, object>();

            if (building.Attributes.TryGetValue("lod", out var oldLod))
            {
                building.Attributes["sourceLod"] = oldLod;
            }

            building.Attributes["lod"] = "3";

            building.Attributes["geometryType"] = "MultiSurface";

            if (!building.Attributes.ContainsKey("name"))
            {
                building.Attributes["name"] = "LOD3_Building";
            }

            if (!building.Attributes.ContainsKey("measuredHeight"))
            {
                if (building.Attributes.TryGetValue("revitBoundingHeight", out var hObj))
                {
                    try
                    {
                        double h = Convert.ToDouble(hObj);
                        building.Attributes["measuredHeight"] = h;
                    }
                    catch
                    {
                        building.Attributes["measuredHeight"] = 10.0;
                    }
                }
                else
                {
                    building.Attributes["measuredHeight"] = 10.0;
                }
            }

            if (!building.Attributes.ContainsKey("function"))
            {
                building.Attributes["function"] = "unknown";
            }
        }

        private void EnhanceGenericAttributes(CityJSONObj obj)
        {
            if (obj.Attributes == null)
                obj.Attributes = new Dictionary<string, object>();

            if (obj.Attributes.TryGetValue("lod", out var oldLod))
            {
                obj.Attributes["sourceLod"] = oldLod;
            }

            obj.Attributes["lod"] = "3";

            obj.Attributes["geometryType"] = "MultiSurface";
        }

        private class AutoCloseInfoForm : System.Windows.Forms.Form
        {
            private readonly System.Windows.Forms.Timer _timer;

            public AutoCloseInfoForm(string message, int milliseconds)
            {
                this.Text = "LOD3 Export";
                this.StartPosition = FormStartPosition.CenterScreen;
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.Size = new Size(420, 260);

                var label = new Label
                {
                    Dock = DockStyle.Fill,
                    Text = message,
                    AutoSize = false,
                    Font = new Font("Segoe UI", 10),
                    TextAlign = ContentAlignment.MiddleLeft,
                    Padding = new Padding(20, 20, 20, 20)
                };
                this.Controls.Add(label);

                _timer = new System.Windows.Forms.Timer();
                _timer.Interval = milliseconds;
                _timer.Tick += (s, e) =>
                {
                    _timer.Stop();
                    this.Close();
                };
            }

            protected override void OnShown(EventArgs e)
            {
                base.OnShown(e);
                _timer.Start();
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing && _timer != null)
                {
                    _timer.Dispose();
                }
                base.Dispose(disposing);
            }
        }
 
        private void AddAnalysisAttributes(CityJSONData data)
        {
            if (data == null || data.CityObjects == null || data.CityObjects.Count == 0)
                return;

            double minZ = double.MaxValue;
            double maxZ = double.MinValue;

            if (data.Vertices != null)
            {
                foreach (var v in data.Vertices)
                {
                    if (v != null && v.Count >= 3)
                    {
                        double z = v[2];
                        if (z < minZ) minZ = z;
                        if (z > maxZ) maxZ = z;
                    }
                }
            }

            double heightMeters = 0.0;
            if (minZ < double.MaxValue && maxZ > double.MinValue)
            {
                heightMeters = maxZ - minZ;
            }

            foreach (var kv in data.CityObjects)
            {
                string objId = kv.Key;
                CityJSONObj obj = kv.Value;

                if (obj.Attributes == null)
                    obj.Attributes = new Dictionary<string, object>();

                if (!obj.Attributes.ContainsKey("revitBoundingHeight"))
                {
                    obj.Attributes["revitBoundingHeight"] = Math.Round(heightMeters, 3);
                }
            }
        }
    }


    public class CityJSONData
    {
        [JsonProperty("type")]
        public string Type { get; set; } = "CityJSON";

        [JsonProperty("version")]
        public string Version { get; set; } = "2.0";

        [JsonProperty("CityObjects")]
        public Dictionary<string, CityJSONObj> CityObjects { get; set; }

        [JsonProperty("vertices")]
        public List<List<double>> Vertices { get; set; }

        [JsonProperty("transform")]
        public CityJSONTrans Transform { get; set; }

        [JsonProperty("metadata", NullValueHandling = NullValueHandling.Ignore)]
        public CityJSONMeta Metadata { get; set; }
    }

    public class CityJSONMeta
    {
        [JsonProperty("referenceSystem")]
        public string ReferenceSystem { get; set; }

        [JsonProperty("geographicalExtent", NullValueHandling = NullValueHandling.Ignore)]
        public List<double> GeographicalExtent { get; set; }

        [JsonProperty("title", NullValueHandling = NullValueHandling.Ignore)]
        public string Title { get; set; }

        [JsonProperty("datasetTitle", NullValueHandling = NullValueHandling.Ignore)]
        public string DatasetTitle { get; set; }

        [JsonProperty("datasetReferenceDate", NullValueHandling = NullValueHandling.Ignore)]
        public string DatasetReferenceDate { get; set; }

        [JsonProperty("source", NullValueHandling = NullValueHandling.Ignore)]
        public string Source { get; set; }

        [JsonProperty("revitInfo", NullValueHandling = NullValueHandling.Ignore)]
        public Dictionary<string, object> RevitInfo { get; set; }
    }

    public class CityJSONObj
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("attributes")]
        public Dictionary<string, object> Attributes { get; set; }

        [JsonProperty("geometry")]
        public List<CityJSONGeom> Geometry { get; set; }

        [JsonProperty("children", NullValueHandling = NullValueHandling.Ignore)]
        public List<string> Children { get; set; }

        [JsonProperty("parents", NullValueHandling = NullValueHandling.Ignore)]
        public List<string> Parents { get; set; }
    }

    public class CityJSONGeom
    {
        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("lod")]
        public string Lod { get; set; }

        [JsonProperty("boundaries")]
        public List<object> Boundaries { get; set; }

        [JsonProperty("semantics", NullValueHandling = NullValueHandling.Ignore)]
        public CityJSONSemantic Semantics { get; set; }
    }

    public class CityJSONSemantic
    {
        [JsonProperty("values")]
        public List<int> Values { get; set; }

        [JsonProperty("surfaces")]
        public List<CityJSONSurf> Surfaces { get; set; }
    }

    public class CityJSONSurf
    {
        [JsonProperty("type")]
        public string Type { get; set; }
    }

    public class CityJSONTrans
    {
        [JsonProperty("scale")]
        public List<double> Scale { get; set; } = new List<double> { 1.0, 1.0, 1.0 };

        [JsonProperty("translate")]
        public List<double> Translate { get; set; } = new List<double> { 0.0, 0.0, 0.0 };
    }
}
