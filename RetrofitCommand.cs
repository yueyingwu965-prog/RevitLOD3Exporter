using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace RevitLOD3Exporter
{
    [Transaction(TransactionMode.Manual)]
    public class RetrofitCommand : IExternalCommand
    {
        private const string PARAM_RETROFIT_MODE = "retrofit_mode";

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {

                string mode = GetProjectInfoString(doc, PARAM_RETROFIT_MODE);
                if (string.IsNullOrWhiteSpace(mode)) mode = "baseline3";
                mode = mode.Trim().ToLowerInvariant();

                RetrofitWallApplier.Apply(doc, mode);

                RetrofitFloorApplier.Apply(doc, mode);

                RetrofitRoofApplier.Apply(doc);

                RetrofitWindowPlanWriter.ExportOpeningsReportTxt(doc);

                TaskDialog.Show("Retrofit", $"Retrofit finished!");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Retrofit Error", ex.ToString());
                return Result.Failed;
            }
        }

        private static string GetProjectInfoString(Document doc, string paramName)
        {
            try
            {
                var pi = doc.ProjectInformation;
                var p = pi?.LookupParameter(paramName);
                if (p == null) return null;
                return p.AsString();
            }
            catch { return null; }
        }
    }
}
