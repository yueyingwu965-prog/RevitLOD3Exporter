using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitLOD3Exporter
{
    [Transaction(TransactionMode.Manual)]
    public class RetrofitCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            try
            {
                // 1) Walls
                RetrofitWallApplier.Apply(doc);

                // 2) Floors
                RetrofitFloorApplier.Apply(doc);

                // 3) Roofs
                RetrofitRoofApplier.Apply(doc);

                // 4) Windows (ONLY: call API -> write text plan; NO geometry/type changes)
                RetrofitWindowPlanWriter.ExportWindowsReportTxt(doc);

                TaskDialog.Show("Retrofit", "Retrofit finished!");
                return Result.Succeeded;
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Retrofit Error", ex.ToString());
                return Result.Failed;
            }
        }
    }
}
