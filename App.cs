using Autodesk.Revit.UI;
using System;
using System.Reflection;

namespace RevitLOD3Exporter
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication app)
        {
            const string tabName = "LOD Tools";
            const string panelName = "CityJSON";

            try { app.CreateRibbonTab(tabName); } catch { } // tab may exist already
            RibbonPanel panel = app.CreateRibbonPanel(tabName, panelName);

            string dllPath = Assembly.GetExecutingAssembly().Location;

            // Button 1: Export LOD3
            var btnExport = new PushButtonData(
                "ExportLOD3",
                "Export\nLOD3",
                dllPath,
                "RevitLOD3Exporter.ExportLOD3"
            );

            // Button 2: Import CSV
            var btnImport = new PushButtonData(
                "ImportRetrofitCSV",
                "Import\nCSV",
                dllPath,
                "RevitLOD3Exporter.ImportRetrofitCSV"
            );

            // Button 3: Retrofit (apply upgrades based on parameters)
            var btnRetrofit = new PushButtonData(
                "Retrofit",
                "Retrofit\nApply",
                dllPath,
                "RevitLOD3Exporter.RetrofitCommand"
            );

            panel.AddItem(btnExport);
            panel.AddItem(btnImport);

            // Add a separator for clarity (optional, safe)
            panel.AddSeparator();

            panel.AddItem(btnRetrofit);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication app)
        {
            return Result.Succeeded;
        }
    }
}
