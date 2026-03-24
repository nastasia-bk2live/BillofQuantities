using System.Reflection;
using Autodesk.Revit.UI;

namespace Test.ExportToExcel
{
    public class App : IExternalApplication
    {
        private const string TabName = "Привет";
        private const string PanelName = "Привет";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                application.CreateRibbonTab(TabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // Tab already exists.
            }

            RibbonPanel panel = application.CreateRibbonPanel(TabName, PanelName);
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            var exportButton = new PushButtonData(
                "ExportToExcelButton",
                "Экспорт в Excel",
                assemblyPath,
                typeof(ExportToExcelCommand).FullName)
            {
                ToolTip = "Выгружает все экземпляры модели и параметры в Excel (.xlsx)."
            };

            panel.AddItem(exportButton);
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
