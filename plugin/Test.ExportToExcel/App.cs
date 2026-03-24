using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI;
using Test.ExportToExcel.Infrastructure;

namespace Test.ExportToExcel
{
    public class App : IExternalApplication
    {
        private const string TabName = "BOQ Tools";
        private const string PanelName = "Export";

        public Result OnStartup(UIControlledApplication application)
        {
            var logger = new FileLogger();

            try
            {
                TryCreateRibbonTab(application, TabName);

                var panel = GetOrCreatePanel(application, TabName, PanelName);
                var assemblyPath = Assembly.GetExecutingAssembly().Location;

                var buttonData = new PushButtonData(
                    "Test.ExportToExcel.Button",
                    "Экспорт в Excel",
                    assemblyPath,
                    "Test.ExportToExcel.Commands.ExportToExcelCommand");

                panel.AddItem(buttonData);
                logger.Info("Ribbon-кнопка 'Экспорт в Excel' успешно добавлена.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                logger.Error("Ошибка инициализации ribbon-компонентов.", ex);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private static void TryCreateRibbonTab(UIControlledApplication app, string tabName)
        {
            try
            {
                app.CreateRibbonTab(tabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // Вкладка уже существует.
            }
        }

        private static RibbonPanel GetOrCreatePanel(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (var panel in app.GetRibbonPanels(tabName))
            {
                if (panel.Name.Equals(panelName, StringComparison.OrdinalIgnoreCase))
                {
                    return panel;
                }
            }

            return app.CreateRibbonPanel(tabName, panelName);
        }
    }
}
