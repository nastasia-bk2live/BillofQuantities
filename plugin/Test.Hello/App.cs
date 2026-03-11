using System;
using System.Reflection;
using Autodesk.Revit.UI;

namespace Test.Hello
{
    public class App : IExternalApplication
    {
        private const string TabName = "Привет";
        private const string PanelName = "Привет";
        private const string ButtonName = "Привет";

        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                application.CreateRibbonTab(TabName);
            }
            catch (Autodesk.Revit.Exceptions.ArgumentException)
            {
                // Таб уже существует.
            }

            RibbonPanel panel = application.CreateRibbonPanel(TabName, PanelName);

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            var buttonData = new PushButtonData(
                "TestHelloButton",
                ButtonName,
                assemblyPath,
                typeof(HelloCommand).FullName);

            panel.AddItem(buttonData);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
