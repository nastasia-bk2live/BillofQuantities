using System;
using System.IO;
using System.IO.IsolatedStorage;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using Test.ExportToExcel.Infrastructure;
using Test.ExportToExcel.Services;
using Test.ExportToExcel.UI;

namespace Test.ExportToExcel.Commands
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ExportToExcelCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var logger = new FileLogger();

            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                    DefaultExt = "xlsx",
                    AddExtension = true,
                    FileName = BuildDefaultFileName(commandData.Application.ActiveUIDocument.Document)
                };

                var saveResult = saveFileDialog.ShowDialog();
                if (saveResult != true)
                {
                    logger.Info("Пользователь отменил выбор файла для экспорта.");
                    return Result.Cancelled;
                }

                var filePath = saveFileDialog.FileName;
                logger.Info("Старт экспорта в файл: " + filePath);

                var dataService = new ElementExportDataService();
                var excelService = new ExcelExportService();

                var progressWindow = new ExportProgressWindow();
                progressWindow.Show();

                try
                {
                    var data = dataService.Collect(commandData.Application.ActiveUIDocument.Document, (current, total) =>
                    {
                        progressWindow.ReportProgress(current, total, "Сбор данных");
                        if (progressWindow.IsCancellationRequested)
                        {
                            throw new OperationCanceledException("Экспорт отменен пользователем на этапе сбора данных.");
                        }
                    });

                    excelService.Export(filePath, data, (current, total) =>
                    {
                        progressWindow.ReportProgress(current, total, "Запись Excel");
                        if (progressWindow.IsCancellationRequested)
                        {
                            throw new OperationCanceledException("Экспорт отменен пользователем на этапе записи в Excel.");
                        }
                    });

                    logger.Info("Экспорт успешно завершен. Записано элементов: " + data.Rows.Count);
                }
                finally
                {
                    progressWindow.Close();
                }
                TaskDialog.Show("Экспорт в Excel", "Экспорт завершен успешно.\nФайл: " + filePath);
                return Result.Succeeded;
            }
            catch (OperationCanceledException ex)
            {
                logger.Info(ex.Message);
                TaskDialog.Show("Экспорт в Excel", "Операция отменена пользователем.");
                return Result.Cancelled;
            }
            catch (IsolatedStorageException ex)
            {
                logger.Error("Ошибка IsolatedStorage при экспорте в Excel.", ex);
                message = ex.Message;
                TaskDialog.Show(
                    "Экспорт в Excel",
                    "Ошибка сохранения XLSX (IsolatedStorage).\n" +
                    "Попробуйте сохранить сначала на локальный диск, например C:\\Temp\\export.xlsx.");
                return Result.Failed;
            }
            catch (Exception ex)
            {
                logger.Error("Ошибка при экспорте в Excel.", ex);
                message = ex.Message;
                TaskDialog.Show("Экспорт в Excel", "Произошла ошибка. Подробности в логе.");
                return Result.Failed;
            }
        }

        private static string BuildDefaultFileName(Document document)
        {
            var modelName = document != null ? document.Title : "RevitModel";
            foreach (var invalid in Path.GetInvalidFileNameChars())
            {
                modelName = modelName.Replace(invalid.ToString(), string.Empty);
            }

            return modelName + "_export.xlsx";
        }
    }
}
