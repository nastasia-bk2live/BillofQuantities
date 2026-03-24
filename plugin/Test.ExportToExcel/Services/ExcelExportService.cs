using System;
using System.Collections.Generic;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using ClosedXML.Excel;
using Test.ExportToExcel.Models;

namespace Test.ExportToExcel.Services
{
    /// <summary>
    /// Сервис записи собранных данных в XLSX.
    /// </summary>
    public class ExcelExportService
    {
        public void Export(string filePath, ExportData data, Action<int, int> progressCallback)
        {
            if (data == null)
            {
                throw new ArgumentNullException(nameof(data));
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Elements");
                var headers = BuildHeaders(data.ParameterColumns);

                for (var i = 0; i < headers.Count; i++)
                {
                    worksheet.Cell(1, i + 1).Value = headers[i];
                    worksheet.Cell(1, i + 1).Style.Font.Bold = true;
                }

                for (var rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++)
                {
                    var row = data.Rows[rowIndex];
                    var excelRow = rowIndex + 2;

                    worksheet.Cell(excelRow, 1).Value = row.Key;
                    worksheet.Cell(excelRow, 2).Value = row.ElementId;
                    worksheet.Cell(excelRow, 3).Value = row.Category;
                    worksheet.Cell(excelRow, 4).Value = row.Family;
                    worksheet.Cell(excelRow, 5).Value = row.Type;

                    for (var paramIndex = 0; paramIndex < data.ParameterColumns.Count; paramIndex++)
                    {
                        var column = data.ParameterColumns[paramIndex];
                        string value;
                        if (!row.Parameters.TryGetValue(column, out value))
                        {
                            value = "no";
                        }

                        worksheet.Cell(excelRow, 6 + paramIndex).Value = value ?? string.Empty;
                    }

                    progressCallback?.Invoke(rowIndex + 1, data.Rows.Count);
                }

                worksheet.Columns().AdjustToContents();

                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                SaveWorkbook(workbook, filePath);
            }
        }

        private static void SaveWorkbook(XLWorkbook workbook, string filePath)
        {
            // В Revit на больших объёмах данных сохранение по пути иногда падает
            // с IsolatedStorageException внутри Packaging API.
            // Сохраняем через FileStream и даём fallback на временный локальный файл.
            try
            {
                using (var outputStream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    workbook.SaveAs(outputStream);
                    outputStream.Flush();
                }
            }
            catch (IsolatedStorageException)
            {
                var tempFile = Path.Combine(Path.GetTempPath(), "Test.ExportToExcel_" + Guid.NewGuid().ToString("N") + ".xlsx");
                try
                {
                    using (var outputStream = new FileStream(tempFile, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                    {
                        workbook.SaveAs(outputStream);
                        outputStream.Flush();
                    }

                    File.Copy(tempFile, filePath, true);
                }
                finally
                {
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                    }
                }
            }
        }

        private static IList<string> BuildHeaders(IList<ParameterColumn> parameterColumns)
        {
            var headers = new List<string>
            {
                "id_имя семейства_имя типа",
                "ElementId",
                "Category",
                "Family",
                "Type"
            };

            var seen = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);

            foreach (var parameterColumn in parameterColumns)
            {
                if (!seen.ContainsKey(parameterColumn.Name))
                {
                    seen[parameterColumn.Name] = 1;
                    headers.Add(parameterColumn.Name);
                    continue;
                }

                var sourceSuffix = parameterColumn.Source.ToString();
                var newHeader = parameterColumn.Name + " [" + sourceSuffix + "]";
                seen[parameterColumn.Name]++;
                headers.Add(newHeader);
            }

            return headers;
        }
    }
}
