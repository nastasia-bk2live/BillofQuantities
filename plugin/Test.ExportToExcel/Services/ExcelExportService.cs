using System;
using System.Collections.Generic;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using ClosedXML.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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

            var directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            try
            {
                ExportWithClosedXml(filePath, data, progressCallback);
            }
            catch (IsolatedStorageException)
            {
                // В некоторых окружениях Revit/.NET Framework ClosedXML может падать
                // из-за internal Packaging API (IsolatedStorage identity issue).
                // В этом случае используем надёжный fallback через NPOI.
                ExportWithNpoi(filePath, data, progressCallback);
            }
        }

        private static void ExportWithClosedXml(string filePath, ExportData data, Action<int, int> progressCallback)
        {
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
                SaveWorkbook(workbook, filePath);
            }
        }

        private static void SaveWorkbook(XLWorkbook workbook, string filePath)
        {
            using (var outputStream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
            {
                workbook.SaveAs(outputStream);
                outputStream.Flush();
            }
        }

        private static void ExportWithNpoi(string filePath, ExportData data, Action<int, int> progressCallback)
        {
            using (var workbook = new XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("Elements");
                var headers = BuildHeaders(data.ParameterColumns);

                var headerRow = sheet.CreateRow(0);
                var headerStyle = workbook.CreateCellStyle();
                var headerFont = workbook.CreateFont();
                headerFont.IsBold = true;
                headerStyle.SetFont(headerFont);

                for (var i = 0; i < headers.Count; i++)
                {
                    var cell = headerRow.CreateCell(i, CellType.String);
                    cell.SetCellValue(headers[i]);
                    cell.CellStyle = headerStyle;
                }

                for (var rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++)
                {
                    var source = data.Rows[rowIndex];
                    var row = sheet.CreateRow(rowIndex + 1);

                    row.CreateCell(0, CellType.String).SetCellValue(source.Key ?? string.Empty);
                    row.CreateCell(1, CellType.String).SetCellValue(source.ElementId ?? string.Empty);
                    row.CreateCell(2, CellType.String).SetCellValue(source.Category ?? string.Empty);
                    row.CreateCell(3, CellType.String).SetCellValue(source.Family ?? string.Empty);
                    row.CreateCell(4, CellType.String).SetCellValue(source.Type ?? string.Empty);

                    for (var paramIndex = 0; paramIndex < data.ParameterColumns.Count; paramIndex++)
                    {
                        var column = data.ParameterColumns[paramIndex];
                        string value;
                        if (!source.Parameters.TryGetValue(column, out value))
                        {
                            value = "no";
                        }

                        row.CreateCell(5 + paramIndex, CellType.String).SetCellValue(value ?? string.Empty);
                    }

                    progressCallback?.Invoke(rowIndex + 1, data.Rows.Count);
                }

                for (var i = 0; i < headers.Count; i++)
                {
                    sheet.AutoSizeColumn(i);
                }

                using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    workbook.Write(stream);
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
