using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Test.ExportToExcel.Models;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Test.ExportToExcel.Services
{
    /// <summary>
    /// Сервис записи собранных данных в XLSX.
    /// </summary>
    public class XlsxExportService
    {
        private const int ProgressReportStep = 200;

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

            var headers = BuildHeaders(data.ParameterColumns);

            using (var workbook = new XSSFWorkbook())
            {
                var sheet = workbook.CreateSheet("Elements");

                var headerStyle = workbook.CreateCellStyle();
                var headerFont = workbook.CreateFont();
                headerFont.IsBold = true;
                headerStyle.SetFont(headerFont);

                var headerRow = sheet.CreateRow(0);
                for (var i = 0; i < headers.Count; i++)
                {
                    var cell = headerRow.CreateCell(i, CellType.String);
                    cell.SetCellValue(headers[i]);
                    cell.CellStyle = headerStyle;
                }

                var total = data.Rows.Count;
                for (var rowIndex = 0; rowIndex < total; rowIndex++)
                {
                    var rowData = data.Rows[rowIndex];
                    var row = sheet.CreateRow(rowIndex + 1);

                    row.CreateCell(0, CellType.String).SetCellValue(rowData.Key ?? string.Empty);
                    row.CreateCell(1, CellType.String).SetCellValue(rowData.ElementId ?? string.Empty);
                    row.CreateCell(2, CellType.String).SetCellValue(rowData.Category ?? string.Empty);
                    row.CreateCell(3, CellType.String).SetCellValue(rowData.Family ?? string.Empty);
                    row.CreateCell(4, CellType.String).SetCellValue(rowData.Type ?? string.Empty);

                    for (var paramIndex = 0; paramIndex < data.ParameterColumns.Count; paramIndex++)
                    {
                        var column = data.ParameterColumns[paramIndex];
                        string value;
                        if (!rowData.Parameters.TryGetValue(column, out value))
                        {
                            value = "no";
                        }

                        row.CreateCell(5 + paramIndex, CellType.String).SetCellValue(value ?? string.Empty);
                    }

                    var current = rowIndex + 1;
                    if (current % ProgressReportStep == 0 || current == total)
                    {
                        progressCallback?.Invoke(current, total);
                    }
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

                headers.Add(parameterColumn.Name + " [" + parameterColumn.Source + "]");
                seen[parameterColumn.Name]++;
            }

            return headers;
        }
    }
}
