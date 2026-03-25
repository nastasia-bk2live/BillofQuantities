using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Test.ExportToExcel.Models;

namespace Test.ExportToExcel.Services
{
    /// <summary>
    /// Сервис записи собранных данных в CSV (UTF-8 BOM).
    /// </summary>
    public class CsvExportService
    {
        private const int ProgressReportStep = 200;
        private const char Separator = ';';

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

            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (var writer = new StreamWriter(stream, new UTF8Encoding(true)))
            {
                writer.WriteLine(string.Join(Separator.ToString(), headers.Select(EscapeCsv)));

                var total = data.Rows.Count;
                for (var rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++)
                {
                    var row = data.Rows[rowIndex];
                    var cells = new List<string>(FixedColumnsCount + data.ParameterColumns.Count)
                    {
                        row.Key ?? string.Empty,
                        row.ElementId ?? string.Empty,
                        row.Category ?? string.Empty,
                        row.Family ?? string.Empty,
                        row.Type ?? string.Empty
                    };

                    for (var paramIndex = 0; paramIndex < data.ParameterColumns.Count; paramIndex++)
                    {
                        var column = data.ParameterColumns[paramIndex];
                        string value;
                        if (!row.Parameters.TryGetValue(column, out value))
                        {
                            value = "no";
                        }

                        cells.Add(value ?? string.Empty);
                    }

                    writer.WriteLine(string.Join(Separator.ToString(), cells.Select(EscapeCsv)));

                    var current = rowIndex + 1;
                    if (current % ProgressReportStep == 0 || current == total)
                    {
                        progressCallback?.Invoke(current, total);
                    }
                }
            }
        }

        private const int FixedColumnsCount = 5;

        private static string EscapeCsv(string input)
        {
            var value = input ?? string.Empty;
            var mustQuote = value.IndexOfAny(new[] { Separator, '"', '\r', '\n' }) >= 0;
            if (!mustQuote)
            {
                return value;
            }

            return '"' + value.Replace("\"", "\"\"") + '"';
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
