using System;
using System.Collections.Generic;
using System.IO;
using System.IO.IsolatedStorage;
using System.Linq;
using System.Text;
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
        private const int MaxExcelRows = 1048576;
        private const int MaxExcelColumns = 16384;
        private const int FixedColumnsCount = 5;
        private const int MaxCellLength = 32767;

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
                WriteData(
                    data,
                    progressCallback,
                    createSheet: name => workbook.Worksheets.Add(name),
                    setCell: (sheet, row, col, value, isHeader) =>
                    {
                        var cell = sheet.Cell(row, col);
                        cell.Value = value;
                        if (isHeader)
                        {
                            cell.Style.Font.Bold = true;
                        }
                    },
                    autoFit: sheet => sheet.Columns().AdjustToContents());

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
                WriteData(
                    data,
                    progressCallback,
                    createSheet: name => workbook.CreateSheet(name),
                    setCell: (sheet, row, col, value, isHeader) =>
                    {
                        var npoiRow = sheet.GetRow(row - 1) ?? sheet.CreateRow(row - 1);
                        var cell = npoiRow.CreateCell(col - 1, CellType.String);
                        cell.SetCellValue(value ?? string.Empty);

                        if (isHeader)
                        {
                            var style = workbook.CreateCellStyle();
                            var font = workbook.CreateFont();
                            font.IsBold = true;
                            style.SetFont(font);
                            cell.CellStyle = style;
                        }
                    },
                    autoFit: sheet =>
                    {
                        var firstRow = sheet.GetRow(0);
                        if (firstRow == null)
                        {
                            return;
                        }

                        var cellCount = firstRow.LastCellNum;
                        for (var i = 0; i < cellCount; i++)
                        {
                            sheet.AutoSizeColumn(i);
                        }
                    });

                using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    workbook.Write(stream);
                }
            }
        }

        private static void WriteData<TSheet>(
            ExportData data,
            Action<int, int> progressCallback,
            Func<string, TSheet> createSheet,
            Action<TSheet, int, int, string, bool> setCell,
            Action<TSheet> autoFit)
        {
            var allHeaders = BuildHeaders(data.ParameterColumns);
            var maxParameterColumnsPerSheet = Math.Max(1, MaxExcelColumns - FixedColumnsCount);
            var parameterChunks = SplitIntoChunks(data.ParameterColumns.Count, maxParameterColumnsPerSheet).ToList();
            var rowChunks = SplitIntoChunks(data.Rows.Count, MaxExcelRows - 1).ToList();

            var progressReported = 0;
            var totalProgress = Math.Max(data.Rows.Count, 1);

            for (var rowChunkIndex = 0; rowChunkIndex < rowChunks.Count; rowChunkIndex++)
            {
                var rowChunk = rowChunks[rowChunkIndex];

                for (var paramChunkIndex = 0; paramChunkIndex < parameterChunks.Count; paramChunkIndex++)
                {
                    var paramChunk = parameterChunks[paramChunkIndex];
                    var sheetName = BuildSheetName(rowChunkIndex, paramChunkIndex);
                    var sheet = createSheet(sheetName);

                    WriteHeadersForChunk(allHeaders, paramChunk, setCell, sheet);

                    for (var localRowIndex = 0; localRowIndex < rowChunk.Count; localRowIndex++)
                    {
                        var sourceRow = data.Rows[rowChunk.Start + localRowIndex];
                        var excelRow = localRowIndex + 2;

                        setCell(sheet, excelRow, 1, SanitizeForExcel(sourceRow.Key), false);
                        setCell(sheet, excelRow, 2, SanitizeForExcel(sourceRow.ElementId), false);
                        setCell(sheet, excelRow, 3, SanitizeForExcel(sourceRow.Category), false);
                        setCell(sheet, excelRow, 4, SanitizeForExcel(sourceRow.Family), false);
                        setCell(sheet, excelRow, 5, SanitizeForExcel(sourceRow.Type), false);

                        for (var localParamIndex = 0; localParamIndex < paramChunk.Count; localParamIndex++)
                        {
                            var globalParamIndex = paramChunk.Start + localParamIndex;
                            var column = data.ParameterColumns[globalParamIndex];
                            string value;
                            if (!sourceRow.Parameters.TryGetValue(column, out value))
                            {
                                value = "no";
                            }

                            setCell(sheet, excelRow, FixedColumnsCount + localParamIndex + 1, SanitizeForExcel(value), false);
                        }

                        if (paramChunkIndex == 0)
                        {
                            progressReported++;
                            progressCallback?.Invoke(progressReported, totalProgress);
                        }
                    }

                    autoFit(sheet);
                }
            }
        }

        private static void WriteHeadersForChunk<TSheet>(
            IList<string> allHeaders,
            Chunk paramChunk,
            Action<TSheet, int, int, string, bool> setCell,
            TSheet sheet)
        {
            for (var fixedCol = 0; fixedCol < FixedColumnsCount; fixedCol++)
            {
                setCell(sheet, 1, fixedCol + 1, SanitizeForExcel(allHeaders[fixedCol]), true);
            }

            for (var localParamIndex = 0; localParamIndex < paramChunk.Count; localParamIndex++)
            {
                var globalHeaderIndex = FixedColumnsCount + paramChunk.Start + localParamIndex;
                var headerName = allHeaders[globalHeaderIndex];
                setCell(sheet, 1, FixedColumnsCount + localParamIndex + 1, SanitizeForExcel(headerName), true);
            }
        }

        private static IEnumerable<Chunk> SplitIntoChunks(int totalItems, int chunkSize)
        {
            if (totalItems <= 0)
            {
                yield return new Chunk(0, 0);
                yield break;
            }

            for (var start = 0; start < totalItems; start += chunkSize)
            {
                var count = Math.Min(chunkSize, totalItems - start);
                yield return new Chunk(start, count);
            }
        }

        private static string BuildSheetName(int rowChunkIndex, int paramChunkIndex)
        {
            // Ограничение Excel: имя листа <= 31 символ.
            var raw = string.Format("Elements_R{0}_P{1}", rowChunkIndex + 1, paramChunkIndex + 1);
            return raw.Length <= 31 ? raw : raw.Substring(0, 31);
        }

        private static string SanitizeForExcel(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }

            var builder = new StringBuilder(value.Length);
            for (var i = 0; i < value.Length; i++)
            {
                var ch = value[i];
                if (IsAllowedXmlChar(ch))
                {
                    builder.Append(ch);
                }
            }

            var sanitized = builder.ToString();
            if (sanitized.Length > MaxCellLength)
            {
                sanitized = sanitized.Substring(0, MaxCellLength);
            }

            return sanitized;
        }

        private static bool IsAllowedXmlChar(char ch)
        {
            return ch == 0x9 || ch == 0xA || ch == 0xD || (ch >= 0x20 && ch <= 0xD7FF) || (ch >= 0xE000 && ch <= 0xFFFD);
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

        private struct Chunk
        {
            public Chunk(int start, int count)
            {
                Start = start;
                Count = count;
            }

            public int Start { get; }

            public int Count { get; }
        }
    }
}
