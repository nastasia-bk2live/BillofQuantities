using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;

namespace Test.ExportToExcel
{
    [Transaction(TransactionMode.Manual)]
    public class ExportToExcelCommand : IExternalCommand
    {
        private sealed class ParameterColumn
        {
            public string Key { get; set; }
            public string Name { get; set; }
            public bool IsBuiltIn { get; set; }
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDocument = commandData.Application.ActiveUIDocument;
            Document document = uiDocument?.Document;

            if (document == null)
            {
                TaskDialog.Show("Экспорт в Excel", "Активный документ не найден.");
                return Result.Cancelled;
            }

            string filePath = SelectOutputPath(document.Title);
            if (string.IsNullOrWhiteSpace(filePath))
            {
                return Result.Cancelled;
            }

            List<Element> modelInstances = new FilteredElementCollector(document)
                .WhereElementIsNotElementType()
                .ToElements()
                .ToList();

            if (!modelInstances.Any())
            {
                TaskDialog.Show("Экспорт в Excel", "В модели нет экземпляров элементов для выгрузки.");
                return Result.Succeeded;
            }

            List<ParameterColumn> columns = BuildParameterColumns(document, modelInstances);

            using (var workbook = new XLWorkbook())
            {
                IXLWorksheet sheet = workbook.Worksheets.Add("Elements");
                WriteHeader(sheet, columns);
                WriteRows(sheet, document, modelInstances, columns);
                sheet.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }

            TaskDialog.Show("Экспорт в Excel", $"Экспорт завершён. Файл сохранён:\n{filePath}");
            return Result.Succeeded;
        }

        private static string SelectOutputPath(string documentTitle)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "Сохранить выгрузку Revit в Excel";
                dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                dialog.AddExtension = true;
                dialog.DefaultExt = "xlsx";
                dialog.FileName = $"{documentTitle}_elements.xlsx";

                DialogResult result = dialog.ShowDialog();
                return result == DialogResult.OK ? dialog.FileName : null;
            }
        }

        private static List<ParameterColumn> BuildParameterColumns(Document document, IEnumerable<Element> elements)
        {
            var byKey = new Dictionary<string, ParameterColumn>(StringComparer.Ordinal);

            foreach (Element element in elements)
            {
                foreach (Parameter parameter in EnumerateInstanceAndTypeParameters(document, element))
                {
                    if (!TryCreateColumn(parameter, out ParameterColumn column))
                    {
                        continue;
                    }

                    if (!byKey.ContainsKey(column.Key))
                    {
                        byKey[column.Key] = column;
                    }
                }
            }

            List<ParameterColumn> builtIn = byKey.Values
                .Where(c => c.IsBuiltIn)
                .OrderBy(c => c.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            List<ParameterColumn> user = byKey.Values
                .Where(c => !c.IsBuiltIn)
                .OrderBy(c => c.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var ordered = builtIn.Concat(user).ToList();
            ApplyDuplicateNameSuffixes(ordered);
            return ordered;
        }

        private static void ApplyDuplicateNameSuffixes(List<ParameterColumn> ordered)
        {
            var groups = ordered.GroupBy(c => c.Name, StringComparer.CurrentCultureIgnoreCase);
            foreach (IGrouping<string, ParameterColumn> group in groups)
            {
                int count = 0;
                foreach (ParameterColumn column in group)
                {
                    count++;
                    if (count > 1)
                    {
                        column.Name = $"{column.Name} ({count})";
                    }
                }
            }
        }

        private static bool TryCreateColumn(Parameter parameter, out ParameterColumn column)
        {
            column = null;
            if (parameter == null || parameter.Definition == null)
            {
                return false;
            }

            string name = parameter.Definition.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            InternalDefinition internalDefinition = parameter.Definition as InternalDefinition;
            if (internalDefinition != null)
            {
                BuiltInParameter builtInParameter = internalDefinition.BuiltInParameter;
                if (builtInParameter != BuiltInParameter.INVALID)
                {
                    column = new ParameterColumn
                    {
                        Key = $"BIP:{(int)builtInParameter}",
                        Name = name,
                        IsBuiltIn = true
                    };
                    return true;
                }

                column = new ParameterColumn
                {
                    Key = $"USR_INT:{name}",
                    Name = name,
                    IsBuiltIn = false
                };
                return true;
            }

            ExternalDefinition externalDefinition = parameter.Definition as ExternalDefinition;
            if (externalDefinition != null)
            {
                column = new ParameterColumn
                {
                    Key = $"USR_EXT:{externalDefinition.GUID}",
                    Name = name,
                    IsBuiltIn = false
                };
                return true;
            }

            column = new ParameterColumn
            {
                Key = $"USR_NAME:{name}",
                Name = name,
                IsBuiltIn = false
            };
            return true;
        }

        private static IEnumerable<Parameter> EnumerateInstanceAndTypeParameters(Document document, Element element)
        {
            if (element == null)
            {
                yield break;
            }

            foreach (Parameter parameter in element.Parameters)
            {
                yield return parameter;
            }

            ElementId typeId = element.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
            {
                yield break;
            }

            Element typeElement = document.GetElement(typeId);
            if (typeElement == null)
            {
                yield break;
            }

            foreach (Parameter parameter in typeElement.Parameters)
            {
                yield return parameter;
            }
        }

        private static void WriteHeader(IXLWorksheet sheet, IReadOnlyList<ParameterColumn> columns)
        {
            sheet.Cell(1, 1).Value = "id_имя семейства_имя типа";
            sheet.Cell(1, 2).Value = "ElementId";
            sheet.Cell(1, 3).Value = "Category";
            sheet.Cell(1, 4).Value = "Family";
            sheet.Cell(1, 5).Value = "Type";

            int startCol = 6;
            for (int i = 0; i < columns.Count; i++)
            {
                sheet.Cell(1, startCol + i).Value = columns[i].Name;
            }

            sheet.Row(1).Style.Font.Bold = true;
        }

        private static void WriteRows(IXLWorksheet sheet, Document document, IReadOnlyList<Element> elements, IReadOnlyList<ParameterColumn> columns)
        {
            int row = 2;
            foreach (Element element in elements)
            {
                GetFamilyAndTypeNames(document, element, out string familyName, out string typeName);
                string id = element.Id.IntegerValue.ToString(CultureInfo.InvariantCulture);

                sheet.Cell(row, 1).Value = $"{id}_{familyName}_{typeName}";
                sheet.Cell(row, 2).Value = id;
                sheet.Cell(row, 3).Value = element.Category?.Name ?? string.Empty;
                sheet.Cell(row, 4).Value = familyName;
                sheet.Cell(row, 5).Value = typeName;

                Dictionary<string, Parameter> available = BuildAvailableParameterMap(document, element);

                int startCol = 6;
                for (int i = 0; i < columns.Count; i++)
                {
                    ParameterColumn column = columns[i];
                    if (!available.TryGetValue(column.Key, out Parameter parameter))
                    {
                        sheet.Cell(row, startCol + i).Value = "no";
                        continue;
                    }

                    sheet.Cell(row, startCol + i).Value = ToCellValue(parameter);
                }

                row++;
            }
        }

        private static Dictionary<string, Parameter> BuildAvailableParameterMap(Document document, Element element)
        {
            var map = new Dictionary<string, Parameter>(StringComparer.Ordinal);

            void addFrom(Element source)
            {
                if (source == null)
                {
                    return;
                }

                foreach (Parameter parameter in source.Parameters)
                {
                    if (!TryCreateColumn(parameter, out ParameterColumn column))
                    {
                        continue;
                    }

                    if (!map.ContainsKey(column.Key))
                    {
                        map[column.Key] = parameter;
                    }
                }
            }

            addFrom(element);

            ElementId typeId = element.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                addFrom(document.GetElement(typeId));
            }

            return map;
        }

        private static string ToCellValue(Parameter parameter)
        {
            if (parameter == null)
            {
                return "";
            }

            switch (parameter.StorageType)
            {
                case StorageType.String:
                    return parameter.AsString() ?? string.Empty;

                case StorageType.Integer:
                    return parameter.AsInteger().ToString(CultureInfo.InvariantCulture);

                case StorageType.Double:
                    return parameter.AsValueString() ?? parameter.AsDouble().ToString(CultureInfo.InvariantCulture);

                case StorageType.ElementId:
                    ElementId valueId = parameter.AsElementId();
                    return valueId == null ? string.Empty : valueId.IntegerValue.ToString(CultureInfo.InvariantCulture);

                case StorageType.None:
                default:
                    return parameter.AsValueString() ?? string.Empty;
            }
        }

        private static void GetFamilyAndTypeNames(Document document, Element element, out string familyName, out string typeName)
        {
            familyName = string.Empty;
            typeName = string.Empty;

            if (element is FamilyInstance familyInstance)
            {
                familyName = familyInstance.Symbol?.FamilyName ?? string.Empty;
                typeName = familyInstance.Symbol?.Name ?? string.Empty;
                return;
            }

            Element typeElement = null;
            ElementId typeId = element.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                typeElement = document.GetElement(typeId);
            }

            if (typeElement is ElementType elementType)
            {
                familyName = elementType.FamilyName ?? string.Empty;
                typeName = elementType.Name ?? string.Empty;
            }
            else
            {
                familyName = element.Category?.Name ?? string.Empty;
                typeName = element.Name ?? string.Empty;
            }
        }
    }
}
