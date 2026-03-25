using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Test.ExportToExcel.Models;

namespace Test.ExportToExcel.Services
{
    /// <summary>
    /// Сервис сбора данных из документа Revit для последующего экспорта в CSV.
    /// </summary>
    public class ElementExportDataService
    {
        private const int ProgressReportStep = 200;

        public ExportData Collect(Document document, Action<int, int> progressCallback)
        {
            var elements = new FilteredElementCollector(document)
                .WhereElementIsNotElementType()
                .ToElements();

            var rows = new List<ElementExportRow>(elements.Count);
            var columnSet = new HashSet<ParameterColumn>();

            for (var index = 0; index < elements.Count; index++)
            {
                var element = elements[index];
                var row = BuildRow(document, element, columnSet);
                rows.Add(row);

                var current = index + 1;
                if (current % ProgressReportStep == 0 || current == elements.Count)
                {
                    progressCallback?.Invoke(current, elements.Count);
                }
            }

            var orderedColumns = columnSet
                .OrderBy(c => c.Source)
                .ThenBy(c => c.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            return new ExportData
            {
                ParameterColumns = orderedColumns,
                Rows = rows
            };
        }

        private ElementExportRow BuildRow(Document document, Element element, ISet<ParameterColumn> columnSet)
        {
            var familyName = ResolveFamilyName(document, element);
            var typeName = ResolveTypeName(document, element);

            var row = new ElementExportRow
            {
                ElementId = element.Id.IntegerValue.ToString(),
                Category = element.Category != null ? element.Category.Name : string.Empty,
                Family = familyName,
                Type = typeName,
                Key = element.Id.IntegerValue + "_" + familyName + "_" + typeName,
                Parameters = new Dictionary<ParameterColumn, string>()
            };

            foreach (Parameter parameter in element.Parameters)
            {
                if (parameter == null || parameter.Definition == null)
                {
                    continue;
                }

                var source = ResolveParameterSource(parameter);
                var column = new ParameterColumn(source, parameter.Definition.Name);
                columnSet.Add(column);
                row.Parameters[column] = ResolveParameterValue(parameter);
            }

            return row;
        }

        private static ParameterSource ResolveParameterSource(Parameter parameter)
        {
            var internalDefinition = parameter.Definition as InternalDefinition;
            if (internalDefinition != null &&
                internalDefinition.BuiltInParameter != BuiltInParameter.INVALID)
            {
                return ParameterSource.BuiltIn;
            }

            if (parameter.IsShared)
            {
                return ParameterSource.Shared;
            }

            return ParameterSource.Project;
        }

        private static string ResolveParameterValue(Parameter parameter)
        {
            if (parameter == null)
            {
                return "no";
            }

            if (!parameter.HasValue)
            {
                return string.Empty;
            }

            if (parameter.StorageType == StorageType.ElementId)
            {
                var idValue = parameter.AsElementId();
                return idValue != null ? idValue.IntegerValue.ToString() : string.Empty;
            }

            var valueAsString = parameter.AsValueString();
            if (!string.IsNullOrEmpty(valueAsString))
            {
                return valueAsString;
            }

            switch (parameter.StorageType)
            {
                case StorageType.String:
                    return parameter.AsString() ?? string.Empty;
                case StorageType.Integer:
                    return parameter.AsInteger().ToString();
                case StorageType.Double:
                    return parameter.AsDouble().ToString(System.Globalization.CultureInfo.InvariantCulture);
                case StorageType.None:
                    return string.Empty;
                default:
                    return string.Empty;
            }
        }

        private static string ResolveTypeName(Document document, Element element)
        {
            var typeElement = document.GetElement(element.GetTypeId()) as ElementType;
            return typeElement != null ? typeElement.Name : element.Name;
        }

        private static string ResolveFamilyName(Document document, Element element)
        {
            var familyInstance = element as FamilyInstance;
            if (familyInstance != null)
            {
                return familyInstance.Symbol != null && familyInstance.Symbol.FamilyName != null
                    ? familyInstance.Symbol.FamilyName
                    : string.Empty;
            }

            var typeElement = document.GetElement(element.GetTypeId()) as ElementType;
            if (typeElement != null && typeElement.FamilyName != null)
            {
                return typeElement.FamilyName;
            }

            return string.Empty;
        }
    }
}
