using System.Collections.Generic;

namespace Test.ExportToExcel.Models
{
    /// <summary>
    /// Строка данных одного экземпляра элемента для экспорта.
    /// </summary>
    public class ElementExportRow
    {
        public string Key { get; set; }

        public string ElementId { get; set; }

        public string Category { get; set; }

        public string Family { get; set; }

        public string Type { get; set; }

        public IDictionary<ParameterColumn, string> Parameters { get; set; }
    }
}
