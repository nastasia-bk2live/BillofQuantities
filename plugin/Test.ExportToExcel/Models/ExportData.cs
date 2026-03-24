using System.Collections.Generic;

namespace Test.ExportToExcel.Models
{
    /// <summary>
    /// Собранный набор данных для выгрузки.
    /// </summary>
    public class ExportData
    {
        public IList<ParameterColumn> ParameterColumns { get; set; }

        public IList<ElementExportRow> Rows { get; set; }
    }
}
