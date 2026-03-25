using System;

namespace Test.ExportToExcel.Models
{
    public enum ParameterSource
    {
        BuiltIn = 0,
        Project = 1,
        Shared = 2
    }

    /// <summary>
    /// Описание колонки параметра в итоговом CSV.
    /// </summary>
    public class ParameterColumn : IEquatable<ParameterColumn>
    {
        public ParameterColumn(ParameterSource source, string name)
        {
            Source = source;
            Name = name;
        }

        public ParameterSource Source { get; }

        public string Name { get; }

        public bool Equals(ParameterColumn other)
        {
            if (other == null)
            {
                return false;
            }

            return Source == other.Source && string.Equals(Name, other.Name, StringComparison.OrdinalIgnoreCase);
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as ParameterColumn);
        }

        public override int GetHashCode()
        {
            return ((int)Source * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(Name ?? string.Empty);
        }
    }
}
