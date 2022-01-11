using System;
using System.Xml.Linq;

namespace FastExcel
{
    /// <summary>
    /// Reads/hold information from XElement representing a stored DefinedName
    /// A defined name is an alias for a cell, multiple cells, a range of cells or multiple ranges of cells
    /// It is also used as an alias for a column
    /// </summary>
    internal class DefinedName
    {
        internal string Name { get; }
        internal int? WorksheetIndex { get; }
        internal string Reference { get; }
        internal string Key { get { return Name + (!WorksheetIndex.HasValue ? "" : ":" + WorksheetIndex); } }

        internal DefinedName(XElement e)
        {
            Name = e.Attribute("name").Value;
            if (e.Attribute("localSheetId") != null)
            {
                try
                {
                    WorksheetIndex = Convert.ToInt32(e.Attribute("localSheetId").Value) + 1;
                }
                catch (Exception exception)
                {
                    // In a well formed file, this should never happen.
                    throw new DefinedNameLoadException("Error reading localSheetId value for DefinedName: '" + Name + "'", exception);
                }
            }

            Reference = e.Value;
        }
    }
}
