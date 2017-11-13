using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        internal int? worksheetIndex { get; }
        internal string Reference { get; }
        internal string Key { get { return Name + (!worksheetIndex.HasValue ? "" : ":" + worksheetIndex); } }

        internal DefinedName(XElement e)
        {
            Name = e.Attribute("name").Value;
            if (e.Attribute("localSheetId") != null)
                try
                {
                    worksheetIndex = Convert.ToInt32(e.Attribute("localSheetId").Value)+1;
                }
                catch (Exception exception)
                {
                    // In a well formed file, this should never happen.
                    throw new DefinedNameLoadException("Error reading localSheetId value for DefinedName: '" + Name + "'", exception);
                }
            Reference = e.Value;
        }
    }
}
