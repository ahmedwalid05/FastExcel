using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace FastExcel
{
    partial class FastExcel
    {
        /// <summary>
        /// Dictionary of defined names
        /// </summary>
        internal IReadOnlyDictionary<string, DefinedName> DefinedNames { get; private set; }

        private void loadDefinedNames()
        {
            using (Stream stream = Archive.GetEntry("xl/workbook.xml").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    throw new Exception("Unable to load workbook.xml");
                }

                var definedNames = new Dictionary<string, DefinedName>();

                foreach (var e in (from d2 in document.Descendants().Where(dx => dx.Name.LocalName == "definedNames").Descendants()
                                   select d2))
                {
                    var currentDefinedName = new DefinedName(e);
                    definedNames.Add(currentDefinedName.Key, currentDefinedName);
                }
                DefinedNames = definedNames;
            }
        }
    }


    /// <summary>
    /// Reads/hold information from XElement representing a stored DefinedName
    /// A defined name is an alias for a cell, multiple cells, a range of cells or multiple ranges of cells
    /// It is also used as an alias for a column
    /// </summary>
    internal class DefinedName
    {
        internal string Name { get; }
        internal int? SheetId { get; }
        internal string Reference { get; }
        internal string Key { get { return Name + (SheetId == null ? "" : ":" + SheetId); } }

        internal DefinedName(XElement e)
        {
            Name = e.Attribute("name").Value;
            if (e.Attribute("localSheetId") != null)
                try
                {
                    SheetId = Convert.ToInt32(e.Attribute("localSheetId").Value);
                }
                catch (Exception exception)
                {
                    throw new DefinedNameLoadException("Error reading localSheetId value for DefinedName: '" + Name + "'", exception);
                }
            Reference = e.Value;
        }
    }
    /// <summary>
    /// Extensions to use on Dictionary of DefinedNames
    /// </summary>
    internal static class DefinedNamesExtensions
    {
        /// <summary>
        /// Finds all the cell names for a given cell
        /// </summary>
        /// <param name="definedNames"></param>
        /// <param name="sheetName">Name of sheet containing cell</param>
        /// <param name="columnLetter">Column letter of cell</param>
        /// <param name="rowNumber">Row number of cell(this is string as thats datatype when reading from XML)</param>
        /// <returns>
        /// List of cell names that is assigned to this cell. Does not include names which this cell is within range.
        /// Empty List if none found
        /// </returns>
        internal static List<string> FindCellNames(this IReadOnlyDictionary<string, DefinedName> definedNames, string sheetName, string columnLetter, string rowNumber)
        {
            return (from e in definedNames where e.Value.Reference.Contains(sheetName + "!$" + columnLetter.ToUpper() + "$" + rowNumber.ToUpper()) select e.Value.Name).ToList();
        }
        /// <summary>
        /// Finds the column name for a given column letter
        /// </summary>
        /// <param name="definedNames"></param>
        /// <param name="sheetName">Name of sheet containing column</param>
        /// <param name="columnLetter">Column letter</param>
        /// <returns></returns>
        internal static string FindColumnName(this IReadOnlyDictionary<string, DefinedName> definedNames, string sheetName, string columnLetter)
        {
            columnLetter = columnLetter.ToUpper();
            return (from e in definedNames where e.Value.Reference == sheetName + "!$" + columnLetter + ":$" + columnLetter select e.Value.Name).FirstOrDefault();
        }

        
    }

    public class DefinedNameLoadException : Exception
    {
        public DefinedNameLoadException(string message, Exception innerException = null)
            : base(message, innerException)
        {

        }
    }
}
