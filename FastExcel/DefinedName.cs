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
    partial class FastExcel
    {
        /// <summary>
        /// Dictionary of defined names
        /// </summary>
        internal IReadOnlyDictionary<string, DefinedName> DefinedNames { get; private set; }

        private void loadDefinedNames()
        {
            XDocument document;
            try
            {
                document = XDocument.Load(Archive.GetEntry("xl/workbook.xml").Open());
            }
            catch (Exception exception)
            {
                throw new DefinedNameLoadException("Unable to open stream to read internal workbook.xml file", exception);
            }
            if (document == null)
            {
                throw new DefinedNameLoadException("Unable to load workbook.xml file stream");
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

        /// <summary>
        /// Retrieves ranges of cells by their defined name
        /// </summary>
        /// <param name="definedName">Defined Name</param>
        /// <param name="sheetId">If scoped to a sheet, the sheetId</param>
        /// <returns>List of cells encapsulated in another list representing seperate ranges</returns>
        public IEnumerable<IEnumerable<Cell>> GetCellRangesByDefinedName(string definedName, int? sheetId = null)
        {
            List<List<Cell>> result = new List<List<Cell>>();

            string key = (sheetId == null) ? definedName : definedName + ":" + sheetId;

            if (!DefinedNames.ContainsKey(key))
                return result;

            string[] references = DefinedNames[key].Reference.Split(',');

            foreach(string reference in references)
            {
                // If not containing these characters then its a reference that's not supported
                if (!reference.Contains("!") || !reference.Contains("$"))
                    continue;

                CellRange cellRange = new CellRange(reference);

                Worksheet worksheet = Read(cellRange.SheetName);

                result.Add(worksheet.GetCellsInRange(cellRange).ToList());
            }

            return result;
        }

        /// <summary>
        /// Returns first range of cells by defined names
        /// Use when you know defined name only represents one range
        /// </summary>
        /// <param name="definedName"></param>
        /// <returns></returns>
        public IEnumerable<Cell> GetCellRangeByDefinedName(string definedName)
        {
            return GetCellRangesByDefinedName(definedName).FirstOrDefault();
        }

        /// <summary>
        /// Returns first cell by defined name
        /// Use when you know defined name only represents one cell
        /// </summary>
        /// <param name="definedName"></param>
        /// <returns></returns>
        public Cell GetFirstCellByDefinedName(string definedName)
        {
            return GetCellRangeByDefinedName(definedName).FirstOrDefault();
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
                    // In a well formed file, this should never happen.
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
        /// <param name="rowNumber">Row number of cell</param>
        /// <returns>
        /// List of cell names that is assigned to this cell. Does not include names which this cell is within range.
        /// Empty List if none found
        /// </returns>
        internal static List<string> FindCellNames(this IReadOnlyDictionary<string, DefinedName> definedNames, string sheetName, string columnLetter, int rowNumber)
        {
            return (from e in definedNames where e.Value.Reference.Contains(sheetName + "!$" + columnLetter.ToUpper() + "$" + rowNumber.ToString()) select e.Value.Name).ToList();
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

    /// <summary>
    /// Exception used during loading process
    /// </summary>
    public class DefinedNameLoadException : Exception
    {
        public DefinedNameLoadException(string message, Exception innerException = null)
            : base(message, innerException)
        {

        }
    }
}
