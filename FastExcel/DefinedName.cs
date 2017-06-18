﻿using System;
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
                throw new DefinedNameLoadException("Unable to load workbook.xml from open stream. Probable corrupt file.");
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
        /// <param name="worksheetIndex">If scoped to a sheet, the worksheetIndex</param>
        /// <returns>List of cells encapsulated in another list representing seperate ranges</returns>
        public IEnumerable<IEnumerable<Cell>> GetCellRangesByDefinedName(string definedName, int? worksheetIndex = null)
        {
            List<List<Cell>> result = new List<List<Cell>>();

            string key = (!worksheetIndex.HasValue ? definedName : definedName + ":" + worksheetIndex);

            if (!DefinedNames.ContainsKey(key))
                return result;

            string[] references = DefinedNames[key].Reference.Split(',');

            foreach(string reference in references)
            {
                // If A) is a formula or B) does not contain ! or $ then this reference is not supported
                if (Regex.IsMatch(reference, @"[\(\)\*\+\-\/]") || !Regex.IsMatch(reference, @"[!$]"))
                    continue;

                CellRange cellRange = new CellRange(reference);

                Worksheet worksheet = Read(cellRange.SheetName);

                result.Add(worksheet.GetCellsInRange(cellRange).ToList());
            }

            return result;
        }

        /// <summary>
        /// Gets all cells by defined name
        /// Like GetCellRangesByCellName, but just retreives all cells in a single list
        /// </summary>
        /// <param name="definedName"></param>
        /// <param name="worksheetIndex"></param>
        /// <returns></returns>
        public IEnumerable<Cell> GetCellsByDefinedName(string definedName, int? worksheetIndex = null)
        {
            var cells = new List<Cell>();
            var cellRanges = GetCellRangesByDefinedName(definedName) as List<List<Cell>>;

            foreach (var cellRange in cellRanges)
                cells.InsertRange(cells.Count, (from cell in cellRange select cell));

            return cells;
        }

        /// <summary>
        /// Returns cell by defined name
        /// If theres more than one, this is the first one.
        /// </summary>
        /// <param name="definedName"></param>
        /// <returns></returns>
        public Cell GetCellByDefinedName(string definedName, int? worksheetIndex = null)
        {
            return GetCellsByDefinedName(definedName, worksheetIndex).FirstOrDefault();
        }


        /// <summary>
        /// Returns all cells in a column by name, within optional row range
        /// </summary>
        /// <param name="columnName"></param>
        /// <param name="rowStart"></param>
        /// <param name="rowEnd"></param>
        /// <returns></returns>
        public IEnumerable<Cell> GetCellsByColumnName(string columnName, int rowStart = 1, int? rowEnd = null)
        {
            var columnCells = GetCellsByDefinedName(columnName) as List<Cell>;
            if (!rowEnd.HasValue)
                rowEnd = columnCells.Last().RowNumber;
            return columnCells.Where(cell=>cell.RowNumber>=rowStart && cell.RowNumber<=rowEnd).ToList();
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
