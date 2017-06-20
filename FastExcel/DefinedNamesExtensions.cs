using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FastExcel
{
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
}
