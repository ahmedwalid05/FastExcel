using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace FastExcel
{
    /// <summary>
    /// Contains the actual value
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// Column Numnber (Starts at 1)
        /// </summary>
        public int ColumnNumber { get; set; }

        /// <summary>
        /// The value that is stored
        /// </summary>
        public object Value { get; set; }

        /// <summary>
        /// Defined name or the column letter(s) for column this cell is in
        /// </summary>
        public string ColumnName { get; }
        
        /// <summary>
        /// Raw underlying XElement of cell
        /// </summary>
        public XElement XElement { get; }
        
        /// <summary>
        /// List of defined names assigned to this cell
        /// *Does not include names of ranges this cell is within*
        /// </summary>
        public IReadOnlyList<string> CellNames { get; }

        /// <summary>
        /// First defined name assigned to cell or cell Letter + row number if no named defined.
        /// </summary>
        public string CellName
        {
            get
            {
                if (CellNames.Any())
                    return CellNames.FirstOrDefault();
                return ColumnName + RowNumber;
            }
        }

        /// <summary>
        /// Number of the row this cell is on
        /// </summary>
        public int RowNumber { get; }

        /// <summary>
        /// Create a new Cell
        /// </summary>
        /// <param name="columnNumber">Column number starting at 1</param>
        /// <param name="value">Cell Value</param>
        public Cell(int columnNumber, object value)
        {
            if (columnNumber <= 0)
            {
                throw new Exception("Column numbers starting at 1");
            }
            ColumnNumber = columnNumber;
            ColumnName = GetExcelColumnName(columnNumber);
            RowNumber = 0;
            Value = value;
        }

        /// <summary>
        /// Create a new Cell
        /// </summary>
        /// <param name="cellElement">Cell</param>
        /// <param name="worksheet">The worksheet that this cell is on</param>
        public Cell(XElement cellElement, Worksheet worksheet)
        {
            bool isTextRow = (from a in cellElement.Attributes("t")
                              where a.Value == "s"
                              select a).Any();
            string columnName = (from a in cellElement.Attributes("r")
                                 select a.Value).FirstOrDefault();

            string columnLetter = Regex.Replace(columnName, @"\d", "");

            RowNumber = Convert.ToInt32(Regex.Replace(columnName, @"[^\d]", ""));

            ColumnName = worksheet.FastExcel.DefinedNames.FindColumnName(worksheet.Name, columnLetter) ?? columnLetter;

            CellNames = worksheet.FastExcel.DefinedNames.FindCellNames(worksheet.Name, columnLetter, RowNumber);

            ColumnNumber = GetExcelColumnNumber(columnName);

            XElement = cellElement;

            if (isTextRow)
            {
                Value = worksheet.FastExcel.SharedStrings.GetString(cellElement.Value);
            }
            else
            {
                // cellElement.Value will give a concatenated Value + reference/calculation

                var node = cellElement.Elements().Where(x => x.Name.LocalName == "v").SingleOrDefault();

                if (node != null)
                {
                    Value = node.Value;
                }
                else
                {
                    Value = cellElement.Value;
                }
                
            }
        }

        internal StringBuilder ToXmlString(SharedStrings sharedStrings, int rowNumber)
        {
            StringBuilder cell = new StringBuilder();

            if (Value != null)
            {
                bool isString = false;
                object value = Value;

                if (Value is int)
                {
                    isString = false;
                }
                else if (Value is double)
                {
                    isString = false;
                }
                else if (Value is string)
                {
                    isString = true;
                }

                if (isString)
                {
                    value = sharedStrings.AddString(value.ToString());
                }

                cell.AppendFormat("<c r=\"{0}{1}\"{2}>", GetExcelColumnName(ColumnNumber), rowNumber, (isString ? " t=\"s\"" : string.Empty));
                cell.AppendFormat("<v>{0}</v>", value);
                cell.Append("</c>");
            }

            return cell;
        }

        //http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
        /// <summary>
        /// Convert Column Number into Column Name - Character(s) eg 1-A, 2-B
        /// </summary>
        /// <param name="columnNumber">Column Number</param>
        /// <returns>Column Name - Character(s)</returns>
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = string.Concat(Convert.ToChar(65 + modulo), columnName);
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        //http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
        /// <summary>
        /// Covert Column Name - Character(s) into a Column Number eg A-1, B-2, A1 - 1, B9 - 2
        /// </summary>
        /// <param name="columnName">Column Name - Character(s) optinally with the Row Number</param>
        /// <param name="includesRowNumber">Specify if the row number is included</param>
        /// <returns>Column Number</returns>
        public static int GetExcelColumnNumber(string columnName, bool includesRowNumber = true)
        {
            if (includesRowNumber)
            {
                columnName = Regex.Replace(columnName, @"\d", "");
            }

            int[] digits = new int[columnName.Length];
            for (int i = 0; i < columnName.Length; ++i)
            {
                digits[i] = Convert.ToInt32(columnName[i]) - 64;
            }
            int mul = 1; int res = 0;
            for (int pos = digits.Length - 1; pos >= 0; --pos)
            {
                res += digits[pos] * mul;
                mul *= 26;
            }
            return res;
        }

        /// <summary>
        /// Merge the parameter cell into this cell
        /// </summary>
        /// <param name="cell">Cell to merge</param>
        public void Merge(Cell cell)
        {
            Value = cell.Value;
        }
    }
}