using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace FastExcel
{
    public class Cell
    {
        public int ColumnNumber { get; set; }
        public object Value { get; set; }

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
            this.ColumnNumber = columnNumber;
            this.Value = value;
        }

        public Cell(XElement cellElement, SharedStrings sharedStrings)
        {
            bool isTextRow = (from a in cellElement.Attributes("t")
                              where a.Value == "s"
                              select a).Any();
            string columnName = (from a in cellElement.Attributes("r")
                                 select a.Value).FirstOrDefault();

            this.ColumnNumber = GetExcelColumnNumber(columnName);

            if (isTextRow)
            {
                this.Value = sharedStrings.GetString(cellElement.Value);
            }
            else
            {
                this.Value = cellElement.Value;
            }
        }

        internal StringBuilder ToString(SharedStrings sharedStrings, int rowNumber)
        {
            StringBuilder cell = new StringBuilder();

            if (this.Value != null)
            {
                bool isString = false;
                object value = this.Value;

                if (this.Value is int)
                {
                    isString = false;
                }
                else if (this.Value is double)
                {
                    isString = false;
                }
                else if (this.Value is string)
                {
                    isString = true;
                }

                if (isString)
                {
                    value = sharedStrings.AddString(value.ToString());
                }

                cell.AppendFormat("<c r=\"{0}{1}\"{2}>", GetExcelColumnName(this.ColumnNumber), rowNumber, (isString ? " t=\"s\"" : string.Empty));
                cell.AppendFormat("<v>{0}</v>", value);
                cell.Append("</c>");
            }

            return cell;
        }

        //http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
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
            this.Value = cell.Value;
        }
    }
}
