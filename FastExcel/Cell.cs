using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastExcel
{
    public class Cell
    {
        public int ColumnNumber { get; set; }
        public object Value { get; set; }

        public Cell(int columnNumber, object value)
        {
            if (columnNumber <= 0)
            {
                // TODO error
            }
            this.ColumnNumber = columnNumber;
            this.Value = value;
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
        private string GetExcelColumnName(int columnNumber)
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
    }
}
