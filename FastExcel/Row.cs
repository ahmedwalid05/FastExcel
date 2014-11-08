using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastExcel
{
    public class Row
    {
        public int RowNumber { get; set; }
        public IEnumerable<Cell> Cells { get; set; }

        /// <summary>
        /// Create a new Row
        /// </summary>
        /// <param name="rowNumber">Row number starting with 1</param>
        /// <param name="cells">Cells on this row</param>
        public Row(int rowNumber, IEnumerable<Cell> cells)
        {
            if (rowNumber <= 0)
            {
                throw new Exception("Row numbers starting at 1");
            }
            this.RowNumber = rowNumber;
            this.Cells = cells;
        }

        internal StringBuilder ToString(SharedStrings sharedStrings)
        {
            StringBuilder row = new StringBuilder();

            if (this.Cells != null && Cells.Any())
            {
                row.AppendFormat("<row r=\"{0}\">", this.RowNumber);
                try
                {
                    foreach (Cell cell in this.Cells)
                    {
                        row.Append(cell.ToString(sharedStrings, this.RowNumber));
                    }
                }
                finally
                {
                    row.Append("</row>");
                }
            }

            return row;
        }
    }
}
