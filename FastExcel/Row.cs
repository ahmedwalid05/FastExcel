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

        public Row(int rowNumber, IEnumerable<Cell> cells)
        {
            if (rowNumber <= 0)
            {
                // TODO error
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
