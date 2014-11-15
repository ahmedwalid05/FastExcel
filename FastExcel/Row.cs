using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

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

        public Row(XElement rowElement, SharedStrings sharedStrings)
        {
            try
            {
                this.RowNumber = (from a in rowElement.Attributes("r")
                             select int.Parse(a.Value)).First();
            }
            catch (Exception ex)
            {
                throw new Exception("Row Number not found", ex);
            }

            List<Cell> cells = new List<Cell>();
            
            if (rowElement.HasElements)
            {
                foreach (XElement cellElement in rowElement.Elements())
                {
                    Cell cell = new Cell(cellElement, sharedStrings);
                    if (cell.Value != null)
                    {
                        cells.Add(cell);
                    }
                }
            }

            this.Cells = cells;
        }

        internal StringBuilder ToXmlString(SharedStrings sharedStrings)
        {
            StringBuilder row = new StringBuilder();

            if (this.Cells != null && Cells.Any())
            {
                row.AppendFormat("<row r=\"{0}\">", this.RowNumber);
                try
                {
                    foreach (Cell cell in this.Cells)
                    {
                        row.Append(cell.ToXmlString(sharedStrings, this.RowNumber));
                    }
                }
                finally
                {
                    row.Append("</row>");
                }
            }

            return row;
        }

        internal void Merge(Row row)
        {
            // Merge cells
            List<Cell> outputList = new List<Cell>();
            foreach (var cell in this.Cells.Union(row.Cells).GroupBy(c => c.ColumnNumber))
            {
                int count = cell.Count();
                if (count == 1)
                {
                    outputList.Add(cell.First());
                }
                else
                {
                    cell.First().Merge(cell.Skip(1).First());

                    outputList.Add(cell.First());
                }
            }

            // Sort
            this.Cells = (from c in outputList
                          orderby c.ColumnNumber
                          select c);
        }
    }
}
