using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace FastExcel
{
    /// <summary>
    /// Row that contains the Cells
    /// </summary>
    public class Row
    {
        /// <summary>
        /// The Row Number (Row numbers start at 1)
        /// </summary>
        public int RowNumber { get; set; }

        /// <summary>
        /// The collection of cells for this row
        /// </summary>
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
                        
            if (rowElement.HasElements)
            {
                this.Cells = GetCells(rowElement, sharedStrings);
            }
        }

        private IEnumerable<Cell> GetCells(XElement rowElement, SharedStrings sharedStrings)
        {
            foreach (XElement cellElement in rowElement.Elements())
            {
                Cell cell = new Cell(cellElement, sharedStrings);
                if (cell.Value != null)
                {
                    yield return cell;
                }
            }
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

        /// <summary>
        /// Merge this row and the passed one togeather
        /// </summary>
        /// <param name="row">Row to be merged into this one</param>
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
