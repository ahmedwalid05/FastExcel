using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastExcel
{
    public class DataSet
    {
        public IEnumerable<Row> Rows { get; set; }

        public IEnumerable<string> Headings { get; set; }

        /// <summary>
        /// This method is slow for large datasets, use the rows property instead
        /// </summary>
        public void AddValue(int rowNumber, int columnNumber, object value)
        {
            if (this.Rows == null)
            {
                this.Rows = new List<Row>();
            }

            Row row = (from r in this.Rows
                       where r.RowNumber == rowNumber
                       select r).FirstOrDefault();
            Cell cell = null;

            if (row == null)
            {
                cell = new Cell(columnNumber, value);
                row = new Row(rowNumber, new List<Cell>{ cell });
                (this.Rows as List<Row>).Add(row);
            }

            if (cell == null)
            {
                cell = (from c in row.Cells
                        where c.ColumnNumber == columnNumber
                        select c).FirstOrDefault();

                if (cell == null)
                {
                    cell = new Cell(columnNumber, value);
                    (row.Cells as List<Cell>).Add(cell);
                }
            }

        }

        // TODO: create function to convert data to dynamic types


        /// <summary>
        /// Merges the parameter into the current DatSet object, the parameter takes precedence
        /// </summary>
        /// <param name="data">A DataSet to merge</param>
        public void Merge(DataSet data)
        {
            // Merge headings
            if (this.Headings == null || !this.Headings.Any())
            {
                this.Headings = data.Headings;
            }

            // Merge rows
            List<Row> outputList = new List<Row>();
            foreach (var row in this.Rows.Union(data.Rows).GroupBy(r => r.RowNumber))
            {
                int count = row.Count();
                if (count == 1)
                {
                    outputList.Add(row.First());
                }
                else
                {
                    row.First().Merge(row.Skip(1).First());

                    outputList.Add(row.First());
                }
            }

            // Sort
            this.Rows = (from r in outputList
                        orderby r.RowNumber
                        select r);
        }
    }
}
