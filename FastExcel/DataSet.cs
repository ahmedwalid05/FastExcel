using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace FastExcel
{
    public class DataSet
    {
        public IEnumerable<Row> Rows { get; set; }

        public IEnumerable<string> Headings { get; set; }

        public DataSet() { }

        public void PopulateRows<T>(IEnumerable<T> objects, bool usePropertiesAsHeadings = false)
        {
            int rowNumber = 1;

            // Get all properties
            PropertyInfo[] properties = typeof(T).GetProperties();
            List<Row> rows = new List<Row>();
            
            if (usePropertiesAsHeadings)
            {
                this.Headings = (from prop in properties
                                 select prop.Name);

                int headingColumnNumber = 1;
                IEnumerable<Cell> headingCells = (from h in this.Headings
                                                   select new Cell(headingColumnNumber++, h)).ToArray();

                Row headingRow = new Row(rowNumber++, headingCells);

                rows.Add(headingRow);
            }

            foreach (T rowObject in objects)
            {
                List<Cell> cells = new List<Cell>();
                
                int columnNumber = 1;

                // Get value from each property
                foreach (PropertyInfo propertyInfo in properties)
                {
                    object value = propertyInfo.GetValue(rowObject, null);
                    if(value != null)
                    {
                        Cell cell = new Cell(columnNumber, value);
                        cells.Add(cell);
                    }
                    columnNumber++;
                }

                Row row = new Row(rowNumber++, cells);
                rows.Add(row);
            }

            this.Rows = rows;
        }


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
