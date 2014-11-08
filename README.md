#Fast Excel

Currently provides a fast way of writing to *.xlsx Excel files.

I am not using the Open XML SDK to interact with the data but going directly and editing the underlying xml files.

.Net version 4.5 is required because it uses System.IO.Compression


Check out the demo project for usage and benchmark testing against EPPlus.
This project is not intended to be a replacement for full featured packages like EPPlus, just light weight fast way of saving data to Excel.

##Demo

```C#
// Get your template and output file paths
FileInfo templateFile = new FileInfo("Template.xlsx");
FileInfo outputFile = new FileInfo("C:\Temp\output.xlsx"));

// Create an instance of the writer
FastExcel.FastExcelWriter writer = new FastExcel.FastExcelWriter(templateFile, outputFile);

//Create a data set some rows data
DataSet data = new DataSet();
List<Row> rows = new List<Row>();

for (int rowNumber = 1; rowNumber < 100000; rowNumber++)
{
	List<Cell> cells = new List<Cell>();
    for (int columnNumber = 1; columnNumber < 13; columnNumber++)
    {
    	cells.Add(new Cell(columnNumber, columnNumber * DateTime.Now.Millisecond));
    }
    cells.Add(new Cell(13, "Hello" + rowNumber));
    cells.Add(new Cell(14, "Some Text"));
    
    rows.Add(new Row(rowNumber, cells));
}
data.Rows = rows;

// Write the data
writer.Write(data, null, "sheet1", 1);
```