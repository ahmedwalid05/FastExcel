#Fast Excel

Currently provides a fast way of reading and writing to *.xlsx Excel files.

More features to come, fell free to suggest something.

I am not using the Open XML SDK to interact with the data but going directly and editing the underlying xml files.

.Net version 4.5 is required because it uses System.IO.Compression

Check out the demo project for usage and benchmarking.

This project is not intended to be a replacement for full featured Excel packages with things like formatting, just light weight fast way of interacting with data in Excel.

Below are a few demos check out https://github.com/mrjono1/FastExcel/blob/master/FastExcelDemo/Program.cs for more.

##Write Demo 1
This demo uses Generic objects, ie any object you wish with public properties
```C#
using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(outputFile))
{
    List<MyObject> objectList = new List<MyObject>();

    for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber++)
    {
        MyObject genericObject = new MyObject();
        genericObject.StringColumn1 = "A string " + rowNumber.ToString();
        genericObject.IntegerColumn2 = 45678854;
        genericObject.DoubleColumn3 = 87.01d;
        genericObject.ObjectColumn4 = DateTime.Now.ToLongTimeString();

        objectList.Add(genericObject);
    }
    fastExcel.Write(objectList, "sheet3", true);
}
```

##Write Demo 2
This demo lets you specify exactly which cell you are writing to

```C#
// Get your template and output file paths
FileInfo templateFile = new FileInfo("Template.xlsx");
FileInfo outputFile = new FileInfo("C:\\Temp\\output.xlsx");

//Create a data set
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


// Create an instance of FastExcel
using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(templateFile, outputFile))
{
    // Write the data
    fastExcel.Write(data, "sheet1");
}
```

##Read Demo

```C#
// Get the input file paths
FileInfo inputFile = new FileInfo("C:\\Temp\\input.xlsx");

//Create a data set
DataSet data = null;

// Create an instance of Fast Excel
using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
{
    // Read the data
    data = fastExcel.Read("sheet1");
}
```

##Update Demo

```C#
// Get the input file paths
FileInfo inputFile = new FileInfo("C:\\Temp\\input.xlsx");

//Create a data set
DataSet data = new DataSet();
List<Row> rows = new List<Row>();
                
for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber+= 50)
{
    List<Cell> cells = new List<Cell>();
    for (int columnNumber = 1; columnNumber < 13; columnNumber+= 2)
    {
        cells.Add(new Cell(columnNumber, rowNumber));
    }
    cells.Add(new Cell(13, "Updated Row"));

    rows.Add(new Row(rowNumber, cells));
}
data.Rows = rows;

// Create an instance of Fast Excel
using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile))
{
    // Read the data
    data = fastExcel.Update("sheet1");
}
```