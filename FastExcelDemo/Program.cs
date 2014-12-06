using FastExcel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastExcelDemo
{
    public class Program
    {
        private int NumberOfRecords = 100000;
        private string DemoDir = "C:\\Temp";
        private bool EPPlusTest = false;

        public static void Main(string[] args)
        {
            new Program();
        }

        public Program()
        {
            Console.WriteLine("Starting Fast Excel Demo");
            Console.WriteLine(string.Format("Demos use {0} rows", NumberOfRecords));
            
            FileInfo outputFile = new FileInfo(Path.Combine(DemoDir,"outputfile.xlsx"));
            FileInfo epplusOutputFile = null;
            FileInfo templateFile = new FileInfo("Template.xlsx");

            if (outputFile.Exists)
            {
                outputFile.Delete();
                outputFile = new FileInfo(Path.Combine(DemoDir,"outputfile.xlsx"));
            }

            if (EPPlusTest)
            {
                epplusOutputFile = new FileInfo(Path.Combine(DemoDir, "epplusOutputfile.xlsx"));

                if (epplusOutputFile.Exists)
                {
                    epplusOutputFile.Delete();
                    epplusOutputFile = new FileInfo(Path.Combine(DemoDir, "epplusOutputfile.xlsx"));
                }
            }
            FastExcelWriteDemo(templateFile, outputFile);
            outputFile.Refresh();
            FastExcelWriteGenericsDemo(outputFile);
            FastExcelWriteAddRowDemo(outputFile);
            FastExcelWrite2DimensionArrayDemo(outputFile);

            FastExcelMergeDemo(outputFile);
              
            FastExcelReadDemo(outputFile);
            FastExcelReadDemo2(outputFile);

          //  FastExcelAddWorksheet(outputFile);

          //  FastExcelDeleteWorkSheet(outputFile);

            if (EPPlusTest)
            {
                EPPlusDemo(templateFile, epplusOutputFile);
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        /*private void FastExcelAddWorksheet(FileInfo outputFile)
        {

            Console.WriteLine();
            Console.WriteLine("DEMO ADD");

            Stopwatch stopwatch = new Stopwatch();

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(outputFile))
            {
                Worksheet worksheet = new Worksheet();
                worksheet.Name = "Sheet77";

                List<GenericObject> objectList = new List<GenericObject>();

                for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber++)
                {
                    GenericObject genericObject = new GenericObject();
                    genericObject.IntegerColumn1 = 1 * DateTime.Now.Millisecond;
                    genericObject.IntegerColumn2 = 2 * DateTime.Now.Millisecond;
                    genericObject.IntegerColumn3 = 3 * DateTime.Now.Millisecond;
                    genericObject.IntegerColumn4 = 4 * DateTime.Now.Millisecond;
                    genericObject.IntegerColumn5 = 45678854;
                    genericObject.DoubleColumn6 = 87.01d;
                    genericObject.StringColumn7 = "Test 3" + rowNumber;
                    genericObject.ObjectColumn8 = DateTime.Now.ToLongTimeString();

                    objectList.Add(genericObject);
                }
                worksheet.PopulateRows(objectList);

                stopwatch.Start();
                Console.WriteLine("Writing using IEnumerable<MyObject>...");
                fastExcel.Add(worksheet, "sheet3");
            }

            Console.WriteLine(string.Format("Writing IEnumerable<MyObject> took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }*/

        #region Write Demos
        private void FastExcelWriteDemo(FileInfo templateFile, FileInfo outputFile)
        {
            Console.WriteLine();
            Console.WriteLine("DEMO WRITE 1");

            Stopwatch stopwatch = new Stopwatch();

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(templateFile, outputFile))
            {
                Worksheet worksheet = new Worksheet();
                List<Row> rows = new List<Row>();
                
                for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber++)
                {
                    List<Cell> cells = new List<Cell>();
                    cells.Add(new Cell(1, 1 * DateTime.Now.Millisecond));
                    cells.Add(new Cell(2, 2 * DateTime.Now.Millisecond));
                    cells.Add(new Cell(3, 3 * DateTime.Now.Millisecond));
                    cells.Add(new Cell(4, 4 * DateTime.Now.Millisecond));
                    cells.Add(new Cell(5, 45678854));
                    cells.Add(new Cell(6, 87.01d));
                    cells.Add(new Cell(7, "Test 1 " + rowNumber));
                    cells.Add(new Cell(8, DateTime.Now.ToLongTimeString()));

                    rows.Add(new Row(rowNumber, cells));
                }
                worksheet.Rows = rows;

                stopwatch.Start();
                Console.WriteLine("Writing data...");
                fastExcel.Write(worksheet, "sheet1");
            }

            Console.WriteLine(string.Format("Writing data took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private void FastExcelWriteGenericsDemo(FileInfo outputFile)
        {
            Console.WriteLine();
            Console.WriteLine("DEMO WRITE 2");

            Stopwatch stopwatch = new Stopwatch();

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(outputFile))
            {
                List<GenericObject> objectList = new List<GenericObject>();

                for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber++)
                {
                    GenericObject genericObject = new GenericObject();
                    genericObject.IntegerColumn1 = 1 * DateTime.Now.Millisecond;
                    genericObject.IntegerColumn2 = 2 * DateTime.Now.Millisecond;
                    genericObject.IntegerColumn3 = 3 * DateTime.Now.Millisecond;
                    genericObject.IntegerColumn4 = 4 * DateTime.Now.Millisecond;
                    genericObject.IntegerColumn5 = 45678854;
                    genericObject.DoubleColumn6 = 87.01d;
                    genericObject.StringColumn7 = "Test 3" + rowNumber;
                    genericObject.ObjectColumn8 = DateTime.Now.ToLongTimeString();

                    objectList.Add(genericObject);
                }
                stopwatch.Start();
                Console.WriteLine("Writing using IEnumerable<MyObject>...");
                fastExcel.Write(objectList, "sheet3", true);
            }

            Console.WriteLine(string.Format("Writing IEnumerable<MyObject> took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private void FastExcelWriteAddRowDemo(FileInfo outputFile)
        {
            Console.WriteLine();
            Console.WriteLine("DEMO WRITE 3");

            Stopwatch stopwatch = new Stopwatch();

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(outputFile))
            {
                Worksheet worksheet = new Worksheet();

                for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber++)
                {
                    worksheet.AddRow(1 * DateTime.Now.Millisecond
                                , 2 * DateTime.Now.Millisecond
                                , 3 * DateTime.Now.Millisecond
                                , 4 * DateTime.Now.Millisecond
                                , 45678854
                                , 87.01d
                                , "Test 2" + rowNumber
                                , DateTime.Now.ToLongTimeString());
                }
                stopwatch.Start();
                Console.WriteLine("Writing using AddRow(params object[])...");
                fastExcel.Write(worksheet, "sheet4");
            }

            Console.WriteLine(string.Format("Writing using AddRow(params object[]) took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private void FastExcelWrite2DimensionArrayDemo(FileInfo outputFile)
        {
            Console.WriteLine();
            Console.WriteLine("DEMO WRITE 4");

            Stopwatch stopwatch = new Stopwatch();

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(outputFile))
            {
                List<object[]> rowData = new List<object[]>();

                // Note rowNumber starts at 2 because we are using existing headings on the sheet
                for (int rowNumber = 2; rowNumber < NumberOfRecords; rowNumber++)
                {
                    rowData.Add( new object[]{ 1 * DateTime.Now.Millisecond
                                , 2 * DateTime.Now.Millisecond
                                , 3 * DateTime.Now.Millisecond
                                , 4 * DateTime.Now.Millisecond
                                , 45678854
                                , 87.01d
                                , "Test 2" + rowNumber
                                , DateTime.Now.ToLongTimeString()});
                }
                stopwatch.Start();
                Console.WriteLine("Writing using IEnumerable<IEnumerable<object>>...");

                // Note existingHeadingRows = 1, because we are keeping existing headings on the sheet
                fastExcel.Write(rowData, "sheet2", 1);
            }

            Console.WriteLine(string.Format("Writing using IEnumerable<IEnumerable<object>> took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }
        #endregion

        #region Update Demos
        private void FastExcelMergeDemo(FileInfo inputFile)
        {
            Console.WriteLine();
            Console.WriteLine("DEMO UPDATE 1");

            Stopwatch stopwatch = new Stopwatch();
            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile))
            {
                Worksheet worksheet = new Worksheet();
                List<Row> rows = new List<Row>();

                for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber += 50)
                {
                    List<Cell> cells = new List<Cell>();
                    for (int columnNumber = 1; columnNumber < 12; columnNumber += 2)
                    {
                        cells.Add(new Cell(columnNumber, rowNumber));
                    }
                    cells.Add(new Cell(13, "Updated Row"));

                    rows.Add(new Row(rowNumber, cells));
                }
                worksheet.Rows = rows;

                stopwatch.Start();
                Console.WriteLine("Updating data every 50th row...");
                fastExcel.Update(worksheet, "sheet1");
            }

            Console.WriteLine(string.Format("Updating data took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }
        #endregion

        #region Reading Demos
        private void FastExcelReadDemo(FileInfo inputFile)
        {
            Console.WriteLine();
            Console.WriteLine("DEMO READ 1");

            Stopwatch stopwatch = Stopwatch.StartNew();

            // Open excel file using read only is much faster, but you cannot perfrom any writes
            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
            {
                Console.WriteLine("Reading data (Read Only Access) still needs enumerating...");
                Worksheet worksheet = fastExcel.Read("sheet1", 1);
            }
            
            Console.WriteLine(string.Format("Reading data took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private void FastExcelReadDemo2(FileInfo inputFile)
        {
            Console.WriteLine();
            Console.WriteLine("DEMO READ 2");

            Stopwatch stopwatch = Stopwatch.StartNew();

            // Open excel file using read/write is slower, but you can also perform writes
            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, false))
            {
                Console.WriteLine("Reading data (Read/Write Access) still needs enumerating...");
                Worksheet worksheet = fastExcel.Read("sheet1", 1);
            }

            Console.WriteLine(string.Format("Reading data took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }
        #endregion

        private class GenericObject
        {
            public int IntegerColumn1 { get; set; }
            public int IntegerColumn2 { get; set; }
            public int IntegerColumn3 { get; set; }
            public int IntegerColumn4 { get; set; }
            public int IntegerColumn5 { get; set; }
            public double DoubleColumn6 { get; set; }
            public string StringColumn7 { get; set; }
            public string ObjectColumn8 { get; set; }
        }

        /*private void FastExcelDeleteWorkSheet(FileInfo outputFile)
        {
            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(outputFile))
            {
                fastExcel.Delete(1);
            }
        }*/

        private void EPPlusDemo(FileInfo templateFile, FileInfo epplusOutputFile)
        {
            Console.WriteLine();
            Console.WriteLine("EPPlus Comparison DEMO");
            Console.WriteLine();
            Console.WriteLine("Preparing EPPlus data");
            using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(epplusOutputFile, templateFile))
            {
                var sheet = package.Workbook.Worksheets["sheet1"];

                List<object[]> epplusData = new List<object[]>();
                for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber++)
                {
                    var list = new List<object>();

                    list.Add(1 * DateTime.Now.Millisecond);
                    list.Add(2 * DateTime.Now.Millisecond);
                    list.Add(3 * DateTime.Now.Millisecond);
                    list.Add(4 * DateTime.Now.Millisecond);
                    list.Add(45678854);
                    list.Add(87.01d);
                    list.Add("EPPlus 1" + rowNumber);
                    list.Add(DateTime.Now.ToLongTimeString());
                    epplusData.Add(list.ToArray());
                }
                Stopwatch epplusStopwatch = Stopwatch.StartNew();
                Console.WriteLine("Adding rows to EPPlus Worksheet...");

                sheet.Cells["A1"].LoadFromArrays(epplusData);
                package.Save();

                Console.WriteLine(string.Format("Saving data took {0} seconds", epplusStopwatch.Elapsed.TotalSeconds));
            }
        }
    }
}