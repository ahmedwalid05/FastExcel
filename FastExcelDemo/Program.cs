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
            FastExcelReadDemo(outputFile);
            FastExcelMergeDemo(outputFile);
            FastExcelWriteGenericsDemo(outputFile);

            if (EPPlusTest)
            {
                EPPlusDemo(templateFile, epplusOutputFile);
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private void FastExcelWriteDemo(FileInfo templateFile, FileInfo outputFile)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(templateFile, outputFile))
            {
                Console.WriteLine(string.Format("Creating {0} rows in Data Set...", NumberOfRecords));
                DataSet data = new DataSet();
                List<Row> rows = new List<Row>();

                /*
                This method is very easy but a lot slower to populate
                for (int rowNumber = 1; rowNumber < 100000; rowNumber++)
                {
                    List<Cell> cells = new List<Cell>();
                    for (int columnNumber = 1; columnNumber < 13; columnNumber++)
                    {
                        data.AddValue(rowNumber, columnNumber, columnNumber * DateTime.Now.Second);
                    }
                }*/

                for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber++)
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

                Console.WriteLine(string.Format("Data Set Creation took {0} seconds", stopwatch.Elapsed.TotalSeconds));
                stopwatch = Stopwatch.StartNew();
                Console.WriteLine("Writing data...");
                fastExcel.Write(data, "sheet1");

                //Write to sheet 2 with headings
                //writer.Write(data, null, "sheet2", 1);
            }

            Console.WriteLine(string.Format("Writing data took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private void FastExcelReadDemo(FileInfo inputFile)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile))
            {
                Console.WriteLine("Reading data...");
                DataSet dataSet = fastExcel.Read("sheet1", 1);
            }

            Console.WriteLine(string.Format("Reading data took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private void FastExcelMergeDemo(FileInfo inputFile)
        {
            Stopwatch stopwatch = new Stopwatch();
            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile))
            {
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

                stopwatch.Start();
                Console.WriteLine("Updating data every 50th row...");
                fastExcel.Update(data, "sheet1");
            }

            Console.WriteLine(string.Format("Updating data took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private void FastExcelWriteGenericsDemo(FileInfo outputFile)
        {
            Stopwatch stopwatch = new Stopwatch();

            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(outputFile))
            {
                List<GenericObject> objectList = new List<GenericObject>();

                for (int rowNumber = 1; rowNumber < NumberOfRecords; rowNumber++)
                {
                    GenericObject genericObject = new GenericObject();
                    genericObject.StringColumn1 = "A string " + rowNumber.ToString();
                    genericObject.IntegerColumn2 = 45678854;
                    genericObject.DoubleColumn3 = 87.01d;
                    genericObject.ObjectColumn4 = DateTime.Now.ToLongTimeString();

                    objectList.Add(genericObject);
                }
                stopwatch.Start();
                Console.WriteLine("Writing Generic object list...");
                fastExcel.Write(objectList, "sheet3", true);
            }

            Console.WriteLine(string.Format("Writing Generic object list took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private class GenericObject
        {
            public string StringColumn1 { get; set; }
            public int IntegerColumn2 { get; set; }
            public double DoubleColumn3 { get; set; }
            public object ObjectColumn4 { get; set; }
        }

        private void EPPlusDemo(FileInfo templateFile, FileInfo epplusOutputFile)
        {
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
                    list.Add(5 * DateTime.Now.Millisecond);
                    list.Add(6 * DateTime.Now.Millisecond);
                    list.Add(7 * DateTime.Now.Millisecond);
                    list.Add(8 * DateTime.Now.Millisecond);
                    list.Add(9 * DateTime.Now.Millisecond);
                    list.Add(10 * DateTime.Now.Millisecond);
                    list.Add(11 * DateTime.Now.Millisecond);
                    list.Add(12 * DateTime.Now.Millisecond);
                    list.Add("Hello" + rowNumber);
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
