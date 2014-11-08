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
            FileInfo outputFile = new FileInfo(Path.Combine(DemoDir,"outpufile.xlsx"));
            FileInfo epplusOutputFile = null;
            FileInfo templateFile = new FileInfo("Template.xlsx");

            if (outputFile.Exists)
            {
                outputFile.Delete();
                outputFile = new FileInfo(Path.Combine(DemoDir,"outpufile.xlsx"));
            }

            if (EPPlusTest)
            {
                epplusOutputFile = new FileInfo(Path.Combine(DemoDir, "epplusOutpufile.xlsx"));

                if (epplusOutputFile.Exists)
                {
                    epplusOutputFile.Delete();
                    epplusOutputFile = new FileInfo(Path.Combine(DemoDir, "epplusOutpufile.xlsx"));
                }
            }

            FastExcelDemo(templateFile, outputFile);

            if (EPPlusTest)
            {
                EPPlusDemo(templateFile, epplusOutputFile);
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private void FastExcelDemo(FileInfo templateFile, FileInfo outputFile)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            using (FastExcel.FastExcelWriter writer = new FastExcel.FastExcelWriter(templateFile, outputFile))
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

                for (int rowNumber = 2; rowNumber < NumberOfRecords; rowNumber++)
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
                writer.Write(data, null, "sheet1", 0);

                //Write to sheet 2 with headings
                //writer.Write(data, null, "sheet2", 1);
            }

            Console.WriteLine(string.Format("Writing data took {0} seconds", stopwatch.Elapsed.TotalSeconds));
        }

        private void EPPlusDemo(FileInfo templateFile, FileInfo epplusOutputFile)
        {
            Console.WriteLine("Preparing EPPlus data");
            using (OfficeOpenXml.ExcelPackage package = new OfficeOpenXml.ExcelPackage(epplusOutputFile, templateFile))
            {
                var sheet3 = package.Workbook.Worksheets["sheet1"];

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

                sheet3.Cells["A1"].LoadFromArrays(epplusData);
                package.Save();

                Console.WriteLine(string.Format("Saving data took {0} seconds", epplusStopwatch.Elapsed.TotalSeconds));
            }
        }
    }
}
