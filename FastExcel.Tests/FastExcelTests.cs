using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Xunit;
using Xunit.Abstractions;

namespace FastExcel.Tests
{
    public class FastExcelTests
    {
        private static readonly string ResourcesPath = Path.Combine(Environment.CurrentDirectory, "ResourcesTests");
        private static readonly string TemplateFilePath = Path.Combine(ResourcesPath, "template.xlsx");

        private static readonly CellRow TestCellRow = new()
        {
            StringColumn1 = "&",
            IntegerColumn2 = 45678854,
            DoubleColumn3 = 87.01d,
            ObjectColumn4 = DateTime.Now
        };

        private readonly ITestOutputHelper output;

        public FastExcelTests(ITestOutputHelper output)
        {
            this.output = output;
        }

        [Fact]
        public void FileNotExist_NewFastExcelWithInvalidInputFile_ThrowsFileNotFoundException()
        {
            var filePath = Path.Combine(Environment.CurrentDirectory, "test_not_exist.xlsx");
            var inputFile = new FileInfo(filePath);

            var action = new Action(() =>
            {
                using FastExcel fastExcel = new(inputFile);
            });

            var exception = Assert.Throws<FileNotFoundException>(action);
            Assert.Equal($"Input file '{filePath}' does not exist", exception.Message);
        }

        [Fact]
        public void FileNotExist_NewFastExcelWithInvalidTemplateFile_ThrowsFileNotFoundException()
        {
            var templateFilePath = Path.Combine(Environment.CurrentDirectory, "templateFilePath_not_exist.xlsx");
            var templateFile = new FileInfo(templateFilePath);

            var outputFilePath = Path.Combine(Environment.CurrentDirectory, "outputFilePath_not_exist.xlsx");
            var outputFile = new FileInfo(outputFilePath);

            var action = new Action(() =>
            {
                using FastExcel fastExcel = new(templateFile, outputFile);
            });

            var exception = Assert.Throws<FileNotFoundException>(action);
            Assert.Equal($"Template file '{templateFilePath}' was not found", exception.Message);
        }

        [Fact]
        public void FilesExist_NewFastExcelWithExistOutputFile_ThrowsFileNotFoundException()
        {
            var templateFilePath = Path.Combine(ResourcesPath, "RouteMaster.xlsx");
            var templateFile = new FileInfo(templateFilePath);

            var outputFilePath = Path.Combine(ResourcesPath, "RouteMaster.xlsx");
            var outputFile = new FileInfo(outputFilePath);

            var action = new Action(() =>
            {
                using FastExcel fastExcel = new(templateFile, outputFile);
            });

            var exception = Assert.Throws<Exception>(action);
            Assert.Equal($"Output file '{outputFilePath}' already exists", exception.Message);
        }

        [Fact]
        public void InputFile_ReadExcelWithNullReference_ExceptionIsNull()
        {
            var inputFilePath = Path.Combine(ResourcesPath, "RouteMaster.xlsx");
            var inputFile = new FileInfo(inputFilePath);

            var action = new Action(() =>
            {
                using FastExcel fastExcel = new(inputFile, true);
                var worksheet = fastExcel.Read(1, 1);
            });

            var exception = Record.Exception(action);
            Assert.Null(exception);
        }

        [Fact]
        public void ThrowsErrorIfInitializedWithStreamAndFileInfoIsAccessed()
        {
            using var inputMemorystream = new MemoryStream(new byte[] { 0x1 });
            using var outputMemorystream = new MemoryStream();

            var fastExcel = new FastExcel(inputMemorystream, outputMemorystream);
            var exception = Assert.Throws<ApplicationException>(() => fastExcel.ExcelFile);

            Assert.Equal("ExcelFile was not provided", exception.Message);
            exception = Assert.Throws<ApplicationException>(() => fastExcel.TemplateFile);

            Assert.Equal("TemplateFile was not provided", exception.Message);
        }

        private string FileRead_ReadingSpecialCharactersCore_Read(FileInfo inputFile)
        {
            inputFile.Refresh();
            using var fastExcel = new FastExcel(inputFile);

            var worksheet = fastExcel.Read("sheet1");
            var rows = worksheet.Rows;

            foreach (var item in rows)
            {
                foreach (var cell in item.Cells)
                {
                    output.WriteLine(cell.ToString());
                }
            }

            var row = rows.ToArray()[1].Cells.ToArray();
            Assert.Equal(TestCellRow.StringColumn1, row[0].Value);
            //TODO - Add tests for data-types when implemented 

            return "Passed";
        }

        [Fact]
        public string FileRead_ReadingSpecialCharacters_Read()
        {
            var inputFilePath = new FileInfo(Path.Combine(ResourcesPath, "special-char.xlsx"));
            return FileRead_ReadingSpecialCharactersCore_Read(inputFilePath);
        }

        [Fact]
        public string FileWrite_WritingOneRow_Wrote()
        {
            var inputFilePath = new FileInfo(Path.Combine(ResourcesPath, "temp.xlsx"));
            if (inputFilePath.Exists)
                inputFilePath.Delete();
            inputFilePath.Refresh();
            var templateFilePath = new FileInfo(TemplateFilePath);
            using (var fastExcel = new FastExcel(templateFilePath, inputFilePath))
            {
                List<CellRow> objectList = new();

                objectList.Add(TestCellRow);
                fastExcel.Write(objectList, "sheet1", true);
            }

            return FileRead_ReadingSpecialCharactersCore_Read(inputFilePath);
        }

        [Fact]
        public string FileUpdate_UpdatingEmptyFile_Updated()
        {
            var worksheet = new Worksheet();
            var cells = new List<CellRow>
            {
                TestCellRow
            };

            worksheet.PopulateRows(cells, usePropertiesAsHeadings: true);
            var templateFile = new FileInfo(TemplateFilePath);
            var inputFile = templateFile.CopyTo(Path.Combine(ResourcesPath, "temp1.xlsx"), true);

            using (var fastExcel = new FastExcel(inputFile))
            {
                // Read the data
                fastExcel.Update(worksheet, "Sheet1");
            }

            return FileRead_ReadingSpecialCharactersCore_Read(inputFile);
        }

        [Fact]
        public string FileUpdate_UpdatingWithOneRow_Updated()
        {
            var worksheet = new Worksheet();
            var cells = new List<CellRow> { TestCellRow };

            worksheet.PopulateRows(cells, usePropertiesAsHeadings: true);

            var templateFile = new FileInfo(Path.Combine(ResourcesPath, "OneRowFile.xlsx"));
            var inputFile = templateFile.CopyTo(Path.Combine(ResourcesPath, "temp.xlsx"), true);
            using (var fastExcel = new FastExcel(inputFile))
            {
                fastExcel.Update(worksheet, "Sheet1");
            }

            return FileRead_ReadingSpecialCharactersCore_Read(inputFile);
        }

        [Fact]
        public string FileUpdate_WriteAndUpdatingWithOneRow_Updated()
        {
            var worksheet = new Worksheet();
            var cells = new List<CellRow>
            {
                TestCellRow
            };

            worksheet.PopulateRows(cells, usePropertiesAsHeadings: true);

            var templateFile = new FileInfo(TemplateFilePath);
            var inputFile = new FileInfo(Path.Combine(ResourcesPath, "temp.xlsx"));
            if (inputFile.Exists)
            {
                inputFile.Delete();
                inputFile.Refresh();
            }

            //Writing Data One Row
            using (var fastExcel = new FastExcel(templateFile, inputFile))
            {
                fastExcel.Write(worksheet, "Sheet1");
            }

            inputFile.Refresh();
            using (var fastExcel = new FastExcel(inputFile))
            {
                fastExcel.Update(worksheet, "Sheet1");
            }

            return FileRead_ReadingSpecialCharactersCore_Read(inputFile);
        }

        [Fact]
        public string FileRead_ReadSameStringKey_Read()
        {
            var inputFile = new FileInfo(Path.Combine(ResourcesPath, "SameKey.xlsx"));

            using var fastExcel = new FastExcel(inputFile, true);
            var sheet = fastExcel.Read(1);
            return sheet.Name;
        }
    }

    public class CellRow
    {
        public string StringColumn1 { get; set; }
        public int IntegerColumn2 { get; set; }
        public double DoubleColumn3 { get; set; }
        public DateTime ObjectColumn4 { get; set; }
    }
}