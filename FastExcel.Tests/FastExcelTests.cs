using System;
using System.IO;
using Xunit;

namespace FastExcel.Tests
{
    public class FastExcelTests
    {
        private static readonly string ResourcesPath = Path.Combine(Environment.CurrentDirectory, "ResourcesTests");

        [Fact]
        public void FileNotExist_NewFastExcelWithInvalidInputFile_ThrowsFileNotFoundException()
        {
            var filePath = Path.Combine(Environment.CurrentDirectory, "test_not_exist.xlsx");
            var inputFile = new FileInfo(filePath);

            var action = new Action(() =>
            {
                using (var fastExcel = new FastExcel(inputFile));
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
                using (var fastExcel = new FastExcel(templateFile, outputFile)) ;
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
                using (var fastExcel = new FastExcel(templateFile, outputFile)) ;
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
                using (var fastExcel = new FastExcel(inputFile, true))
                {
                    var worksheet = fastExcel.Read(1, 1);
                }
            });

            var exception = Record.Exception(action);
            Assert.Null(exception);
        }

        [Fact]
        public void ThrowsErrorIfInitializedWithStreamAndFileInfoIsAccessed()
        {
            using var inputMemorystream = new MemoryStream(new byte[] {0x1});
            using var outputMemorystream = new MemoryStream();
            var fastExcel = new FastExcel(inputMemorystream, outputMemorystream);
            var exception = Assert.Throws<ApplicationException>(() => fastExcel.ExcelFile);
            Assert.Equal($"ExcelFile was not provided", exception.Message);
            exception = Assert.Throws<ApplicationException>(() => fastExcel.TemplateFile);
            Assert.Equal($"TemplateFile was not provided", exception.Message);
        }
    }
}