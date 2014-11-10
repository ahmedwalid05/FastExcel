using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace FastExcel
{
    public class FastExcelReader: IDisposable
    {
        public FileInfo InputFile { get; private set; }
        private SharedStrings SharedStrings { get; set; }
        private ZipArchive Archive { get; set; }

        public FastExcelReader(FileInfo inputFile)
        {
            this.InputFile = inputFile;

            CheckFile();
        }

        /// <summary>
        /// Ensure files are ready for use
        /// </summary>
        private void CheckFile()
        {
            if (this.InputFile == null)
            {
                throw new Exception("No input file name was supplied");
            }
            else if (!this.InputFile.Exists)
            {
                this.InputFile = null;
                throw new Exception(string.Format("Input file '{0}' does not exist", this.InputFile.FullName));
            }
        }

        public DataSet Read(int sheetNumber, int existingHeadingRows = 0)
        {
            return Read(sheetNumber, null, existingHeadingRows);
        }

        public DataSet Read(string sheetName, int existingHeadingRows = 0)
        {
            return Read(null, sheetName, existingHeadingRows);
        }

        private DataSet Read(int? sheetNumber = null, string sheetName = null, int existingHeadingRows = 0)
        {
            CheckFile();

            if (this.Archive == null)
            {
                Archive = ZipFile.Open(this.InputFile.FullName, ZipArchiveMode.Update);
            }
            
            // Get Strings file
            if (this.SharedStrings == null)
            {
                this.SharedStrings = new SharedStrings(this.Archive);
            }
                
            // Open worksheet
            Worksheet worksheet = null;
            if (sheetNumber.HasValue)
            {
                worksheet = new Worksheet(this.Archive, SharedStrings, sheetNumber.Value);
            }
            else if (!string.IsNullOrEmpty(sheetName))
            {
                worksheet = new Worksheet(this.Archive, SharedStrings, sheetName);
            }
            else
            {
                throw new Exception("No worksheet name or number was specified");
            }

            worksheet.ExistingHeadingRows = existingHeadingRows;

            //Write Data
            return worksheet.Read();
        }
        
        public void Dispose()
        {
            if (this.Archive == null)
            {
                return;
            }
            
            this.Archive.Dispose();
        }
    }
}
