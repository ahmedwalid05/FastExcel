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
    public partial class FastExcel
    {
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
            CheckFiles();

            if (this.Archive == null)
            {
                Archive = ZipFile.Open(this.ExcelFile.FullName, ZipArchiveMode.Update);
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
    }
}
