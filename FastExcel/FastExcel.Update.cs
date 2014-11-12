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
        public void Update(DataSet data, int sheetNumber)
        {
            this.Update(data, sheetNumber, null);
        }

        public void Update(DataSet data, string sheetName)
        {
            this.Update(data, null, sheetName);
        }

        private void Update(DataSet data, int? sheetNumber = null, string sheetName = null)
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

            //Update Data
            worksheet.Update(data);
        }
    }
}
