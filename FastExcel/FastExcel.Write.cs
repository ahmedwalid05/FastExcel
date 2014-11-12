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
        /// <summary>
        /// Write data to a sheet
        /// </summary>
        /// <param name="data">A dataset</param>
        /// <param name="sheetNumber">The number of the sheet starting at 1</param>
        /// <param name="existingHeadingRows">How many rows in the template sheet you would like to keep</param>
        public void Write(DataSet data, int sheetNumber, int existingHeadingRows = 0)
        {
            this.Write(data, sheetNumber, null, existingHeadingRows);
        }

        /// <summary>
        /// Write data to a sheet
        /// </summary>
        /// <param name="data">A dataset</param>
        /// <param name="sheetName">The display name of the sheet</param>
        /// <param name="existingHeadingRows">How many rows in the template sheet you would like to keep</param>
        public void Write(DataSet data, string sheetName, int existingHeadingRows = 0)
        {
            this.Write(data, null, sheetName, existingHeadingRows);
        }

        private void Write(DataSet data, int? sheetNumber = null, string sheetName = null, int existingHeadingRows = 0)
        {
            CheckFiles();

            try
            {
                if (!this.UpdateExisting)
                {
                    File.Copy(this.TemplateFile.FullName, this.ExcelFile.FullName);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Could not copy template to output file path", ex);
            }

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
            worksheet.Write(data);
        }
    }
}
