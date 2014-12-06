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
        /// Deletes the selected sheet Note:delete happens on Dispose
        /// </summary>
        /// <param name="sheetNumber">sheet number, starts at 1</param>
        public void Delete(int sheetNumber)
        {
            this.Delete(sheetNumber, null);
        }

        /// <summary>
        /// Deletes the selected sheet Note:delete happens on Dispose
        /// </summary>
        /// <param name="sheetName">Worksheet name</param>
        public void Delete(string sheetName)
        {
            this.Update(null, sheetName);
        }

        private void Delete(int? sheetNumber = null, string sheetName = null)
        {
            CheckFiles();

            PrepareArchive(false);

            // Get worksheet details
            Worksheet worksheet = new Worksheet();
            worksheet.GetWorksheetProperties(this, sheetNumber, sheetName);

            // Delete the file
            if (!string.IsNullOrEmpty(worksheet.FileName))
            {
                ZipArchiveEntry entry = this.Archive.GetEntry(worksheet.FileName);
                if (entry != null)
                {
                    entry.Delete();
                }

                if (this.DeleteWorksheets == null)
                {
                    this.DeleteWorksheets = new List<int>();
                }
                this.DeleteWorksheets.Add(worksheet.Index);
            }
        }
    }
}
