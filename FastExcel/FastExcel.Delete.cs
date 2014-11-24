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
        public void Delete(int sheetNumber)
        {
            this.Delete(sheetNumber, null);
        }

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
            worksheet.GetWorksheetProperties(this.Archive, sheetNumber, sheetName);

            if (this.DeleteWorksheets == null)
            {
                this.DeleteWorksheets = new List<int>();
            }
            this.DeleteWorksheets.Add(worksheet.Index);
        }
    }
}
