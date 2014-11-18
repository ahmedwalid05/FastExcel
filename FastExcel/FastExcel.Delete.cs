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

            // Open worksheet
            Worksheet worksheet = null;
            if (sheetNumber.HasValue)
            {
                worksheet = new Worksheet(this, null, sheetNumber.Value);
            }
            else if (!string.IsNullOrEmpty(sheetName))
            {
                worksheet = new Worksheet(this, null, sheetName);
            }

            if (worksheet.Exists)
            {
                if (this.WorksheetReferenceUpdates == null)
                {
                    this.WorksheetReferenceUpdates = new Dictionary<int, bool>();
                }
                this.WorksheetReferenceUpdates.Add(worksheet.Index, false);
            }
        }
    }
}
