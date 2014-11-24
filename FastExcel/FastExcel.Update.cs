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
        public void Update(Worksheet data, int sheetNumber)
        {
            this.Update(data, sheetNumber, null);
        }

        public void Update(Worksheet data, string sheetName)
        {
            this.Update(data, null, sheetName);
        }

        private void Update(Worksheet data, int? sheetNumber = null, string sheetName = null)
        {
            CheckFiles();
            PrepareArchive();

            Worksheet currentData = this.Read(sheetNumber, sheetName);
            currentData.Merge(data);
            this.Write(currentData);
        }
    }
}
