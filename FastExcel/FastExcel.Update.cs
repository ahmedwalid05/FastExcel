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
        /// Update the worksheet
        /// </summary>
        /// <param name="data">The worksheet</param>
        /// <param name="sheetNumber">eg 1,2,4</param>
        public void Update(Worksheet data, int sheetNumber)
        {
            this.Update(data, sheetNumber, null);
        }

        /// <summary>
        /// Update the worksheet
        /// </summary>
        /// <param name="data">The worksheet</param>
        /// <param name="sheetName">eg. Sheet1, Sheet2</param>
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
