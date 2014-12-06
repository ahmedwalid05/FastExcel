using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FastExcel
{
    internal class WorksheetAddSettings
    {
        public string Name { get; set; }

        public int SheetId { get; set; }

        public int InsertAfterSheetId { get; set; }
    }
}
