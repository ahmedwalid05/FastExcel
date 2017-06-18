using System;
using System.Collections.Generic;
using System.Text;

namespace FastExcel
{
    /// <summary>
    /// Extra properties for a worksheet
    /// </summary>
    public class WorksheetProperties
    {
        /// <summary>
        /// Sheet index
        /// </summary>
        public int CurrentIndex { get; set; }
        /// <summary>
        /// Sheet Id
        /// </summary>
        public int SheetId { get; set; }
        /// <summary>
        /// Sheet name
        /// </summary>
        public string Name { get; set; }
    }
}
