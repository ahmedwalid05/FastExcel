using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace FastExcel
{
    public partial class FastExcel
    {
        public Worksheet Read(int sheetNumber, int existingHeadingRows = 0)
        {
            return Read(sheetNumber, null, existingHeadingRows);
        }

        public Worksheet Read(string sheetName, int existingHeadingRows = 0)
        {
            return Read(null, sheetName, existingHeadingRows);
        }

        private Worksheet Read(int? sheetNumber = null, string sheetName = null, int existingHeadingRows = 0)
        {
            Worksheet worksheet = null;
            if (_worksheets == null)
            {
                worksheet = new Worksheet(this);
                worksheet.Read(sheetNumber, sheetName, existingHeadingRows);
            }
            else
            {
                worksheet = (from w in Worksheets
                             where (sheetNumber.HasValue && sheetNumber.Value == w.Index) ||
                                    (sheetName == w.Name)
                             select w).SingleOrDefault();
                worksheet.Read(existingHeadingRows);
            }
            return worksheet;
        }
    }
}
