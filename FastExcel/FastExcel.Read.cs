using System.Linq;

namespace FastExcel
{
    /// <summary>
    /// Fast Excel
    /// </summary>
    public partial class FastExcel
    {
        /// <summary>
        /// Read a sheet by sheet number
        /// </summary>
        public Worksheet Read(int sheetNumber, int existingHeadingRows = 0)
        {
            return Read(sheetNumber, null, existingHeadingRows);
        }

        /// <summary>
        /// Read a sheet by sheet name
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="existingHeadingRows"></param>
        /// <returns></returns>
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