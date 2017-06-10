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
            Update(data, sheetNumber, null);
        }

        /// <summary>
        /// Update the worksheet
        /// </summary>
        /// <param name="data">The worksheet</param>
        /// <param name="sheetName">eg. Sheet1, Sheet2</param>
        public void Update(Worksheet data, string sheetName)
        {
            Update(data, null, sheetName);
        }

        private void Update(Worksheet data, int? sheetNumber = null, string sheetName = null)
        {
            CheckFiles();
            PrepareArchive();

            Worksheet currentData = Read(sheetNumber, sheetName);
            currentData.Merge(data);
            Write(currentData);
        }
    }
}
