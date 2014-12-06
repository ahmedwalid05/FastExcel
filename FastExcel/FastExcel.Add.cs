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
        /// Append new worksheet
        /// </summary>
        /// <param name="worksheet">New worksheet</param>
        public void Add(Worksheet worksheet)
        {
            this.Add(worksheet, null, null);
        }

        public void Add(Worksheet worksheet, int insertAfterSheetNumber)
        {
            this.Add(worksheet, insertAfterSheetNumber, null);
        }

        public void Add(Worksheet worksheet, string insertAfterSheetName)
        {
            this.Add(worksheet, null, insertAfterSheetName);
        }

        private void Add(Worksheet worksheet, int? insertAfterSheetNumber = null, string insertAfterSheetName = null)
        {
            CheckFiles();

            PrepareArchive(true);

            worksheet.ValidateNewWorksheet(this, insertAfterSheetNumber, insertAfterSheetName);

            if (this.AddWorksheets == null)
            {
                this.AddWorksheets = new List<WorksheetAddSettings>();
            }

            this.AddWorksheets.Add(worksheet.AddSettings);


            if (!this.ReadOnly)
            {
                throw new Exception("FastExcel is in ReadOnly mode so cannot perform a write");
            }

            // Check if ExistingHeadingRows will be overridden by the dataset
            if (worksheet.ExistingHeadingRows != 0 && worksheet.Rows.Where(r => r.RowNumber <= worksheet.ExistingHeadingRows).Any())
            {
                throw new Exception("Existing Heading Rows was specified but some or all will be overridden by data rows. Check DataSet.Row.RowNumber against ExistingHeadingRows");
            }

            using (StreamWriter streamWriter = null)//new StreamWriter(this.Archive.CreateEntry(worksheet.FileName).Open()))
            {
                streamWriter.Write(worksheet.Headers);
                if (!worksheet.Template)
                {
                    worksheet.Headers = null;
                }

                this.SharedStrings.ReadWriteMode = true;

                // Add Rows
                foreach (Row row in worksheet.Rows)
                {
                    streamWriter.Write(row.ToXmlString(this.SharedStrings));
                }
                this.SharedStrings.ReadWriteMode = false;

                //Add Footers
                streamWriter.Write(worksheet.Footers);
                if (!worksheet.Template)
                {
                    worksheet.Footers = null;
                }
                streamWriter.Flush();
            }
        }
    }
}
