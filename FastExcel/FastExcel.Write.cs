using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;

namespace FastExcel
{
    public partial class FastExcel
    {
        /// <summary>
        /// Write data to a sheet
        /// </summary>
        /// <param name="worksheet">A dataset</param>
        public void Write(Worksheet worksheet)
        {
            Write(worksheet, null, null);
        }

        /// <summary>
        /// Write data to a sheet
        /// </summary>
        /// <param name="worksheet">A dataset</param>
        /// <param name="sheetNumber">The number of the sheet starting at 1</param>
        /// <param name="existingHeadingRows">How many rows in the template sheet you would like to keep</param>
        public void Write(Worksheet worksheet, int sheetNumber, int existingHeadingRows = 0)
        {
            Write(worksheet, sheetNumber, null, existingHeadingRows);
        }

        /// <summary>
        /// Write data to a sheet
        /// </summary>
        /// <param name="worksheet">A dataset</param>
        /// <param name="sheetName">The display name of the sheet</param>
        /// <param name="existingHeadingRows">How many rows in the template sheet you would like to keep</param>
        public void Write(Worksheet worksheet, string sheetName, int existingHeadingRows = 0)
        {
            Write(worksheet, null, sheetName, existingHeadingRows);
        }

        /// <summary>
        /// Write a list of objects to a sheet
        /// </summary>
        /// <typeparam name="T">Row Object</typeparam>
        /// <param name="rows">IEnumerable list of objects</param>
        /// <param name="sheetNumber">The number of the sheet starting at 1</param>
        /// <param name="existingHeadingRows">How many rows in the template sheet you would like to keep</param>
        public void Write<T>(IEnumerable<T> rows, int sheetNumber, int existingHeadingRows = 0)
        {
            Worksheet data = new Worksheet();
            data.PopulateRows<T>(rows);
            Write(data, sheetNumber, null, existingHeadingRows);
        }

        /// <summary>
        /// Write a list of objects to a sheet
        /// </summary>
        /// <typeparam name="T">Row Object</typeparam>
        /// <param name="rows">IEnumerable list of objects</param>
        /// <param name="sheetName">The display name of the sheet</param>
        /// <param name="existingHeadingRows">How many rows in the template sheet you would like to keep</param>
        public void Write<T>(IEnumerable<T> rows, string sheetName, int existingHeadingRows = 0)
        {
            Worksheet data = new Worksheet();
            data.PopulateRows<T>(rows, existingHeadingRows);
            Write(data, null, sheetName, existingHeadingRows);
        }

        /// <summary>
        /// Write a list of objects to a sheet
        /// </summary>
        /// <typeparam name="T">Row Object</typeparam>
        /// <param name="objectList">IEnumerable list of objects</param>
        /// <param name="sheetNumber">The number of the sheet starting at 1</param>
        /// <param name="usePropertiesAsHeadings">Use property names from object list as headings</param>
        public void Write<T>(IEnumerable<T> objectList, int sheetNumber, bool usePropertiesAsHeadings)
        {
            Worksheet data = new Worksheet();
            data.PopulateRows<T>(objectList, 0, usePropertiesAsHeadings);
            Write(data, sheetNumber, null, 0);
        }

        /// <summary>
        /// Write a list of objects to a sheet
        /// </summary>
        /// <typeparam name="T">Row Object</typeparam>
        /// <param name="rows">IEnumerable list of objects</param>
        /// <param name="sheetName">The display name of the sheet</param>
        /// <param name="usePropertiesAsHeadings">Use property names from object list as headings</param>
        public void Write<T>(IEnumerable<T> rows, string sheetName, bool usePropertiesAsHeadings)
        {
            Worksheet data = new Worksheet();
            data.PopulateRows<T>(rows, 0,usePropertiesAsHeadings);
            Write(data, null, sheetName, 0);
        }

        private void Write(Worksheet worksheet, int? sheetNumber = null, string sheetName = null, int existingHeadingRows = 0)
        {
            CheckFiles();

            try
            {
                if (!UpdateExisting)
                {
                    File.Copy(TemplateFile.FullName, ExcelFile.FullName);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Could not copy template to output file path", ex);
            }

            PrepareArchive();

            // Open worksheet
            worksheet.GetWorksheetProperties(this, sheetNumber, sheetName);
            worksheet.ExistingHeadingRows = existingHeadingRows;

            if (Archive.Mode != ZipArchiveMode.Update)
            {
                throw new Exception("FastExcel is in ReadOnly mode so cannot perform a write");
            }

            // Check if ExistingHeadingRows will be overridden by the dataset
            if (worksheet.ExistingHeadingRows != 0 && worksheet.Rows.Where(r => r.RowNumber <= worksheet.ExistingHeadingRows).Any())
            {
                throw new Exception("Existing Heading Rows was specified but some or all will be overridden by data rows. Check DataSet.Row.RowNumber against ExistingHeadingRows");
            }

            using (Stream stream = Archive.GetEntry(worksheet.FileName).Open())
            {
                // Open worksheet and read the data at the top and bottom of the sheet
                StreamReader streamReader = new StreamReader(stream);
                worksheet.ReadHeadersAndFooters(streamReader, ref worksheet);
                
                //Set the stream to the start
                stream.Position = 0;

                // Open the stream so we can override all content of the sheet
                StreamWriter streamWriter = new StreamWriter(stream);

                // TODO instead of saving the headers then writing them back get position where the headers finish then write from there
                streamWriter.Write(worksheet.Headers);
                if (!worksheet.Template)
                {
                    worksheet.Headers = null;
                }

                SharedStrings.ReadWriteMode = true;

                // Add Rows
                foreach (Row row in worksheet.Rows)
                {
                    streamWriter.Write(row.ToXmlString(SharedStrings));
                }
                SharedStrings.ReadWriteMode = false;

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
