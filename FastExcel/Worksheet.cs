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
    public class Worksheet
    {
        private ZipArchive Archive { get; set; }
        private string FileName { get; set; }
        private SharedStrings SharedStrings { get; set; }
        public int ExistingHeadingRows { get; set; }

        public Worksheet(ZipArchive archive, SharedStrings sharedStrings, int sheetNumber)
        {
            this.Archive = archive;
            this.SharedStrings = sharedStrings;
            this.FileName = GetWorksheetName(sheetNumber, null);
            this.ExistingHeadingRows = 0;
        }

        public Worksheet(ZipArchive archive, SharedStrings sharedStrings, string sheetName)
        {
            this.Archive = archive;
            this.SharedStrings = sharedStrings;
            this.FileName = GetWorksheetName(null, sheetName);
            this.ExistingHeadingRows = 0;
        }

        /// <summary>
        /// Get worksheet file name from xl/workbook.xml
        /// </summary>
        private string GetWorksheetName(int? sheetNumber = null, string sheetName = null)
        {
            string result = null;

            // TODO: May be able to speed up by only loading the sheets element
            using (Stream stream = this.Archive.GetEntry("xl/workbook.xml").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    //TODO error
                }

                List<XElement> sheetsElements = document.Descendants().Where(d => d.Name.LocalName == "sheet").ToList();

                XElement sheetElement = null;

                if (sheetNumber.HasValue)
                {
                    sheetElement = sheetsElements[sheetNumber.Value];
                }
                else if (!string.IsNullOrEmpty(sheetName))
                {
                    sheetElement = (from sheet in sheetsElements
                                    from attribute in sheet.Attributes()
                                    where attribute.Name == "name" && attribute.Value.Equals(sheetName, StringComparison.InvariantCultureIgnoreCase)
                                    select sheet).FirstOrDefault();
                }

                if (sheetElement != null)
                {
                    result = (from attribute in sheetElement.Attributes()
                              where attribute.Name == "sheetId"
                              select string.Format("xl/worksheets/sheet{0}.xml", attribute.Value)).FirstOrDefault();
                }
                else
                {
                    // TODO: raise error
                }
            }

            if (string.IsNullOrEmpty(result))
            {
                // TODO: raise error
            }

            return result;
        }

        private void ReadHeadersAndFooters(StreamReader stream, out StringBuilder headers, out StringBuilder footers)
        {
            headers = new StringBuilder();
            footers = new StringBuilder();
            
            bool headersComplete = false;
            bool rowsComplete = false;

            int existingHeadingRows = this.ExistingHeadingRows;

            while (stream.Peek() >= 0)
            {
                string line = stream.ReadLine();
                int currentLineIndex = 0;

                if (!headersComplete)
                {
                    if (line.Contains("<sheetData/>"))
                    {
                        currentLineIndex = line.IndexOf("<sheetData/>");
                        headers.Append(line.Substring(0, currentLineIndex));
                        //remove the read section from line
                        line = line.Substring(currentLineIndex, line.Length - currentLineIndex);

                        headers.Append("<sheetData>");

                        // Headers complete now skip any content and start footer
                        headersComplete = true;
                        footers = new StringBuilder();
                        footers.Append("</sheetData>");

                        //There is no rows
                        rowsComplete = true;
                    }
                    else if (line.Contains("<sheetData>"))
                    {
                        currentLineIndex = line.IndexOf("<sheetData>");
                        headers.Append(line.Substring(0, currentLineIndex));
                        //remove the read section from line
                        line = line.Substring(currentLineIndex, line.Length - currentLineIndex);

                        headers.Append("<sheetData>");

                        // Headers complete now skip any content and start footer
                        headersComplete = true;
                        footers = new StringBuilder();
                        footers.Append("</sheetData>");
                    }
                    else
                    {
                        headers.Append(line);
                    }
                }

                if (headersComplete && !rowsComplete)
                {
                    if (existingHeadingRows == 0)
                    {
                        rowsComplete = true;
                    }

                    if (!rowsComplete)
                    {
                        while (!string.IsNullOrEmpty(line) && existingHeadingRows != 0)
                        {
                            if (line.Contains("<row"))
                            {
                                if (line.Contains("</row>"))
                                {
                                    int index = line.IndexOf("<row");
                                    currentLineIndex = line.IndexOf("</row>") + "</row>".Length;
                                    headers.Append(line.Substring(index, currentLineIndex - index));

                                    //remove the read section from line
                                    line = line.Substring(currentLineIndex, line.Length - currentLineIndex);
                                    existingHeadingRows--;
                                }
                                else
                                {
                                    int index = line.IndexOf("<row");
                                    headers.Append(line.Substring(index, line.Length - index));
                                    line = string.Empty;
                                }
                            }
                            else if (line.Contains("</row>"))
                            {
                                currentLineIndex = line.IndexOf("</row>") + "</row>".Length;
                                headers.Append(line.Substring(0, currentLineIndex));

                                //remove the read section from line
                                line = line.Substring(currentLineIndex, line.Length - currentLineIndex);
                                existingHeadingRows--;
                            }
                        }
                    }

                    if (existingHeadingRows == 0)
                    {
                        rowsComplete = true;
                    }
                }

                if (rowsComplete)
                {
                    if (line.Contains("</sheetData>"))
                    {
                        int index = line.IndexOf("</sheetData>") + "</sheetData>".Length;
                        footers.Append(line.Substring(index, line.Length - index));
                    }
                    else if (line.Contains("<sheetData/>"))
                    {
                        int index = line.IndexOf("<sheetData/>") + "<sheetData/>".Length;
                        footers.Append(line.Substring(index, line.Length - index));
                    }
                    else
                    {
                        footers.Append(line);
                    }
                }
            }
        }

        public void Write(DataSet data)
        {
            // Check if ExistingHeadingRows will be overridden by the dataset
            if (this.ExistingHeadingRows != 0 && data.Rows.Where(r => r.RowNumber <= this.ExistingHeadingRows).Any())
            {
                //TODO complete error message
                throw new Exception();
            }

            using (Stream stream = this.Archive.GetEntry(this.FileName).Open())
            {
                StringBuilder worksheetHeaders = null;
                StringBuilder worksheetFooters = null;

                // Open worksheet and read the data at the top and bottom of the sheet
                StreamReader streamReader = new StreamReader(stream);
                ReadHeadersAndFooters(streamReader, out worksheetHeaders, out worksheetFooters);

                //Set the stream to the start
                stream.Position = 0;

                // Open the stream so we can override all content of the sheet
                StreamWriter streamWriter = new StreamWriter(stream);

                // TODO instead of saving the headers then writing them back get position where the headers finish then write from there
                streamWriter.Write(worksheetHeaders);

                // Add Rows
                foreach (Row row in data.Rows)
                {
                    streamWriter.Write(row.ToString(this.SharedStrings));
                }

                //Add Footers
                streamWriter.Write(worksheetFooters);
                streamWriter.Flush();
            }
        }
    }
}
