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
                    throw new Exception("Unable to load workbook.xml");
                }

                List<XElement> sheetsElements = document.Descendants().Where(d => d.Name.LocalName == "sheet").ToList();

                XElement sheetElement = null;

                if (sheetNumber.HasValue)
                {
                    if (sheetNumber.Value <= sheetsElements.Count)
                    {
                        sheetElement = sheetsElements[sheetNumber.Value - 1];
                    }
                    else
                    {
                        throw new Exception(string.Format("There is no sheet at index '{0}'", sheetNumber));
                    }
                }
                else if (!string.IsNullOrEmpty(sheetName))
                {
                    sheetElement = (from sheet in sheetsElements
                                    from attribute in sheet.Attributes()
                                    where attribute.Name == "name" && attribute.Value.Equals(sheetName, StringComparison.InvariantCultureIgnoreCase)
                                    select sheet).FirstOrDefault();

                    if (sheetElement == null)
                    {
                        throw new Exception(string.Format("There is no sheet named '{0}'", sheetName));
                    }
                }

                result = (from attribute in sheetElement.Attributes()
                            where attribute.Name == "sheetId"
                            select string.Format("xl/worksheets/sheet{0}.xml", attribute.Value)).FirstOrDefault();
                
            }

            if (string.IsNullOrEmpty(result))
            {
                throw new Exception("Unable to resolve internal sheet name");
            }

            return result;
        }

        /// <summary>
        /// Read the existing sheet and copy some of the existing content
        /// </summary>
        /// <param name="stream">Worksheet stream</param>
        /// <param name="headers">Content at top of document</param>
        /// <param name="footers">Content at bottom of document</param>
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

        internal void Write(DataSet data)
        {
            // Check if ExistingHeadingRows will be overridden by the dataset
            if (this.ExistingHeadingRows != 0 && data.Rows.Where(r => r.RowNumber <= this.ExistingHeadingRows).Any())
            {
                throw new Exception("Existing Heading Rows was specified but some or all will be overridden by data rows. Check DataSet.Row.RowNumber against ExistingHeadingRows");
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
                    streamWriter.Write(row.ToXmlString(this.SharedStrings));
                }

                //Add Footers
                streamWriter.Write(worksheetFooters);
                streamWriter.Flush();
            }
        }

        internal DataSet Read()
        {
            DataSet dataSet = new DataSet();
            List<Row> rows = new List<Row>();
            List<string> headings = new List<string>();

            using (Stream stream = this.Archive.GetEntry(this.FileName).Open())
            {
                XDocument document = XDocument.Load(stream);

                var rowElements = document.Descendants().Where(d => d.Name.LocalName == "row");

                foreach (var rowElement in rowElements)
                {
                    Row row = new Row(rowElement, this.SharedStrings);
                    if (this.ExistingHeadingRows == 1 && row.RowNumber == 1)
                    {
                        foreach (Cell headerCell in row.Cells)
                        {
                            headings.Add(headerCell.Value.ToString());
                        }
                    }
                    else if (row.Cells.Any())
                    {
                        rows.Add(row);
                    }
                }
            }

            dataSet.Headings = headings;
            dataSet.Rows = rows;

            return dataSet;
        }

        internal void Update(DataSet data)
        {
            DataSet currentData = this.Read();
            currentData.Merge(data);
            this.Write(currentData);
        }

        private List<string> GetHeadings(string lineBuffer, out bool headingsComplete, out bool rowsComplete, out string newLineBuffer)
        {
            List<string> headings = new List<string>();

            headingsComplete = false;
            rowsComplete = false;

            while (!string.IsNullOrEmpty(lineBuffer))
            {
                if (lineBuffer.Contains("<row"))
                {
                    if (lineBuffer.Contains("</row>"))
                    {
                        int index = lineBuffer.IndexOf("<row");
                        int currentLineIndex = lineBuffer.IndexOf("</row>") + "</row>".Length;
                        XElement rowElement = XElement.Parse(lineBuffer.Substring(index, currentLineIndex - index));
                        bool isFirstRow = (from a in rowElement.Attributes("r")
                                           where a.Value == "1"
                                           select a).Any();

                        if (rowElement.HasElements && isFirstRow)
                        {
                            foreach (XElement cell in rowElement.Elements())
                            {
                                bool isTextRow = (from a in cell.Attributes("t")
                                                  where a.Value == "s"
                                                  select a).Any();

                                if (isTextRow)
                                {
                                    headings.Add(this.SharedStrings.GetString(cell.Value));
                                }
                                else
                                {
                                    headings.Add(cell.Value);
                                }
                            }
                        }

                        //remove the read section from line
                        lineBuffer = lineBuffer.Substring(currentLineIndex, lineBuffer.Length - currentLineIndex);

                        headingsComplete = true;
                        break;
                    }
                    else
                    {
                        // Keep reading
                    }
                }
                else
                {
                    headingsComplete = true;
                    rowsComplete = true;
                    break;
                }
            }

            newLineBuffer = lineBuffer;
            return headings;
        }

    }
}
