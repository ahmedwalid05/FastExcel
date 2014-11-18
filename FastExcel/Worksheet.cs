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
        private FastExcel FastExcel { get; set; }
        
        private SharedStrings SharedStrings { get; set; }
        public int ExistingHeadingRows { get; set; }

        internal int Index { get; set; }
        internal int Number { get; set; }
        public string Name { get; set; }
        internal string FileName { get; set; }

        public Worksheet(FastExcel fastExcel, SharedStrings sharedStrings, int sheetNumber) : this(fastExcel, sharedStrings, sheetNumber, null){}

        public Worksheet(FastExcel fastExcel, SharedStrings sharedStrings, string sheetName) : this(fastExcel, sharedStrings, null, sheetName){}

        private Worksheet(FastExcel fastExcel, SharedStrings sharedStrings, int? sheetNumber = null, string sheetName = null)
        {
            this.FastExcel = fastExcel;
            this.SharedStrings = sharedStrings;
            Tuple<int, int, string, string> worksheetProperties = this.FastExcel.GetWorksheetName(sheetNumber, sheetName);
            this.Index = worksheetProperties.Item1;
            this.Number = worksheetProperties.Item2;
            this.Name = worksheetProperties.Item3;
            this.FileName = worksheetProperties.Item4;
            this.ExistingHeadingRows = 0;
        }

        public bool Exists
        {
            get
            {
                return !string.IsNullOrEmpty(this.FileName);
            }
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
            if (this.FastExcel.Archive.Mode != ZipArchiveMode.Update)
            {
                throw new Exception("FastExcel is in ReadOnly mode so cannot perform a write");
            }

            // Check if ExistingHeadingRows will be overridden by the dataset
            if (this.ExistingHeadingRows != 0 && data.Rows.Where(r => r.RowNumber <= this.ExistingHeadingRows).Any())
            {
                throw new Exception("Existing Heading Rows was specified but some or all will be overridden by data rows. Check DataSet.Row.RowNumber against ExistingHeadingRows");
            }

            using (Stream stream = this.FastExcel.Archive.GetEntry(this.FileName).Open())
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
            IEnumerable<Row> rows = null;

            List<string> headings = new List<string>();
            using (Stream stream = this.FastExcel.Archive.GetEntry(this.FileName).Open())
            {
                XDocument document = XDocument.Load(stream);
                int skipRows = 0;

                Row possibleHeadingRow = new Row(document.Descendants().Where(d => d.Name.LocalName == "row").FirstOrDefault(), this.SharedStrings);
                if (this.ExistingHeadingRows == 1 && possibleHeadingRow.RowNumber == 1)
                {
                    foreach (Cell headerCell in possibleHeadingRow.Cells)
                    {
                        headings.Add(headerCell.Value.ToString());
                    }
                }
                rows = GetRows(document.Descendants().Where(d => d.Name.LocalName == "row").Skip(skipRows));
            }

            dataSet.Headings = headings;

            dataSet.Rows = rows;

            return dataSet;
        }

        private IEnumerable<Row> GetRows(IEnumerable<XElement> rowElements)
        {
            foreach (var rowElement in rowElements)
            {
                yield return new Row(rowElement, this.SharedStrings);
            }
        }

        internal void Update(DataSet data)
        {
            DataSet currentData = this.Read();
            this.SharedStrings.ReadWriteMode = true;
            currentData.Merge(data);
            this.Write(currentData);
            this.SharedStrings.ReadWriteMode = false;
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
