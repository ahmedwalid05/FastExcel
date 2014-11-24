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
            CheckFiles();
            PrepareArchive();

            Worksheet worksheet = new Worksheet();
            worksheet.ExistingHeadingRows = existingHeadingRows;
            worksheet.GetWorksheetProperties(this.Archive, sheetNumber, sheetName);

            IEnumerable<Row> rows = null;

            List<string> headings = new List<string>();
            using (Stream stream = this.Archive.GetEntry(worksheet.FileName).Open())
            {
                XDocument document = XDocument.Load(stream);
                int skipRows = 0;

                Row possibleHeadingRow = new Row(document.Descendants().Where(d => d.Name.LocalName == "row").FirstOrDefault(), this.SharedStrings);
                if (worksheet.ExistingHeadingRows == 1 && possibleHeadingRow.RowNumber == 1)
                {
                    foreach (Cell headerCell in possibleHeadingRow.Cells)
                    {
                        headings.Add(headerCell.Value.ToString());
                    }
                }
                rows = GetRows(document.Descendants().Where(d => d.Name.LocalName == "row").Skip(skipRows));
            }

            worksheet.Headings = headings;
            worksheet.Rows = rows;

            return worksheet;
        }

        private IEnumerable<Row> GetRows(IEnumerable<XElement> rowElements)
        {
            foreach (var rowElement in rowElements)
            {
                yield return new Row(rowElement, this.SharedStrings);
            }
        }

        /// <summary>
        /// Read the existing sheet and copy some of the existing content
        /// </summary>
        /// <param name="stream">Worksheet stream</param>
        /// <param name="worksheet">Saves the header and footer to the worksheet</param>
        private void ReadHeadersAndFooters(StreamReader stream, ref Worksheet worksheet)
        {
            StringBuilder headers = new StringBuilder();
            StringBuilder footers = new StringBuilder();

            bool headersComplete = false;
            bool rowsComplete = false;

            int existingHeadingRows = worksheet.ExistingHeadingRows;

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
            worksheet.Headers = headers.ToString();
            worksheet.Footers = footers.ToString();
        }
    }
}
