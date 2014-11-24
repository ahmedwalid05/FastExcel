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
    public partial class FastExcel: IDisposable
    {
        public FileInfo ExcelFile { get; private set; }
        public FileInfo TemplateFile { get; private set; }
        public bool ReadOnly { get; private set; }
        public bool CacheWorksheetReferences { get; set; }

        private SharedStrings SharedStrings { get; set; }
        internal ZipArchive Archive { get; set; }
        private bool UpdateExisting { get; set; }

        /// <summary>
        /// A list of worksheet indexs to delete
        /// </summary>
        private List<int> DeleteWorksheets { get; set; }

        /// <summary>
        /// Update an existing excel file
        /// </summary>
        /// <param name="excelFile">location of an existing excel file</param>
        public FastExcel(FileInfo excelFile, bool readOnly = false) : this(null, excelFile, true, readOnly) {}
        
        /// <summary>
        /// Create a new excel file from a template
        /// </summary>
        /// <param name="templateFile">template location</param>
        /// <param name="excelFile">location of where a new excel file will be saved to</param>
        public FastExcel(FileInfo templateFile, FileInfo excelFile) :this(templateFile, excelFile, false, false) {}

        private FastExcel(FileInfo templateFile, FileInfo excelFile, bool updateExisting, bool readOnly = false)
        {
            this.TemplateFile = templateFile;
            this.ExcelFile = excelFile;
            this.UpdateExisting = updateExisting;
            this.ReadOnly = readOnly;

            CheckFiles();
        }

        private void PrepareArchive(bool openSharedStrings = true)
        {
            if (this.Archive == null)
            {
                if (this.ReadOnly)
                {
                    Archive = ZipFile.Open(this.ExcelFile.FullName, ZipArchiveMode.Read);
                }
                else
                {
                    Archive = ZipFile.Open(this.ExcelFile.FullName, ZipArchiveMode.Update);
                }
            }

            // Get Strings file
            if (this.SharedStrings == null && openSharedStrings)
            {
                this.SharedStrings = new SharedStrings(this.Archive);
            }
        }

        /// <summary>
        /// Ensure files are ready for use
        /// </summary>
        private void CheckFiles()
        {
            if (this.UpdateExisting)
            {
                if (this.ExcelFile == null)
                {
                    throw new Exception("No input file name was supplied");
                }
                else if (!this.ExcelFile.Exists)
                {
                    this.ExcelFile = null;
                    throw new Exception(string.Format("Input file '{0}' does not exist", this.ExcelFile.FullName));
                }
            }
            else
            {
                if (this.TemplateFile == null)
                {
                    throw new Exception("No Template file was supplied");
                }
                else if (!this.TemplateFile.Exists)
                {
                    this.TemplateFile = null;
                    throw new FileNotFoundException(string.Format("Template file '{0}' was not found", this.TemplateFile.FullName));
                }

                if (this.ExcelFile == null)
                {
                    throw new Exception("No Ouput file name was supplied");
                }
                else if (this.ExcelFile.Exists)
                {
                    this.ExcelFile = null;
                    throw new Exception(string.Format("Output file '{0}' already exists", this.ExcelFile.FullName));
                }
            }
        }

        /// <summary>
        /// Update xl/_rels/workbook.xml.rels file
        /// </summary>
        private void UpdateRelations(bool ensureStrings)
        {
            if (!(ensureStrings || (this.DeleteWorksheets != null && this.DeleteWorksheets.Any())))
            {
                // Nothing to update
                return;
            }

            using (Stream stream = this.Archive.GetEntry("xl/_rels/workbook.xml.rels").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    //TODO error
                }
                bool update = false;

                List<XElement> relationshipElements = document.Descendants().Where(d => d.Name.LocalName == "Relationship").ToList();
                if (ensureStrings)
                {
                    //Ensure SharedStrings
                    XElement relationshipElement = (from element in relationshipElements
                                                    from attribute in element.Attributes()
                                                    where attribute.Name == "Target" && attribute.Value.Equals("sharedStrings.xml", StringComparison.InvariantCultureIgnoreCase)
                                                    select element).FirstOrDefault();

                    if (relationshipElement == null)
                    {
                        relationshipElement = new XElement(document.Root.GetDefaultNamespace() + "Relationship");
                        relationshipElement.Add(new XAttribute("Target", "sharedStrings.xml"));
                        relationshipElement.Add(new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"));
                        relationshipElement.Add(new XAttribute("Id", string.Format("rId{0}", relationshipElements.Count + 1)));

                        document.Root.Add(relationshipElement);
                        update = true;
                    }
                }
                if (this.DeleteWorksheets != null && this.DeleteWorksheets.Any())
                {
                    foreach (var item in this.DeleteWorksheets)
                    {
                        string fileName = string.Format("worksheets/sheet{0}.xml", item);

                        XElement relationshipElement = (from element in relationshipElements
                                                        from attribute in element.Attributes()
                                                        where attribute.Name == "Target" && attribute.Value == fileName
                                                        select element).FirstOrDefault();
                        if (relationshipElement != null)
                        {
                            relationshipElement.Remove();
                            update = true;
                        }
                    }
                }

                if (update)
                {
                    // Set the stream to the start
                    stream.Position = 0;
                    // Clear the stream
                    stream.SetLength(0);

                    // Open the stream so we can override all content of the sheet
                    StreamWriter streamWriter = new StreamWriter(stream, Encoding.UTF8);
                    document.Save(streamWriter);
                    streamWriter.Flush();
                }
            }
        }

        /// <summary>
        /// Update [Content_Types].xml file
        /// </summary>
        private void UpdateContentTypes(bool ensureStrings)
        {
            if (!(ensureStrings || (this.DeleteWorksheets != null && this.DeleteWorksheets.Any())))
            {
                // Nothing to update
                return;
            }

            using (Stream stream = this.Archive.GetEntry("[Content_Types].xml").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    //TODO error
                }
                bool update = false;
                List<XElement> overrideElements = document.Descendants().Where(d => d.Name.LocalName == "Override").ToList();

                //Ensure SharedStrings
                if (ensureStrings)
                {
                    XElement overrideElement = (from element in overrideElements
                                                from attribute in element.Attributes()
                                                where attribute.Name == "PartName" && attribute.Value.Equals("/xl/sharedStrings.xml", StringComparison.InvariantCultureIgnoreCase)
                                                select element).FirstOrDefault();

                    if (overrideElement == null)
                    {
                        overrideElement = new XElement(document.Root.GetDefaultNamespace() + "Override");
                        overrideElement.Add(new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"));
                        overrideElement.Add(new XAttribute("PartName", "/xl/sharedStrings.xml"));

                        document.Root.Add(overrideElement);
                        update = true;
                    }
                }
                if (this.DeleteWorksheets != null && this.DeleteWorksheets.Any())
                {
                    foreach (var item in this.DeleteWorksheets)
                    {
                        // TODO resuse filename saved on worksheet
                        string fileName = string.Format("/xl/worksheets/sheet{0}.xml", item);

                        XElement overrideElement = (from element in overrideElements
                                                    from attribute in element.Attributes()
                                                    where attribute.Name == "PartName" && attribute.Value == fileName
                                                    select element).FirstOrDefault();
                        if (overrideElement != null)
                        {
                            overrideElement.Remove();
                            update = true;
                        }
                    }
                }
                if (update)
                {
                    // Set the stream to the start
                    stream.Position = 0;
                    // Clear the stream
                    stream.SetLength(0);
                    
                    // Open the stream so we can override all content of the sheet
                    StreamWriter streamWriter = new StreamWriter(stream, Encoding.UTF8);
                    document.Save(streamWriter);
                    streamWriter.Flush();
                }
            }
        }
        
        /// <summary>
        /// Update xl/workbook.xml file
        /// </summary>
        private void UpdateWorkbook()
        {
            if (this.DeleteWorksheets == null || !this.DeleteWorksheets.Any())
            {
                // Nothing to update
                return;
            }

            using (Stream stream = this.Archive.GetEntry("xl/workbook.xml").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    throw new Exception("Unable to load workbook.xml");
                }
                bool update = false;

                foreach (var item in this.DeleteWorksheets)
                {
                    XElement sheetElement = (from sheet in document.Descendants()
                                             where sheet.Name.LocalName == "sheet"
                                             from attribute in sheet.Attributes()
                                             where attribute.Name == "sheetId" && attribute.Value == item.ToString()
                                             select sheet).FirstOrDefault();
                    if (sheetElement != null)
                    {
                        sheetElement.Remove();
                        update = true;
                    }
                }

                if (update)
                {
                    //Set the stream to the start
                    stream.Position = 0;
                    // Clear the stream
                    stream.SetLength(0);

                    // Open the stream so we can override all content of the sheet
                    StreamWriter streamWriter = new StreamWriter(stream);
                    document.Save(streamWriter);
                    streamWriter.Flush();
                }
            }
        }
        
        /// <summary>
        /// Saves any pending changes to the Excel stream and adds/updates associated files if needed
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (this.Archive == null)
            {
                return;
            }

            if (this.Archive.Mode != ZipArchiveMode.Read)
            {
                bool ensureSharedStrings = false;

                // Update or create xl/sharedStrings.xml file
                if (this.SharedStrings != null)
                {
                    ensureSharedStrings = this.SharedStrings.PendingChanges;
                    this.SharedStrings.Write();
                }

                // Update xl/_rels/workbook.xml.rels file
                UpdateRelations(ensureSharedStrings);
                // Update [Content_Types].xml file
                UpdateContentTypes(ensureSharedStrings);
                // Update xl/workbook.xml file
                UpdateWorkbook();
            }

            this.Archive.Dispose();
        }
    }
}
