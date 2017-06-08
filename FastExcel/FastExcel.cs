using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace FastExcel
{
    public partial class FastExcel: IDisposable
    {
        public FileInfo ExcelFile { get; private set; }
        public FileInfo TemplateFile { get; private set; }
        public bool ReadOnly { get; private set; }
        
        internal SharedStrings SharedStrings { get; set; }
        internal ZipArchive Archive { get; set; }
        private bool UpdateExisting { get; set; }
        private bool _filesChecked;

        /// <summary>
        /// Maximum sheet number, obtained when a sheet is added
        /// </summary>
        internal int MaxSheetNumber { get; set; }

        /// <summary>
        /// A list of worksheet indexs to delete
        /// </summary>
        private List<int> DeleteWorksheets { get; set; }

        /// <summary>
        /// A list of worksheet indexs to insert
        /// </summary>
        private List<WorksheetAddSettings> AddWorksheets { get; set; }

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

        internal void PrepareArchive(bool openSharedStrings = true)
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
        internal void CheckFiles()
        {
            if (_filesChecked)
            {
                return;
            }

            if (this.UpdateExisting)
            {
                if (this.ExcelFile == null)
                {
                    throw new Exception("No input file name was supplied");
                }
                else if (!this.ExcelFile.Exists)
                {
                    string exceptionMessage = string.Format("Input file '{0}' does not exist", this.ExcelFile.FullName);
                    this.ExcelFile = null;
                    throw new Exception(exceptionMessage);
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
                    string exceptionMessage = string.Format("Template file '{0}' was not found", this.TemplateFile.FullName);
                    this.TemplateFile = null;
                    throw new FileNotFoundException(exceptionMessage);
                }

                if (this.ExcelFile == null)
                {
                    throw new Exception("No Ouput file name was supplied");
                }
                else if (this.ExcelFile.Exists)
                {
                    string exceptionMessage = string.Format("Output file '{0}' already exists", this.ExcelFile.FullName);
                    this.ExcelFile = null;
                    throw new Exception(exceptionMessage);
                }
            }
            _filesChecked = true;
        }

        /// <summary>
        /// Update xl/_rels/workbook.xml.rels file
        /// </summary>
        private void UpdateRelations(bool ensureStrings)
        {
            if (!(ensureStrings || 
                (this.DeleteWorksheets != null && this.DeleteWorksheets.Any()) || 
                (this.AddWorksheets != null && this.AddWorksheets.Any())))
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
                int id = relationshipElements.Count;
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
                        relationshipElement.Add(new XAttribute("Id", string.Format("rId{0}", ++id)));

                        document.Root.Add(relationshipElement);
                        update = true;
                    }
                }

                // Remove all references to sheets from this file as they are not requried
                if ((this.DeleteWorksheets != null && this.DeleteWorksheets.Any()) ||
                (this.AddWorksheets != null && this.AddWorksheets.Any()))
                {
                    XElement[] worksheetElements = (from element in relationshipElements
                                                    from attribute in element.Attributes()
                                                    where attribute.Name == "Type" && attribute.Value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                                                    select element).ToArray();
                    for (int i = worksheetElements.Length -1; i > 0; i--)
                    {
                        worksheetElements[i].Remove();
                        update = true;
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
        private string[] UpdateWorkbook()
        {
            if (!(this.DeleteWorksheets != null && this.DeleteWorksheets.Any() ||
                (this.AddWorksheets != null && this.AddWorksheets.Any())))
            {
                // Nothing to update
                return null;
            }

            List<string> sheetNames = new List<string>();
            using (Stream stream = this.Archive.GetEntry("xl/workbook.xml").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    throw new Exception("Unable to load workbook.xml");
                }

                bool update = false;

                RenameAndRebildWorksheetProperties((from sheet in document.Descendants()
                                              where sheet.Name.LocalName == "sheet"
                                              select sheet).ToArray());
                
                if (update)
                {
                    // Re number sheet ids
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    int id = 1;
                    foreach (XElement sheetElement in (from sheet in document.Descendants()
                                                       where sheet.Name.LocalName == "sheet"
                                                       select sheet))
                    {
                        sheetElement.SetAttributeValue(r + "id", string.Format("rId{0}", id++));
                        sheetNames.Add(sheetElement.Attribute("name").Value);
                    }

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
            return sheetNames.ToArray();
        }


        /// <summary>
        /// If sheets have been added or deleted, sheets need to be renamed
        /// </summary>
        private void RenameAndRebildWorksheetProperties(XElement[] sheets)
        {
            if (!((this.DeleteWorksheets != null && this.DeleteWorksheets.Any()) ||
                (this.AddWorksheets != null && this.AddWorksheets.Any())))
            {
                // Nothing to update
                return;
            }
            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            List<WorksheetProperties> sheetProperties = (from sheet in sheets
                                                         select new WorksheetProperties
                                                         ()
                                                         {
                                                             SheetId = int.Parse(sheet.Attribute("sheetId").Value),
                                                             Name = sheet.Attribute("name").Value,
                                                             CurrentIndex = int.Parse(sheet.Attribute(r + "id").Value)
                                                         }).ToList();

            // Remove deleted worksheets to sheetProperties
            if (this.DeleteWorksheets != null && this.DeleteWorksheets.Any())
            {
                foreach (var item in this.DeleteWorksheets)
                {
                    WorksheetProperties sheetToDelete = (from sp in sheetProperties
                                        where sp.SheetId == item
                                        select sp).FirstOrDefault();

                    if (sheetToDelete != null)
                    {
                        sheetProperties.Remove(sheetToDelete);
                    }
                }
            }

            // Add new worksheets to sheetProperties
            if (this.AddWorksheets != null && this.AddWorksheets.Any())
            {
                // Add the sheets in reverse, this will add them correctly with less work
                foreach (var item in this.AddWorksheets.Reverse<WorksheetAddSettings>())
                {
                    WorksheetProperties previousSheet = (from sp in sheetProperties
                                                where sp.SheetId == item.InsertAfterSheetId
                                        select sp).FirstOrDefault();
                    
                    if (previousSheet == null)
                    {
                        throw new Exception(string.Format("Sheet name {0} cannot be added because the insertAfterSheetNumber or insertAfterSheetName is now invalid", item.Name));
                    }

                    WorksheetProperties newWorksheet = new WorksheetProperties();
                    newWorksheet.SheetId = item.SheetId;
                    newWorksheet.Name = item.Name;
                    newWorksheet.CurrentIndex = 0;// TODO Something??

                    sheetProperties.Insert(sheetProperties.IndexOf(previousSheet), newWorksheet);
                }
            }

            int index = 1;
            foreach (WorksheetProperties worksheet in sheetProperties)
            {
                if (worksheet.CurrentIndex != index)
                {
                    ZipArchiveEntry entry = this.Archive.GetEntry(Worksheet.GetFileName(worksheet.CurrentIndex));
                    if (entry == null)
                    {
                        // TODO better message
                        throw new Exception("Worksheets could not be rebuilt");
                    }


                }
                index++;
            }
        }
        
        public class WorksheetProperties
        {
            public int CurrentIndex { get; set; }
            public int SheetId { get; set; }
            public string Name { get; set; }
        }

        /// <summary>
        /// Update [Content_Types].xml file
        /// </summary>
        private void UpdateContentTypes(bool ensureStrings)
        {
            if (!(ensureStrings ||
                (this.DeleteWorksheets != null && this.DeleteWorksheets.Any()) ||
                (this.AddWorksheets != null && this.AddWorksheets.Any())))
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
                        // the file name is different for each xml file
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

                if (this.AddWorksheets != null && this.AddWorksheets.Any())
                {
                    foreach (var item in this.AddWorksheets)
                    {
                        // the file name is different for each xml file
                        string fileName = string.Format("/xl/worksheets/sheet{0}.xml", item.SheetId);

                        XElement overrideElement = new XElement(document.Root.GetDefaultNamespace() + "Override");
                        overrideElement.Add(new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                        overrideElement.Add(new XAttribute("PartName", fileName));

                        document.Root.Add(overrideElement);
                        update = true;
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
        /// Update docProps/app.xml file
        /// </summary>
        private void UpdateDocPropsApp(string[] sheetNames)
        {
           /* if (sheetNames == null || !sheetNames.Any())
            {
                // Nothing to update
                return;
            }

            using (Stream stream = this.Archive.GetEntry("docProps/app.xml ").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    throw new Exception("Unable to load app.xml");
                }
                
                // Update TilesOfParts



                // Update HeadingPairs

                if (this.AddWorksheets != null && this.AddWorksheets.Any())
                {
                    // Add the sheets in reverse, this will add them correctly with less work
                    foreach (var item in this.AddWorksheets.Reverse<WorksheetAddSettings>())
                    {
                        XElement previousSheetElement = (from sheet in document.Descendants()
                                                         where sheet.Name.LocalName == "sheet"
                                                         from attribute in sheet.Attributes()
                                                         where attribute.Name == "sheetId" && attribute.Value == item.InsertAfterIndex.ToString()
                                                         select sheet).FirstOrDefault();

                        if (previousSheetElement == null)
                        {
                            throw new Exception(string.Format("Sheet name {0} cannot be added because the insertAfterSheetNumber or insertAfterSheetName is now invalid", item.Name));
                        }

                        XElement newSheetElement = new XElement(document.Root.GetDefaultNamespace() + "sheet");
                        newSheetElement.Add(new XAttribute("name", item.Name));
                        newSheetElement.Add(new XAttribute("sheetId", item.Index));

                        previousSheetElement.AddAfterSelf(newSheetElement);
                        update = true;
                    }
                }

                if (update)
                {
                    // Re number sheet ids
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    int id = 1;
                    foreach (XElement sheetElement in (from sheet in document.Descendants()
                                                       where sheet.Name.LocalName == "sheet"
                                                       select sheet))
                    {
                        sheetElement.SetAttributeValue(r + "id", string.Format("rId{0}", id++));
                    }

                    //Set the stream to the start
                    stream.Position = 0;
                    // Clear the stream
                    stream.SetLength(0);

                    // Open the stream so we can override all content of the sheet
                    StreamWriter streamWriter = new StreamWriter(stream);
                    document.Save(streamWriter);
                    streamWriter.Flush();
                }
            }*/
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

                // Update xl/workbook.xml file
                string[] sheetNames = UpdateWorkbook();

                // Update [Content_Types].xml file
                UpdateContentTypes(ensureSharedStrings);

                // Update docProps/app.xml file
                UpdateDocPropsApp(sheetNames);
            }

            this.Archive.Dispose();
        }
    }
}
