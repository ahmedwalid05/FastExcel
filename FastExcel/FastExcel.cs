using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace FastExcel {
    /// <summary>
    /// Fast Excel
    /// </summary>
    public partial class FastExcel : IDisposable {

        /// <summary>
        ///  Output excel file
        /// </summary>
        public FileInfo ExcelFile
        {
            get {
                if (_excelFile == null)
                {
                    throw new ApplicationException("ExcelFile was not provided");
                }
                return _excelFile;
            }
        }
        /// <summary>
        /// The template excel file
        /// </summary>
        public FileInfo TemplateFile{
            get
            {
                if (_templateFile == null)
                {
                    throw new ApplicationException("TemplateFile was not provided");
                }

                return _templateFile;
            }
        }
        private Stream ExcelFileStream { get; set; }
        private Stream TemplateFileStream { get; set; }

        /// <summary>
        /// Is the excel file read only
        /// </summary>
        public bool ReadOnly { get; private set; }

        internal SharedStrings SharedStrings { get; set; }
        internal ZipArchive Archive { get; set; }
        private bool UpdateExisting { get; set; }
        private bool _filesChecked;
        private readonly FileInfo _excelFile;
        private readonly FileInfo _templateFile;

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
        /// <param name="readOnly">is the file read only</param>
        public FastExcel(FileInfo excelFile, bool readOnly = false) : this(null, excelFile, true, readOnly) {
        }

        /// <summary>
        /// Create a new excel file from a template
        /// </summary>
        /// <param name="templateFile">template location</param>
        /// <param name="excelFile">location of where a new excel file will be saved to</param>
        public FastExcel(FileInfo templateFile, FileInfo excelFile) : this(templateFile, excelFile, false, false) {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="templateFile"></param>
        /// <param name="excelFile"></param>
        /// <param name="updateExisting"></param>
        /// <param name="readOnly"></param>
        private FastExcel(FileInfo templateFile, FileInfo excelFile, bool updateExisting, bool readOnly = false) {
            if (updateExisting) {
                if (!excelFile.Exists) {
                    var exceptionMessage = $"Input file '{excelFile.FullName}' does not exist";
                    throw new FileNotFoundException(exceptionMessage);
                }
            }
            else {
                if (excelFile.Exists) {
                    var exceptionMessage = $"Output file '{excelFile.FullName}' already exists";
                    throw new Exception(exceptionMessage);
                }
                if (!templateFile.Exists) {
                    var exceptionMessage = $"Template file '{templateFile.FullName}' was not found";
                    throw new FileNotFoundException(exceptionMessage);
                }
            }
            _templateFile = templateFile;
            _excelFile = excelFile;
            TemplateFileStream = templateFile !=  null ? new FileStream(templateFile.FullName, FileMode.Open, FileAccess.Read) : null;
            ExcelFileStream = updateExisting
                ? new FileStream(excelFile.FullName, FileMode.Open, readOnly ? FileAccess.Read : FileAccess.ReadWrite)
                : new FileStream(excelFile.FullName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            UpdateExisting = updateExisting;
            ReadOnly = readOnly;

            CheckFiles();
        }

        /// <summary>
        /// Update an existing excel file stream
        /// </summary>
        /// <param name="excelStream"></param>
        public FastExcel(Stream excelStream) : this(null, excelStream, true) {
        }

        /// <summary>
        /// Create a new excel file from a template
        /// </summary>
        /// <param name="templateStream">Input Template Stream</param>
        /// <param name="excelStream">Output Excel Stream</param>
        /// <param name="updateExisting"></param>
        /// <param name="readOnly"></param>
        public FastExcel(Stream templateStream, Stream excelStream, bool updateExisting = false, bool readOnly = false) {
            if (templateStream is FileStream templatefileStream)
            {
                _templateFile = new FileInfo(templatefileStream.Name);
            }
            if (excelStream is FileStream excelFileStream)
            {
                _excelFile = new FileInfo(excelFileStream.Name);
            }
            TemplateFileStream = templateStream;
            ExcelFileStream = excelStream;
            UpdateExisting = updateExisting;
            ReadOnly = readOnly;
            CheckFiles();
        }

        internal void PrepareArchive(bool openSharedStrings = true) {
            if (Archive == null) {
                if (ReadOnly) {
                    Archive = new ZipArchive(ExcelFileStream, ZipArchiveMode.Read);
                }
                else {
                    Archive = new ZipArchive(ExcelFileStream, ZipArchiveMode.Update);
                }
            }

            // Get Strings file
            if (SharedStrings == null && openSharedStrings) {
                SharedStrings = new SharedStrings(Archive);
            }
        }

        /// <summary>
        /// Ensure files are ready for use
        /// </summary>
        internal void CheckFiles() {
            if (_filesChecked) {
                return;
            }

            if (UpdateExisting) {
                if (ExcelFileStream?.Length == 0) {
                    throw new Exception("No input file name was supplied");
                }
            }
            else {
                if (TemplateFileStream == null) {
                    throw new Exception("No Template file was supplied");
                }

                if (ExcelFileStream == null) {
                    throw new Exception("No Ouput file name was supplied");
                }
                else if (ExcelFileStream.Length > 0) {
                    var exceptionMessage = $"Output file  already exists";
                    throw new Exception(exceptionMessage);
                }
            }

            _filesChecked = true;
        }

        /// <summary>
        /// Update xl/_rels/workbook.xml.rels file
        /// </summary>
        private void UpdateRelations(bool ensureStrings) {
            if (!(ensureStrings ||
                  (DeleteWorksheets != null && DeleteWorksheets.Any()) ||
                  (AddWorksheets != null && AddWorksheets.Any()))) {
                // Nothing to update
                return;
            }

            using (Stream stream = Archive.GetEntry("xl/_rels/workbook.xml.rels").Open()) {
                XDocument document = XDocument.Load(stream);

                if (document == null) {
                    //TODO error
                }

                bool update = false;

                List<XElement> relationshipElements = document.Descendants().Where(d => d.Name.LocalName == "Relationship").ToList();
                int id = relationshipElements.Count;
                if (ensureStrings) {
                    //Ensure SharedStrings
                    XElement relationshipElement = (from element in relationshipElements
                        from attribute in element.Attributes()
                        where attribute.Name == "Target" && attribute.Value.Equals("sharedStrings.xml", StringComparison.OrdinalIgnoreCase)
                        select element).FirstOrDefault();

                    if (relationshipElement == null) {
                        relationshipElement = new XElement(document.Root.GetDefaultNamespace() + "Relationship");
                        relationshipElement.Add(new XAttribute("Target", "sharedStrings.xml"));
                        relationshipElement.Add(new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"));
                        relationshipElement.Add(new XAttribute("Id", string.Format("rId{0}", ++id)));

                        document.Root.Add(relationshipElement);
                        update = true;
                    }
                }

                // Remove all references to sheets from this file as they are not requried
                if ((DeleteWorksheets != null && DeleteWorksheets.Any()) ||
                    (AddWorksheets != null && AddWorksheets.Any())) {
                    XElement[] worksheetElements = (from element in relationshipElements
                        from attribute in element.Attributes()
                        where attribute.Name == "Type" && attribute.Value == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
                        select element).ToArray();
                    for (int i = worksheetElements.Length - 1; i > 0; i--) {
                        worksheetElements[i].Remove();
                        update = true;
                    }
                }

                if (update) {
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
        private string[] UpdateWorkbook() {
            if (!(DeleteWorksheets != null && DeleteWorksheets.Any() ||
                  (AddWorksheets != null && AddWorksheets.Any()))) {
                // Nothing to update
                return null;
            }

            List<string> sheetNames = new List<string>();
            using (Stream stream = Archive.GetEntry("xl/workbook.xml").Open()) {
                XDocument document = XDocument.Load(stream);

                if (document == null) {
                    throw new Exception("Unable to load workbook.xml");
                }

                bool update = false;

                RenameAndRebildWorksheetProperties((from sheet in document.Descendants()
                    where sheet.Name.LocalName == "sheet"
                    select sheet).ToArray());

                if (update) {
                    // Re number sheet ids
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    int id = 1;
                    foreach (XElement sheetElement in (from sheet in document.Descendants()
                        where sheet.Name.LocalName == "sheet"
                        select sheet)) {
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
        private void RenameAndRebildWorksheetProperties(XElement[] sheets) {
            if (!((DeleteWorksheets != null && DeleteWorksheets.Any()) ||
                  (AddWorksheets != null && AddWorksheets.Any()))) {
                // Nothing to update
                return;
            }

            XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            List<WorksheetProperties> sheetProperties = (from sheet in sheets
                select new WorksheetProperties
                    () {
                        SheetId = int.Parse(sheet.Attribute("sheetId").Value),
                        Name = sheet.Attribute("name").Value,
                        CurrentIndex = int.Parse(sheet.Attribute(r + "id").Value)
                    }).ToList();

            // Remove deleted worksheets to sheetProperties
            if (DeleteWorksheets != null && DeleteWorksheets.Any()) {
                foreach (var item in DeleteWorksheets) {
                    WorksheetProperties sheetToDelete = (from sp in sheetProperties
                        where sp.SheetId == item
                        select sp).FirstOrDefault();

                    if (sheetToDelete != null) {
                        sheetProperties.Remove(sheetToDelete);
                    }
                }
            }

            // Add new worksheets to sheetProperties
            if (AddWorksheets != null && AddWorksheets.Any()) {
                // Add the sheets in reverse, this will add them correctly with less work
                foreach (var item in AddWorksheets.Reverse<WorksheetAddSettings>()) {
                    WorksheetProperties previousSheet = (from sp in sheetProperties
                        where sp.SheetId == item.InsertAfterSheetId
                        select sp).FirstOrDefault();

                    if (previousSheet == null) {
                        throw new Exception(string.Format("Sheet name {0} cannot be added because the insertAfterSheetNumber or insertAfterSheetName is now invalid", item.Name));
                    }

                    WorksheetProperties newWorksheet = new WorksheetProperties() {
                        SheetId = item.SheetId,
                        Name = item.Name,
                        CurrentIndex = 0 // TODO Something??
                    };
                    sheetProperties.Insert(sheetProperties.IndexOf(previousSheet), newWorksheet);
                }
            }

            int index = 1;
            foreach (WorksheetProperties worksheet in sheetProperties) {
                if (worksheet.CurrentIndex != index) {
                    ZipArchiveEntry entry = Archive.GetEntry(Worksheet.GetFileName(worksheet.CurrentIndex));
                    if (entry == null) {
                        // TODO better message
                        throw new Exception("Worksheets could not be rebuilt");
                    }
                }

                index++;
            }
        }

        /// <summary>
        /// Update [Content_Types].xml file
        /// </summary>
        private void UpdateContentTypes(bool ensureStrings) {
            if (!(ensureStrings ||
                  (DeleteWorksheets != null && DeleteWorksheets.Any()) ||
                  (AddWorksheets != null && AddWorksheets.Any()))) {
                // Nothing to update
                return;
            }

            using (Stream stream = Archive.GetEntry("[Content_Types].xml").Open()) {
                XDocument document = XDocument.Load(stream);

                if (document == null) {
                    //TODO error
                }

                bool update = false;
                List<XElement> overrideElements = document.Descendants().Where(d => d.Name.LocalName == "Override").ToList();

                //Ensure SharedStrings
                if (ensureStrings) {
                    XElement overrideElement = (from element in overrideElements
                        from attribute in element.Attributes()
                        where attribute.Name == "PartName" && attribute.Value.Equals("/xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase)
                        select element).FirstOrDefault();

                    if (overrideElement == null) {
                        overrideElement = new XElement(document.Root.GetDefaultNamespace() + "Override");
                        overrideElement.Add(new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"));
                        overrideElement.Add(new XAttribute("PartName", "/xl/sharedStrings.xml"));

                        document.Root.Add(overrideElement);
                        update = true;
                    }
                }

                if (DeleteWorksheets != null && DeleteWorksheets.Any()) {
                    foreach (var item in DeleteWorksheets) {
                        // the file name is different for each xml file
                        string fileName = string.Format("/xl/worksheets/sheet{0}.xml", item);

                        XElement overrideElement = (from element in overrideElements
                            from attribute in element.Attributes()
                            where attribute.Name == "PartName" && attribute.Value == fileName
                            select element).FirstOrDefault();
                        if (overrideElement != null) {
                            overrideElement.Remove();
                            update = true;
                        }
                    }
                }

                if (AddWorksheets != null && AddWorksheets.Any()) {
                    foreach (var item in AddWorksheets) {
                        // the file name is different for each xml file
                        string fileName = string.Format("/xl/worksheets/sheet{0}.xml", item.SheetId);

                        XElement overrideElement = new XElement(document.Root.GetDefaultNamespace() + "Override");
                        overrideElement.Add(new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
                        overrideElement.Add(new XAttribute("PartName", fileName));

                        document.Root.Add(overrideElement);
                        update = true;
                    }
                }

                if (update) {
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
        /// Retrieves the index for given worksheet name
        /// </summary>
        /// <param name="name"></param>
        /// <returns>1 based index of sheet or 0 if not found</returns>
        public int GetWorksheetIndexFromName(string name) {
            return (from worksheet in Worksheets where worksheet.Name == name select worksheet.Index).FirstOrDefault();
        }

        /// <summary>
        /// Update docProps/app.xml file
        /// </summary>
        private void UpdateDocPropsApp(string[] sheetNames) {
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
        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Main disposal function
        /// </summary>
        protected virtual void Dispose(bool disposing) {
            if (Archive == null) {
                return;
            }

            if (Archive.Mode != ZipArchiveMode.Read) {
                bool ensureSharedStrings = false;

                // Update or create xl/sharedStrings.xml file
                if (SharedStrings != null) {
                    ensureSharedStrings = SharedStrings.PendingChanges;
                    SharedStrings.Write();
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

            Archive.Dispose();
        }
    }
}