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

        private SharedStrings SharedStrings { get; set; }
        private ZipArchive Archive { get; set; }
        private bool UpdateExisting { get; set; }

        /// <summary>
        /// Update an existing excel file
        /// </summary>
        /// <param name="excelFile">location of an existing excel file</param>
        public FastExcel(FileInfo excelFile) : this(null, excelFile, true) {}
        
        /// <summary>
        /// Create a new excel file from a template
        /// </summary>
        /// <param name="templateFile">template location</param>
        /// <param name="excelFile">location of where a new excel file will be saved to</param>
        public FastExcel(FileInfo templateFile, FileInfo excelFile) :this(templateFile, excelFile, false) {}

        private FastExcel(FileInfo templateFile, FileInfo excelFile, bool updateExisting)
        {
            this.TemplateFile = templateFile;
            this.ExcelFile = excelFile;
            this.UpdateExisting = updateExisting;

            CheckFiles();
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
        private void UpdateRelations()
        {
            //I'm keeping UpdateRelations in this class because it might need to update more than shared strings eventually

            using (Stream stream = this.Archive.GetEntry("xl/_rels/workbook.xml.rels").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    //TODO error
                }

                List<XElement> relationshipElements = document.Descendants().Where(d => d.Name.LocalName == "Relationship").ToList();

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

                    //Set the stream to the start
                    stream.Position = 0;

                    // Open the stream so we can override all content of the sheet
                    StreamWriter streamWriter = new StreamWriter(stream);
                    document.Save(streamWriter);
                    streamWriter.Flush();
                }
            }
        }

        /// <summary>
        /// Update [Content_Types].xml file
        /// </summary>
        private void UpdateContentTypes()
        {
            //I'm keeping UpdateContentTypes in this class because it might need to update more than shared strings eventually

            using (Stream stream = this.Archive.GetEntry("[Content_Types].xml").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    //TODO error
                }

                List<XElement> overrideElements = document.Descendants().Where(d => d.Name.LocalName == "Override").ToList();

                //Ensure SharedStrings
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
                    //stream.Position = 0;

                    //Set the stream to the start
                    stream.Position = 0;

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
            if (this.Archive == null)
            {
                return;
            }

            bool ensureSharedStrings = false;

            // Update or create xl/sharedStrings.xml file
            if (this.SharedStrings != null)
            {
                ensureSharedStrings = this.SharedStrings.PendingChanges;
                this.SharedStrings.Write();
            }

            if (ensureSharedStrings)
            {
                // Update xl/_rels/workbook.xml.rels file
                UpdateRelations();
                // Update [Content_Types].xml file
                UpdateContentTypes();
            }

            this.Archive.Dispose();
        }
    }
}
