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
    public class FastExcelWriter
    {
        public FileInfo TemplateFile { get; private set;}
        public FileInfo OutpuFile { get; private set; }
        private SharedStrings SharedStrings { get; set; }

        public FastExcelWriter(FileInfo outputFile)
            :this(null, outputFile)
        {
        }

        public FastExcelWriter(FileInfo templateFile, FileInfo outputFile)
        {
            this.TemplateFile = templateFile;
            this.OutpuFile = outputFile;

            CheckFiles();
        }

        private bool CheckFiles()
        {
            if (!this.TemplateFile.Exists)
            {
                this.TemplateFile = null;
                throw new Exception();
            }

            if (this.OutpuFile.Exists)
            {
                this.OutpuFile = null;

                //throw new Exception();
            }

            return true;
        }

        public void Write(DataSet data, int? sheetNumber = null, string sheetName = null, int existingHeadingRows = 0)
        {
            if (!CheckFiles())
            {
                return;
            }

            try
            {
                File.Copy(this.TemplateFile.FullName, this.OutpuFile.FullName);
            }
            catch (Exception) { }

            using (ZipArchive archive = ZipFile.Open(this.OutpuFile.FullName, ZipArchiveMode.Update))
            {
                // Get Strings file
                this.SharedStrings = new SharedStrings(archive);
                
                // Open worksheet
                Worksheet worksheet = null;
                if (sheetNumber.HasValue)
                {
                    worksheet = new Worksheet(archive, SharedStrings, sheetNumber.Value);
                }
                else if (!string.IsNullOrEmpty(sheetName))
                {
                    worksheet = new Worksheet(archive, SharedStrings, sheetName);
                }
                else
                {
                    //TODO Thow exception
                }

                worksheet.ExistingHeadingRows = existingHeadingRows;

                //Write Data
                worksheet.Write(data);

                // Update or create xl/sharedStrings.xml file
                this.SharedStrings.Write();

                // Update xl/_rels/workbook.xml.rels file
                UpdateRelations(archive);
                // Update [Content_Types].xml file
                UpdateContentTypes(archive);
            }
        }

        /// <summary>
        /// Update xl/_rels/workbook.xml.rels file
        /// </summary>
        private void UpdateRelations(ZipArchive archive)
        {
            using (Stream stream = archive.GetEntry("xl/_rels/workbook.xml.rels").Open())
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
        private void UpdateContentTypes(ZipArchive archive)
        {
            using (Stream stream = archive.GetEntry("[Content_Types].xml").Open())
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
        
    }
}
