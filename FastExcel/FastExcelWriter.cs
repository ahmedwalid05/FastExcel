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
    public class FastExcelWriter : IDisposable
    {
        public FileInfo TemplateFile { get; private set;}
        public FileInfo OutpuFile { get; private set; }
        private SharedStrings SharedStrings { get; set; }
        private ZipArchive Archive { get; set; }
        
        public FastExcelWriter(FileInfo templateFile, FileInfo outputFile)
        {
            this.TemplateFile = templateFile;
            this.OutpuFile = outputFile;

            CheckFiles();
        }

        /// <summary>
        /// Ensure files are ready for use
        /// </summary>
        private void CheckFiles()
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

            if (this.OutpuFile == null)
            {
                throw new Exception("No Ouput file name was supplied");
            }
            else if (this.OutpuFile.Exists)
            {
                this.OutpuFile = null;
                throw new Exception(string.Format("Output file '{0}' already exists", this.OutpuFile.FullName));
            }
        }

        /// <summary>
        /// Write data to a sheet
        /// </summary>
        /// <param name="data">A dataset</param>
        /// <param name="sheetNumber">The number of the sheet starting at 1</param>
        /// <param name="existingHeadingRows">How many rows in the template sheet you would like to keep</param>
        public void Write(DataSet data, int sheetNumber, int existingHeadingRows = 0)
        {
            Write(data, sheetNumber, null, existingHeadingRows);
        }

        /// <summary>
        /// Write data to a sheet
        /// </summary>
        /// <param name="data">A dataset</param>
        /// <param name="sheetName">The display name of the sheet</param>
        /// <param name="existingHeadingRows">How many rows in the template sheet you would like to keep</param>
        public void Write(DataSet data, string sheetName, int existingHeadingRows = 0)
        {
            Write(data, null, sheetName, existingHeadingRows);
        }

        private void Write(DataSet data, int? sheetNumber = null, string sheetName = null, int existingHeadingRows = 0)
        {
            CheckFiles();

            try
            {
                File.Copy(this.TemplateFile.FullName, this.OutpuFile.FullName);
            }
            catch (Exception ex) 
            {
                throw new Exception("Could not copy template to output file path", ex);
            }

            if (this.Archive == null)
            {
                Archive = ZipFile.Open(this.OutpuFile.FullName, ZipArchiveMode.Update);
            }
            
            // Get Strings file
            if (this.SharedStrings == null)
            {
                this.SharedStrings = new SharedStrings(this.Archive);
            }
                
            // Open worksheet
            Worksheet worksheet = null;
            if (sheetNumber.HasValue)
            {
                worksheet = new Worksheet(this.Archive, SharedStrings, sheetNumber.Value);
            }
            else if (!string.IsNullOrEmpty(sheetName))
            {
                worksheet = new Worksheet(this.Archive, SharedStrings, sheetName);
            }
            else
            {
                throw new Exception("No worksheet name or number was specified");
            }

            worksheet.ExistingHeadingRows = existingHeadingRows;

            //Write Data
            worksheet.Write(data);
        }

        /// <summary>
        /// Update xl/_rels/workbook.xml.rels file
        /// </summary>
        private void UpdateRelations()
        {
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

        public void Dispose()
        {
            if (this.Archive == null)
            {
                return;
            }

            // Update or create xl/sharedStrings.xml file
            if (this.SharedStrings != null)
            {
                this.SharedStrings.Write();
            }

            // Update xl/_rels/workbook.xml.rels file
            UpdateRelations();
            // Update [Content_Types].xml file
            UpdateContentTypes();

            this.Archive.Dispose();
        }
    }
}
