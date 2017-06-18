using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace FastExcel
{
    /// <summary>
    /// Fast Excel
    /// </summary>
    public partial class FastExcel
    {
        private Worksheet[] _worksheets;

        /// <summary>
        /// List of worksheets, loaded on first access of property
        /// </summary>
        public Worksheet[] Worksheets
        {
            get
            {
                if (_worksheets != null)
                {
                    return _worksheets;
                }
                else
                {
                    _worksheets = GetWorksheetProperties();
                    return _worksheets;
                }
            }
        }

        private Worksheet[] GetWorksheetProperties()
        {
            CheckFiles();
            PrepareArchive(false);

            var worksheets = new List<Worksheet>();
            using (Stream stream = Archive.GetEntry("xl/workbook.xml").Open())
            {
                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    throw new Exception("Unable to load workbook.xml");
                }

                List<XElement> sheetsElements = document.Descendants().Where(d => d.Name.LocalName == "sheet").ToList();

                foreach (var sheetElement in sheetsElements)
                {
                    var worksheet = new Worksheet(this);
                    worksheet.Index = sheetsElements.IndexOf(sheetElement) + 1;

                    worksheet.Name = (from attribute in sheetElement.Attributes()
                                 where attribute.Name == "name"
                                 select attribute.Value).FirstOrDefault();

                    worksheets.Add(worksheet);
                }
            }
            return worksheets.ToArray();
        }
    }
}
