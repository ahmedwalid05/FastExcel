using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace FastExcel
{
    /// <summary>
    /// Read and update xl/sharedStrings.xml file
    /// </summary>
    public class SharedStrings
    {
        //A dictionary is a lot faster than a list
        private Dictionary<string, int> StringDictionary { get; set; }
        private Dictionary<int, string> StringArray { get; set; }

        private bool SharedStringsExists { get; set; }
        private ZipArchive ZipArchive { get; set; }

        /// <summary>
        /// Is there any pending changes
        /// </summary>
        public bool PendingChanges { get; private set; }

        /// <summary>
        /// Is in read/write mode
        /// </summary>
        public bool ReadWriteMode { get; set; }

        internal SharedStrings(ZipArchive archive)
        {
            ZipArchive = archive;

            SharedStringsExists = true;

            if (!ZipArchive.Entries.Where(entry => entry.FullName == "xl/sharedStrings.xml").Any())
            {
                StringDictionary = new Dictionary<string, int>();
                SharedStringsExists = false;
                return;
            }

            using (Stream stream = ZipArchive.GetEntry("xl/sharedStrings.xml").Open())
            {
                if (stream == null)
                {
                    StringDictionary = new Dictionary<string, int>();
                    SharedStringsExists = false;
                    return;
                }

                XDocument document = XDocument.Load(stream);

                if (document == null)
                {
                    StringDictionary = new Dictionary<string, int>();
                    SharedStringsExists = false;
                    return;
                }

                int i = 0;
                StringDictionary = document.Descendants().Where(d => d.Name.LocalName == "t").Select(e => e.Value).ToDictionary(k => k, v => i++);
            }
        }

        internal int AddString(string stringValue)
        {
            if (StringDictionary.ContainsKey(stringValue))
            {
                return StringDictionary[stringValue];
            }
            else
            {
                PendingChanges = true;
                StringDictionary.Add(stringValue, StringDictionary.Count);

                // Clear String Array used for retrieval
                if (ReadWriteMode && StringArray != null)
                {
                    StringArray.Add(StringDictionary.Count - 1, stringValue);
                }
                else
                {
                    StringArray = null;
                }

                return StringDictionary.Count - 1;
            }
        }

        internal void Write()
        {
            // Only update if changes were made
            if (!PendingChanges)
            {
                return;
            }

            StreamWriter streamWriter = null;
            try
            {
                if (SharedStringsExists)
                {
                    streamWriter = new StreamWriter(ZipArchive.GetEntry("xl/sharedStrings.xml").Open());
                }
                else
                {
                    streamWriter = new StreamWriter(ZipArchive.CreateEntry("xl/sharedStrings.xml").Open());
                }

                // TODO instead of saving the headers then writing them back get position where the headers finish then write from there

                /* Note: the count attribute value is wrong, it is the number of times strings are used thoughout the workbook it is different to the unique count 
                 *       but because this library is about speed and Excel does not seem to care I am not going to fix it because I would need to read the whole workbook
                 */

                streamWriter.Write(string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                            "<sst uniqueCount=\"{0}\" count=\"{0}\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">", StringDictionary.Count));

                // Add Rows
                foreach (var stringValue in StringDictionary)
                {
                    streamWriter.Write(string.Format("<si><t>{0}</t></si>", stringValue.Key));
                }

                //Add Footers
                streamWriter.Write("</sst>");
                streamWriter.Flush();
            }
            finally
            {
                streamWriter.Dispose();
                PendingChanges = false;
            }
        }

        internal string GetString(string position)
        {
            int pos;
            if (int.TryParse(position, out pos))
            {
                return GetString(pos + 1);
            }
            else
            {
                // TODO: should I throw an error? this is a corrupted excel document
                return string.Empty;
            }
        }

        internal string GetString(int position)
        {
            if (StringArray == null)
            {
                StringArray = StringDictionary.ToDictionary(kv => kv.Value, kv => kv.Key);
            }

            return StringArray[position - 1];
        }
    }
}