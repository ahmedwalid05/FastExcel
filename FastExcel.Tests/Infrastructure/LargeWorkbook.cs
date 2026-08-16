using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// Generates large workbooks for performance work, writing the package directly rather
    /// than through FastExcel so that generation cost never contaminates a measurement of the
    /// library itself.
    /// <para>
    /// Files are deterministic and cached in the temp directory by their parameters, so a
    /// benchmark or test run pays the generation cost once per machine rather than per run.
    /// </para>
    /// </summary>
    public static class LargeWorkbook
    {
        private const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>What kind of cell payload to fill the sheet with.</summary>
        public enum Payload
        {
            /// <summary>Cells reference the shared string table (t="s").</summary>
            SharedStrings,

            /// <summary>Cells carry inline numeric values, as a data export typically would.</summary>
            Numbers
        }

        private static readonly string CacheRoot =
            Path.Combine(Path.GetTempPath(), "fastexcel-largefixtures");

        /// <summary>
        /// Returns a workbook of the requested shape, generating it if it is not already cached.
        /// </summary>
        /// <param name="definedNameCount">
        /// Number of defined names to declare. Reading cost is expected to be independent of
        /// this, but the per-cell lookup currently scans every defined name, so it is a
        /// parameter worth being able to vary.
        /// </param>
        public static FileInfo Get(int rows, int columns, Payload payload = Payload.SharedStrings, int definedNameCount = 0)
        {
            Directory.CreateDirectory(CacheRoot);

            var name = $"r{rows}-c{columns}-{payload}-n{definedNameCount}.xlsx";
            var path = Path.Combine(CacheRoot, name);
            var file = new FileInfo(path);
            if (file.Exists && file.Length > 0) return file;

            // Write to a unique temp name first, then move into place, so concurrent runs
            // cannot observe a half-written package.
            var staging = Path.Combine(CacheRoot, Guid.NewGuid().ToString("n") + ".tmp");
            Generate(staging, rows, columns, payload, definedNameCount);

            try { File.Move(staging, path); }
            catch (IOException) { File.Delete(staging); } // lost the race; the winner's file is fine

            return new FileInfo(path);
        }

        /// <summary>Total cells a workbook of this shape contains.</summary>
        public static long CellCount(int rows, int columns) => (long)rows * columns;

        private static void Generate(string path, int rows, int columns, Payload payload, int definedNameCount)
        {
            using var stream = File.Create(path);
            using var zip = new ZipArchive(stream, ZipArchiveMode.Create);

            void WritePart(string partName, string content)
            {
                using var entry = zip.CreateEntry(partName).Open();
                using var writer = new StreamWriter(entry, new UTF8Encoding(false));
                writer.Write(content);
            }

            WritePart("[Content_Types].xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>" +
                "</Types>");

            WritePart("_rels/.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<Relationships xmlns=\"{NsPkgRel}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{NsRel}/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                "</Relationships>");

            var definedNames = new StringBuilder();
            if (definedNameCount > 0)
            {
                definedNames.Append("<definedNames>");
                for (var i = 0; i < definedNameCount; i++)
                {
                    var column = ColumnName((i % columns) + 1);
                    definedNames.Append(
                        $"<definedName name=\"Name{i}\">Sheet1!${column}${(i % rows) + 1}</definedName>");
                }
                definedNames.Append("</definedNames>");
            }

            WritePart("xl/workbook.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<workbook xmlns=\"{NsMain}\" xmlns:r=\"{NsRel}\">" +
                "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                definedNames +
                "</workbook>");

            WritePart("xl/_rels/workbook.xml.rels",
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                $"<Relationships xmlns=\"{NsPkgRel}\">" +
                $"<Relationship Id=\"rId1\" Type=\"{NsRel}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                $"<Relationship Id=\"rId2\" Type=\"{NsRel}/sharedStrings\" Target=\"sharedStrings.xml\"/>" +
                "</Relationships>");

            // One distinct shared string per column keeps the table small and realistic;
            // real exports repeat category values down each column.
            var sst = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sst.Append($"<sst xmlns=\"{NsMain}\">");
            for (var c = 0; c < Math.Max(columns, 1); c++) sst.Append($"<si><t>value {c}</t></si>");
            sst.Append("</sst>");
            WritePart("xl/sharedStrings.xml", sst.ToString());

            using var sheetEntry = zip.CreateEntry("xl/worksheets/sheet1.xml").Open();
            using var sheet = new StreamWriter(sheetEntry, new UTF8Encoding(false), 1 << 20);

            sheet.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sheet.Write($"<worksheet xmlns=\"{NsMain}\"><sheetData>");

            var columnNames = new string[columns];
            for (var c = 0; c < columns; c++) columnNames[c] = ColumnName(c + 1);

            for (var r = 1; r <= rows; r++)
            {
                sheet.Write("<row r=\"");
                sheet.Write(r.ToString(CultureInfo.InvariantCulture));
                sheet.Write("\">");

                for (var c = 0; c < columns; c++)
                {
                    sheet.Write("<c r=\"");
                    sheet.Write(columnNames[c]);
                    sheet.Write(r.ToString(CultureInfo.InvariantCulture));

                    if (payload == Payload.SharedStrings)
                    {
                        sheet.Write("\" t=\"s\"><v>");
                        sheet.Write(c.ToString(CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        sheet.Write("\"><v>");
                        sheet.Write(((r * 31 + c) % 9973).ToString(CultureInfo.InvariantCulture));
                    }

                    sheet.Write("</v></c>");
                }

                sheet.Write("</row>");
            }

            sheet.Write("</sheetData></worksheet>");
        }

        private static string ColumnName(int number)
        {
            var name = string.Empty;
            while (number > 0)
            {
                var modulo = (number - 1) % 26;
                name = (char)('A' + modulo) + name;
                number = (number - modulo) / 26;
            }
            return name;
        }
    }
}
