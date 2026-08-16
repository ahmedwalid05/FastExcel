using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// Builds a minimal but valid .xlsx package in memory with exactly the XML we want.
    /// <para>
    /// Binary fixtures make it impossible to see what a test is actually exercising. Most of
    /// the interesting defects in this library are about how a very specific piece of
    /// spreadsheet XML is interpreted — a shared-string table containing duplicates, a row
    /// with gaps in it, a drawing element among the cells — so the tests spell that XML out
    /// literally and build a package around it.
    /// </para>
    /// </summary>
    public sealed class XlsxBuilder
    {
        private const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string NsRel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NsPkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string NsCt = "http://schemas.openxmlformats.org/package/2006/content-types";

        private sealed class SheetSpec
        {
            public string Name;
            public string PartName;      // e.g. "worksheets/sheet1.xml"
            public string SheetDataInner; // raw <row> elements
        }

        private readonly List<SheetSpec> _sheets = new List<SheetSpec>();
        private readonly List<string> _definedNames = new List<string>();
        private string _sharedStringsInner;   // raw <si> elements, or null for no part

        /// <summary>Adds a worksheet whose &lt;sheetData&gt; contains the supplied raw XML.</summary>
        /// <param name="partName">
        /// Package part name, relative to xl/. Defaults to worksheets/sheet{N}.xml matching the
        /// declaration order. Override it to build a file whose sheet parts are NOT named after
        /// their position — which is what Excel does after a sheet is deleted or reordered.
        /// </param>
        public XlsxBuilder WithSheet(string name, string sheetDataInner = "", string partName = null)
        {
            _sheets.Add(new SheetSpec
            {
                Name = name,
                PartName = partName ?? $"worksheets/sheet{_sheets.Count + 1}.xml",
                SheetDataInner = sheetDataInner ?? string.Empty
            });
            return this;
        }

        /// <summary>Supplies the shared string table verbatim as a sequence of &lt;si&gt; elements.</summary>
        public XlsxBuilder WithSharedStringsXml(string siElements)
        {
            _sharedStringsInner = siElements;
            return this;
        }

        /// <summary>Supplies a plain shared string table, one &lt;si&gt;&lt;t&gt; per value.</summary>
        public XlsxBuilder WithSharedStrings(params string[] values)
            => WithSharedStringsXml(string.Concat(values.Select(v => $"<si><t>{Escape(v)}</t></si>")));

        /// <summary>
        /// Declares a defined name, e.g. WithDefinedName("Totals", "Sheet1!$A$1:$A$3").
        /// </summary>
        /// <param name="scopedToSheetIndex">
        /// 1-based sheet index to scope the name to. Null declares a workbook-global name.
        /// Excel stores this as a 0-based localSheetId, which is what gets written here.
        /// </param>
        public XlsxBuilder WithDefinedName(string name, string reference, int? scopedToSheetIndex = null)
        {
            var localSheetId = scopedToSheetIndex.HasValue
                ? $" localSheetId=\"{scopedToSheetIndex.Value - 1}\""
                : string.Empty;

            _definedNames.Add($"<definedName name=\"{Escape(name)}\"{localSheetId}>{Escape(reference)}</definedName>");
            return this;
        }

        public byte[] ToBytes()
        {
            if (_sheets.Count == 0) WithSheet("Sheet1");

            using var buffer = new MemoryStream();
            using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
            {
                Write(zip, "[Content_Types].xml", ContentTypes());
                Write(zip, "_rels/.rels",
                    $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    $"<Relationships xmlns=\"{NsPkgRel}\">" +
                    $"<Relationship Id=\"rId1\" Type=\"{NsRel}/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    $"</Relationships>");
                Write(zip, "xl/workbook.xml", Workbook());
                Write(zip, "xl/_rels/workbook.xml.rels", WorkbookRels());

                foreach (var sheet in _sheets)
                {
                    Write(zip, "xl/" + sheet.PartName,
                        $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        $"<worksheet xmlns=\"{NsMain}\" xmlns:r=\"{NsRel}\">" +
                        $"<sheetData>{sheet.SheetDataInner}</sheetData>" +
                        $"</worksheet>");
                }

                if (_sharedStringsInner != null)
                {
                    Write(zip, "xl/sharedStrings.xml",
                        $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                        $"<sst xmlns=\"{NsMain}\">{_sharedStringsInner}</sst>");
                }
            }
            return buffer.ToArray();
        }

        /// <summary>Materialises the package on disk and returns it.</summary>
        public FileInfo ToFile(TempWorkspace workspace, string fileName = "book.xlsx")
        {
            var path = workspace.Path(fileName);
            File.WriteAllBytes(path, ToBytes());
            return new FileInfo(path);
        }

        private string ContentTypes()
        {
            var overrides = new StringBuilder();
            overrides.Append($"<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            foreach (var s in _sheets)
                overrides.Append($"<Override PartName=\"/xl/{s.PartName}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            if (_sharedStringsInner != null)
                overrides.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");

            return $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   $"<Types xmlns=\"{NsCt}\">" +
                   $"<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                   $"<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                   overrides +
                   $"</Types>";
        }

        private string Workbook()
        {
            var sheets = string.Concat(_sheets.Select((s, i) =>
                $"<sheet name=\"{Escape(s.Name)}\" sheetId=\"{i + 1}\" r:id=\"rId{i + 1}\"/>"));

            var definedNames = _definedNames.Count == 0
                ? string.Empty
                : $"<definedNames>{string.Concat(_definedNames)}</definedNames>";

            return $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   $"<workbook xmlns=\"{NsMain}\" xmlns:r=\"{NsRel}\">" +
                   $"<sheets>{sheets}</sheets>{definedNames}</workbook>";
        }

        private string WorkbookRels()
        {
            var rels = string.Concat(_sheets.Select((s, i) =>
                $"<Relationship Id=\"rId{i + 1}\" Type=\"{NsRel}/worksheet\" Target=\"{s.PartName}\"/>"));

            if (_sharedStringsInner != null)
                rels += $"<Relationship Id=\"rId{_sheets.Count + 1}\" Type=\"{NsRel}/sharedStrings\" Target=\"sharedStrings.xml\"/>";

            return $"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   $"<Relationships xmlns=\"{NsPkgRel}\">{rels}</Relationships>";
        }

        private static void Write(ZipArchive zip, string path, string content)
        {
            using var stream = zip.CreateEntry(path).Open();
            using var writer = new StreamWriter(stream, new UTF8Encoding(false));
            writer.Write(content);
        }

        private static string Escape(string value) => value
            .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
            .Replace("\"", "&quot;").Replace("'", "&apos;");

        // ---- small helpers for composing sheetData ------------------------------------

        /// <summary>A &lt;row&gt; containing the supplied cell XML.</summary>
        public static string Row(int rowNumber, params string[] cells)
            => $"<row r=\"{rowNumber}\">{string.Concat(cells)}</row>";

        /// <summary>A cell referencing the shared string table, e.g. SharedCell("A1", 0).</summary>
        public static string SharedCell(string reference, int sharedStringIndex)
            => $"<c r=\"{reference}\" t=\"s\"><v>{sharedStringIndex}</v></c>";

        /// <summary>A numeric cell, e.g. NumberCell("B2", "42.5").</summary>
        public static string NumberCell(string reference, string value)
            => $"<c r=\"{reference}\"><v>{value}</v></c>";

        /// <summary>A cell carrying an inline string rather than a shared one.</summary>
        public static string InlineStringCell(string reference, string value)
            => $"<c r=\"{reference}\" t=\"inlineStr\"><is><t>{Escape(value)}</t></is></c>";
    }
}
