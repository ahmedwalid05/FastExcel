using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// The shared string table is the single richest source of correctness bugs in this
    /// library, because it is a <b>positional array</b> that the implementation treats as a
    /// set. Index N means "the Nth &lt;si&gt; element" — duplicates are legal and Excel emits
    /// them freely, and a single &lt;si&gt; may be split across several &lt;r&gt;&lt;t&gt;
    /// formatting runs that together form one string.
    /// </summary>
    public class SharedStringsTests
    {
        /// <summary>Reads every cell into a reference-keyed map, e.g. "A2" -> "Country".</summary>
        private static Dictionary<string, object> ReadCells(FileInfo file, int sheet = 1)
        {
            var values = new Dictionary<string, object>();
            using var fastExcel = new FastExcel(file, true);
            var worksheet = fastExcel.Read(sheet);
            foreach (var row in worksheet.Rows)
            {
                foreach (var cell in row.Cells)
                {
                    values[Cell.GetExcelColumnName(cell.ColumnNumber) + row.RowNumber] = cell.Value;
                }
            }
            return values;
        }

        // ---------------------------------------------------------------- baseline

        [Fact]
        public void SimpleTable_EachIndexResolvesToItsOwnEntry()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("Alpha", "Beta", "Gamma")
                .WithSheet("Sheet1", XlsxBuilder.Row(1,
                    XlsxBuilder.SharedCell("A1", 0),
                    XlsxBuilder.SharedCell("B1", 1),
                    XlsxBuilder.SharedCell("C1", 2)))
                .ToFile(workspace);

            var cells = ReadCells(file);

            Assert.Equal("Alpha", cells["A1"]);
            Assert.Equal("Beta", cells["B1"]);
            Assert.Equal("Gamma", cells["C1"]);
        }

        [Fact]
        public void NoSharedStringsPart_DoesNotThrow()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.NumberCell("A1", "42")))
                .ToFile(workspace);

            var cells = ReadCells(file);

            Assert.Equal("42", cells["A1"]);
        }

        // ---------------------------------------------------- #87 / #81 duplicate entries

        [Fact]
        public void DuplicateEntries_LaterIndicesStillResolveCorrectly()
        {
            using var workspace = new TempWorkspace();
            // "Alpha" appears at index 0 and again at index 2. Excel does this constantly —
            // most often when the same text carries different formatting.
            var file = new XlsxBuilder()
                .WithSharedStringsXml(
                    "<si><t>Alpha</t></si>" +
                    "<si><t>Beta</t></si>" +
                    "<si><t>Alpha</t></si>" +
                    "<si><t>Delta</t></si>")
                .WithSheet("Sheet1", XlsxBuilder.Row(1,
                    XlsxBuilder.SharedCell("A1", 2),
                    XlsxBuilder.SharedCell("B1", 3)))
                .ToFile(workspace);

            KnownBug.StillBroken("#87 / #81",
                "a duplicate <si> must still consume an index, so index 2 is \"Alpha\" and index 3 is \"Delta\"",
                () =>
                {
                    var cells = ReadCells(file);
                    Assert.Equal("Alpha", cells["A1"]);
                    Assert.Equal("Delta", cells["B1"]);
                });
        }

        [Fact]
        public void SameKeyFixture_ReturnsTheValueExcelShows()
        {
            // The checked-in fixture has four <si> entries where #0 and #2 are both "Country".
            // Row 2 reads indices 2 and 3.
            KnownBug.StillBroken("#87",
                "A2 is shared index 2 (\"Country\") and B2 is index 3 (\"Service Area\")",
                () =>
                {
                    var cells = ReadCells(TestFixtures.SameKey);
                    Assert.Equal("Country", cells["A2"]);
                    Assert.Equal("Service Area", cells["B2"]);
                });
        }

        [Fact]
        public void SameKeyFixture_EnumeratingCellsDoesNotThrow()
        {
            // Distinct from the assertion above: past the shifted values, the table runs out of
            // entries entirely and the read throws KeyNotFoundException rather than returning
            // anything at all.
            KnownBug.StillBroken("#81",
                "reading every cell of a workbook with duplicate shared strings must not throw",
                () => ReadCells(TestFixtures.SameKey));
        }

        // ------------------------------------------------------------ #10 rich text runs

        [Fact]
        public void RichTextRuns_FormOneEntryNotSeveral()
        {
            using var workspace = new TempWorkspace();
            // "Yellow" split across two formatting runs is ONE shared string at index 0,
            // so "Red" is index 1.
            var file = new XlsxBuilder()
                .WithSharedStringsXml(
                    "<si><r><t>Yel</t></r><r><t>low</t></r></si>" +
                    "<si><t>Red</t></si>")
                .WithSheet("Sheet1", XlsxBuilder.Row(1,
                    XlsxBuilder.SharedCell("A1", 0),
                    XlsxBuilder.SharedCell("B1", 1)))
                .ToFile(workspace);

            KnownBug.StillBroken("#10 / #87",
                "the <r><t> runs inside one <si> concatenate to \"Yellow\" and occupy a single index",
                () =>
                {
                    var cells = ReadCells(file);
                    Assert.Equal("Yellow", cells["A1"]);
                    Assert.Equal("Red", cells["B1"]);
                });
        }

        // -------------------------------------------------------------- text fidelity

        [Theory]
        [InlineData("plain")]
        [InlineData("with spaces")]
        [InlineData("ampersand & more")]
        [InlineData("angle <brackets>")]
        [InlineData("quote \" and apostrophe '")]
        [InlineData("naïve café Straße")]
        [InlineData("日本語のテキスト")]
        [InlineData("emoji 🎉 tail")]
        [InlineData("trailing space ")]
        [InlineData("  leading space")]
        public void ReadingAnEntry_PreservesTheTextExactly(string text)
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings(text)
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace);

            var cells = ReadCells(file);

            Assert.Equal(text, cells["A1"]);
        }

        [Fact]
        public void EmptyEntry_ReadsAsEmptyStringNotNull()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStringsXml("<si><t></t></si>")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace);

            var cells = ReadCells(file);

            Assert.Equal(string.Empty, cells["A1"]);
        }

        // ------------------------------------------------- #76 escaping on the write path

        /// <summary>Pulls xl/sharedStrings.xml out of a package exactly as another reader would see it.</summary>
        private static string RawSharedStringsXml(FileInfo file)
        {
            using var zip = ZipFile.OpenRead(file.FullName);
            var entry = zip.GetEntry("xl/sharedStrings.xml");
            Assert.NotNull(entry);
            using var reader = new StreamReader(entry.Open());
            return reader.ReadToEnd();
        }

        [Theory]
        [InlineData("Hello World")]
        [InlineData("two  spaces")]
        [InlineData("tab\tseparated")]
        public void WritingAStringWithWhitespace_DoesNotEscapeItIntoTheText(string text)
        {
            using var workspace = new TempWorkspace();
            var output = workspace.NewFile();

            using (var fastExcel = new FastExcel(TestFixtures.Template, output))
            {
                var worksheet = new Worksheet();
                worksheet.AddRow(text);
                fastExcel.Write(worksheet, 1);
            }

            KnownBug.StillBroken("#76",
                "cell text is written as XML character data, so a space stays a space rather " +
                "than becoming the literal _x0020_ that every other reader will display",
                () => Assert.DoesNotContain("_x00", RawSharedStringsXml(output)));
        }

        [Fact]
        public void WritingXmlSpecialCharacters_UsesXmlEntitiesNotNameEscapes()
        {
            using var workspace = new TempWorkspace();
            var output = workspace.NewFile();

            using (var fastExcel = new FastExcel(TestFixtures.Template, output))
            {
                var worksheet = new Worksheet();
                worksheet.AddRow("a & b < c");
                fastExcel.Write(worksheet, 1);
            }

            KnownBug.StillBroken("#76",
                "\"a & b < c\" is escaped as &amp; and &lt; entities, which is what the OOXML " +
                "spec requires and what Excel renders back as the original text",
                () =>
                {
                    var xml = RawSharedStringsXml(output);
                    Assert.Contains("&amp;", xml);
                    Assert.Contains("&lt;", xml);
                    Assert.DoesNotContain("_x0026_", xml);
                });
        }

        [Fact]
        public void WrittenStringsRoundTripThroughTheLibrary()
        {
            // This one passes today. It is the reason #76 went unnoticed for five years: the
            // library escapes and unescapes symmetrically, so a FastExcel -> FastExcel round
            // trip is clean even though the file on disk is wrong for everyone else.
            using var workspace = new TempWorkspace();
            var output = workspace.NewFile();
            const string text = "Hello World & <friends>";

            using (var fastExcel = new FastExcel(TestFixtures.Template, output))
            {
                var worksheet = new Worksheet();
                worksheet.AddRow(text);
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            var cells = ReadCells(output);
            Assert.Equal(text, cells["A1"]);
        }
    }
}
