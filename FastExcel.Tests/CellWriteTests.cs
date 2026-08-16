using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Covers how values become cell XML. Two whole classes of corruption live here: the
    /// writer formats with the ambient culture (so a comma-decimal locale emits numbers Excel
    /// rejects), and it recognises only int, double and string (so every other type is
    /// stringified into a cell declared as numeric).
    /// </summary>
    public class CellWriteTests
    {
        /// <summary>Writes one row and returns the resulting sheet1.xml &lt;sheetData&gt; content.</summary>
        private static string WriteRowAndGetSheetXml(TempWorkspace workspace, params object[] values)
        {
            var output = workspace.NewFile(Guid.NewGuid().ToString("n") + ".xlsx");
            using (var fastExcel = new FastExcel(TestFixtures.Template, output))
            {
                var worksheet = new Worksheet();
                worksheet.AddRow(values);
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            using var zip = ZipFile.OpenRead(output.FullName);
            using var reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml").Open());
            var xml = reader.ReadToEnd();

            var start = xml.IndexOf("<sheetData>", StringComparison.Ordinal);
            var end = xml.IndexOf("</sheetData>", StringComparison.Ordinal);
            return start < 0 || end < 0 ? xml : xml.Substring(start, end - start + "</sheetData>".Length);
        }

        /// <summary>The text inside the first cell's &lt;v&gt; element.</summary>
        private static string FirstCellValue(string sheetXml)
        {
            var match = Regex.Match(sheetXml, @"<v>(?<v>.*?)</v>", RegexOptions.Singleline);
            Assert.True(match.Success, $"no <v> element found in: {sheetXml}");
            return match.Groups["v"].Value;
        }

        // ------------------------------------------------------- #55 culture independence

        [Fact]
        public void Double_UnderAPointDecimalCulture_UsesADecimalPoint()
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("en-US");

            Assert.Equal("1.2", FirstCellValue(WriteRowAndGetSheetXml(workspace, 1.2d)));
        }

        [Theory]
        [InlineData("de-DE", "1,2")]
        [InlineData("ru-RU", "1,2")]
        [InlineData("fr-FR", "1,2")]
        [InlineData("tr-TR", "1,2")]
        [InlineData("ar-SA", "1٫2")] // Arabic decimal separator, not even a comma
        public void Double_UnderAnyCulture_UsesADecimalPoint(string culture, string whatItActuallyWrites)
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope(culture);

            // The xlsx format is culture-neutral: the stored value must always use '.',
            // regardless of how Excel later displays it to the user.
            KnownBug.StillBroken("#55",
                $"under {culture} the value 1.2 is stored as \"1.2\"; today it is written as " +
                $"\"{whatItActuallyWrites}\", which makes Excel report the workbook as corrupt",
                () => Assert.Equal("1.2", FirstCellValue(WriteRowAndGetSheetXml(workspace, 1.2d))));
        }

        [Theory]
        [InlineData("de-DE")]
        [InlineData("ru-RU")]
        [InlineData("fr-FR")]
        public void Decimal_IsWrittenWithAnInvariantDecimalPoint(string culture)
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope(culture);

            KnownBug.StillBroken("#55",
                "decimal values are formatted with CultureInfo.InvariantCulture, so 3.5 is " +
                "written as \"3.5\" even under a comma-decimal locale",
                () => Assert.Equal("3.5", FirstCellValue(WriteRowAndGetSheetXml(workspace, 3.5m))));
        }

        [Fact]
        public void Double_UnderACommaDecimalLocale_DoesNotEmitAComma()
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("ru-RU");

            // A comma here is what makes Excel declare the whole workbook corrupt.
            KnownBug.StillBroken("#55",
                "no cell value contains a comma as a decimal separator under any culture",
                () => Assert.DoesNotContain("<v>1,2</v>", WriteRowAndGetSheetXml(workspace, 1.2d)));
        }

        // ------------------------------------------------------------ #77 supported types

        [Theory]
        [InlineData(42)]
        [InlineData(-7)]
        [InlineData(0)]
        [InlineData(int.MaxValue)]
        [InlineData(int.MinValue)]
        public void Int_IsWrittenAsANumericCell(int value)
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("en-US");

            var xml = WriteRowAndGetSheetXml(workspace, value);

            Assert.Equal(value.ToString(CultureInfo.InvariantCulture), FirstCellValue(xml));
            Assert.DoesNotContain("t=\"s\"", xml);
        }

        [Fact]
        public void String_IsWrittenAsASharedStringCell()
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("en-US");

            var xml = WriteRowAndGetSheetXml(workspace, "hello");

            Assert.Contains("t=\"s\"", xml);
        }

        [Fact]
        public void Null_ProducesNoCell()
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("en-US");

            var xml = WriteRowAndGetSheetXml(workspace, new object[] { null });

            Assert.DoesNotContain("<c ", xml);
        }

        [Theory]
        [InlineData((long)9_000_000_000)]
        [InlineData((short)12)]
        [InlineData((byte)200)]
        public void OtherIntegerTypes_AreWrittenAsNumericCells(object value)
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("en-US");

            // These happen to round-trip correctly today because their invariant and
            // culture-specific representations are identical, but they are only reaching the
            // numeric branch by accident — the writer never actually recognises the type.
            var xml = WriteRowAndGetSheetXml(workspace, value);

            Assert.Equal(Convert.ToString(value, CultureInfo.InvariantCulture), FirstCellValue(xml));
        }

        [Fact]
        public void Float_IsWrittenAsANumericCell()
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("en-US");

            var xml = WriteRowAndGetSheetXml(workspace, 2.5f);

            Assert.Equal("2.5", FirstCellValue(xml));
        }

        [Fact]
        public void Bool_IsWrittenAsABooleanCell()
        {
            using var workspace = new TempWorkspace();

            KnownBug.StillBroken("#77",
                "a bool is written as t=\"b\" with the value 1 or 0, per the OOXML spec; " +
                "currently it lands as the literal text \"True\" inside a numeric cell",
                () =>
                {
                    var xml = WriteRowAndGetSheetXml(workspace, true);
                    Assert.Contains("t=\"b\"", xml);
                    Assert.Equal("1", FirstCellValue(xml));
                });
        }

        [Fact]
        public void DateTime_IsWrittenAsASerialNumber()
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("en-US");
            var date = new DateTime(2024, 5, 6);

            KnownBug.StillBroken("#77",
                "a DateTime is written as its OLE Automation serial (45418 for 2024-05-06) " +
                "with a date style, not as a formatted date string in a numeric cell",
                () => Assert.Equal(
                    date.ToOADate().ToString(CultureInfo.InvariantCulture),
                    FirstCellValue(WriteRowAndGetSheetXml(workspace, date))));
        }

        [Fact]
        public void DateTime_DoesNotEmitAFormattedDateStringIntoANumericCell()
        {
            using var workspace = new TempWorkspace();
            using var _ = new CultureScope("en-US");

            KnownBug.StillBroken("#77",
                "a numeric cell never contains a human-readable date; \"5/6/2024 12:00:00 AM\" " +
                "in a <v> makes the workbook unreadable",
                () =>
                {
                    var xml = WriteRowAndGetSheetXml(workspace, new DateTime(2024, 5, 6));
                    Assert.DoesNotContain("/", FirstCellValue(xml));
                });
        }

        [Fact]
        public void UnrecognisedType_FallsBackToAStringCell()
        {
            using var workspace = new TempWorkspace();

            KnownBug.StillBroken("#77 / #61",
                "a type the writer does not understand falls back to a shared string cell " +
                "rather than being stringified into a cell declared as numeric",
                () =>
                {
                    var xml = WriteRowAndGetSheetXml(workspace, new Uri("https://example.com"));
                    Assert.Contains("t=\"s\"", xml);
                });
        }

        // ------------------------------------------------------------- #72 cell ordering

        [Fact]
        public void CellsSuppliedOutOfOrder_AreWrittenInAscendingColumnOrder()
        {
            using var workspace = new TempWorkspace();
            var output = workspace.NewFile();

            using (var fastExcel = new FastExcel(TestFixtures.Template, output))
            {
                var worksheet = new Worksheet();
                worksheet.AddRow(); // ensure Rows is initialised
                (worksheet.Rows as List<Row>).Clear();
                (worksheet.Rows as List<Row>).Add(new Row(1, new List<Cell>
                {
                    new Cell(1, "first"),
                    new Cell(3, "third"),
                    new Cell(2, "second"),
                }));
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            using var zip = ZipFile.OpenRead(output.FullName);
            using var reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml").Open());
            var xml = reader.ReadToEnd();

            KnownBug.StillBroken("#72",
                "cells are sorted by column before being written; Excel treats a row whose " +
                "cell references are not ascending as a damaged file",
                () =>
                {
                    var order = Regex.Matches(xml, @"<c r=""(?<ref>[A-Z]+)\d+""")
                                     .Cast<Match>()
                                     .Select(m => m.Groups["ref"].Value)
                                     .ToList();
                    Assert.Equal(new[] { "A", "B", "C" }, order);
                });
        }
    }
}
