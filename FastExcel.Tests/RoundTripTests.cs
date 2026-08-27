using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Write a value, read it back, and see what survives — plus the boundaries of the grid.
    ///
    /// The type half of this mirrors, deliberately, the test a downstream project wrote before
    /// abandoning this library for another one. Their assertion was that a value written as an
    /// <c>int</c>, <c>bool</c>, <c>string</c>, <c>DateTime</c> or <c>decimal</c> comes back as
    /// that same type at a stable position. It is the clearest statement of the contract
    /// consumers expect, so it belongs in this repo whether or not the library meets it today.
    ///
    /// It does not. Everything comes back as <see cref="string"/>, because the reader never
    /// converts and the writer never records what the value was.
    /// </summary>
    public class RoundTripTests
    {
        /// <summary>Writes one row of values into a fresh workbook and reads the row straight back.</summary>
        private static List<object> RoundTrip(params object[] values)
        {
            using var workspace = new TempWorkspace();
            var template = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace, "template.xlsx");
            var output = workspace.NewFile();

            var worksheet = new Worksheet();
            worksheet.AddRow(values);

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            XlsxAssert.IsValidPackage(output);

            using var reopened = new FastExcel(output, true);
            return reopened.Read(1).Rows.Single().Cells.Select(c => c.Value).ToList();
        }

        private static object RoundTripOne(object value) => RoundTrip(value).Single();

        // ------------------------------------------------------------------ what survives

        [Fact]
        public void AString_ComesBackUnchanged()
        {
            Assert.Equal("hello", RoundTripOne("hello"));
        }

        [Theory]
        [InlineData("with spaces")]
        [InlineData("ampersand & angle <")]
        [InlineData("café")]
        [InlineData("你好")]
        public void TextWithAwkwardCharacters_ComesBackUnchanged(string text)
        {
            // Passes because the library escapes and unescapes with the same wrong scheme. The
            // file is still wrong for Excel - SharedStringsTests pins that separately by looking
            // at the bytes. This test only says the library is self-consistent.
            Assert.Equal(text, RoundTripOne(text));
        }

        [Fact]
        public void AnEmptyString_ComesBackAsAnEmptyString()
        {
            Assert.Equal(string.Empty, RoundTripOne(string.Empty));
        }

        // ------------------------------------------------------------------ what does not

        [Fact]
        public void AnInt_ComesBackAsAString()
        {
            var value = RoundTripOne(42);

            Assert.Equal("42", value);

            KnownBug.StillBroken("#58",
                "a value written as an int is read back as an int, so a caller can cast it " +
                "without parsing",
                () => Assert.IsType<int>(value));
        }

        [Fact]
        public void ADouble_ComesBackAsAString()
        {
            var value = RoundTripOne(1.5d);

            Assert.Equal("1.5", value);

            KnownBug.StillBroken("#58",
                "a value written as a double is read back as a double",
                () => Assert.IsType<double>(value));
        }

        [Fact]
        public void ADecimal_ComesBackAsAStringAndLosesItsType()
        {
            var value = RoundTripOne(12345.6789m);

            KnownBug.StillBroken("#58",
                "a value written as a decimal is read back as a decimal with its precision intact",
                () => Assert.Equal(12345.6789m, value));

            Assert.Equal("12345.6789", value);
        }

        [Fact]
        public void ABool_ComesBackAsTheWordTrue()
        {
            // The writer does not recognise bool, so it takes the numeric branch and writes
            // <v>True</v> into a cell Excel expects a number in. The file is still readable by
            // this library, which is why nothing noticed.
            var value = RoundTripOne(true);

            Assert.Equal("True", value);

            KnownBug.StillBroken("#77",
                "a value written as a bool is read back as a bool",
                () => Assert.IsType<bool>(value));
        }

        [Fact]
        public void ADateTime_ComesBackAsAFormattedDateString()
        {
            var when = new DateTime(2024, 5, 6);
            var value = RoundTripOne(when);

            KnownBug.StillBroken("#77",
                "a value written as a DateTime is read back as a DateTime",
                () => Assert.Equal(when, value));

            // What actually happens: the current culture's ToString() lands in a numeric cell.
            // The exact text depends on the machine's culture, so assert the shape, not the text.
            var text = Assert.IsType<string>(value);
            Assert.Contains("2024", text);
        }

        [Fact]
        public void AMixedRow_KeepsEveryColumnInPlace()
        {
            // Position stability matters more than type here: consumers index cells by ordinal,
            // so even while the types are wrong the columns must not shift.
            var values = RoundTrip("text", 1, 2.5d, true, new DateTime(2024, 1, 1));

            Assert.Equal(5, values.Count);
            Assert.Equal("text", values[0]);
            Assert.Equal("1", values[1]);
            Assert.Equal("2.5", values[2]);
            Assert.Equal("True", values[3]);
        }

        [Fact]
        public void ANullInTheMiddle_LeavesAGapWithoutShiftingLaterColumns()
        {
            using var workspace = new TempWorkspace();
            var template = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace, "template.xlsx");
            var output = workspace.NewFile();

            var worksheet = new Worksheet();
            worksheet.AddRow("a", null, "c");

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            using var reopened = new FastExcel(output, true);
            var cells = reopened.Read(1).Rows.Single().Cells.ToList();

            // Only the written cells come back, but each knows its real column - so a caller
            // reading by ColumnNumber is fine, while one reading by ordinal is not (#88).
            Assert.Equal(new[] { 1, 3 }, cells.Select(c => c.ColumnNumber));
            Assert.Equal("c", cells.Last().Value);
        }

        // ------------------------------------------------------------------ grid boundaries

        [Fact]
        public void TheLastColumn_CanBeWrittenAndReadBack()
        {
            // XFD is column 16,384 - the last one Excel has.
            using var workspace = new TempWorkspace();
            var template = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace, "template.xlsx");
            var output = workspace.NewFile();

            var worksheet = new Worksheet
            {
                Rows = new List<Row> { new Row(1, new List<Cell> { new Cell(16384, "last") }) }
            };

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            XlsxAssert.IsValidPackage(output);

            using var reopened = new FastExcel(output, true);
            var cell = reopened.Read(1).Rows.Single().Cells.Single();

            Assert.Equal(16384, cell.ColumnNumber);
            Assert.Equal("XFD", cell.ColumnName);
            Assert.Equal("last", cell.Value);
        }

        [Fact]
        public void AColumnPastTheLastOne_IsAcceptedWithoutComplaint()
        {
            // 16,385 does not exist. Nothing rejects it, so the file is written with a
            // reference Excel cannot parse and the failure surfaces when a user opens it.
            var cell = new Cell(16385, "past the end");

            Assert.Equal("XFE", cell.ColumnName);

            KnownBug.StillBroken("#18",
                "a column number beyond Excel's last column (16,384 / XFD) is rejected when the " +
                "cell is constructed, rather than written into a file Excel cannot open",
                () => Assert.ThrowsAny<ArgumentException>(() => new Cell(16385, "past the end")));
        }

        [Fact]
        public void TheLastRow_CanBeWrittenAndReadBack()
        {
            using var workspace = new TempWorkspace();
            var template = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace, "template.xlsx");
            var output = workspace.NewFile();

            var worksheet = new Worksheet
            {
                Rows = new List<Row> { new Row(1_048_576, new List<Cell> { new Cell(1, "bottom") }) }
            };

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            XlsxAssert.IsValidPackage(output);

            using var reopened = new FastExcel(output, true);
            var row = reopened.Read(1).Rows.Single();

            Assert.Equal(1_048_576, row.RowNumber);
            Assert.Equal("bottom", row.Cells.Single().Value);
        }

        [Fact]
        public void ARowPastTheLastOne_IsAcceptedWithoutComplaint()
        {
            var row = new Row(1_048_577, new List<Cell> { new Cell(1, "past the end") });

            Assert.Equal(1_048_577, row.RowNumber);

            KnownBug.StillBroken("#18",
                "a row number beyond Excel's last row (1,048,576) is rejected when the row is " +
                "constructed, rather than written into a file Excel cannot open",
                () => Assert.ThrowsAny<ArgumentException>(
                    () => new Row(1_048_577, new List<Cell> { new Cell(1, "past the end") })));
        }

        [Fact]
        public void AStringAtTheCellLengthLimit_RoundTrips()
        {
            // 32,767 characters is Excel's maximum for one cell.
            var atTheLimit = new string('x', 32_767);

            Assert.Equal(atTheLimit, RoundTripOne(atTheLimit));
        }

        [Fact]
        public void AStringPastTheCellLengthLimit_IsAcceptedWithoutComplaint()
        {
            var tooLong = new string('x', 32_768);

            // Written and read back happily. Excel truncates or refuses the cell, so the data
            // loss happens in the user's spreadsheet rather than at the call that caused it.
            Assert.Equal(tooLong, RoundTripOne(tooLong));

            KnownBug.StillBroken("#18",
                "a cell value longer than Excel's 32,767-character limit is rejected on write, " +
                "rather than producing a file Excel will silently truncate",
                () => Assert.ThrowsAny<ArgumentException>(() => RoundTripOne(tooLong)));
        }

        // ------------------------------------------------------------------ repeated writes

        [Fact]
        public void WritingTheSameSheetTwice_KeepsTheSecondWrite()
        {
            // Overwriting one sheet through a single instance works, and the later write wins.
            using var workspace = new TempWorkspace();
            var template = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace, "template.xlsx");
            var output = workspace.NewFile();

            var first = new Worksheet();
            first.AddRow("first");
            var second = new Worksheet();
            second.AddRow("second");

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(first, 1);
                fastExcel.Write(second, 1);
            }

            output.Refresh();
            XlsxAssert.IsValidPackage(output);

            using var reopened = new FastExcel(output, true);
            Assert.Equal("second", reopened.Read(1).Rows.Single().Cells.Single().Value);
        }

        [Fact]
        public void WritingTwoDifferentSheetsThroughOneInstance_KeepsBoth()
        {
            // Worth stating plainly, because it contradicts a reasonable reading of #52
            // ("Update multiple worksheets at once"): writing two different sheets through one
            // instance works, and both sheets keep their data.
            //
            // #52's actual symptom is different - it is about the template copy failing with
            // "Could not copy template to output file path" when a second FastExcel instance is
            // opened over the same output. So this is a regression guard for behaviour that
            // works, not a reproduction of that report.
            using var workspace = new TempWorkspace();
            var template = new XlsxBuilder()
                .WithSheet("Sheet1")
                .WithSheet("Sheet2")
                .ToFile(workspace, "template.xlsx");
            var output = workspace.NewFile();

            var first = new Worksheet();
            first.AddRow("into-sheet-one");
            var second = new Worksheet();
            second.AddRow("into-sheet-two");

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(first, 1);
                fastExcel.Write(second, 2);
            }

            output.Refresh();
            XlsxAssert.IsValidPackage(output);

            using var reopened = new FastExcel(output, true);
            var sheetTwo = reopened.Read(2).Rows.ToList();

            Assert.Equal("into-sheet-two", sheetTwo.Single().Cells.Single().Value);

            var sheetOne = reopened.Read(1).Rows.ToList();

            Assert.Equal("into-sheet-one", sheetOne.Single().Cells.Single().Value);
        }
    }
}
