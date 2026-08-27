using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Finding a sheet, and the eight <c>Write</c> overloads.
    ///
    /// <c>FastExcel.Worksheets</c> had no coverage at all, which is worth more than it sounds:
    /// merely touching that property changes how a later <c>Read</c> of a missing sheet fails,
    /// because the two code paths report the failure differently. A caller that enumerates the
    /// workbook first gets a <see cref="NullReferenceException"/>; one that does not gets a
    /// message naming the sheet.
    ///
    /// Every write test here ends at <see cref="XlsxAssert.IsValidPackage"/>. Asserting the value
    /// came back is not enough — this library reads its own output happily whether or not anyone
    /// else can.
    /// </summary>
    public class WorkbookNavigationTests
    {
        private static XlsxBuilder ThreeSheets() =>
            new XlsxBuilder()
                .WithSharedStrings("alpha", "beta", "gamma")
                .WithSheet("First", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .WithSheet("Second", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 1)))
                .WithSheet("Third", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 2)));

        // ------------------------------------------------------------- Worksheets

        [Fact]
        public void Worksheets_ListsEverySheetWithItsNameAndPosition()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.Equal(new[] { "First", "Second", "Third" }, fastExcel.Worksheets.Select(w => w.Name));
            Assert.Equal(new[] { 1, 2, 3 }, fastExcel.Worksheets.Select(w => w.Index));
        }

        [Fact]
        public void Worksheets_IsLoadedOnceAndReused()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.Same(fastExcel.Worksheets, fastExcel.Worksheets);
        }

        [Fact]
        public void Worksheets_HandsOutTheLiveArraySoACallerCanCorruptIt()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            fastExcel.Worksheets[0] = null;

            // The property returns the cached array itself rather than a copy, so anything a
            // caller does to it is permanent for the lifetime of the instance. Reading sheet 1
            // afterwards then fails inside the library rather than at the point of the mistake.
            KnownBug.StillBroken("#124",
                "Worksheets hands back a copy, so a caller cannot corrupt the workbook's own " +
                "view of its sheets",
                () => Assert.NotNull(fastExcel.Worksheets[0]));
        }

        [Fact]
        public void Worksheets_OnAWorkbookWithOneSheet_ReturnsThatSheet()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder().WithSheet("Only").ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            var only = Assert.Single(fastExcel.Worksheets);
            Assert.Equal("Only", only.Name);
            Assert.Equal(1, only.Index);
        }

        // ------------------------------------------------------------- GetWorksheetIndexFromName

        [Fact]
        public void GetWorksheetIndexFromName_ReturnsThePosition()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.Equal(2, fastExcel.GetWorksheetIndexFromName("Second"));
        }

        [Fact]
        public void GetWorksheetIndexFromName_ForAnUnknownSheet_ReturnsZero()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            // Zero is not a sentinel a caller can distinguish from a real answer without knowing
            // that sheet indexes are 1-based, and nothing in the signature says so. Feeding the
            // result straight back into Read(int) produces a confusing failure much later.
            KnownBug.StillBroken("#123",
                "an unknown sheet name is reported distinguishably - either a negative sentinel " +
                "or an exception - rather than as 0, which reads like a valid index",
                () => Assert.NotEqual(0, fastExcel.GetWorksheetIndexFromName("Nope")));
        }

        [Fact]
        public void GetWorksheetIndexFromName_IsCaseSensitive_WhileReadIsNot()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            // Read resolves sheet names case-insensitively; this method compares with ==. The
            // same string therefore finds a sheet through one entry point and not the other.
            Assert.Equal("beta", fastExcel.Read("second").Rows.First().Cells.First().Value);

            KnownBug.StillBroken("#123",
                "sheet names are matched the same way everywhere; today Read is " +
                "case-insensitive and GetWorksheetIndexFromName is not",
                () => Assert.Equal(2, fastExcel.GetWorksheetIndexFromName("second")));
        }

        // ------------------------------------------------------------- missing sheets

        [Fact]
        public void ReadingAMissingSheetByName_ExplainsItself()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            var thrown = Record.Exception(() => fastExcel.Read("Nope"));

            Assert.NotNull(thrown);
            Assert.Contains("Nope", thrown.Message);
        }

        [Fact]
        public void ReadingAMissingSheet_FailsDifferentlyDependingOnWhatYouTouchedFirst()
        {
            // Read has two branches. Before Worksheets has been loaded it delegates to
            // Worksheet.Read, which throws a message naming the sheet. Afterwards it runs a LINQ
            // query, gets null from SingleOrDefault, and dereferences it - so the identical call
            // produces a NullReferenceException instead. Whether a caller sees a useful error
            // depends on whether they happened to enumerate the workbook earlier.
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            _ = fastExcel.Worksheets;  // the only difference from the test above

            var thrown = Record.Exception(() => fastExcel.Read("Nope"));

            Assert.NotNull(thrown);
            KnownBug.StillBroken("#119",
                "a missing sheet reports the same useful error whether or not Worksheets was " +
                "read first, instead of degrading to a NullReferenceException",
                () => Assert.IsNotType<NullReferenceException>(thrown));
        }

        [Fact]
        public void ReadingAnOutOfRangeSheetNumber_ExplainsItself()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            var thrown = Record.Exception(() => fastExcel.Read(9));

            Assert.NotNull(thrown);
            Assert.Contains("9", thrown.Message);
        }

        [Fact]
        public void ReadingSheetNumberZero_ExplainsItself()
        {
            using var workspace = new TempWorkspace();
            var file = ThreeSheets().ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            var thrown = Record.Exception(() => fastExcel.Read(0));

            Assert.NotNull(thrown);

            // Sheet numbers are 1-based, so 0 is out of range - but it passes the upper-bound
            // guard and reaches an index of -1, so the caller gets an argument exception from
            // deep inside a list rather than the message the other out-of-range cases produce.
            KnownBug.StillBroken("#123",
                "sheet number 0 is rejected with the same message as any other out-of-range " +
                "sheet, rather than surfacing an index error from inside the library",
                () => Assert.IsNotType<ArgumentOutOfRangeException>(thrown));
        }

        // ------------------------------------------------------------- the Write overloads

        /// <summary>A template with one empty sheet, ready to be written into.</summary>
        private static FastExcel OpenAgainstTemplate(TempWorkspace workspace, out System.IO.FileInfo output,
            string sheetName = "Sheet1")
        {
            var template = new XlsxBuilder().WithSheet(sheetName).ToFile(workspace, "template.xlsx");
            output = workspace.NewFile();
            return new FastExcel(template, output);
        }

        private static Worksheet OneRow(string value)
        {
            var worksheet = new Worksheet();
            worksheet.AddRow(value);
            return worksheet;
        }

        private static object FirstValue(System.IO.FileInfo file)
        {
            using var fastExcel = new FastExcel(file, true);
            return fastExcel.Read(1).Rows.First().Cells.First().Value;
        }

        [Fact]
        public void Write_BySheetNumber_ProducesAValidFile()
        {
            using var workspace = new TempWorkspace();
            using (var fastExcel = OpenAgainstTemplate(workspace, out var output))
            {
                fastExcel.Write(OneRow("by-number"), 1);
            }

            var written = workspace.Path("out.xlsx");
            var file = new System.IO.FileInfo(written);
            XlsxAssert.IsValidPackage(file);
            Assert.Equal("by-number", FirstValue(file));
        }

        [Fact]
        public void Write_BySheetName_ProducesAValidFile()
        {
            using var workspace = new TempWorkspace();
            using (var fastExcel = OpenAgainstTemplate(workspace, out var output))
            {
                fastExcel.Write(OneRow("by-name"), "Sheet1");
            }

            var file = new System.IO.FileInfo(workspace.Path("out.xlsx"));
            XlsxAssert.IsValidPackage(file);
            Assert.Equal("by-name", FirstValue(file));
        }

        [Fact]
        public void Write_WithNoSheetSpecified_Throws()
        {
            // The parameterless overload passes null for both the sheet number and the sheet
            // name, so there is nothing to resolve. It is unusable as published.
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenAgainstTemplate(workspace, out _);

            Assert.ThrowsAny<Exception>(() => fastExcel.Write(OneRow("nowhere")));
        }

        private class Line
        {
            public string Label { get; set; }
            public int Count { get; set; }
        }

        [Fact]
        public void WriteGeneric_BySheetName_ProducesAValidFile()
        {
            using var workspace = new TempWorkspace();
            using (var fastExcel = OpenAgainstTemplate(workspace, out _))
            {
                fastExcel.Write(new[] { new Line { Label = "a", Count = 1 } }, "Sheet1");
            }

            var file = new System.IO.FileInfo(workspace.Path("out.xlsx"));
            XlsxAssert.IsValidPackage(file);
            Assert.Equal("a", FirstValue(file));
        }

        [Fact]
        public void WriteGeneric_BySheetNumber_ProducesAValidFile()
        {
            using var workspace = new TempWorkspace();
            using (var fastExcel = OpenAgainstTemplate(workspace, out _))
            {
                fastExcel.Write(new[] { new Line { Label = "a", Count = 1 } }, 1);
            }

            var file = new System.IO.FileInfo(workspace.Path("out.xlsx"));
            XlsxAssert.IsValidPackage(file);
            Assert.Equal("a", FirstValue(file));
        }

        [Fact]
        public void WriteGeneric_WithHeadings_PutsThePropertyNamesInRowOne()
        {
            using var workspace = new TempWorkspace();
            using (var fastExcel = OpenAgainstTemplate(workspace, out _))
            {
                fastExcel.Write(new[] { new Line { Label = "a", Count = 1 } }, "Sheet1",
                    usePropertiesAsHeadings: true);
            }

            var file = new System.IO.FileInfo(workspace.Path("out.xlsx"));
            XlsxAssert.IsValidPackage(file);
            Assert.Equal("Label", FirstValue(file));
        }

        [Fact]
        public void WriteDataTable_ProducesAValidFile()
        {
            // The path the CDC runs in production, and until now the only test of it was that
            // the library could read its own output back.
            var table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Rows.Add("from-a-datatable");

            using var workspace = new TempWorkspace();
            using (var fastExcel = OpenAgainstTemplate(workspace, out _))
            {
                fastExcel.Write(table, "Sheet1");
            }

            var file = new System.IO.FileInfo(workspace.Path("out.xlsx"));
            XlsxAssert.IsValidPackage(file);

            using var reopened = new FastExcel(file, true);
            var rows = reopened.Read(1).Rows.ToList();

            Assert.Equal(2, rows.Count);
            Assert.Equal("Name", rows[0].Cells.First().Value);
            Assert.Equal("from-a-datatable", rows[1].Cells.First().Value);
        }

        [Fact]
        public void Write_ToAMissingSheet_Throws()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenAgainstTemplate(workspace, out _);

            var thrown = Record.Exception(() => fastExcel.Write(OneRow("x"), "NoSuchSheet"));

            Assert.NotNull(thrown);
            Assert.Contains("NoSuchSheet", thrown.Message);
        }

        [Fact]
        public void Write_OnAReadOnlyInstance_Throws()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.ThrowsAny<Exception>(() => fastExcel.Write(OneRow("x"), 1));
        }

        [Fact]
        public void Write_WhereDataWouldLandOnAHeadingRow_Throws()
        {
            // A guard worth keeping: writing row 1 while claiming row 1 is a template header
            // would silently destroy the header, so the library refuses instead.
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenAgainstTemplate(workspace, out _);

            var thrown = Record.Exception(() => fastExcel.Write(OneRow("collides"), 1, existingHeadingRows: 1));

            Assert.NotNull(thrown);
            Assert.Contains("Heading", thrown.Message);
        }
    }
}
