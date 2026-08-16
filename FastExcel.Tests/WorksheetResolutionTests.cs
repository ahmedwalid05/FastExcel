using System;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// How a sheet name or number becomes a package part.
    /// <para>
    /// The library derives the part name arithmetically — the Nth &lt;sheet&gt; element is
    /// assumed to live in xl/worksheets/sheet{N}.xml. That holds for a freshly created
    /// workbook and stops holding as soon as a sheet is deleted or reordered, because Excel
    /// keeps the original part names and only rewrites the r:id relationships. The correct
    /// resolution is r:id → xl/_rels/workbook.xml.rels → Target.
    /// </para>
    /// </summary>
    public class WorksheetResolutionTests
    {
        // -------------------------------------------------------------------- baseline

        [Fact]
        public void SheetCanBeReadByIndex()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("one", "two")
                .WithSheet("First", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .WithSheet("Second", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 1)))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.Equal("one", fastExcel.Read(1).Rows.Single().Cells.Single().Value);
        }

        [Fact]
        public void SheetCanBeReadByName()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("one", "two")
                .WithSheet("First", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .WithSheet("Second", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 1)))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.Equal("two", fastExcel.Read("Second").Rows.Single().Cells.Single().Value);
        }

        [Fact]
        public void SheetNameMatchingIsCaseInsensitive()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("value")
                .WithSheet("MySheet", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.Equal("value", fastExcel.Read("mysheet").Rows.Single().Cells.Single().Value);
        }

        [Fact]
        public void SheetNameIsExposedOnTheWorksheet()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSheet("Quarterly Results")
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.Equal("Quarterly Results", fastExcel.Read(1).Name);
        }

        [Fact]
        public void UnknownSheetName_ThrowsWithAUsefulMessage()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            var exception = Assert.Throws<Exception>(() => fastExcel.Read("NoSuchSheet"));
            Assert.Contains("NoSuchSheet", exception.Message);
        }

        [Fact]
        public void OutOfRangeSheetNumber_ThrowsWithAUsefulMessage()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            var exception = Assert.Throws<Exception>(() => fastExcel.Read(5));
            Assert.Contains("5", exception.Message);
        }

        // ------------------------------------------- #82 / PR #62 part-name resolution

        [Fact]
        public void SheetWhosePartIsNotNamedAfterItsPosition_IsResolvedViaItsRelationship()
        {
            using var workspace = new TempWorkspace();
            // One sheet, declared first, but stored as sheet3.xml — which is exactly what a
            // workbook looks like after its first two sheets have been deleted.
            var file = new XlsxBuilder()
                .WithSharedStrings("survivor")
                .WithSheet("Remaining", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)),
                           partName: "worksheets/sheet3.xml")
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            KnownBug.StillBroken("#82 / PR #62",
                "the worksheet part is located by following the sheet's r:id through " +
                "xl/_rels/workbook.xml.rels, rather than assuming the Nth sheet lives in " +
                "sheet{N}.xml — the assumption breaks for any workbook whose sheets have been " +
                "deleted or reordered, which is most real-world files",
                () => Assert.Equal("survivor", fastExcel.Read(1).Rows.Single().Cells.Single().Value));
        }

        [Fact]
        public void ReorderedSheets_ReadTheContentTheirNamesPromise()
        {
            using var workspace = new TempWorkspace();
            // Declaration order is Beta, Alpha; storage order is the original sheet1, sheet2.
            // Reading "Alpha" must return Alpha's content, not whatever sits in position 2.
            var file = new XlsxBuilder()
                .WithSharedStrings("alpha content", "beta content")
                .WithSheet("Beta", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 1)),
                           partName: "worksheets/sheet2.xml")
                .WithSheet("Alpha", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)),
                           partName: "worksheets/sheet1.xml")
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            KnownBug.StillBroken("#82",
                "a sheet's content is found through its relationship, so reading \"Alpha\" " +
                "returns Alpha's cells even when the sheets were reordered in the workbook; " +
                "today position and part name are conflated and the sheets are swapped",
                () => Assert.Equal("alpha content", fastExcel.Read("Alpha").Rows.Single().Cells.Single().Value));
        }

        [Fact]
        public void MultipleSheets_AreIndependentlyReadable()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("a", "b", "c")
                .WithSheet("One", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .WithSheet("Two", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 1)))
                .WithSheet("Three", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 2)))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            Assert.Equal("a", fastExcel.Read(1).Rows.Single().Cells.Single().Value);
            Assert.Equal("b", fastExcel.Read(2).Rows.Single().Cells.Single().Value);
            Assert.Equal("c", fastExcel.Read(3).Rows.Single().Cells.Single().Value);
        }
    }
}
