using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Defined names are aliases for a cell, a range, several ranges, or a whole column, and
    /// they may be scoped to a single worksheet. The scoping is where the API goes wrong: the
    /// public methods accept a worksheetIndex and then drop it on the way through.
    /// </summary>
    public class DefinedNameTests
    {
        private const string ThreeRows =
            "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>" +
            "<row r=\"2\"><c r=\"A2\" t=\"s\"><v>1</v></c></row>" +
            "<row r=\"3\"><c r=\"A3\" t=\"s\"><v>2</v></c></row>";

        private static FileInfo SingleSheetWorkbook(TempWorkspace workspace)
            => new XlsxBuilder()
                .WithSharedStrings("one", "two", "three")
                .WithSheet("Sheet1", ThreeRows)
                .WithDefinedName("Totals", "Sheet1!$A$1:$A$3")
                .ToFile(workspace);

        /// <summary>
        /// Opens a workbook for reading with its archive already initialised.
        /// <para>
        /// Calling Read first is a WORKAROUND, not incidental setup: LoadDefinedNames uses
        /// FastExcel.Archive without ever calling PrepareArchive, so on a freshly constructed
        /// instance the archive is still null and every defined-name lookup throws. Read is
        /// what happens to initialise it. See
        /// <see cref="DefinedNameApi_WorksWithoutReadingASheetFirst"/>.
        /// </para>
        /// </summary>
        private static FastExcel OpenForDefinedNames(FileInfo file)
        {
            var fastExcel = new FastExcel(file, true);
            fastExcel.Read(1);
            return fastExcel;
        }

        [Fact]
        public void DefinedNameApi_WorksWithoutReadingASheetFirst()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = new FastExcel(SingleSheetWorkbook(workspace), true);

            KnownBug.StillBroken("#49 / #89",
                "LoadDefinedNames calls PrepareArchive before touching FastExcel.Archive, so " +
                "the defined-name API is usable as the first call on an instance; today it " +
                "throws DefinedNameLoadException wrapping a NullReferenceException unless a " +
                "sheet happens to have been read first",
                () => fastExcel.GetCellsByDefinedName("Totals").ToList());
        }

        // ------------------------------------------------------------------- baseline

        [Fact]
        public void GlobalDefinedName_ResolvesToItsCells()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenForDefinedNames(SingleSheetWorkbook(workspace));

            var cells = fastExcel.GetCellsByDefinedName("Totals").ToList();

            Assert.Equal(new object[] { "one", "two", "three" }, cells.Select(c => c.Value));
        }

        [Fact]
        public void GetCellByDefinedName_ReturnsTheFirstCellOfTheRange()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenForDefinedNames(SingleSheetWorkbook(workspace));

            Assert.Equal("one", fastExcel.GetCellByDefinedName("Totals").Value);
        }

        [Fact]
        public void UnknownDefinedName_ReturnsNoCells()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenForDefinedNames(SingleSheetWorkbook(workspace));

            Assert.Empty(fastExcel.GetCellsByDefinedName("NoSuchName"));
        }

        [Fact]
        public void GetCellRangesByDefinedName_GroupsCellsByRange()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("one", "two", "three")
                .WithSheet("Sheet1", ThreeRows)
                // A single name may cover several comma-separated ranges.
                .WithDefinedName("Ends", "Sheet1!$A$1:$A$1,Sheet1!$A$3:$A$3")
                .ToFile(workspace);

            using var fastExcel = OpenForDefinedNames(file);
            var ranges = fastExcel.GetCellRangesByDefinedName("Ends").ToList();

            Assert.Equal(2, ranges.Count);
            Assert.Equal("one", ranges[0].Single().Value);
            Assert.Equal("three", ranges[1].Single().Value);
        }

        [Fact]
        public void RepeatedLookups_ReturnConsistentResults()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenForDefinedNames(SingleSheetWorkbook(workspace));

            // #86 is about this re-reading and re-parsing the whole worksheet every time. That
            // is a performance defect rather than a correctness one, so this pins the
            // correctness half — repeated calls must at least agree with each other.
            var first = fastExcel.GetCellsByDefinedName("Totals").Select(c => c.Value).ToList();
            var second = fastExcel.GetCellsByDefinedName("Totals").Select(c => c.Value).ToList();
            var third = fastExcel.GetCellByDefinedName("Totals").Value;

            Assert.Equal(first, second);
            Assert.Equal(first.First(), third);
        }

        // --------------------------------------------------- #84 sheet-scoped defined names

        /// <summary>
        /// Two sheets each declare a name spelled "Region", scoped to themselves — legal in
        /// Excel, and the reason the API takes a worksheetIndex at all.
        /// </summary>
        private static FileInfo ScopedNameWorkbook(TempWorkspace workspace)
            => new XlsxBuilder()
                .WithSharedStrings("north", "south")
                .WithSheet("Sheet1", "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>")
                .WithSheet("Sheet2", "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>1</v></c></row>")
                .WithDefinedName("Region", "Sheet1!$A$1:$A$1", scopedToSheetIndex: 1)
                .WithDefinedName("Region", "Sheet2!$A$1:$A$1", scopedToSheetIndex: 2)
                .ToFile(workspace);

        [Fact]
        public void GetCellsByDefinedName_HonoursTheWorksheetIndex()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenForDefinedNames(ScopedNameWorkbook(workspace));

            KnownBug.StillBroken("#84",
                "GetCellsByDefinedName forwards its worksheetIndex argument to " +
                "GetCellRangesByDefinedName, so asking for \"Region\" on sheet 2 returns " +
                "sheet 2's cell; today the argument is silently discarded",
                () =>
                {
                    Assert.Equal("north", fastExcel.GetCellsByDefinedName("Region", 1).Single().Value);
                    Assert.Equal("south", fastExcel.GetCellsByDefinedName("Region", 2).Single().Value);
                });
        }

        [Fact]
        public void GetCellByDefinedName_HonoursTheWorksheetIndex()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenForDefinedNames(ScopedNameWorkbook(workspace));

            KnownBug.StillBroken("#84",
                "GetCellByDefinedName resolves a sheet-scoped name against the requested sheet",
                () => Assert.Equal("south", fastExcel.GetCellByDefinedName("Region", 2).Value));
        }

        [Fact]
        public void GetCellRangesByDefinedName_HonoursTheWorksheetIndex()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenForDefinedNames(ScopedNameWorkbook(workspace));

            // This overload does thread the index through, so scoped lookups work here. It is
            // the two convenience wrappers above that drop it.
            var sheet2 = fastExcel.GetCellRangesByDefinedName("Region", 2).ToList();

            Assert.Equal("south", sheet2.Single().Single().Value);
        }

        // ------------------------------------------------------------- #49 column aliases

        [Fact]
        public void GetCellsByColumnName_ReturnsTheCellsOfANamedColumn()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("one", "two", "three")
                .WithSheet("Sheet1", ThreeRows)
                .WithDefinedName("AllColumns", "Sheet1!$A:$A")
                .ToFile(workspace);

            using var fastExcel = OpenForDefinedNames(file);

            var cells = fastExcel.GetCellsByColumnName("AllColumns").ToList();

            Assert.Equal(new object[] { "one", "two", "three" }, cells.Select(c => c.Value));
        }

        [Fact]
        public void GetCellsByColumnName_RespectsARowRange()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("one", "two", "three")
                .WithSheet("Sheet1", ThreeRows)
                .WithDefinedName("AllColumns", "Sheet1!$A:$A")
                .ToFile(workspace);

            using var fastExcel = OpenForDefinedNames(file);

            var cells = fastExcel.GetCellsByColumnName("AllColumns", rowStart: 2, rowEnd: 3).ToList();

            Assert.Equal(new object[] { "two", "three" }, cells.Select(c => c.Value));
        }

        [Fact]
        public void GetCellsByColumnName_WithAnUnknownName_DoesNotThrow()
        {
            using var workspace = new TempWorkspace();
            using var fastExcel = OpenForDefinedNames(SingleSheetWorkbook(workspace));

            KnownBug.StillBroken("#49",
                "an unknown column name yields an empty sequence; today the method calls " +
                "Last() on the empty result and throws InvalidOperationException",
                () => Assert.Empty(fastExcel.GetCellsByColumnName("NoSuchColumn")));
        }

        [Fact]
        public void MalformedReference_IsSkippedRatherThanThrowing()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("one")
                .WithSheet("Sheet1", "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c></row>")
                .WithDefinedName("Broken", "#REF!")
                .ToFile(workspace);

            using var fastExcel = OpenForDefinedNames(file);

            // A deleted range leaves #REF! behind in real workbooks, so this must not be fatal.
            Assert.Empty(fastExcel.GetCellsByDefinedName("Broken"));
        }
    }
}
