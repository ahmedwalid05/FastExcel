using System;
using System.Collections.Generic;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// The small public types and properties that had no coverage: the column-name attribute,
    /// the worksheet-properties record, a cell's identity properties, and the single-cell branch
    /// of reference parsing.
    ///
    /// Small does not mean unimportant. <c>Cell.CellName</c> throws for every cell a caller
    /// constructs by hand, which is the only way a caller ever constructs one.
    /// </summary>
    public class SmallSurfaceTests
    {
        // --------------------------------------------------------- ExcelColumnAttribute

        private class Labelled
        {
            [ExcelColumn(Name = "Product name")]
            public string Product { get; set; }

            [ExcelColumn(Name = "Units sold")]
            public int Units { get; set; }

            public string Unlabelled { get; set; }
        }

        [Fact]
        public void ExcelColumnAttribute_RenamesTheHeading()
        {
            var worksheet = new Worksheet();
            worksheet.PopulateRows(
                new[] { new Labelled { Product = "widget", Units = 2, Unlabelled = "x" } },
                existingHeadingRows: 0,
                usePropertiesAsHeadings: true);

            Assert.Equal(new[] { "Product name", "Units sold", "Unlabelled" }, worksheet.Headings.ToArray());
        }

        [Fact]
        public void ExcelColumnAttribute_NameDefaultsToNull()
        {
            Assert.Null(new ExcelColumnAttribute().Name);
        }

        [Fact]
        public void ExcelColumnAttribute_OnAField_IsIgnored()
        {
            // The attribute declares AttributeTargets.Field, so this compiles and a reader would
            // reasonably expect it to work - but headings are only ever read from properties, so
            // the field spelling is silently inert. Either the target should be removed or the
            // lookup should honour it.
            var usage = (AttributeUsageAttribute)Attribute.GetCustomAttribute(
                typeof(ExcelColumnAttribute), typeof(AttributeUsageAttribute));

            Assert.True(usage.ValidOn.HasFlag(AttributeTargets.Field));

            KnownBug.StillBroken("#18",
                "ExcelColumnAttribute is only valid where it is honoured; today it advertises " +
                "AttributeTargets.Field while headings are read from properties alone",
                () => Assert.False(usage.ValidOn.HasFlag(AttributeTargets.Field)));
        }

        // --------------------------------------------------------- WorksheetProperties

        [Fact]
        public void WorksheetProperties_RoundTripsItsValues()
        {
            var properties = new WorksheetProperties
            {
                Name = "Quarterly",
                SheetId = 4,
                CurrentIndex = 2
            };

            Assert.Equal("Quarterly", properties.Name);
            Assert.Equal(4, properties.SheetId);
            Assert.Equal(2, properties.CurrentIndex);
        }

        [Fact]
        public void WorksheetProperties_IsPublicButNoPublicMemberEverReturnsOne()
        {
            // Public API surface that a consumer can name but never receive. It exists for the
            // add-worksheet machinery, which is unreachable - see PublicApiTests.
            var returnedAnywhere = typeof(FastExcel).Assembly.GetExportedTypes()
                .SelectMany(t => t.GetMembers())
                .Any(m => (m as System.Reflection.MethodInfo)?.ReturnType == typeof(WorksheetProperties)
                       || (m as System.Reflection.PropertyInfo)?.PropertyType == typeof(WorksheetProperties));

            Assert.False(returnedAnywhere);
        }

        // --------------------------------------------------------- Cell identity

        [Fact]
        public void ConstructedCell_KnowsItsColumnLetter()
        {
            Assert.Equal("A", new Cell(1, "x").ColumnName);
            Assert.Equal("Z", new Cell(26, "x").ColumnName);
            Assert.Equal("AA", new Cell(27, "x").ColumnName);
        }

        [Fact]
        public void ConstructedCell_HasNoUnderlyingElement()
        {
            Assert.Null(new Cell(1, "x").XElement);
        }

        [Fact]
        public void ConstructedCell_CellName_Throws()
        {
            var cell = new Cell(1, "x");

            // CellNames is only assigned when a cell is parsed from a file, so it is null here,
            // and CellName calls .Any() on it. Every cell a caller builds by hand hits this.
            KnownBug.StillBroken("#18",
                "CellName on a constructed cell returns its address (\"A0\" or similar) rather " +
                "than throwing a NullReferenceException from an unassigned CellNames",
                () => Assert.NotNull(cell.CellName));
        }

        [Fact]
        public void ConstructedCell_ToStringIsItsValue()
        {
            Assert.Equal("hello", new Cell(1, "hello").ToString());
        }

        [Fact]
        public void ConstructedCell_WithANullValue_ToStringThrows()
        {
            var cell = new Cell(1, null);

            KnownBug.StillBroken("#18",
                "ToString on a cell with no value returns an empty string rather than throwing",
                () => Assert.NotNull(cell.ToString()));
        }

        [Fact]
        public void ParsedCell_ExposesItsUnderlyingElementAndAddress()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("value")
                .WithSheet("Sheet1", XlsxBuilder.Row(3, XlsxBuilder.SharedCell("B3", 0)))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);
            var cell = fastExcel.Read(1).Rows.Single().Cells.Single();

            Assert.NotNull(cell.XElement);
            Assert.Equal(2, cell.ColumnNumber);
            Assert.Equal("B", cell.ColumnName);
            Assert.Equal(3, cell.RowNumber);
            Assert.Equal("B3", cell.CellName);
            Assert.Empty(cell.CellNames);
        }

        [Fact]
        public void ParsedCell_InsideADefinedName_ReportsThatNameAsItsCellName()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("value")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .WithDefinedName("Total", "Sheet1!$A$1")
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);
            var cell = fastExcel.Read(1).Rows.Single().Cells.Single();

            Assert.Equal(new[] { "Total" }, cell.CellNames.ToArray());
            Assert.Equal("Total", cell.CellName);

            // ColumnName is *not* affected: it is resolved from whole-column names like
            // "Sheet1!$A:$A", not from a name pointing at a single cell. So a cell can be called
            // "Total" while still living in column "A".
            Assert.Equal("A", cell.ColumnName);
        }

        [Fact]
        public void ParsedCell_InAWholeColumnDefinedName_ReportsThatNameAsItsColumnName()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("value")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .WithDefinedName("Amount", "Sheet1!$A:$A")
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);
            var cell = fastExcel.Read(1).Rows.Single().Cells.Single();

            Assert.Equal("Amount", cell.ColumnName);
        }

        [Fact]
        public void Cell_Merge_TakesTheOtherCellsValue()
        {
            var target = new Cell(1, "original");
            target.Merge(new Cell(1, "incoming"));

            // Documents the direction, which is the mechanism behind the Update precedence bug:
            // Merge copies *from* the argument, and Worksheet.MergeRows calls it with the
            // arguments the other way round from what the doc comment promises.
            Assert.Equal("incoming", target.Value);
        }

        // --------------------------------------------------------- single-cell references

        [Fact]
        public void DefinedNameForASingleCell_ResolvesToThatOneCell()
        {
            // Exercises the non-range branch of reference parsing: "Sheet1!$B$2" has no colon,
            // so start and end collapse to the same cell.
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("a", "b", "c", "d")
                .WithSheet("Sheet1",
                    XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0), XlsxBuilder.SharedCell("B1", 1)) +
                    XlsxBuilder.Row(2, XlsxBuilder.SharedCell("A2", 2), XlsxBuilder.SharedCell("B2", 3)))
                .WithDefinedName("Corner", "Sheet1!$B$2")
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);

            // Reading a sheet first is not incidental: the defined-name API does not prepare the
            // archive itself, so calling it first throws (#49 / #89). DefinedNameTests pins that
            // separately; here we just need past it.
            _ = fastExcel.Read(1);

            var cells = fastExcel.GetCellsByDefinedName("Corner").ToList();

            var only = Assert.Single(cells);
            Assert.Equal("d", only.Value);
        }

        [Fact]
        public void CellRange_PublicConstructor_KeepsWhatItIsGiven()
        {
            var range = new CellRange("B", "D", 2, 7);

            Assert.Equal("B", range.ColumnStart);
            Assert.Equal("D", range.ColumnEnd);
            Assert.Equal(2, range.RowStart);
            Assert.Equal(7, range.RowEnd);
        }

        [Fact]
        public void CellRange_PublicConstructor_ValidatesNothing()
        {
            // Reversed columns, a row start below one and an end before the start are all
            // accepted without complaint; the mistake only shows up later as an empty result.
            var nonsense = new CellRange("Z", "A", -5, -9);

            Assert.Equal(-5, nonsense.RowStart);
            Assert.Equal(-9, nonsense.RowEnd);

            KnownBug.StillBroken("#18",
                "CellRange rejects a reversed column pair and a row number below one at " +
                "construction, rather than silently matching nothing much later",
                () => Assert.ThrowsAny<ArgumentException>(() => new CellRange("Z", "A", -5, -9)));
        }

        // --------------------------------------------------------- remaining Write overload

        [Fact]
        public void WriteGeneric_BySheetNumber_WithHeadings_ProducesAValidFile()
        {
            using var workspace = new TempWorkspace();
            var template = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace, "template.xlsx");
            var output = workspace.NewFile();

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(new[] { new Labelled { Product = "widget", Units = 2 } }, 1,
                    usePropertiesAsHeadings: true);
            }

            output.Refresh();
            XlsxAssert.IsValidPackage(output);

            using var reopened = new FastExcel(output, true);
            Assert.Equal("Product name", reopened.Read(1).Rows.First().Cells.First().Value);
        }
    }
}
