using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// The half of <see cref="Worksheet"/> that builds a sheet in memory before it is written:
    /// <c>PopulateRowsFromDataTable</c>, <c>PopulateRows</c>, <c>AddRow</c>, <c>AddValue</c> and
    /// <c>GetCellsInRange</c>.
    ///
    /// Until now none of it had a single test, which matters more than the coverage number
    /// suggests: the DataTable path is the one running in production at the CDC, and
    /// <c>AddRow</c>/<c>AddValue</c> are the API the README teaches first.
    ///
    /// Several tests here pin behaviour that is arguably wrong rather than asserting what a
    /// caller would want. Where that is the case the test says so and explains what a caller
    /// actually gets, because pinning it is what stops it changing by accident before anyone has
    /// decided to change it deliberately.
    /// </summary>
    public class WorksheetAuthoringTests
    {
        private static DataTable TwoColumnTable()
        {
            var table = new DataTable();
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Amount", typeof(int));
            table.Rows.Add("first", 1);
            table.Rows.Add("second", 2);
            return table;
        }

        private static List<Row> Materialise(Worksheet worksheet) => worksheet.Rows.ToList();

        private static object[] ValuesOf(Row row) => row.Cells.Select(c => c.Value).ToArray();

        // ------------------------------------------------- PopulateRowsFromDataTable

        [Fact]
        public void DataTable_BecomesAHeadingRowFollowedByTheData()
        {
            var worksheet = new Worksheet();
            worksheet.PopulateRowsFromDataTable(TwoColumnTable());

            var rows = Materialise(worksheet);

            Assert.Equal(3, rows.Count);
            Assert.Equal(new[] { 1, 2, 3 }, rows.Select(r => r.RowNumber));
            Assert.Equal(new object[] { "Name", "Amount" }, ValuesOf(rows[0]));
            Assert.Equal(new object[] { "first", 1 }, ValuesOf(rows[1]));
            Assert.Equal(new object[] { "second", 2 }, ValuesOf(rows[2]));
        }

        [Fact]
        public void DataTable_WithExistingHeadingRows_StartsBelowThem()
        {
            var worksheet = new Worksheet();
            worksheet.PopulateRowsFromDataTable(TwoColumnTable(), existingHeadingRows: 3);

            var rows = Materialise(worksheet);

            // Row numbering correctly leaves room for the template's own header.
            Assert.Equal(new[] { 4, 5, 6 }, rows.Select(r => r.RowNumber));
        }

        [Fact]
        public void DataTable_WithExistingHeadingRows_StillEmitsItsOwnHeadingRow()
        {
            var worksheet = new Worksheet();
            worksheet.PopulateRowsFromDataTable(TwoColumnTable(), existingHeadingRows: 3);

            var rows = Materialise(worksheet);

            // Pinning a wart. `existingHeadingRows` says "the template already has headers", yet a
            // second header is emitted anyway - so the caller gets the column names twice, once
            // from the template and once at row 4. Callers work around it by trimming the table.
            Assert.Equal(new object[] { "Name", "Amount" }, ValuesOf(rows[0]));
        }

        [Fact]
        public void DataTable_WithNoRows_StillEmitsTheHeadingRow()
        {
            var table = new DataTable();
            table.Columns.Add("Only", typeof(string));

            var worksheet = new Worksheet();
            worksheet.PopulateRowsFromDataTable(table);

            var rows = Materialise(worksheet);

            Assert.Single(rows);
            Assert.Equal(new object[] { "Only" }, ValuesOf(rows[0]));
        }

        [Fact]
        public void DataTable_WithNoColumns_ProducesAnEmptyHeadingRow()
        {
            var worksheet = new Worksheet();
            worksheet.PopulateRowsFromDataTable(new DataTable());

            var rows = Materialise(worksheet);

            Assert.Single(rows);
            Assert.Empty(rows[0].Cells);
        }

        [Fact]
        public void DataTable_EmptyCell_BecomesDbNullRatherThanBeingSkipped()
        {
            var table = new DataTable();
            table.Columns.Add("A", typeof(string));
            table.Columns.Add("B", typeof(string));
            table.Rows.Add("left", null);   // ADO.NET stores this as DBNull.Value, not null

            var worksheet = new Worksheet();
            worksheet.PopulateRowsFromDataTable(table);

            var dataRow = Materialise(worksheet)[1];
            var cells = dataRow.Cells.ToList();

            // The null guard in the loop tests `value == null`, and DBNull.Value is not null, so
            // the cell survives carrying DBNull. That matters downstream: the writer only
            // understands int, double and string, so a DBNull is formatted with ToString() - an
            // empty string - into a numeric <v> element.
            Assert.Equal(2, cells.Count);
            Assert.Equal("left", cells[0].Value);
            Assert.Equal(DBNull.Value, cells[1].Value);
        }

        [Fact]
        public void DataTable_ColumnsKeepTheirPositionEvenWhenAValueIsMissing()
        {
            var table = new DataTable();
            table.Columns.Add("A", typeof(string));
            table.Columns.Add("B", typeof(string));
            table.Columns.Add("C", typeof(string));
            table.Rows.Add("a", null, "c");

            var worksheet = new Worksheet();
            worksheet.PopulateRowsFromDataTable(table);

            var cells = Materialise(worksheet)[1].Cells.ToList();

            // Column numbers come from the table's column index, so unlike the read path
            // (#88/#83) a missing value here does not shift the columns after it.
            Assert.Equal(new[] { 1, 2, 3 }, cells.Select(c => c.ColumnNumber));
            Assert.Equal("c", cells[2].Value);
        }

        [Fact]
        public void DataTable_Null_Throws()
        {
            var worksheet = new Worksheet();
            Assert.ThrowsAny<Exception>(() => worksheet.PopulateRowsFromDataTable(null));
        }

        // ------------------------------------------------- PopulateRows<T>

        private class Sale
        {
            public string Product { get; set; }
            public int Units { get; set; }
        }

        [Fact]
        public void PopulateRows_FromObjects_UsesPublicPropertiesAsColumns()
        {
            var worksheet = new Worksheet();
            worksheet.PopulateRows(new[]
            {
                new Sale { Product = "widget", Units = 3 },
                new Sale { Product = "gadget", Units = 5 }
            });

            var rows = Materialise(worksheet);

            Assert.Equal(2, rows.Count);
            Assert.Equal(new object[] { "widget", 3 }, ValuesOf(rows[0]));
            Assert.Equal(new object[] { "gadget", 5 }, ValuesOf(rows[1]));
        }

        [Fact]
        public void PopulateRows_WithHeadings_EmitsThePropertyNamesFirst()
        {
            var worksheet = new Worksheet();
            worksheet.PopulateRows(new[] { new Sale { Product = "widget", Units = 3 } },
                existingHeadingRows: 0, usePropertiesAsHeadings: true);

            var rows = Materialise(worksheet);

            Assert.Equal(new object[] { "Product", "Units" }, ValuesOf(rows[0]));
            Assert.Equal(new[] { "Product", "Units" }, worksheet.Headings.ToArray());
        }

        [Fact]
        public void PopulateRows_FromSequencesOfValues_UsesThemAsCellsDirectly()
        {
            var worksheet = new Worksheet();
            worksheet.PopulateRows(new List<IEnumerable<object>>
            {
                new object[] { "a", 1 },
                new object[] { "b", 2 }
            });

            var rows = Materialise(worksheet);

            Assert.Equal(new object[] { "a", 1 }, ValuesOf(rows[0]));
            Assert.Equal(new object[] { "b", 2 }, ValuesOf(rows[1]));
        }

        [Fact]
        public void PopulateRows_FromANonReplayableSequence_StillProducesEveryRow()
        {
            // PopulateRows calls FirstOrDefault() to decide which branch to take, and the branch
            // it picks then enumerates the sequence again. For a source that can only be walked
            // once - a reader, or any yield-based generator - the first element is consumed by
            // the decision and never reaches the output.
            IEnumerable<Sale> OnlyWalkableOnce()
            {
                yield return new Sale { Product = "first", Units = 1 };
                yield return new Sale { Product = "second", Units = 2 };
            }

            var worksheet = new Worksheet();
            worksheet.PopulateRows(OnlyWalkableOnce());

            // A C# iterator method restarts on each enumeration, so this happens to survive.
            // The test exists to catch the day the source is a DataReader instead.
            Assert.Equal(2, Materialise(worksheet).Count);
        }

        // ------------------------------------------------- AddRow

        [Fact]
        public void AddRow_OnAFreshWorksheet_StartsAtRowOne()
        {
            var worksheet = new Worksheet();
            worksheet.AddRow("a", "b");

            var rows = Materialise(worksheet);

            Assert.Single(rows);
            Assert.Equal(1, rows[0].RowNumber);
            Assert.Equal(new object[] { "a", "b" }, ValuesOf(rows[0]));
        }

        [Fact]
        public void AddRow_NumbersEachRowAfterTheLast()
        {
            var worksheet = new Worksheet();
            worksheet.AddRow("one");
            worksheet.AddRow("two");
            worksheet.AddRow("three");

            Assert.Equal(new[] { 1, 2, 3 }, Materialise(worksheet).Select(r => r.RowNumber));
        }

        [Fact]
        public void AddRow_NullValue_LeavesAGapRatherThanShiftingLaterColumns()
        {
            var worksheet = new Worksheet();
            worksheet.AddRow("a", null, "c");

            var cells = Materialise(worksheet)[0].Cells.ToList();

            // No cell is emitted for the null, but the column counter still advances, so "c"
            // correctly lands in column 3 rather than sliding into column 2.
            Assert.Equal(2, cells.Count);
            Assert.Equal(new[] { 1, 3 }, cells.Select(c => c.ColumnNumber));
        }

        [Fact]
        public void AddRow_WithNoValues_AddsAnEmptyRow()
        {
            var worksheet = new Worksheet();
            worksheet.AddRow();

            var rows = Materialise(worksheet);
            Assert.Single(rows);
            Assert.Empty(rows[0].Cells);
        }

        [Fact]
        public void AddRow_AfterPopulatingWithHeadingRows_MisnumbersTheRow()
        {
            // The new row number is `Rows.Count() + 1`, which assumes rows are contiguous and
            // start at 1. After PopulateRowsFromDataTable reserved space for a template header
            // they start at 4, so counting produces a number that collides with a row that
            // already exists.
            var worksheet = new Worksheet();
            worksheet.PopulateRowsFromDataTable(TwoColumnTable(), existingHeadingRows: 3);

            worksheet.AddRow("appended");

            var rows = Materialise(worksheet);
            var appended = rows.Last();

            KnownBug.StillBroken("#125",
                "AddRow appends after the highest row number in the sheet; today it counts rows " +
                "instead, so appending to a sheet that starts at row 4 produces row 4 again",
                () => Assert.Equal(7, appended.RowNumber));

            // Whatever the number, the duplicate is real and worth pinning.
            Assert.Equal(4, appended.RowNumber);
            Assert.Equal(2, rows.Count(r => r.RowNumber == 4));
        }

        // ------------------------------------------------- AddValue

        [Fact]
        public void AddValue_CreatesTheRowAndTheCell()
        {
            var worksheet = new Worksheet();
            worksheet.AddValue(2, 3, "here");

            var rows = Materialise(worksheet);

            Assert.Single(rows);
            Assert.Equal(2, rows[0].RowNumber);
            var cell = Assert.Single(rows[0].Cells);
            Assert.Equal(3, cell.ColumnNumber);
            Assert.Equal("here", cell.Value);
        }

        [Fact]
        public void AddValue_AddsToARowThatAlreadyExists()
        {
            var worksheet = new Worksheet();
            worksheet.AddValue(1, 1, "a");
            worksheet.AddValue(1, 2, "b");

            var rows = Materialise(worksheet);

            Assert.Single(rows);
            Assert.Equal(new object[] { "a", "b" }, ValuesOf(rows[0]));
        }

        [Fact]
        public void AddValue_OnACellThatAlreadyHasAValue_SilentlyDoesNothing()
        {
            var worksheet = new Worksheet();
            worksheet.AddValue(1, 1, "original");
            worksheet.AddValue(1, 1, "replacement");

            var cell = Materialise(worksheet)[0].Cells.Single();

            // The method looks up the cell, finds it, and then only assigns inside an
            // `if (cell == null)` branch - so the second call is a no-op. No exception, no
            // return value, nothing to check: the caller has no way to learn the write was
            // dropped. Naming it AddValue rather than SetValue is the only warning there is.
            Assert.Equal("original", cell.Value);
        }

        [Fact]
        public void AddValue_RejectsAColumnBelowOne()
        {
            var worksheet = new Worksheet();
            Assert.ThrowsAny<Exception>(() => worksheet.AddValue(1, 0, "x"));
        }

        [Fact]
        public void AddValue_RejectsARowBelowOne()
        {
            var worksheet = new Worksheet();
            Assert.ThrowsAny<Exception>(() => worksheet.AddValue(0, 1, "x"));
        }

        // ------------------------------------------------- authoring onto a sheet that was read

        [Fact]
        public void AddRow_AfterReadingASheet_Throws()
        {
            // Read() assigns a lazy iterator to Rows, and AddRow casts it to List<Row>. The cast
            // yields null and the call dereferences it, so the caller gets a bare
            // NullReferenceException with nothing naming the real problem.
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("existing")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);
            var worksheet = fastExcel.Read(1);

            KnownBug.StillBroken("#119",
                "a worksheet that was read can also be appended to, or the attempt fails with " +
                "an error that explains itself rather than a NullReferenceException",
                () =>
                {
                    var thrown = Record.Exception(() => worksheet.AddRow("appended"));
                    Assert.IsNotType<NullReferenceException>(thrown);
                });
        }

        [Fact]
        public void AddValue_AfterReadingASheet_Throws()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("existing")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace);

            using var fastExcel = new FastExcel(file, true);
            var worksheet = fastExcel.Read(1);

            KnownBug.StillBroken("#119",
                "a worksheet that was read can also have values added to it, or the attempt " +
                "fails with an error that explains itself rather than a NullReferenceException",
                () =>
                {
                    var thrown = Record.Exception(() => worksheet.AddValue(5, 1, "appended"));
                    Assert.IsNotType<NullReferenceException>(thrown);
                });
        }

        // ------------------------------------------------- GetCellsInRange

        private static Worksheet ThreeByThree()
        {
            var worksheet = new Worksheet();
            worksheet.AddRow("a1", "b1", "c1");
            worksheet.AddRow("a2", "b2", "c2");
            worksheet.AddRow("a3", "b3", "c3");
            return worksheet;
        }

        [Fact]
        public void GetCellsInRange_ReturnsOnlyTheCellsInsideTheBox()
        {
            var cells = ThreeByThree()
                .GetCellsInRange(new CellRange("A", "B", 1, 2))
                .Select(c => c.Value)
                .ToArray();

            Assert.Equal(new object[] { "a1", "b1", "a2", "b2" }, cells);
        }

        [Fact]
        public void GetCellsInRange_ASingleCell_ReturnsJustThatCell()
        {
            var cells = ThreeByThree().GetCellsInRange(new CellRange("B", "B", 2, 2)).ToList();

            var cell = Assert.Single(cells);
            Assert.Equal("b2", cell.Value);
        }

        [Fact]
        public void GetCellsInRange_WithNoRowEnd_RunsToTheEndOfTheSheet()
        {
            var cells = ThreeByThree()
                .GetCellsInRange(new CellRange("C", "C", 1))
                .Select(c => c.Value)
                .ToArray();

            Assert.Equal(new object[] { "c1", "c2", "c3" }, cells);
        }

        [Fact]
        public void GetCellsInRange_ThatMatchesNothing_ReturnsEmpty()
        {
            Assert.Empty(ThreeByThree().GetCellsInRange(new CellRange("A", "C", 10, 20)));
        }

        [Fact]
        public void GetCellsInRange_WithTheColumnsReversed_ReturnsEmptyRatherThanThrowing()
        {
            // ColumnStart > ColumnEnd is nonsense, and the public CellRange constructor performs
            // no validation, so it is accepted and quietly matches nothing.
            Assert.Empty(ThreeByThree().GetCellsInRange(new CellRange("C", "A", 1, 3)));
        }

        // ------------------------------------------------- small surface

        [Fact]
        public void AFreshWorksheet_HasNoNameAndIndexZero()
        {
            var worksheet = new Worksheet();

            Assert.Null(worksheet.Name);
            Assert.Equal(0, worksheet.Index);
            Assert.Null(worksheet.Rows);
        }

        [Fact]
        public void TemplateAndExistingHeadingRows_RoundTrip()
        {
            var worksheet = new Worksheet { Template = true, ExistingHeadingRows = 2 };

            Assert.True(worksheet.Template);
            Assert.Equal(2, worksheet.ExistingHeadingRows);
        }

        [Fact]
        public void Exists_IsTrueEvenForAWorksheetBackedByNothing()
        {
            // Exists is derived from the part name, which is built by string.Format and is
            // therefore never empty - so the property is a constant true and tells a caller
            // nothing about whether the sheet is really in the package.
            var worksheet = new Worksheet();

            Assert.True(worksheet.Exists);
        }
    }
}
