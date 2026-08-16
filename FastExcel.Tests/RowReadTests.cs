using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Covers how a &lt;row&gt; becomes a sequence of cells. The xlsx format omits empty cells
    /// entirely, so a row holding A, C and D stores only three &lt;c&gt; elements. Every other
    /// reader (Excel, EPPlus, Sylvan) materialises the gaps; this library does not, so callers
    /// indexing the sequence positionally silently land on the wrong column.
    /// </summary>
    public class RowReadTests
    {
        /// <summary>
        /// Reads a sheet and fully materialises it before the archive closes.
        /// <para>
        /// Both Worksheet.Rows and Row.Cells are lazy iterators over the open ZipArchive, so
        /// calling .ToList() on Rows alone is not enough — the cells are still unread, and
        /// enumerating them after disposal throws. See
        /// <see cref="EnumeratingCellsAfterDisposal_Throws"/>, which pins that behaviour.
        /// </para>
        /// </summary>
        private static List<Row> ReadRows(FileInfo file)
        {
            using var fastExcel = new FastExcel(file, true);
            var rows = fastExcel.Read(1).Rows.ToList();
            foreach (var row in rows)
            {
                row.Cells = (row.Cells ?? Enumerable.Empty<Cell>()).ToList();
            }
            return rows;
        }

        [Fact]
        public void EnumeratingCellsAfterDisposal_Throws()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("value")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace);

            List<Row> rows;
            using (var fastExcel = new FastExcel(file, true))
            {
                // .ToList() materialises the rows but NOT their cells — Row.Cells is a separate
                // lazy iterator that has not touched the archive yet.
                rows = fastExcel.Read(1).Rows.ToList();
            }

            // Documents a real trap for callers: a Worksheet does not outlive the FastExcel
            // instance that produced it, even though nothing in the API signals that.
            Assert.Throws<DefinedNameLoadException>(() => rows.Single().Cells.ToList());
        }

        // ------------------------------------------------------- #88 / #83 / #78 cell gaps

        [Fact]
        public void RowWithAGap_ReportsCorrectColumnNumbersForThePresentCells()
        {
            using var workspace = new TempWorkspace();
            // A1 and C1 are present; B1 is absent, as Excel stores it.
            var file = new XlsxBuilder()
                .WithSharedStrings("first", "third")
                .WithSheet("Sheet1", XlsxBuilder.Row(1,
                    XlsxBuilder.SharedCell("A1", 0),
                    XlsxBuilder.SharedCell("C1", 1)))
                .ToFile(workspace);

            var cells = ReadRows(file).Single().Cells.ToList();

            // This part is right today: the column numbers come from the r attribute.
            Assert.Equal(1, cells[0].ColumnNumber);
            Assert.Equal(3, cells[1].ColumnNumber);
        }

        [Fact]
        public void RowWithAGap_MaterialisesTheMissingCell()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("first", "third")
                .WithSheet("Sheet1", XlsxBuilder.Row(1,
                    XlsxBuilder.SharedCell("A1", 0),
                    XlsxBuilder.SharedCell("C1", 1)))
                .ToFile(workspace);

            KnownBug.StillBroken("#88 / #83 / #78",
                "a row spanning A..C yields three cells, with an empty placeholder for the " +
                "absent B, so that indexing the sequence by position matches the spreadsheet",
                () =>
                {
                    var cells = ReadRows(file).Single().Cells.ToList();
                    Assert.Equal(3, cells.Count);
                    Assert.Equal(2, cells[1].ColumnNumber);
                    Assert.Null(cells[1].Value);
                });
        }

        [Fact]
        public void RowWithAGap_PositionalIndexingLandsOnTheRightColumn()
        {
            using var workspace = new TempWorkspace();
            // Y is column 25. With four earlier columns missing, positional indexing is off by
            // four — this is exactly the offset reported in #88.
            var file = new XlsxBuilder()
                .WithSharedStrings("a", "target")
                .WithSheet("Sheet1", XlsxBuilder.Row(1,
                    XlsxBuilder.SharedCell("A1", 0),
                    XlsxBuilder.SharedCell("Y1", 1)))
                .ToFile(workspace);

            KnownBug.StillBroken("#88",
                "cell Y is reachable at index 24 of the row's cell sequence",
                () =>
                {
                    var cells = ReadRows(file).Single().Cells.ToList();
                    Assert.Equal("target", cells[24].Value);
                });
        }

        [Fact]
        public void GetCellByColumnName_FindsAPresentCell()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("first", "third")
                .WithSheet("Sheet1", XlsxBuilder.Row(1,
                    XlsxBuilder.SharedCell("A1", 0),
                    XlsxBuilder.SharedCell("C1", 1)))
                .ToFile(workspace);

            var row = ReadRows(file).Single();

            Assert.Equal("first", row.GetCellByColumnName("A").Value);
            Assert.Equal("third", row.GetCellByColumnName("C").Value);
        }

        [Fact]
        public void GetCellByColumnName_ReturnsNullForAnAbsentColumn()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSharedStrings("first")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace);

            var row = ReadRows(file).Single();

            // Documents current behaviour: callers must null-check. Worth revisiting alongside
            // #88, since returning an empty cell would be friendlier and matches other readers.
            Assert.Null(row.GetCellByColumnName("B"));
        }

        // ------------------------------------------------------ #22 non-cell row content

        [Fact]
        public void RowContainingADrawingElement_DoesNotThrow()
        {
            using var workspace = new TempWorkspace();
            // A row is not guaranteed to contain only <c> elements. Anchored drawings and other
            // foreign content appear here, and they have no r attribute to parse.
            var file = new XlsxBuilder()
                .WithSharedStrings("real cell")
                .WithSheet("Sheet1",
                    "<row r=\"1\">" +
                    XlsxBuilder.SharedCell("A1", 0) +
                    "<xdr:twoCellAnchor xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"><xdr:from><xdr:col>1</xdr:col></xdr:from></xdr:twoCellAnchor>" +
                    "</row>")
                .ToFile(workspace);

            KnownBug.StillBroken("#22",
                "elements inside a row that are not cells are skipped rather than parsed as " +
                "cells; today the missing r attribute makes Regex.Replace throw on null",
                () => ReadRows(file).Single().Cells.ToList());
        }

        [Fact]
        public void CellWithoutAReferenceAttribute_DoesNotThrow()
        {
            using var workspace = new TempWorkspace();
            // The r attribute is optional in the spec; readers are expected to infer position.
            var file = new XlsxBuilder()
                .WithSheet("Sheet1", "<row r=\"1\"><c><v>1</v></c></row>")
                .ToFile(workspace);

            KnownBug.StillBroken("#22",
                "a cell with no r attribute is handled rather than throwing, since the " +
                "attribute is optional and position can be inferred from the cell's index",
                () => ReadRows(file).Single().Cells.ToList());
        }

        // ------------------------------------------------------------------ row basics

        [Fact]
        public void EmptySheet_YieldsNoRows()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace);

            Assert.Empty(ReadRows(file));
        }

        [Fact]
        public void RowNumbers_ArePreservedIncludingGapsBetweenRows()
        {
            using var workspace = new TempWorkspace();
            // Rows 1 and 5 are populated; 2-4 do not exist in the file at all.
            var file = new XlsxBuilder()
                .WithSharedStrings("one", "five")
                .WithSheet("Sheet1",
                    XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)) +
                    XlsxBuilder.Row(5, XlsxBuilder.SharedCell("A5", 1)))
                .ToFile(workspace);

            var rows = ReadRows(file);

            Assert.Equal(new[] { 1, 5 }, rows.Select(r => r.RowNumber));
            Assert.Equal("one", rows[0].Cells.Single().Value);
            Assert.Equal("five", rows[1].Cells.Single().Value);
        }

        [Fact]
        public void NumericCellWithACachedFormula_ReturnsTheCachedValueNotTheFormula()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder()
                .WithSheet("Sheet1", "<row r=\"1\"><c r=\"A1\"><f>1+1</f><v>2</v></c></row>")
                .ToFile(workspace);

            var cell = ReadRows(file).Single().Cells.Single();

            Assert.Equal("2", cell.Value);
        }

        [Fact]
        public void RowWithNoCells_IsReadWithoutThrowing()
        {
            using var workspace = new TempWorkspace();
            var file = new XlsxBuilder().WithSheet("Sheet1", "<row r=\"1\"/>").ToFile(workspace);

            var row = ReadRows(file).Single();

            Assert.Equal(1, row.RowNumber);
            Assert.True(row.Cells == null || !row.Cells.Any());
        }
    }
}
