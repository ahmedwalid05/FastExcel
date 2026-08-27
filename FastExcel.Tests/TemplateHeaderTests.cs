using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Writing into a template that already has content above the data.
    ///
    /// This is the headline use of the library — take a formatted workbook, drop rows into it,
    /// keep the styling — and it is the least tested part of it. The mechanism is a hand-rolled
    /// scan that splits the worksheet part into a "header" prefix and a "footer" suffix, keeps
    /// both verbatim, and writes new rows between them. It never parses anything, so every case
    /// here is really a test of that scan.
    ///
    /// <see cref="WorksheetXmlLayoutTests"/> covers what happens when the layout defeats the scan
    /// entirely. These cover the shapes it is supposed to handle.
    /// </summary>
    public class TemplateHeaderTests
    {
        private const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private static FileInfo TemplateWithSheetXml(TempWorkspace workspace, string worksheetXml) =>
            new XlsxBuilder().WithSheetXml("Sheet1", worksheetXml).ToFile(workspace, "template.xlsx");

        /// <summary>Writes rows into a template and hands back the reopened output.</summary>
        private static FileInfo WriteInto(TempWorkspace workspace, FileInfo template,
            int existingHeadingRows, params Row[] rows)
        {
            var output = workspace.NewFile();
            var worksheet = new Worksheet { Rows = rows.ToList() };

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(worksheet, 1, existingHeadingRows);
            }

            output.Refresh();
            XlsxAssert.IsValidPackage(output);
            return output;
        }

        private static Row RowOf(int number, string value) =>
            new Row(number, new List<Cell> { new Cell(1, value) });

        /// <summary>
        /// Reads every row as plain values.
        ///
        /// Materialising here is not tidiness. <c>Row.Cells</c> is a lazy iterator over an
        /// XDocument held open by the archive, so returning rows and touching their cells later
        /// throws once the FastExcel instance is disposed. Doing it inside the using block is the
        /// only safe shape, and it is the trap RowReadTests pins deliberately.
        /// </summary>
        private static List<List<object>> ReadAll(FileInfo file)
        {
            using var fastExcel = new FastExcel(file, true);
            return fastExcel.Read(1).Rows
                .Select(r => r.Cells.Select(c => c.Value).ToList())
                .ToList();
        }

        // ------------------------------------------------------------------ empty templates

        [Fact]
        public void ATemplateWithASelfClosingSheetData_CanBeWrittenInto()
        {
            // <sheetData/> is what Excel writes for an empty sheet. The scan has a separate
            // branch for it, because it has to invent the opening and closing tags that a
            // self-closing element does not have.
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace,
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"" + Ns + "\"><sheetData/></worksheet>");

            var output = WriteInto(workspace, template, 0, RowOf(1, "written"));

            Assert.Equal("written", ReadAll(output).Single().Single());
        }

        [Fact]
        public void ATemplateWithAnEmptySheetDataPair_CanBeWrittenInto()
        {
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace,
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"" + Ns + "\"><sheetData></sheetData></worksheet>");

            var output = WriteInto(workspace, template, 0, RowOf(1, "written"));

            Assert.Equal("written", ReadAll(output).Single().Single());
        }

        // ------------------------------------------------------------------ keeping heading rows

        private static string TemplateWithHeadings(int count)
        {
            var rows = string.Concat(Enumerable.Range(1, count).Select(n =>
                "<row r=\"" + n + "\"><c r=\"A" + n + "\" t=\"inlineStr\"><is><t>heading" + n +
                "</t></is></c></row>"));

            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                   "<worksheet xmlns=\"" + Ns + "\"><sheetData>" + rows + "</sheetData></worksheet>";
        }

        [Fact]
        public void OneHeadingRow_SurvivesTheWrite()
        {
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace, TemplateWithHeadings(1));

            var output = WriteInto(workspace, template, 1, RowOf(2, "data"));
            var rows = ReadAll(output);

            Assert.Equal(2, rows.Count);
            Assert.Equal("heading1", rows[0].Single());
            Assert.Equal("data", rows[1].Single());
        }

        [Fact]
        public void SeveralHeadingRows_AllSurviveTheWrite()
        {
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace, TemplateWithHeadings(3));

            var output = WriteInto(workspace, template, 3, RowOf(4, "data"));
            var rows = ReadAll(output);

            Assert.Equal(4, rows.Count);
            Assert.Equal(new object[] { "heading1", "heading2", "heading3", "data" },
                rows.Select(r => r.Single()).ToArray());
        }

        [Fact]
        public void HeadingRowsSpreadAcrossLines_AreStillKept()
        {
            // A row split so that its opening tag and its closing tag land on different lines.
            // The scan handles this: a line containing "<row" but not "</row>" is consumed
            // whole, and the closing tag is picked up on the next pass.
            //
            // This is the shape immediately next door to the #80 hang: the difference is that
            // every line here contains one of the two markers. A line with neither - which is
            // what indenting produces - never advances the loop.
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace,
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"" + Ns + "\"><sheetData>" +
                "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>heading1</t></is></c>\n" +
                "</row></sheetData></worksheet>");

            var output = WriteInto(workspace, template, 1, RowOf(2, "data"));
            var rows = ReadAll(output);

            Assert.Equal(2, rows.Count);
            Assert.Equal("heading1", rows[0].Single());
            Assert.Equal("data", rows[1].Single());
        }

        [Fact]
        public void ClaimingMoreHeadingRowsThanTheTemplateHas_Hangs()
        {
            // A second way into the #80 spin loop, and one the issue does not mention.
            //
            // The scan consumes the single heading row, decrements the counter to 3, and is left
            // holding "</sheetData></worksheet>" - a non-empty line containing neither "<row" nor
            // "</row>". Neither branch fires, nothing advances, and it never returns.
            //
            // So #80 is not only about indented templates: passing an existingHeadingRows larger
            // than the template's real header count hangs a perfectly ordinary single-line file.
            // Since the count usually comes from configuration rather than from the file, this is
            // the easier of the two to hit by accident.
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace, TemplateWithHeadings(1));
            var output = workspace.NewFile();

            var worksheet = new Worksheet { Rows = new List<Row> { RowOf(5, "data") } };

            KnownBug.StillBroken("#80",
                "asking to keep more heading rows than the template contains is rejected or " +
                "handled, rather than spinning forever",
                () => Assert.True(
                    Timebox.Completes(TimeSpan.FromSeconds(5), () =>
                    {
                        using var fastExcel = new FastExcel(template, output);
                        fastExcel.Write(worksheet, 1, existingHeadingRows: 4);
                    }, "over-claimed-heading-rows"),
                    "writing never returned - the header scan is spinning"));
        }

        [Fact]
        public void WritingDataOverAHeadingRow_IsRefused()
        {
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace, TemplateWithHeadings(2));
            var output = workspace.NewFile();

            var worksheet = new Worksheet { Rows = new List<Row> { RowOf(2, "collides") } };

            using var fastExcel = new FastExcel(template, output);
            var thrown = Record.Exception(() => fastExcel.Write(worksheet, 1, existingHeadingRows: 2));

            Assert.NotNull(thrown);
            Assert.Contains("Heading", thrown.Message);
        }

        // ------------------------------------------------------------------ what the template loses

        [Fact]
        public void ContentBelowTheDataIsKept()
        {
            // Everything after the last heading row becomes the "footer" and is written back
            // verbatim, so sheet-level settings that live after <sheetData> survive.
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace,
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"" + Ns + "\">" +
                "<sheetData></sheetData>" +
                "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>" +
                "</worksheet>");

            var output = WriteInto(workspace, template, 0, RowOf(1, "data"));

            Assert.Contains("pageMargins", XlsxAssert.RawPart(output, "xl/worksheets/sheet1.xml"));
        }

        [Fact]
        public void ContentAboveTheDataIsKept()
        {
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace,
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"" + Ns + "\">" +
                "<sheetFormatPr defaultRowHeight=\"15\"/>" +
                "<cols><col min=\"1\" max=\"1\" width=\"40\" customWidth=\"1\"/></cols>" +
                "<sheetData></sheetData></worksheet>");

            var output = WriteInto(workspace, template, 0, RowOf(1, "data"));
            var xml = XlsxAssert.RawPart(output, "xl/worksheets/sheet1.xml");

            Assert.Contains("sheetFormatPr", xml);
            Assert.Contains("customWidth", xml);
        }

        [Fact]
        public void TheDimensionElementIsNotUpdatedToMatchTheDataWritten()
        {
            // <dimension> states the used range. It sits above <sheetData>, so it is copied into
            // the header untouched - meaning a template that declared A1:A1 still says A1:A1
            // after ten rows are written. Excel recalculates on open, but other readers trust it
            // and see a truncated sheet.
            using var workspace = new TempWorkspace();
            var template = TemplateWithSheetXml(workspace,
                "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                "<worksheet xmlns=\"" + Ns + "\">" +
                "<dimension ref=\"A1:A1\"/><sheetData></sheetData></worksheet>");

            var output = WriteInto(workspace, template, 0,
                RowOf(1, "one"), RowOf(2, "two"), RowOf(3, "three"));

            var xml = XlsxAssert.RawPart(output, "xl/worksheets/sheet1.xml");

            Assert.Contains("A1:A1", xml);

            KnownBug.StillBroken("#71",
                "the dimension element is rewritten to cover the rows actually written, so a " +
                "reader that trusts it does not see a truncated sheet",
                () => Assert.DoesNotContain("A1:A1", xml));
        }
    }
}
