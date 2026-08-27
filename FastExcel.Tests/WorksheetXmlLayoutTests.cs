using System;
using System.Collections.Generic;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// The write path does not parse the worksheet part. It reads it line by line and looks for
    /// substrings — <c>line.Contains("&lt;sheetData&gt;")</c>, <c>line.Contains("&lt;row")</c> —
    /// which means the *layout* of the XML changes the outcome even when the document is
    /// identical. Two files that any XML parser considers equal behave differently here.
    ///
    /// That makes these tests unusual: the input differs only in whitespace and namespace
    /// prefixes, both of which are semantically meaningless in XML and both of which real
    /// producers emit freely. LibreOffice indents its output; Excel itself is happy to write
    /// prefixed elements.
    ///
    /// One of these does not merely produce a wrong answer — it never returns at all.
    /// </summary>
    public class WorksheetXmlLayoutTests
    {
        private const string Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        /// <summary>The same one-row sheet, written the way a writer that never pretty-prints emits it.</summary>
        private const string SingleLineSheet =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<worksheet xmlns=\"" + Ns + "\"><sheetData>" +
            "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Heading</t></is></c></row>" +
            "</sheetData></worksheet>";

        /// <summary>
        /// Byte-for-byte the same document, indented. This is what LibreOffice writes, and what
        /// anything that has been through an XML formatter looks like.
        /// </summary>
        private const string PrettyPrintedSheet =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
            "<worksheet xmlns=\"" + Ns + "\">\n" +
            "  <sheetData>\n" +
            "    <row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Heading</t></is></c></row>\n" +
            "  </sheetData>\n" +
            "</worksheet>";

        /// <summary>Again the same document, with the main namespace bound to a prefix.</summary>
        private const string PrefixedSheet =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<x:worksheet xmlns:x=\"" + Ns + "\"><x:sheetData>" +
            "<x:row r=\"1\"><x:c r=\"A1\" t=\"inlineStr\"><x:is><x:t>Heading</x:t></x:is></x:c></x:row>" +
            "</x:sheetData></x:worksheet>";

        /// <summary>
        /// Appends one row to a template whose worksheet part is supplied verbatim, so the test
        /// controls the exact layout the write path will scan.
        /// </summary>
        private static void AppendRowToTemplate(string worksheetXml, int existingHeadingRows)
        {
            using var workspace = new TempWorkspace();

            var template = new XlsxBuilder()
                .WithSheetXml("Sheet1", worksheetXml)
                .ToFile(workspace, "template.xlsx");

            var output = workspace.NewFile();

            var worksheet = new Worksheet
            {
                Rows = new List<Row>
                {
                    new Row(2, new List<Cell> { new Cell(1, "data") })
                }
            };

            using var fastExcel = new FastExcel(template, output);
            fastExcel.Write(worksheet, "Sheet1", existingHeadingRows);
        }

        // ------------------------------------------------------------------ the baseline

        [Fact]
        public void SingleLineWorksheetXml_IsWritten()
        {
            // The control. Establishes that everything except the layout is sound, so the
            // pretty-printed case below cannot be blamed on the fixture.
            Assert.True(
                Timebox.Completes(TimeSpan.FromSeconds(10), () => AppendRowToTemplate(SingleLineSheet, 1)),
                "the single-line control case did not finish, so this fixture proves nothing");
        }

        // ------------------------------------------------------------------ #80

        [Fact]
        public void PrettyPrintedWorksheetXml_DoesNotHang()
        {
            // Reproduces #80 ("The program is looping when I use existingHeadingRows").
            //
            // Once the scan is past <sheetData>, the loop advances only on a line containing
            // "<row" or "</row>". A line holding neither — which is exactly what indenting
            // produces — leaves both the line and the remaining heading count untouched, so the
            // condition that entered the loop is still true. It spins at 100% CPU forever.
            //
            // Note this needs existingHeadingRows > 0: with 0 the loop is never entered, which
            // is why the defect looks intermittent to whoever hits it.
            KnownBug.StillBroken("#80",
                "a worksheet part that is indented rather than written on one line is still " +
                "written, because the header scan advances on every line rather than only on " +
                "lines that happen to contain row markup",
                () => Assert.True(
                    Timebox.Completes(TimeSpan.FromSeconds(5), () => AppendRowToTemplate(PrettyPrintedSheet, 1)),
                    "writing never returned - the header scan is spinning"));
        }

        [Fact]
        public void PrettyPrintedWorksheetXml_WithNoHeadingRows_IsUnaffected()
        {
            // Pins the boundary of #80 so a future fix cannot narrow it by accident: with no
            // heading rows the same file is fine today, and must stay fine.
            Assert.True(
                Timebox.Completes(TimeSpan.FromSeconds(10), () => AppendRowToTemplate(PrettyPrintedSheet, 0)),
                "an indented template with no heading rows should never have been affected");
        }

        // ------------------------------------------------------------------ namespace prefixes

        [Fact]
        public void PrefixedWorksheetXml_ProducesAValidFile()
        {
            // A prefixed part is ordinary OOXML. Because the scan looks for the literal
            // "<sheetData>", it never matches "<x:sheetData>", the whole document is swallowed
            // into the header buffer and the produced part is silently malformed.
            //
            // This one does not throw and does not hang, so only a validity check catches it -
            // which is the reason XlsxAssert exists.
            using var workspace = new TempWorkspace();

            var template = new XlsxBuilder()
                .WithSheetXml("Sheet1", PrefixedSheet)
                .ToFile(workspace, "template.xlsx");

            var output = workspace.NewFile();

            var worksheet = new Worksheet
            {
                Rows = new List<Row> { new Row(2, new List<Cell> { new Cell(1, "data") }) }
            };

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(worksheet, "Sheet1");
            }

            KnownBug.StillBroken("#71",
                "a worksheet part whose elements carry a namespace prefix is written back as a " +
                "well-formed package, because the writer locates sheetData by name rather than " +
                "by matching the literal string \"<sheetData>\"",
                () => XlsxAssert.IsValidPackage(output));
        }
    }
}
