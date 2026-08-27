using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// Reads FastExcel's output with somebody else's library, and reads somebody else's output
    /// with FastExcel.
    ///
    /// Every other test in this project asks whether the library agrees with itself. That is a
    /// weak question here, because the library escapes and unescapes with the same scheme and
    /// parses with the same assumptions it writes with — so a file can be wrong for the entire
    /// rest of the world and still round-trip perfectly. These tests ask the question that
    /// actually matters: <b>can anyone else read this?</b>
    ///
    /// The oracles are ClosedXML, MiniExcel and Sylvan, all test-only references. None of them is
    /// Excel, so agreement is evidence rather than proof — but three independent implementations
    /// dropping the same cell is about as close as automated testing gets.
    ///
    /// Two things worth knowing about the oracles, both verified rather than assumed:
    /// Sylvan consumes the first row as a header by default, so a fixture needs a header row plus
    /// data; and ClosedXML's <c>CellsUsed()</c> simply omits cells it could not interpret, which
    /// is exactly the failure mode this file exists to catch.
    /// </summary>
    public class InteropTests
    {
        /// <summary>Writes a workbook with FastExcel and returns it, having checked it is a valid package.</summary>
        private static FileInfo WriteWithFastExcel(TempWorkspace workspace, Action<Worksheet> build)
        {
            var template = new XlsxBuilder().WithSheet("Sheet1").ToFile(workspace, "template.xlsx");
            var output = workspace.NewFile();

            var worksheet = new Worksheet();
            build(worksheet);

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            XlsxAssert.IsValidPackage(output);
            return output;
        }

        /// <summary>Every cell ClosedXML can make sense of, keyed by address.</summary>
        private static Dictionary<string, string> ReadWithClosedXml(FileInfo file)
        {
            using var workbook = new XLWorkbook(file.FullName);
            return workbook.Worksheet(1).CellsUsed()
                .ToDictionary(c => c.Address.ToString(), c => c.Value.ToString());
        }

        /// <summary>Row values as MiniExcel sees them, header row disabled.</summary>
        private static List<List<object>> ReadWithMiniExcel(FileInfo file) =>
            MiniExcelLibs.MiniExcel.Query(file.FullName, useHeaderRow: false)
                .Cast<IDictionary<string, object>>()
                .Select(row => row.Values.ToList())
                .ToList();

        /// <summary>Data rows as Sylvan sees them. Sylvan consumes row 1 as a header.</summary>
        private static List<List<string>> ReadWithSylvan(FileInfo file)
        {
            var rows = new List<List<string>>();
            using var reader = Sylvan.Data.Excel.ExcelDataReader.Create(file.FullName);
            while (reader.Read())
            {
                rows.Add(Enumerable.Range(0, reader.FieldCount).Select(reader.GetString).ToList());
            }
            return rows;
        }

        // ============================================================ what survives the trip out

        [Fact]
        public void AStringWrittenByFastExcel_IsReadBackByEveryOtherLibrary()
        {
            using var workspace = new TempWorkspace();
            var file = WriteWithFastExcel(workspace, ws =>
            {
                ws.AddRow("header");
                ws.AddRow("payload");
            });

            Assert.Equal("payload", ReadWithClosedXml(file)["A2"]);
            Assert.Equal("payload", ReadWithMiniExcel(file)[1][0]);
            Assert.Equal("payload", ReadWithSylvan(file).Single()[0]);
        }

        [Fact]
        public void NumbersWrittenByFastExcel_AreReadBackAsNumbers()
        {
            using var workspace = new TempWorkspace();
            var file = WriteWithFastExcel(workspace, ws =>
            {
                ws.AddRow("header");
                ws.AddRow(42, 1.5d);
            });

            using var workbook = new XLWorkbook(file.FullName);
            var sheet = workbook.Worksheet(1);

            Assert.Equal(XLDataType.Number, sheet.Cell("A2").DataType);
            Assert.Equal(42d, sheet.Cell("A2").GetDouble());
            Assert.Equal(1.5d, sheet.Cell("B2").GetDouble());
        }

        [Theory]
        [InlineData("with spaces")]
        [InlineData("ampersand & angle <")]
        [InlineData("quote \" apostrophe '")]
        [InlineData("café")]
        [InlineData("你好")]
        [InlineData("  leading and trailing  ")]
        public void TextWrittenByFastExcel_ReachesClosedXmlIntact(string text)
        {
            // This is the test that decides what #76 actually is.
            //
            // FastExcel writes a space as the escape sequence _x0020_, which is not required —
            // a space is perfectly legal XML text. But the escape *is* part of the OOXML string
            // type, and a conforming reader decodes it, so the text arrives intact anyway.
            //
            // So the over-escaping is ugly and non-conforming rather than immediately destructive.
            // The genuinely destructive half is the opposite direction, covered by
            // CellTypeTests.SharedStringText_SurvivesTheReadIntact: a user string that merely
            // *looks* like an escape is silently decoded into something else.
            using var workspace = new TempWorkspace();
            var file = WriteWithFastExcel(workspace, ws =>
            {
                ws.AddRow("header");
                ws.AddRow(text);
            });

            Assert.Equal(text, ReadWithClosedXml(file)["A2"]);
            Assert.Equal(text, ReadWithSylvan(file).Single()[0]);
        }

        // ============================================================ what does not survive

        [Fact]
        public void ABoolWrittenByFastExcel_DisappearsForOtherReaders()
        {
            // The cell is written as <c r="A2"><v>True</v></c>: no type attribute, so it claims
            // to be a number, and its value is the word "True". ClosedXML cannot make a number
            // of that and drops the cell entirely — it does not throw, does not warn, and does
            // not appear in CellsUsed().
            //
            // For a consumer this is silent data loss: the column is simply empty.
            using var workspace = new TempWorkspace();
            var file = WriteWithFastExcel(workspace, ws =>
            {
                ws.AddRow("header");
                ws.AddRow(true);
            });

            Assert.Contains("<v>True</v>", XlsxAssert.RawPart(file, "xl/worksheets/sheet1.xml"));

            var closedXml = ReadWithClosedXml(file);

            KnownBug.StillBroken("#77",
                "a bool written by FastExcel is visible to other readers; today it is written " +
                "as the word True in a numeric cell and ClosedXML silently drops it",
                () => Assert.True(closedXml.ContainsKey("A2"), "ClosedXML did not see the cell at all"));
        }

        [Fact]
        public void ADateTimeWrittenByFastExcel_DisappearsForOtherReaders()
        {
            // Same shape, worse content: a numeric cell containing "5/6/2024 12:00:00 AM".
            // ClosedXML drops it. The value is also culture-formatted, so the same code produces
            // different bytes on different machines.
            using var workspace = new TempWorkspace();
            var file = WriteWithFastExcel(workspace, ws =>
            {
                ws.AddRow("header");
                ws.AddRow(new DateTime(2024, 5, 6));
            });

            var closedXml = ReadWithClosedXml(file);

            KnownBug.StillBroken("#77",
                "a DateTime written by FastExcel is visible to other readers as a date; today " +
                "it is written as a formatted string in a numeric cell and ClosedXML drops it",
                () => Assert.True(closedXml.ContainsKey("A2"), "ClosedXML did not see the cell at all"));
        }

        [Fact]
        public void OneBadCellDoesNotCostTheWholeRow()
        {
            // Worth knowing how far the damage spreads: a bool between two good strings loses
            // only its own cell, and the columns either side survive at their own addresses.
            // So the failure is a hole, not a shift - which is at least recoverable.
            using var workspace = new TempWorkspace();
            var file = WriteWithFastExcel(workspace, ws =>
            {
                ws.AddRow("header");
                ws.AddRow("before", true, "after");
            });

            var closedXml = ReadWithClosedXml(file);

            Assert.Equal("before", closedXml["A2"]);
            Assert.Equal("after", closedXml["C2"]);
            Assert.False(closedXml.ContainsKey("B2"));
        }

        // ============================================================ the culture bug, proven

        [Theory]
        [InlineData("de-DE")]
        [InlineData("ru-RU")]
        [InlineData("fr-FR")]
        public void ANumberWrittenUnderACommaLocale_IsUnreadableToOtherLibraries(string culture)
        {
            // #55, demonstrated rather than described.
            //
            // Under these cultures the writer emits <v>1,5</v>. Numbers in a worksheet are always
            // invariant, so "1,5" is not a number at all - and the cell suffers the same fate as
            // the bool above. The user's file loses the column, on their machine only, for
            // reasons that have nothing to do with their data.
            using var workspace = new TempWorkspace();

            FileInfo file;
            using (new CultureScope(culture))
            {
                file = WriteWithFastExcel(workspace, ws =>
                {
                    ws.AddRow("header");
                    ws.AddRow(1.5d);
                });
            }

            var sheetXml = XlsxAssert.RawPart(file, "xl/worksheets/sheet1.xml");
            File.WriteAllText(Path.Combine(Path.GetTempPath(), "formula-probe.xml"), sheetXml);
            var closedXml = ReadWithClosedXml(file);

            KnownBug.StillBroken("#55",
                "a number written under " + culture + " reaches other readers as 1.5; today the " +
                "cell holds a locale-formatted value that is not a number, and ClosedXML drops it",
                () =>
                {
                    Assert.Contains("<v>1.5</v>", sheetXml);
                    Assert.True(closedXml.ContainsKey("A2"), "ClosedXML did not see the cell at all");
                });
        }

        [Fact]
        public void ANumberWrittenUnderAnInvariantLocale_IsReadableEverywhere()
        {
            // The control for the theory above: same code, English machine, everything works.
            // That contrast is the whole reason #55 goes unnoticed by maintainers.
            using var workspace = new TempWorkspace();

            FileInfo file;
            using (new CultureScope("en-US"))
            {
                file = WriteWithFastExcel(workspace, ws =>
                {
                    ws.AddRow("header");
                    ws.AddRow(1.5d);
                });
            }

            Assert.Contains("<v>1.5</v>", XlsxAssert.RawPart(file, "xl/worksheets/sheet1.xml"));
            Assert.Equal(1.5d, ReadWithClosedXml(file).Count > 0
                ? double.Parse(ReadWithClosedXml(file)["A2"], CultureInfo.InvariantCulture)
                : double.NaN);
        }

        // ============================================================ reading real files

        /// <summary>Writes a workbook with ClosedXML — a real, conforming writer.</summary>
        private static FileInfo WriteWithClosedXml(TempWorkspace workspace, Action<IXLWorksheet> build)
        {
            var path = workspace.Path("closedxml.xlsx");
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Sheet1");
                build(sheet);
                workbook.SaveAs(path);
            }
            return new FileInfo(path);
        }

        private static List<List<object>> ReadWithFastExcel(FileInfo file)
        {
            using var fastExcel = new FastExcel(file, true);
            return fastExcel.Read(1).Rows
                .Select(r => r.Cells.Select(c => c.Value).ToList())
                .ToList();
        }

        [Fact]
        public void TextWrittenByClosedXml_IsReadCorrectly()
        {
            using var workspace = new TempWorkspace();
            var file = WriteWithClosedXml(workspace, sheet =>
            {
                sheet.Cell("A1").Value = "plain";
                sheet.Cell("B1").Value = "with spaces";
                sheet.Cell("C1").Value = "café";
            });

            var row = ReadWithFastExcel(file).Single();

            Assert.Equal(new object[] { "plain", "with spaces", "café" }, row.ToArray());
        }

        [Fact]
        public void ANumberWrittenByClosedXml_IsReadAsItsDigits()
        {
            using var workspace = new TempWorkspace();
            var file = WriteWithClosedXml(workspace, sheet => sheet.Cell("A1").Value = 1.5);

            Assert.Equal("1.5", ReadWithFastExcel(file).Single().Single());
        }

        [Fact]
        public void ADateWrittenByClosedXml_IsReadAsARawSerialNumber()
        {
            // A real, conforming date: a number carrying a date number format. FastExcel never
            // opens styles.xml, so the caller receives the serial and no indication it is a date.
            // 45418 is 2024-05-06.
            using var workspace = new TempWorkspace();
            var file = WriteWithClosedXml(workspace, sheet => sheet.Cell("A1").Value = new DateTime(2024, 5, 6));

            var value = ReadWithFastExcel(file).Single().Single();

            Assert.Equal("45418", value);

            KnownBug.StillBroken("#58",
                "a date written by a conforming writer is read back as a DateTime rather than " +
                "as its underlying serial number",
                () => Assert.IsType<DateTime>(value));
        }

        [Fact]
        public void ABooleanWrittenByClosedXml_IsReadAsADigit()
        {
            // ClosedXML writes t="b" with 1, which is correct. FastExcel does not recognise the
            // type and hands back "1", indistinguishable from the number one.
            using var workspace = new TempWorkspace();
            var file = WriteWithClosedXml(workspace, sheet => sheet.Cell("A1").Value = true);

            var value = ReadWithFastExcel(file).Single().Single();

            Assert.Equal("1", value);

            KnownBug.StillBroken("#77",
                "a boolean written by a conforming writer is read back as a bool",
                () => Assert.IsType<bool>(value));
        }

        [Fact]
        public void RepeatedTextWrittenByClosedXml_IsReadCorrectly()
        {
            // Repeated values are the shared-string duplicate case (#87 / #81) as a real writer
            // produces it, rather than as hand-written XML.
            using var workspace = new TempWorkspace();
            var file = WriteWithClosedXml(workspace, sheet =>
            {
                sheet.Cell("A1").Value = "alpha";
                sheet.Cell("B1").Value = "beta";
                sheet.Cell("C1").Value = "alpha";
                sheet.Cell("D1").Value = "gamma";
            });

            var row = ReadWithFastExcel(file).Single();

            Assert.Equal(new object[] { "alpha", "beta", "alpha", "gamma" }, row.ToArray());
        }

        [Fact]
        public void AGapWrittenByClosedXml_IsReportedWithTheRightColumnNumbers()
        {
            // The empty-cell family (#88 / #83 / #78) against a real file. Reading by
            // ColumnNumber is correct; reading by position is not - which is what consumers do.
            using var workspace = new TempWorkspace();
            var file = WriteWithClosedXml(workspace, sheet =>
            {
                sheet.Cell("A1").Value = "first";
                sheet.Cell("C1").Value = "third";
            });

            using var fastExcel = new FastExcel(file, true);
            var cells = fastExcel.Read(1).Rows.Single().Cells.ToList();

            Assert.Equal(new[] { 1, 3 }, cells.Select(c => c.ColumnNumber));

            KnownBug.StillBroken("#88",
                "a gap yields a placeholder cell, so index 1 of the sequence is column B rather " +
                "than skipping straight to column C",
                () => Assert.Equal(3, cells.Count));
        }

        [Fact]
        public void AFormulaWrittenByClosedXml_IsReadAsItsCachedResult()
        {
            using var workspace = new TempWorkspace();
            var file = WriteWithClosedXml(workspace, sheet =>
            {
                sheet.Cell("A1").Value = 2;
                sheet.Cell("B1").Value = 3;
                sheet.Cell("C1").FormulaA1 = "A1+B1";
                sheet.RecalculateAllFormulas();
            });

            var sheetXml = XlsxAssert.RawPart(file, "xl/worksheets/sheet1.xml");
            var row = ReadWithFastExcel(file).Single();

            // ClosedXML saves the formula but not a cached result, which is legal - Excel
            // recalculates on open. FastExcel has no calculation engine and no way to say
            // "not computed", so its fallback hands back the formula's own text. A caller
            // summing a column silently gets the string "A1+B1" where a number belongs.
            // Note the prefix: ClosedXML writes <x:f>, not <f>. Both are the same element.
            Assert.Contains("<x:f>A1+B1</x:f>", sheetXml);
            Assert.DoesNotContain("<x:v>", sheetXml.Substring(sheetXml.IndexOf("<x:f>", StringComparison.Ordinal)));

            KnownBug.StillBroken("#68",
                "a formula with no cached result is distinguishable from a cell whose text " +
                "happens to be a formula; today both read back as \"A1+B1\"",
                () => Assert.NotEqual("A1+B1", row[2]));
        }

        [Fact]
        public void AFileWrittenByClosedXml_CannotBeUsedAsATemplate()
        {
            // The most consequential interop finding here, and it is not in any issue.
            //
            // ClosedXML binds the spreadsheet namespace to the prefix "x" and writes
            // <x:worksheet>, <x:sheetData>, <x:row>, <x:c>. That is ordinary, conforming OOXML -
            // a prefix is just a name, and FastExcel's *reader* handles it correctly because it
            // compares LocalName.
            //
            // The *writer* does not. It locates the insertion point by searching each line for
            // the literal string "<sheetData>", which never matches "<x:sheetData>". The scan
            // runs off the end of the file, the whole document is swallowed into the header
            // buffer, and the new rows are appended after a document that was already closed -
            // producing a part with two root elements.
            //
            // Nothing throws. The call returns, the zip is well formed, and the file is broken.
            // Since "take a workbook someone else produced and add rows to it" is the library's
            // headline use, and ClosedXML is one of the most widely used .NET writers, this is a
            // real path to a corrupt file.
            using var workspace = new TempWorkspace();

            var template = WriteWithClosedXml(workspace, sheet => sheet.Cell("A1").Value = "from-closedxml");
            var output = workspace.NewFile();

            var worksheet = new Worksheet();
            worksheet.AddRow("appended-by-fastexcel");

            using (var fastExcel = new FastExcel(template, output))
            {
                fastExcel.Write(worksheet, 1);
            }
            output.Refresh();

            KnownBug.StillBroken("#71",
                "a workbook produced by another conforming writer can be used as a template; " +
                "today a namespace-prefixed worksheet part defeats the literal string search " +
                "for \"<sheetData>\" and the output is not well-formed XML",
                () => XlsxAssert.IsValidPackage(output));
        }

        [Fact]
        public void AFileWrittenByClosedXml_IsStillReadableByFastExcel()
        {
            // The control for the test above: reading a prefixed file is fine, because the
            // reader compares local names instead of matching strings. So the two halves of the
            // library disagree about what a worksheet looks like.
            using var workspace = new TempWorkspace();
            var file = WriteWithClosedXml(workspace, sheet => sheet.Cell("A1").Value = "from-closedxml");

            Assert.Contains("<x:sheetData>", XlsxAssert.RawPart(file, "xl/worksheets/sheet1.xml"));
            Assert.Equal("from-closedxml", ReadWithFastExcel(file).Single().Single());
        }

        [Fact]
        public void AWorkbookWrittenByClosedXml_HasItsSheetNamesResolvedCorrectly()
        {
            using var workspace = new TempWorkspace();
            var path = workspace.Path("multi.xlsx");
            using (var workbook = new XLWorkbook())
            {
                workbook.Worksheets.Add("Alpha").Cell("A1").Value = "in-alpha";
                workbook.Worksheets.Add("Beta").Cell("A1").Value = "in-beta";
                workbook.SaveAs(path);
            }

            var file = new FileInfo(path);
            using var fastExcel = new FastExcel(file, true);

            Assert.Equal(new[] { "Alpha", "Beta" }, fastExcel.Worksheets.Select(w => w.Name));
            Assert.Equal("in-beta", fastExcel.Read("Beta").Rows.Single().Cells.Single().Value);
        }
    }
}
