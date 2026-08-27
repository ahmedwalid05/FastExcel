using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// What happens when the input is not a well-formed workbook.
    ///
    /// Nothing exercised this before. That matters because an xlsx is usually something the
    /// program was handed rather than something it made: an upload, a file off a share, output
    /// from a reporting tool. The library will meet damaged input in production whether or not
    /// it is ready for it.
    ///
    /// Most of these assert only that the failure is <em>diagnosable</em> — the right exception
    /// type, or a message naming the problem. That is a deliberately low bar, and several cases
    /// do not clear it: a workbook missing an internal part fails with a bare
    /// <see cref="NullReferenceException"/>, and in one case it does so from inside
    /// <c>Dispose</c>, where it also abandons a half-written file.
    /// </summary>
    public class MalformedInputTests
    {
        /// <summary>Writes arbitrary bytes to a .xlsx path so the library is handed real rubbish.</summary>
        private static FileInfo FileOfBytes(TempWorkspace workspace, byte[] bytes, string name = "broken.xlsx")
        {
            var path = workspace.Path(name);
            File.WriteAllBytes(path, bytes);
            return new FileInfo(path);
        }

        /// <summary>Rebuilds a package with one entry removed, to simulate a truncated or stripped file.</summary>
        private static FileInfo PackageWithout(TempWorkspace workspace, FileInfo source, string entryToDrop)
        {
            var path = workspace.Path("without.xlsx");
            using (var input = ZipFile.OpenRead(source.FullName))
            using (var output = new FileStream(path, FileMode.Create))
            using (var zip = new ZipArchive(output, ZipArchiveMode.Create))
            {
                foreach (var entry in input.Entries)
                {
                    if (string.Equals(entry.FullName, entryToDrop, StringComparison.OrdinalIgnoreCase)) continue;

                    using var from = entry.Open();
                    using var to = zip.CreateEntry(entry.FullName).Open();
                    from.CopyTo(to);
                }
            }
            return new FileInfo(path);
        }

        private static FileInfo ValidWorkbook(TempWorkspace workspace) =>
            new XlsxBuilder()
                .WithSharedStrings("value")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace, "valid.xlsx");

        // ------------------------------------------------------------- not a package at all

        [Fact]
        public void AFileThatIsNotAZip_IsRejectedAsBadData()
        {
            using var workspace = new TempWorkspace();
            var file = FileOfBytes(workspace, Encoding.UTF8.GetBytes("this is plain text, not a spreadsheet"));

            Assert.Throws<InvalidDataException>(() =>
            {
                using var fastExcel = new FastExcel(file, true);
                fastExcel.Read(1);
            });
        }

        [Fact]
        public void AnEmptyFile_IsRejectedWithAMessageAboutTheInput()
        {
            using var workspace = new TempWorkspace();
            var file = FileOfBytes(workspace, Array.Empty<byte>());

            var thrown = Record.Exception(() =>
            {
                using var fastExcel = new FastExcel(file, true);
                fastExcel.Read(1);
            });

            Assert.NotNull(thrown);
            Assert.IsNotType<NullReferenceException>(thrown);
        }

        [Fact]
        public void ATruncatedZip_IsRejectedAsBadData()
        {
            using var workspace = new TempWorkspace();
            var valid = ValidWorkbook(workspace);
            var bytes = File.ReadAllBytes(valid.FullName);

            // Keeping only the first half destroys the central directory at the end of the file.
            var file = FileOfBytes(workspace, bytes.Take(bytes.Length / 2).ToArray(), "truncated.xlsx");

            Assert.Throws<InvalidDataException>(() =>
            {
                using var fastExcel = new FastExcel(file, true);
                fastExcel.Read(1);
            });
        }

        [Fact]
        public void AZipThatIsNotAWorkbook_FailsWithoutExplainingWhy()
        {
            using var workspace = new TempWorkspace();
            var path = workspace.Path("notaworkbook.xlsx");
            using (var stream = new FileStream(path, FileMode.Create))
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Create))
            using (var writer = new StreamWriter(zip.CreateEntry("readme.txt").Open()))
            {
                writer.Write("a perfectly good zip, just not a spreadsheet");
            }

            var thrown = Record.Exception(() =>
            {
                using var fastExcel = new FastExcel(new FileInfo(path), true);
                fastExcel.Read(1);
            });

            Assert.NotNull(thrown);

            // Every part lookup is GetEntry(...) with no null check, so a zip without
            // xl/workbook.xml dereferences null. The caller is told nothing about what was wrong
            // with their file.
            KnownBug.StillBroken("#119",
                "a zip that is not a workbook is rejected with a message naming the missing " +
                "part, rather than a bare NullReferenceException",
                () => Assert.IsNotType<NullReferenceException>(thrown));
        }

        // ------------------------------------------------------------- missing internal parts

        [Theory]
        [InlineData("xl/workbook.xml")]
        [InlineData("xl/worksheets/sheet1.xml")]
        public void AWorkbookMissingAPartTheReaderNeeds_FailsWithoutExplainingWhy(string part)
        {
            using var workspace = new TempWorkspace();
            var broken = PackageWithout(workspace, ValidWorkbook(workspace), part);

            var thrown = Record.Exception(() =>
            {
                using var fastExcel = new FastExcel(broken, true);
                fastExcel.Read(1).Rows.ToList();
            });

            Assert.NotNull(thrown);

            KnownBug.StillBroken("#119",
                "a workbook missing " + part + " is rejected with a message naming the missing " +
                "part, rather than a bare NullReferenceException",
                () => Assert.IsNotType<NullReferenceException>(thrown));
        }

        [Theory]
        [InlineData("xl/_rels/workbook.xml.rels")]
        [InlineData("[Content_Types].xml")]
        public void AWorkbookMissingItsPackagingParts_IsReadAnywayWithoutComplaint(string part)
        {
            // Strip the relationship graph, or the content types, and the read still succeeds
            // and returns the right value. That is not robustness - it is the mechanism behind
            // #82.
            //
            // A conforming reader locates a sheet by following its r:id through
            // xl/_rels/workbook.xml.rels. This one skips that entirely and computes the part
            // name from the sheet's position: the third <sheet> element is assumed to live at
            // xl/worksheets/sheet3.xml. That assumption holds for files Excel has only ever
            // appended to, and breaks the moment a sheet is deleted or reordered - at which
            // point asking for one sheet quietly returns another.
            //
            // So this test passing is the bad news. When #82 is fixed the reader will have to
            // consult the relationships, a workbook without them will stop being readable, and
            // this test should be rewritten to expect that.
            using var workspace = new TempWorkspace();
            var stripped = PackageWithout(workspace, ValidWorkbook(workspace), part);

            using var fastExcel = new FastExcel(stripped, true);
            var value = fastExcel.Read(1).Rows.Single().Cells.Single().Value;

            Assert.Equal("value", value);
        }

        [Fact]
        public void AWorkbookMissingItsRelationships_FailsInsideDisposeAndAbandonsTheOutput()
        {
            // The worst of the missing-part cases. The relationship and content-type parts are
            // only touched while finishing a write, which happens in Dispose. So the failure
            // arrives from a using-block's closing brace, after the caller believes the write
            // succeeded, and Archive.Dispose is never reached - leaving an unflushed, truncated
            // file on disk with no exception at the point of the actual write.
            using var workspace = new TempWorkspace();
            var template = PackageWithout(workspace, ValidWorkbook(workspace), "xl/_rels/workbook.xml.rels");
            var output = workspace.NewFile();

            var worksheet = new Worksheet();
            worksheet.AddRow("data");

            var thrown = Record.Exception(() =>
            {
                using var fastExcel = new FastExcel(template, output);
                fastExcel.Write(worksheet, 1);
            });

            Assert.NotNull(thrown);

            KnownBug.StillBroken("#120",
                "a write against a damaged template fails at the write, not from inside " +
                "Dispose, and the archive is still closed cleanly so no half-written file is left",
                () => Assert.IsNotType<NullReferenceException>(thrown));
        }

        // ------------------------------------------------------------- damaged sheet XML

        /// <summary>Builds a workbook whose only sheet holds exactly the given rows XML.</summary>
        private static FileInfo SheetWith(TempWorkspace workspace, string rowsXml, params string[] sharedStrings)
        {
            var builder = new XlsxBuilder();
            if (sharedStrings.Length > 0) builder = builder.WithSharedStrings(sharedStrings);
            return builder.WithSheet("Sheet1", rowsXml).ToFile(workspace, "sheet.xlsx");
        }

        [Fact]
        public void ARowWhoseNumberIsNotANumber_IsReportedAsARowNumberProblem()
        {
            using var workspace = new TempWorkspace();
            var file = SheetWith(workspace, "<row r=\"not-a-number\"><c r=\"A1\"><v>1</v></c></row>");

            using var fastExcel = new FastExcel(file, true);
            var thrown = Record.Exception(() => fastExcel.Read(1).Rows.ToList());

            Assert.NotNull(thrown);
            Assert.Contains("Row Number", thrown.Message);
        }

        [Fact]
        public void ARowNumberTooLargeForAnInt_IsReportedMisleadingly()
        {
            using var workspace = new TempWorkspace();
            var file = SheetWith(workspace, "<row r=\"99999999999\"><c r=\"A1\"><v>1</v></c></row>");

            using var fastExcel = new FastExcel(file, true);
            var thrown = Record.Exception(() => fastExcel.Read(1).Rows.ToList());

            Assert.NotNull(thrown);

            // The parse is wrapped in catch(Exception) and rethrown as "Row Number not found",
            // so an overflow is reported as an absence. The real cause survives only as the
            // inner exception.
            Assert.Contains("Row Number", thrown.Message);
            Assert.IsType<OverflowException>(thrown.InnerException);
        }

        [Fact]
        public void ARowWithNoNumberAtAll_IsReportedAsARowNumberProblem()
        {
            using var workspace = new TempWorkspace();
            var file = SheetWith(workspace, "<row><c r=\"A1\"><v>1</v></c></row>");

            using var fastExcel = new FastExcel(file, true);
            var thrown = Record.Exception(() => fastExcel.Read(1).Rows.ToList());

            Assert.NotNull(thrown);
            Assert.Contains("Row Number", thrown.Message);
        }

        // ------------------------------------------------------------- damaged shared strings

        [Fact]
        public void ASharedStringIndexThatIsNotANumber_ReadsAsEmptyRatherThanFailing()
        {
            using var workspace = new TempWorkspace();
            var file = SheetWith(workspace,
                "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>not-an-index</v></c></row>",
                "value");

            using var fastExcel = new FastExcel(file, true);
            var value = fastExcel.Read(1).Rows.Single().Cells.Single().Value;

            // The library sees an unparseable index and returns "" - the source even carries a
            // TODO asking whether it should throw. A caller cannot tell this apart from a cell
            // that genuinely holds an empty string, so corrupt input reads as clean data.
            Assert.Equal(string.Empty, value);

            KnownBug.StillBroken("#121",
                "a shared-string index that is not a number is reported as corrupt input " +
                "rather than silently read as an empty cell",
                () => Assert.NotEqual(string.Empty, value));
        }

        [Fact]
        public void ASharedStringIndexPastTheEndOfTheTable_Throws()
        {
            using var workspace = new TempWorkspace();
            var file = SheetWith(workspace,
                "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>99</v></c></row>",
                "only-one-entry");

            using var fastExcel = new FastExcel(file, true);
            var thrown = Record.Exception(() => fastExcel.Read(1).Rows.Single().Cells.ToList());

            Assert.NotNull(thrown);

            KnownBug.StillBroken("#121",
                "a shared-string index past the end of the table is reported as a workbook " +
                "problem rather than surfacing a raw collection lookup failure",
                () => Assert.IsNotType<KeyNotFoundException>(thrown));
        }

        [Fact]
        public void ANegativeSharedStringIndex_Throws()
        {
            using var workspace = new TempWorkspace();
            var file = SheetWith(workspace,
                "<row r=\"1\"><c r=\"A1\" t=\"s\"><v>-1</v></c></row>",
                "entry");

            using var fastExcel = new FastExcel(file, true);

            Assert.ThrowsAny<Exception>(() => fastExcel.Read(1).Rows.Single().Cells.ToList());
        }

        // ------------------------------------------------------------- damaged cell references

        [Fact]
        public void ACellReferenceWithNoRowNumber_Throws()
        {
            using var workspace = new TempWorkspace();
            var file = SheetWith(workspace, "<row r=\"1\"><c r=\"A\"><v>1</v></c></row>");

            using var fastExcel = new FastExcel(file, true);
            var thrown = Record.Exception(() => fastExcel.Read(1).Rows.Single().Cells.ToList());

            Assert.NotNull(thrown);

            // "A" with no digits leaves an empty string for the row number, and Convert.ToInt32
            // of "" is a FormatException from deep inside the parse rather than a statement
            // about the cell.
            KnownBug.StillBroken("#121",
                "a cell reference missing its row number is reported as a bad reference, " +
                "naming the cell, rather than as a raw FormatException",
                () => Assert.IsNotType<FormatException>(thrown));
        }

        [Fact]
        public void ACellReferenceWithAnAbsoluteMarker_IsMisread()
        {
            using var workspace = new TempWorkspace();
            var file = SheetWith(workspace, "<row r=\"1\"><c r=\"$B$1\"><v>7</v></c></row>");

            using var fastExcel = new FastExcel(file, true);
            var cell = fastExcel.Read(1).Rows.Single().Cells.Single();

            // The reference is stripped with a regex that removes digits, so the dollar signs
            // survive into the column letters and the column number comes out wrong rather than
            // being rejected. $B should be column 2.
            KnownBug.StillBroken("#121",
                "a cell reference containing absolute markers resolves to the column it names",
                () => Assert.Equal(2, cell.ColumnNumber));
        }

        // ------------------------------------------------------------- constructor guards

        [Fact]
        public void AMissingInputFile_IsRejectedByName()
        {
            using var workspace = new TempWorkspace();
            var missing = new FileInfo(workspace.Path("does-not-exist.xlsx"));

            var thrown = Assert.Throws<FileNotFoundException>(() => new FastExcel(missing, true));
            Assert.Contains("does-not-exist.xlsx", thrown.Message);
        }

        [Fact]
        public void AMissingTemplate_IsRejectedByName()
        {
            using var workspace = new TempWorkspace();
            var missing = new FileInfo(workspace.Path("no-template.xlsx"));
            var output = workspace.NewFile();

            var thrown = Assert.Throws<FileNotFoundException>(() => new FastExcel(missing, output));
            Assert.Contains("no-template.xlsx", thrown.Message);
        }

        [Fact]
        public void AnOutputFileThatAlreadyExists_IsRefused()
        {
            using var workspace = new TempWorkspace();
            var template = ValidWorkbook(workspace);

            var existing = new FileInfo(workspace.Path("taken.xlsx"));
            File.WriteAllText(existing.FullName, "already here");

            var thrown = Record.Exception(() => new FastExcel(template, existing));

            Assert.NotNull(thrown);
            Assert.Contains("already exists", thrown.Message);
        }

        [Fact]
        public void ANullFile_IsRejectedWithoutExplainingWhy()
        {
            var thrown = Record.Exception(() => new FastExcel((FileInfo)null, true));

            Assert.NotNull(thrown);

            KnownBug.StillBroken("#119",
                "a null file argument is rejected with an ArgumentNullException naming the " +
                "parameter, rather than a bare NullReferenceException",
                () => Assert.IsType<ArgumentNullException>(thrown));
        }
    }
}
