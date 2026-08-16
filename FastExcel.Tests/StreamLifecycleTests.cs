using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;

namespace FastExcel.Tests
{
    /// <summary>
    /// How FastExcel acquires, shares and releases the underlying package. Reading from a
    /// stream you do not own — a blob download, an HTTP response, an embedded resource — is a
    /// mainstream scenario that the current constructors cannot express.
    /// </summary>
    public class StreamLifecycleTests
    {
        private static FileInfo SimpleWorkbook(TempWorkspace workspace, string name = "book.xlsx")
            => new XlsxBuilder()
                .WithSharedStrings("hello")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToFile(workspace, name);

        // ------------------------------------------------------------ #75 read-only input

        [Fact]
        public void ReadOnlyStream_CanBeRead()
        {
            using var workspace = new TempWorkspace();
            var file = SimpleWorkbook(workspace);

            using var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);

            KnownBug.StillBroken("#75",
                "a workbook can be read from a stream that is not writable — the single-stream " +
                "constructor should expose the readOnly flag the four-argument one already has, " +
                "so Azure blob and HTTP response streams work without being copied first",
                () =>
                {
                    using var fastExcel = new FastExcel(stream);
                    Assert.Equal("hello", fastExcel.Read(1).Rows.Single().Cells.Single().Value);
                });
        }

        [Fact]
        public void ReadOnlyStream_WorksViaTheExplicitReadOnlyConstructor()
        {
            using var workspace = new TempWorkspace();
            var file = SimpleWorkbook(workspace);

            using var stream = new FileStream(file.FullName, FileMode.Open, FileAccess.Read);

            // The capability exists; only the convenience overload is missing it. This is the
            // workaround to document in the meantime.
            using var fastExcel = new FastExcel(null, stream, updateExisting: true, readOnly: true);

            Assert.Equal("hello", fastExcel.Read(1).Rows.Single().Cells.Single().Value);
        }

        [Fact]
        public void SeekableMemoryStream_CanBeRead()
        {
            using var workspace = new TempWorkspace();
            var bytes = new XlsxBuilder()
                .WithSharedStrings("in-memory")
                .WithSheet("Sheet1", XlsxBuilder.Row(1, XlsxBuilder.SharedCell("A1", 0)))
                .ToBytes();

            using var stream = new MemoryStream(bytes, writable: true);
            using var fastExcel = new FastExcel(null, stream, updateExisting: true, readOnly: true);

            Assert.Equal("in-memory", fastExcel.Read(1).Rows.Single().Cells.Single().Value);
        }

        // --------------------------------------------------------- #69 read-then-update

        [Fact]
        public void UpdatingASheetAfterReadingIt_DoesNotThrow()
        {
            using var workspace = new TempWorkspace();
            var file = SimpleWorkbook(workspace);

            KnownBug.StillBroken("#69",
                "reading a sheet, changing a cell and calling Update on the same instance " +
                "works; today Update re-reads the sheet, which reopens a zip entry while an " +
                "earlier one is still open and throws \"Entries cannot be opened multiple " +
                "times in Update mode\"",
                () =>
                {
                    using var fastExcel = new FastExcel(file);
                    var sheet = fastExcel.Read(1);
                    var rows = sheet.Rows.ToList();
                    foreach (var row in rows)
                    {
                        row.Cells = row.Cells.ToList();
                    }
                    sheet.Rows = rows;

                    fastExcel.Update(sheet, 1);
                });
        }

        [Fact]
        public void UpdatingACellThatAlreadyHasAValue_ReplacesIt()
        {
            using var workspace = new TempWorkspace();
            var file = SimpleWorkbook(workspace); // A1 already contains "hello"

            using (var fastExcel = new FastExcel(file))
            {
                var worksheet = new Worksheet();
                worksheet.AddRow("replacement");
                fastExcel.Update(worksheet, 1);
            }

            file.Refresh();
            using var verify = new FastExcel(file, true);
            var actual = verify.Read(1).Rows.First().Cells.First().Value;

            KnownBug.StillBroken("#69 / #71",
                "Update replaces the value of a cell that already has one. Worksheet.Merge " +
                "documents that \"the parameter takes precedence\", but MergeRows is called on " +
                "the incoming worksheet with the existing rows as its argument, so " +
                "Cell.Merge copies the OLD value over the NEW one and the update is silently " +
                "discarded for every cell that was already populated",
                () => Assert.Equal("replacement", actual));
        }

        [Fact]
        public void UpdatingAddsCellsThatDidNotExistBefore()
        {
            using var workspace = new TempWorkspace();
            var file = SimpleWorkbook(workspace); // only A1 is populated

            using (var fastExcel = new FastExcel(file))
            {
                var worksheet = new Worksheet();
                // Row 2 does not exist in the original, so there is nothing to merge against.
                worksheet.AddRow("first row");
                worksheet.AddRow("brand new row");
                fastExcel.Update(worksheet, 1);
            }

            file.Refresh();
            using var verify = new FastExcel(file, true);
            var rows = verify.Read(1).Rows.ToList();

            // Additive updates land correctly; it is only overwrites that are dropped. That
            // asymmetry is what makes the bug above so easy to miss.
            Assert.Equal("brand new row", rows.Single(r => r.RowNumber == 2).Cells.First().Value);
        }

        // ------------------------------------------------------- construction guard rails

        [Fact]
        public void OutputFileThatAlreadyExists_IsRejected()
        {
            using var workspace = new TempWorkspace();
            var template = SimpleWorkbook(workspace, "template.xlsx");
            var existing = SimpleWorkbook(workspace, "already-here.xlsx");

            var exception = Assert.Throws<Exception>(() =>
            {
                using (new FastExcel(template, existing)) { }
            });

            Assert.Contains("already exists", exception.Message);
        }

        [Fact]
        public void MissingInputFile_IsRejected()
        {
            using var workspace = new TempWorkspace();
            var missing = new FileInfo(workspace.Path("nope.xlsx"));

            Assert.Throws<FileNotFoundException>(() =>
            {
                using (new FastExcel(missing, true)) { }
            });
        }

        [Fact]
        public void MissingTemplateFile_IsRejected()
        {
            using var workspace = new TempWorkspace();
            var missingTemplate = new FileInfo(workspace.Path("no-template.xlsx"));
            var output = workspace.NewFile();

            Assert.Throws<FileNotFoundException>(() =>
            {
                using (new FastExcel(missingTemplate, output)) { }
            });
        }

        [Fact]
        public void CreatingAWorkbookWithNoTemplate_FailsWithAClearMessage()
        {
            using var workspace = new TempWorkspace();
            using var output = new MemoryStream();

            // Reported in #74 as a NullReferenceException. On the stream path the failure is
            // already clear and actionable, so this pins that behaviour rather than treating it
            // as broken. What #74 actually asks for is the ability to create a workbook with no
            // template at all, which remains an open feature request.
            var exception = Record.Exception(() =>
            {
                using var fastExcel = new FastExcel(null, output);
                var worksheet = new Worksheet();
                worksheet.AddRow("value");
                fastExcel.Write(worksheet, 1);
            });

            Assert.NotNull(exception);
            Assert.IsNotType<NullReferenceException>(exception);
            Assert.Contains("template", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        // ------------------------------------------------------------------- disposal

        [Fact]
        public void DisposingTwice_IsSafe()
        {
            using var workspace = new TempWorkspace();
            var file = SimpleWorkbook(workspace);

            var fastExcel = new FastExcel(file, true);
            fastExcel.Read(1);
            fastExcel.Dispose();

            Assert.Null(Record.Exception(() => fastExcel.Dispose()));
        }

        [Fact]
        public void ReadOnlyInstance_DoesNotModifyTheFile()
        {
            using var workspace = new TempWorkspace();
            var file = SimpleWorkbook(workspace);
            var before = File.ReadAllBytes(file.FullName);

            using (var fastExcel = new FastExcel(file, true))
            {
                var rows = fastExcel.Read(1).Rows.ToList();
                foreach (var row in rows) row.Cells.ToList();
            }

            Assert.Equal(before, File.ReadAllBytes(file.FullName));
        }
    }
}
