using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using FastExcel.Tests.Infrastructure;
using MiniExcelLibs;

namespace FastExcel.Benchmarks
{
    /// <summary>
    /// FastExcel against the actively maintained alternatives.
    /// <para>
    /// The package describes itself as "a fast way to read and write" xlsx files, and that
    /// claim was made in 2015. MiniExcel and Sylvan did not exist yet. Measuring against them
    /// is the only way to know whether the claim still holds, and it is the evidence the
    /// revive-versus-deprecate decision actually turns on — so it belongs in the repository
    /// rather than in someone's head.
    /// </para>
    /// <para>
    /// Every library is asked to do the same job: visit every cell and materialise its value.
    /// They are not identical in what they offer — Sylvan is a forward-only reader and does
    /// not write, ClosedXML builds a full mutable object model — so read these as "cost of
    /// getting the data out", not as a feature comparison.
    /// </para>
    /// </summary>
    [MemoryDiagnoser]
    public class ReadComparisonBenchmarks
    {
        // Kept modest on purpose: ClosedXML builds a complete object model and becomes very
        // expensive well before the sizes FastExcel is benchmarked at on its own.
        [Params(50_000, 200_000)]
        public int Cells;

        private FileInfo _file;

        [GlobalSetup]
        public void Setup() => _file = LargeWorkbook.Get(Cells / 10, 10);

        [Benchmark(Baseline = true, Description = "FastExcel")]
        public long FastExcelRead()
        {
            long cells = 0;
            using var fastExcel = new FastExcel(_file, true);
            foreach (var row in fastExcel.Read(1).Rows)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.Value != null) cells++;
                }
            }
            return cells;
        }

        [Benchmark(Description = "MiniExcel")]
        public long MiniExcelRead()
        {
            long cells = 0;
            foreach (IDictionary<string, object> row in MiniExcel.Query(_file.FullName, useHeaderRow: false))
            {
                foreach (var value in row.Values)
                {
                    if (value != null) cells++;
                }
            }
            return cells;
        }

        [Benchmark(Description = "ClosedXML")]
        public long ClosedXmlRead()
        {
            long cells = 0;
            using var workbook = new XLWorkbook(_file.FullName);
            foreach (var row in workbook.Worksheet(1).RowsUsed())
            {
                foreach (var cell in row.CellsUsed())
                {
                    if (!cell.Value.IsBlank) cells++;
                }
            }
            return cells;
        }

        [Benchmark(Description = "Sylvan.Data.Excel")]
        public long SylvanRead()
        {
            long cells = 0;
            using var reader = Sylvan.Data.Excel.ExcelDataReader.Create(_file.FullName);
            while (reader.Read())
            {
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    if (!reader.IsDBNull(i)) cells++;
                }
            }
            return cells;
        }
    }

    /// <summary>
    /// Write throughput against the same alternatives. Sylvan is absent because it is a reader.
    /// </summary>
    [MemoryDiagnoser]
    public class WriteComparisonBenchmarks
    {
        [Params(10_000, 50_000)]
        public int Rows;

        private const int Columns = 10;

        private string _directory;
        private FileInfo _template;
        private List<object[]> _rows;
        private List<Dictionary<string, object>> _rowsAsDictionaries;

        [GlobalSetup]
        public void Setup()
        {
            _directory = Path.Combine(Path.GetTempPath(), "fastexcel-write-comparison", Guid.NewGuid().ToString("n"));
            Directory.CreateDirectory(_directory);

            // FastExcel cannot create a workbook without a template (#74), so one is provided.
            // The other libraries create their own; that asymmetry is itself a finding.
            _template = LargeWorkbook.Get(1, Columns);

            _rows = Enumerable.Range(1, Rows)
                .Select(r => Enumerable.Range(1, Columns).Select(c => (object)(r * c)).ToArray())
                .ToList();

            _rowsAsDictionaries = _rows
                .Select(r => r.Select((v, i) => new { Key = "col" + i, Value = v })
                              .ToDictionary(x => x.Key, x => x.Value))
                .ToList();
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            try { Directory.Delete(_directory, recursive: true); } catch (IOException) { }
        }

        private string NewPath() => Path.Combine(_directory, Guid.NewGuid().ToString("n") + ".xlsx");

        private static long SizeThenDelete(string path)
        {
            var info = new FileInfo(path);
            var length = info.Exists ? info.Length : 0;
            if (info.Exists) info.Delete();
            return length;
        }

        [Benchmark(Baseline = true, Description = "FastExcel")]
        public long FastExcelWrite()
        {
            var path = NewPath();
            using (var fastExcel = new FastExcel(_template, new FileInfo(path)))
            {
                var worksheet = new Worksheet();
                foreach (var row in _rows) worksheet.AddRow(row);
                fastExcel.Write(worksheet, 1);
            }
            return SizeThenDelete(path);
        }

        [Benchmark(Description = "MiniExcel")]
        public long MiniExcelWrite()
        {
            var path = NewPath();
            MiniExcel.SaveAs(path, _rowsAsDictionaries);
            return SizeThenDelete(path);
        }

        [Benchmark(Description = "ClosedXML")]
        public long ClosedXmlWrite()
        {
            var path = NewPath();
            using (var workbook = new XLWorkbook())
            {
                var sheet = workbook.Worksheets.Add("Sheet1");
                for (var r = 0; r < _rows.Count; r++)
                {
                    var row = _rows[r];
                    for (var c = 0; c < row.Length; c++)
                    {
                        sheet.Cell(r + 1, c + 1).Value = Convert.ToInt32(row[c]);
                    }
                }
                workbook.SaveAs(path);
            }
            return SizeThenDelete(path);
        }
    }
}
