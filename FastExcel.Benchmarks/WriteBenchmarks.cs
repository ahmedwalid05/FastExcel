using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using FastExcel.Tests.Infrastructure;

namespace FastExcel.Benchmarks
{
    /// <summary>
    /// Write throughput across row counts. Writing needs a template workbook, which is created
    /// once in setup rather than per iteration.
    /// </summary>
    [MemoryDiagnoser]
    public class WriteBenchmarks
    {
        [Params(1_000, 10_000, 50_000)]
        public int Rows;

        private const int Columns = 10;

        private string _workingDirectory;
        private FileInfo _template;
        private List<object[]> _numericRows;
        private List<object[]> _stringRows;

        [GlobalSetup]
        public void Setup()
        {
            _workingDirectory = Path.Combine(Path.GetTempPath(), "fastexcel-write-bench", Guid.NewGuid().ToString("n"));
            Directory.CreateDirectory(_workingDirectory);

            // A one-row workbook doubles as an empty write template.
            _template = LargeWorkbook.Get(1, Columns);

            _numericRows = Enumerable.Range(1, Rows)
                .Select(r => Enumerable.Range(1, Columns).Select(c => (object)(r * c)).ToArray())
                .ToList();

            _stringRows = Enumerable.Range(1, Rows)
                .Select(r => Enumerable.Range(1, Columns).Select(c => (object)$"row {r} col {c}").ToArray())
                .ToList();
        }

        [GlobalCleanup]
        public void Cleanup()
        {
            try { Directory.Delete(_workingDirectory, recursive: true); } catch (IOException) { }
        }

        private long Write(List<object[]> rows)
        {
            var output = new FileInfo(Path.Combine(_workingDirectory, Guid.NewGuid().ToString("n") + ".xlsx"));

            using (var fastExcel = new FastExcel(_template, output))
            {
                var worksheet = new Worksheet();
                foreach (var row in rows) worksheet.AddRow(row);
                fastExcel.Write(worksheet, 1);
            }

            output.Refresh();
            var length = output.Length;
            output.Delete();
            return length;
        }

        [Benchmark(Baseline = true, Description = "Write numeric rows")]
        public long WriteNumbers() => Write(_numericRows);

        [Benchmark(Description = "Write string rows")]
        public long WriteStrings() => Write(_stringRows);
    }
}
