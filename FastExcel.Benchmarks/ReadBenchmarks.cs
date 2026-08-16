using System;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using FastExcel.Tests.Infrastructure;

namespace FastExcel.Benchmarks
{
    /// <summary>
    /// Read throughput and memory across workbook sizes.
    /// <para>
    /// MemoryDiagnoser is the point of this suite as much as the timings: the library's stated
    /// selling point is a "small memory footprint while running", and #70 reports a 50 MB
    /// workbook consuming 12 GB. The Allocated column is what makes that concrete, and what
    /// will show the Phase 3 work landing.
    /// </para>
    /// </summary>
    [MemoryDiagnoser]
    public class ReadBenchmarks
    {
        [Params(1_000, 10_000, 50_000)]
        public int Rows;

        [Params(10, 20)]
        public int Columns;

        private FileInfo _sharedStrings;
        private FileInfo _numbers;

        [GlobalSetup]
        public void Setup()
        {
            // Generated once and cached on disk, so generation never lands inside a measurement.
            _sharedStrings = LargeWorkbook.Get(Rows, Columns, LargeWorkbook.Payload.SharedStrings);
            _numbers = LargeWorkbook.Get(Rows, Columns, LargeWorkbook.Payload.Numbers);
        }

        [Benchmark(Baseline = true, Description = "Read all cells (shared strings)")]
        public long ReadSharedStrings() => ReadEveryCell(_sharedStrings);

        [Benchmark(Description = "Read all cells (numeric)")]
        public long ReadNumbers() => ReadEveryCell(_numbers);

        [Benchmark(Description = "Enumerate rows without touching cells")]
        public long ReadRowsOnly()
        {
            long rows = 0;
            using var fastExcel = new FastExcel(_sharedStrings, true);
            foreach (var _ in fastExcel.Read(1).Rows) rows++;
            return rows;
        }

        internal static long ReadEveryCell(FileInfo file)
        {
            long cells = 0;
            using var fastExcel = new FastExcel(file, true);

            foreach (var row in fastExcel.Read(1).Rows)
            {
                foreach (var cell in row.Cells)
                {
                    if (cell.Value != null) cells++;
                }
            }

            return cells;
        }
    }

    /// <summary>
    /// Isolates the cost of declaring defined names.
    /// <para>
    /// Cell construction calls FindColumnName and FindCellNames for every cell, and both run a
    /// LINQ scan across the whole defined-name dictionary, allocating strings and a list each
    /// time. That makes reading O(cells x names) rather than O(cells) — and the per-cell cost
    /// is paid even when a workbook declares no names at all. Print areas and tables produce
    /// defined names routinely, so this is not an exotic configuration.
    /// </para>
    /// </summary>
    [MemoryDiagnoser]
    public class DefinedNameScalingBenchmarks
    {
        [Params(0, 10, 100)]
        public int DefinedNames;

        private FileInfo _file;

        [GlobalSetup]
        public void Setup() =>
            _file = LargeWorkbook.Get(10_000, 10, LargeWorkbook.Payload.SharedStrings, DefinedNames);

        [Benchmark(Description = "Read 100k cells")]
        public long Read() => ReadBenchmarks.ReadEveryCell(_file);
    }
}
