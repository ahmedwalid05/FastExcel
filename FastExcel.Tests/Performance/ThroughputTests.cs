using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;
using Xunit.Abstractions;

namespace FastExcel.Tests.Performance
{
    /// <summary>
    /// Wall-clock assertions. These are skipped by default and run with FASTEXCEL_PERF=1,
    /// because a shared CI runner's speed varies enough that gating merges on elapsed time
    /// means either flaky builds or thresholds too loose to catch anything. The
    /// allocation-based tests carry the CI signal; these are for a quiet machine.
    /// <para>
    /// Thresholds are set well above measured behaviour on purpose — they exist to catch an
    /// order-of-magnitude regression, not to police small fluctuations.
    /// </para>
    /// </summary>
    public class ThroughputTests
    {
        private readonly ITestOutputHelper _output;

        public ThroughputTests(ITestOutputHelper output) => _output = output;

        private static (TimeSpan Elapsed, long Cells) TimeRead(FileInfo file)
        {
            AllocationProbe.ReadEveryCell(file); // warm up JIT and the file cache

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            var stopwatch = Stopwatch.StartNew();
            var cells = AllocationProbe.ReadEveryCell(file);
            stopwatch.Stop();

            return (stopwatch.Elapsed, cells);
        }

        [PerformanceFact]
        public void ReadingOneMillionCells_CompletesWithinAGenerousBudget()
        {
            var file = LargeWorkbook.Get(rows: 50_000, columns: 20);

            var (elapsed, cells) = TimeRead(file);
            var perSecond = cells / elapsed.TotalSeconds;

            _output.WriteLine($"{cells:N0} cells in {elapsed.TotalSeconds:N1}s ({perSecond:N0} cells/sec)");

            // Measured at roughly 230,000 cells/sec. A 10x slowdown is a real regression.
            Assert.True(perSecond > 23_000,
                $"Read throughput fell to {perSecond:N0} cells/sec, an order of magnitude below " +
                "the ~230,000 cells/sec this managed when the budget was set.");
        }

        [PerformanceFact]
        public void ReadTime_ScalesLinearlyWithCellCount()
        {
            var small = TimeRead(LargeWorkbook.Get(rows: 5_000, columns: 20));
            var large = TimeRead(LargeWorkbook.Get(rows: 40_000, columns: 20));

            var cellRatio = large.Cells / (double)small.Cells;
            var timeRatio = large.Elapsed.TotalSeconds / small.Elapsed.TotalSeconds;

            _output.WriteLine($"{small.Cells:N0} cells: {small.Elapsed.TotalMilliseconds:N0} ms");
            _output.WriteLine($"{large.Cells:N0} cells: {large.Elapsed.TotalMilliseconds:N0} ms");
            _output.WriteLine($"cells x{cellRatio:N1}, time x{timeRatio:N1}");

            // 8x the cells taking much more than 8x the time would mean super-linear parsing.
            // The allowance is wide because this is timing on an unknown machine.
            Assert.True(timeRatio < cellRatio * 2.0,
                $"Reading {cellRatio:N0}x more cells took {timeRatio:N1}x longer, which suggests " +
                "the read path is super-linear in the number of cells.");
        }

        [PerformanceFact]
        public void DefinedNames_DoNotDominateReadTime()
        {
            var without = TimeRead(LargeWorkbook.Get(rows: 5_000, columns: 10, definedNameCount: 0));
            var with = TimeRead(LargeWorkbook.Get(rows: 5_000, columns: 10, definedNameCount: 100));

            var slowdown = with.Elapsed.TotalSeconds / without.Elapsed.TotalSeconds;

            _output.WriteLine($"0 names:   {without.Elapsed.TotalMilliseconds:N0} ms");
            _output.WriteLine($"100 names: {with.Elapsed.TotalMilliseconds:N0} ms");
            _output.WriteLine($"slowdown:  {slowdown:N1}x");

            KnownBug.StillBroken("#70 / #86",
                "adding 100 defined names does not meaningfully slow down reading a workbook. " +
                "Cell construction currently runs a LINQ scan over every defined name for " +
                "every cell, making the read O(cells x names) instead of O(cells)",
                () => Assert.True(slowdown < 1.5,
                    $"reading was {slowdown:N1}x slower with 100 defined names declared"));
        }

        [PerformanceFact]
        public void WritingOneHundredThousandCells_CompletesWithinAGenerousBudget()
        {
            using var workspace = new TempWorkspace();
            var rows = Enumerable.Range(1, 10_000)
                .Select(r => Enumerable.Range(1, 10).Select(c => (object)(r * c)).ToArray())
                .ToList();

            var output = workspace.NewFile();
            var stopwatch = Stopwatch.StartNew();

            using (var fastExcel = new FastExcel(TestFixtures.Template, output))
            {
                var worksheet = new Worksheet();
                foreach (var row in rows) worksheet.AddRow(row);
                fastExcel.Write(worksheet, 1);
            }

            stopwatch.Stop();
            var perSecond = 100_000 / stopwatch.Elapsed.TotalSeconds;

            _output.WriteLine($"100,000 cells written in {stopwatch.Elapsed.TotalSeconds:N1}s ({perSecond:N0} cells/sec)");

            Assert.True(perSecond > 10_000,
                $"Write throughput fell to {perSecond:N0} cells/sec.");
        }
    }
}
