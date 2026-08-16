using System;
using System.Diagnostics;
using Xunit;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// Measures how much a piece of work allocates.
    /// <para>
    /// Allocated bytes are the right currency for a performance assertion in CI. Wall-clock
    /// time on a shared runner varies by several times depending on the neighbours, so a
    /// duration threshold either flakes or is set so loose it catches nothing. Allocation
    /// counts are essentially deterministic for the same input on any machine, which makes
    /// "this read must not cost more than N bytes per cell" a threshold that actually holds.
    /// </para>
    /// </summary>
    public static class AllocationProbe
    {
        public sealed class Measurement
        {
            public long AllocatedBytes { get; set; }
            public long Items { get; set; }
            public TimeSpan Elapsed { get; set; }

            public double BytesPerItem => Items == 0 ? 0 : AllocatedBytes / (double)Items;
            public double AllocatedMegabytes => AllocatedBytes / 1048576.0;

            public override string ToString() =>
                $"{Items:N0} items, {AllocatedMegabytes:N1} MB allocated, " +
                $"{BytesPerItem:N0} bytes/item, {Elapsed.TotalMilliseconds:N0} ms";
        }

        /// <summary>
        /// Runs <paramref name="work"/> and reports what it allocated. The delegate returns the
        /// number of items it processed, so the result can be normalised per cell.
        /// </summary>
        public static Measurement Measure(Func<long> work, int warmupRuns = 1)
        {
            // The first execution pays for JIT compilation of everything on the path, which
            // would otherwise be attributed to the workload.
            for (var i = 0; i < warmupRuns; i++) work();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            var before = GC.GetAllocatedBytesForCurrentThread();
            var stopwatch = Stopwatch.StartNew();

            var items = work();

            stopwatch.Stop();
            var allocated = GC.GetAllocatedBytesForCurrentThread() - before;

            return new Measurement
            {
                AllocatedBytes = allocated,
                Items = items,
                Elapsed = stopwatch.Elapsed
            };
        }

        /// <summary>Reads every cell of a workbook and returns how many it saw.</summary>
        public static long ReadEveryCell(System.IO.FileInfo file)
        {
            long cells = 0;
            using var fastExcel = new FastExcel(file, true);
            var worksheet = fastExcel.Read(1);

            foreach (var row in worksheet.Rows)
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
    /// A test whose result depends on wall-clock timing. Skipped unless FASTEXCEL_PERF=1.
    /// <para>
    /// These are genuinely useful locally and on a quiet machine, but a shared CI runner can
    /// be several times slower without anything being wrong, so gating merges on them would
    /// mean either flaky builds or thresholds too loose to detect a regression. The
    /// allocation-based assertions run everywhere and carry the CI signal instead.
    /// </para>
    /// </summary>
    public sealed class PerformanceFactAttribute : FactAttribute
    {
        public PerformanceFactAttribute()
        {
            if (Environment.GetEnvironmentVariable("FASTEXCEL_PERF") != "1")
            {
                Skip = "Timing-sensitive. Set FASTEXCEL_PERF=1 to run.";
            }
        }
    }
}
