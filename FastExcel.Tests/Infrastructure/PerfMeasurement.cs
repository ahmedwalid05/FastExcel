using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// One measured scenario.
    /// <para>
    /// Three memory numbers are recorded because they answer three different questions and can
    /// move independently. An optimisation that removes garbage churn can cut
    /// <see cref="AllocatedBytes"/> several-fold while leaving <see cref="RetainedBytes"/> —
    /// the memory a caller cannot get back while holding the worksheet — completely unchanged.
    /// Reporting only allocations, as BenchmarkDotNet's MemoryDiagnoser does, makes that
    /// invisible, which is why #70 could not be validated from the previous output.
    /// </para>
    /// </summary>
    public sealed class PerfResult
    {
        public string Scenario { get; set; }
        public long Cells { get; set; }
        public long FileBytes { get; set; }

        /// <summary>Total bytes allocated over the read, including everything the GC reclaimed.</summary>
        public long AllocatedBytes { get; set; }

        /// <summary>
        /// Live bytes still held once the whole worksheet is materialised. This is the figure a
        /// user sees in Task Manager and the one #70 is about. Measured after a forced
        /// collection, which makes it deterministic to the byte.
        /// </summary>
        public long RetainedBytes { get; set; }

        public double ElapsedMilliseconds { get; set; }

        // ---- derived; not serialised, computed on demand ----
        public double BytesPerCellAllocated => Cells == 0 ? 0 : AllocatedBytes / (double)Cells;
        public double BytesPerCellRetained => Cells == 0 ? 0 : RetainedBytes / (double)Cells;
        public double AllocatedToFileRatio => FileBytes == 0 ? 0 : AllocatedBytes / (double)FileBytes;
        public double RetainedToFileRatio => FileBytes == 0 ? 0 : RetainedBytes / (double)FileBytes;
        public double CellsPerSecond => ElapsedMilliseconds <= 0 ? 0 : Cells / (ElapsedMilliseconds / 1000.0);

        public static double Megabytes(long bytes) => bytes / 1048576.0;
    }

    /// <summary>
    /// Measures a scenario end to end.
    /// </summary>
    public static class PerfMeasurement
    {
        /// <summary>
        /// Reads a scenario's workbook, fully materialising it, and reports what it cost.
        /// </summary>
        /// <param name="warmup">
        /// Run once first so JIT compilation and the OS file cache are not attributed to the
        /// measurement. Worth skipping only when deliberately measuring a cold start.
        /// </param>
        public static PerfResult MeasureRead(PerfScenario scenario, bool warmup = true)
        {
            var file = scenario.Workbook();

            if (warmup) ReadFully(file, out _);

            Settle();
            var baselineLive = GC.GetTotalMemory(forceFullCollection: true);
            var allocatedBefore = GC.GetAllocatedBytesForCurrentThread();

            var stopwatch = Stopwatch.StartNew();
            var rows = ReadFully(file, out var cells);
            stopwatch.Stop();

            var allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;

            // Everything is still reachable through `rows`, so this is peak retention.
            var retained = GC.GetTotalMemory(forceFullCollection: true) - baselineLive;
            GC.KeepAlive(rows);

            return new PerfResult
            {
                Scenario = scenario.Name,
                Cells = cells,
                FileBytes = file.Length,
                AllocatedBytes = allocated,
                RetainedBytes = retained,
                ElapsedMilliseconds = stopwatch.Elapsed.TotalMilliseconds
            };
        }

        /// <summary>
        /// Materialises every row and cell, as a caller who keeps the worksheet around would.
        /// Both Rows and Cells are lazy iterators, so each has to be forced explicitly.
        /// </summary>
        private static List<Row> ReadFully(FileInfo file, out long cells)
        {
            long counted = 0;
            using var fastExcel = new FastExcel(file, true);

            var rows = fastExcel.Read(1).Rows.ToList();
            foreach (var row in rows)
            {
                var materialised = (row.Cells ?? Enumerable.Empty<Cell>()).ToList();
                row.Cells = materialised;
                counted += materialised.Count;
            }

            cells = counted;
            return rows;
        }

        private static void Settle()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
