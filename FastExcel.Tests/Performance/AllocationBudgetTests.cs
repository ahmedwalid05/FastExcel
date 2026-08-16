using System;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;
using Xunit.Abstractions;

namespace FastExcel.Tests.Performance
{
    /// <summary>
    /// Performance assertions expressed in allocated bytes rather than elapsed time, so they
    /// hold on any machine and can gate CI.
    /// <para>
    /// The library's stated selling points are speed and a "small memory footprint while
    /// running", and #70 reports a 50 MB workbook consuming 12 GB. Neither claim had any
    /// coverage, so nothing stopped the footprint drifting.
    /// </para>
    /// </summary>
    public class AllocationBudgetTests
    {
        private readonly ITestOutputHelper _output;

        public AllocationBudgetTests(ITestOutputHelper output) => _output = output;

        /// <summary>
        /// Ceiling for what reading one cell may allocate.
        /// <para>
        /// Measured at roughly 1,280 bytes per cell today, near-constant across workbook sizes
        /// and both payload types. The budget sits above that to leave room for ordinary
        /// variation between runtimes, while still failing loudly on a real regression. It is
        /// deliberately a ratchet: when the read path gets cheaper, lower this number.
        /// </para>
        /// </summary>
        private const double CurrentBytesPerCellBudget = 1_800;

        /// <summary>
        /// What a lightweight reader ought to cost per cell. Not met today — see
        /// <see cref="ReadingACell_MeetsTheTargetFootprint"/>.
        /// </summary>
        private const double TargetBytesPerCell = 256;

        [Fact]
        public void ReadingACell_StaysWithinTheCurrentBudget()
        {
            var file = LargeWorkbook.Get(rows: 2_000, columns: 10);

            var measurement = AllocationProbe.Measure(() => AllocationProbe.ReadEveryCell(file));
            _output.WriteLine(measurement.ToString());

            Assert.Equal(20_000, measurement.Items);
            Assert.True(measurement.BytesPerItem < CurrentBytesPerCellBudget,
                $"Reading a cell allocated {measurement.BytesPerItem:N0} bytes, over the " +
                $"{CurrentBytesPerCellBudget:N0} byte budget. Something on the read path got " +
                $"more expensive. Full measurement: {measurement}");
        }

        [Fact]
        public void ReadingACell_MeetsTheTargetFootprint()
        {
            var file = LargeWorkbook.Get(rows: 2_000, columns: 10);
            var measurement = AllocationProbe.Measure(() => AllocationProbe.ReadEveryCell(file));
            _output.WriteLine(measurement.ToString());

            KnownBug.StillBroken("#70",
                $"reading a cell allocates under {TargetBytesPerCell:N0} bytes. Today it costs " +
                "about 1,280 — roughly 400x the size of the file being read — because every " +
                "cell retains its XElement, re-runs two regexes, and performs two LINQ scans " +
                "over the defined-name dictionary that allocate a list and several strings " +
                "even when the workbook declares no defined names at all",
                () => Assert.True(measurement.BytesPerItem < TargetBytesPerCell,
                    $"still {measurement.BytesPerItem:N0} bytes per cell"));
        }

        [Fact]
        public void AllocationPerCell_DoesNotGrowWithWorkbookSize()
        {
            // The cost of reading a cell must not depend on how many other cells there are.
            // If it does, the read path is quadratic and large files are unusable.
            var small = AllocationProbe.Measure(
                () => AllocationProbe.ReadEveryCell(LargeWorkbook.Get(rows: 1_000, columns: 10)));
            var large = AllocationProbe.Measure(
                () => AllocationProbe.ReadEveryCell(LargeWorkbook.Get(rows: 8_000, columns: 10)));

            _output.WriteLine($"1,000 rows: {small}");
            _output.WriteLine($"8,000 rows: {large}");

            var growth = large.BytesPerItem / small.BytesPerItem;

            Assert.True(growth < 1.25,
                $"Per-cell allocation grew {growth:N2}x when the workbook got 8x bigger " +
                $"({small.BytesPerItem:N0} -> {large.BytesPerItem:N0} bytes/cell), which means " +
                "the read path is super-linear in the number of cells.");
        }

        [Fact]
        public void TotalAllocation_ScalesLinearlyWithCellCount()
        {
            var small = AllocationProbe.Measure(
                () => AllocationProbe.ReadEveryCell(LargeWorkbook.Get(rows: 1_000, columns: 10)));
            var large = AllocationProbe.Measure(
                () => AllocationProbe.ReadEveryCell(LargeWorkbook.Get(rows: 8_000, columns: 10)));

            var cellRatio = large.Items / (double)small.Items;
            var allocationRatio = large.AllocatedBytes / (double)small.AllocatedBytes;

            _output.WriteLine($"cells x{cellRatio:N1}, allocation x{allocationRatio:N1}");

            // 8x the cells should cost about 8x the memory, never 64x.
            Assert.InRange(allocationRatio, cellRatio * 0.75, cellRatio * 1.25);
        }

        [Theory]
        [InlineData(LargeWorkbook.Payload.SharedStrings)]
        [InlineData(LargeWorkbook.Payload.Numbers)]
        public void BothPayloadTypes_StayWithinBudget(LargeWorkbook.Payload payload)
        {
            var file = LargeWorkbook.Get(rows: 2_000, columns: 10, payload: payload);

            var measurement = AllocationProbe.Measure(() => AllocationProbe.ReadEveryCell(file));
            _output.WriteLine($"{payload}: {measurement}");

            Assert.True(measurement.BytesPerItem < CurrentBytesPerCellBudget,
                $"{payload} cells allocated {measurement.BytesPerItem:N0} bytes each. {measurement}");
        }

        [Fact]
        public void ReadingAWorkbookWithNoDefinedNames_DoesNotPayForDefinedNameLookup()
        {
            // Cell construction calls FindColumnName and FindCellNames for every cell. Both
            // scan the defined-name dictionary and allocate regardless of whether it is empty,
            // so a workbook that declares no names still pays the full per-cell cost.
            var withoutNames = AllocationProbe.Measure(
                () => AllocationProbe.ReadEveryCell(
                    LargeWorkbook.Get(rows: 2_000, columns: 10, definedNameCount: 0)));

            _output.WriteLine($"no defined names: {withoutNames}");

            KnownBug.StillBroken("#70",
                $"a workbook declaring no defined names skips the per-cell name lookup " +
                $"entirely and reads for well under {TargetBytesPerCell:N0} bytes per cell",
                () => Assert.True(withoutNames.BytesPerItem < TargetBytesPerCell,
                    $"still {withoutNames.BytesPerItem:N0} bytes per cell with an empty name table"));
        }

        [Fact]
        public void AddingDefinedNames_DoesNotMultiplyTheCostOfReadingEveryCell()
        {
            // FindCellNames runs a LINQ scan across all defined names for every single cell,
            // making the read O(cells x definedNames) rather than O(cells).
            var noNames = AllocationProbe.Measure(
                () => AllocationProbe.ReadEveryCell(
                    LargeWorkbook.Get(rows: 1_000, columns: 10, definedNameCount: 0)));
            var manyNames = AllocationProbe.Measure(
                () => AllocationProbe.ReadEveryCell(
                    LargeWorkbook.Get(rows: 1_000, columns: 10, definedNameCount: 100)));

            _output.WriteLine($"0 defined names:   {noNames}");
            _output.WriteLine($"100 defined names: {manyNames}");

            var growth = manyNames.BytesPerItem / noNames.BytesPerItem;
            _output.WriteLine($"growth: {growth:N2}x");

            KnownBug.StillBroken("#70 / #86",
                "declaring 100 defined names does not measurably change the cost of reading a " +
                "cell — name lookup should be a dictionary hit, not a linear scan per cell",
                () => Assert.True(growth < 1.5,
                    $"per-cell allocation grew {growth:N2}x when 100 defined names were added"));
        }
    }
}
