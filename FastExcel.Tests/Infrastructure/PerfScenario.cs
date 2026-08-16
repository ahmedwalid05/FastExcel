using System;
using System.Collections.Generic;
using System.Linq;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// The catalogue of workbook shapes that performance work measures.
    /// <para>
    /// Shared by the gated tests and the BenchmarkDotNet suite so both measure exactly the same
    /// files, and defined in one place so adding a shape adds it everywhere.
    /// </para>
    /// </summary>
    public sealed class PerfScenario
    {
        /// <summary>Stable identifier. Used as the key in the committed baseline, so renaming one resets its history.</summary>
        public string Name { get; }

        public int Rows { get; }
        public int Columns { get; }
        public LargeWorkbook.Payload Payload { get; }
        public int DefinedNames { get; }

        /// <summary>Whether this shape is cheap enough to run on every build.</summary>
        public PerfTier Tier { get; }

        /// <summary>Why this shape is in the catalogue.</summary>
        public string Rationale { get; }

        public long Cells => (long)Rows * Columns;

        public PerfScenario(string name, int rows, int columns, PerfTier tier, string rationale,
            LargeWorkbook.Payload payload = LargeWorkbook.Payload.SharedStrings, int definedNames = 0)
        {
            Name = name;
            Rows = rows;
            Columns = columns;
            Tier = tier;
            Rationale = rationale;
            Payload = payload;
            DefinedNames = definedNames;
        }

        public System.IO.FileInfo Workbook() => LargeWorkbook.Get(Rows, Columns, Payload, DefinedNames);

        public override string ToString() => Name;

        /// <summary>
        /// Every shape worth measuring.
        /// <para>
        /// Shape matters as much as size. The workbook in #70 is 4000x4000 — far wider than a
        /// typical export — and width exercises different code from height: a wide row holds
        /// thousands of cells whose references run to three letters, so column-name conversion
        /// and reference parsing are hit far harder than in a tall narrow sheet.
        /// </para>
        /// </summary>
        public static IReadOnlyList<PerfScenario> All { get; } = new[]
        {
            // ---- tall and narrow: the ordinary data-export shape ----
            new PerfScenario("tall-small",   1_000, 10, PerfTier.Smoke,
                "Baseline shape; small enough to run on every build."),
            new PerfScenario("tall-medium", 20_000, 10, PerfTier.Smoke,
                "200k cells — enough for per-cell costs to dominate startup noise."),
            new PerfScenario("tall-large",  50_000, 20, PerfTier.Full,
                "1M cells in the shape most callers actually have."),

            // ---- wide: the shape reported in #70 ----
            new PerfScenario("wide-medium",    500,  500, PerfTier.Smoke,
                "250k cells, wide. Column references reach two letters."),
            new PerfScenario("wide-large",   2_000, 2_000, PerfTier.Full,
                "4M cells at the aspect ratio reported in #70, where references reach three letters."),

            // ---- payload ----
            new PerfScenario("numeric-medium", 20_000, 10, PerfTier.Smoke,
                "Numeric cells skip the shared string table entirely.",
                payload: LargeWorkbook.Payload.Numbers),
            new PerfScenario("numeric-large",  50_000, 20, PerfTier.Full,
                "1M numeric cells; the payload used in the #70 report.",
                payload: LargeWorkbook.Payload.Numbers),

            // ---- defined names: cost is O(cells x names), so this is the sharpest axis ----
            new PerfScenario("names-none", 10_000, 10, PerfTier.Smoke,
                "Control for the defined-name comparison."),
            new PerfScenario("names-10",   10_000, 10, PerfTier.Smoke,
                "A handful of names, as Excel creates for print areas.", definedNames: 10),
            new PerfScenario("names-100",  10_000, 10, PerfTier.Smoke,
                "100 names. Cost per cell currently scales linearly with this.", definedNames: 100),
        };

        public static IEnumerable<PerfScenario> Smoke => All.Where(s => s.Tier == PerfTier.Smoke);

        public static PerfScenario ByName(string name) =>
            All.FirstOrDefault(s => s.Name == name)
            ?? throw new ArgumentException($"No performance scenario named '{name}'.", nameof(name));
    }

    public enum PerfTier
    {
        /// <summary>Cheap enough to run on every build alongside the tests.</summary>
        Smoke,

        /// <summary>Large. Run in the benchmark job rather than on every unit-test run.</summary>
        Full
    }
}
