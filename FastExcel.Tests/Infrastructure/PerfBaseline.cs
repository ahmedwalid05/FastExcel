using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>One scenario's recorded numbers, as stored in the committed baseline.</summary>
    public sealed class PerfBaselineEntry
    {
        public string Scenario { get; set; }
        public long Cells { get; set; }
        public long FileBytes { get; set; }
        public long AllocatedBytes { get; set; }
        public long RetainedBytes { get; set; }

        // Elapsed time is deliberately NOT stored. It varies several-fold between CI runners,
        // so a committed number would be meaningless and any threshold built on it would
        // either flake or be too loose to detect anything. Time is reported, never compared.
    }

    public sealed class PerfRegression
    {
        public string Scenario { get; set; }
        public string Metric { get; set; }
        public long Baseline { get; set; }
        public long Current { get; set; }
        public double ChangeFraction => Baseline == 0 ? 0 : (Current - Baseline) / (double)Baseline;

        public override string ToString() =>
            $"{Scenario}: {Metric} {PerfResult.Megabytes(Baseline):N1} MB -> " +
            $"{PerfResult.Megabytes(Current):N1} MB ({ChangeFraction:+0.0%;-0.0%;0%})";
    }

    /// <summary>
    /// Reads, writes and compares the committed performance baseline.
    /// <para>
    /// A benchmark run that reports only its own numbers cannot tell anyone whether things got
    /// worse, which is what makes an unattended run on every merge worth having at all. The
    /// baseline is committed to the repository so the comparison is against a reviewed,
    /// deliberate number rather than whatever the previous run happened to produce.
    /// </para>
    /// </summary>
    public static class PerfBaseline
    {
        /// <summary>
        /// How much a memory metric may grow before it counts as a regression.
        /// <para>
        /// Retained bytes measured identically to the byte across repeated runs, and allocation
        /// counts are nearly as stable, so this does not need to absorb measurement noise — it
        /// only absorbs genuine differences between runtime versions and platforms.
        /// </para>
        /// </summary>
        public const double RegressionTolerance = 0.10;

        private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.Never
        };

        /// <summary>Locates the committed baseline by walking up from the test assembly to the repo.</summary>
        public static string DefaultPath
        {
            get
            {
                var directory = new DirectoryInfo(AppContext.BaseDirectory);
                while (directory != null && !File.Exists(Path.Combine(directory.FullName, "FastExcel.sln")))
                {
                    directory = directory.Parent;
                }

                var root = directory?.FullName ?? Directory.GetCurrentDirectory();
                return Path.Combine(root, "FastExcel.Tests", "perf-baseline.json");
            }
        }

        public static IReadOnlyDictionary<string, PerfBaselineEntry> Load(string path = null)
        {
            path ??= DefaultPath;
            if (!File.Exists(path)) return new Dictionary<string, PerfBaselineEntry>();

            var entries = JsonSerializer.Deserialize<List<PerfBaselineEntry>>(File.ReadAllText(path), JsonOptions)
                          ?? new List<PerfBaselineEntry>();

            return entries.ToDictionary(e => e.Scenario, StringComparer.Ordinal);
        }

        public static void Save(IEnumerable<PerfResult> results, string path = null)
        {
            path ??= DefaultPath;
            var entries = results
                .OrderBy(r => r.Scenario, StringComparer.Ordinal)
                .Select(r => new PerfBaselineEntry
                {
                    Scenario = r.Scenario,
                    Cells = r.Cells,
                    FileBytes = r.FileBytes,
                    AllocatedBytes = r.AllocatedBytes,
                    RetainedBytes = r.RetainedBytes
                })
                .ToList();

            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, JsonSerializer.Serialize(entries, JsonOptions) + Environment.NewLine);
        }

        /// <summary>
        /// Memory metrics that grew beyond the tolerance. Scenarios absent from the baseline are
        /// not regressions — a newly added shape has nothing to compare against yet.
        /// </summary>
        public static IReadOnlyList<PerfRegression> Compare(
            IEnumerable<PerfResult> results,
            IReadOnlyDictionary<string, PerfBaselineEntry> baseline,
            double tolerance = RegressionTolerance)
        {
            var regressions = new List<PerfRegression>();

            foreach (var result in results)
            {
                if (!baseline.TryGetValue(result.Scenario, out var recorded)) continue;

                void Check(string metric, long before, long now)
                {
                    if (before <= 0) return;
                    if (now > before * (1 + tolerance))
                    {
                        regressions.Add(new PerfRegression
                        {
                            Scenario = result.Scenario,
                            Metric = metric,
                            Baseline = before,
                            Current = now
                        });
                    }
                }

                Check("retained", recorded.RetainedBytes, result.RetainedBytes);
                Check("allocated", recorded.AllocatedBytes, result.AllocatedBytes);
            }

            return regressions;
        }

        /// <summary>Renders the results as a markdown table, with movement against the baseline.</summary>
        public static string ToMarkdown(
            IReadOnlyList<PerfResult> results,
            IReadOnlyDictionary<string, PerfBaselineEntry> baseline)
        {
            var report = new StringBuilder();

            report.AppendLine("| Scenario | Cells | File | Time | Cells/sec | Allocated | alloc B/cell | Retained | ret B/cell | ret:file | Δ retained |");
            report.AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |");

            foreach (var r in results.OrderBy(r => r.Scenario, StringComparer.Ordinal))
            {
                var delta = "n/a";
                if (baseline.TryGetValue(r.Scenario, out var recorded) && recorded.RetainedBytes > 0)
                {
                    var change = (r.RetainedBytes - recorded.RetainedBytes) / (double)recorded.RetainedBytes;
                    delta = Math.Abs(change) < 0.005
                        ? "—"
                        : change.ToString("+0.0%;-0.0%", CultureInfo.InvariantCulture);
                }

                report.AppendLine(string.Join(" | ", new[]
                {
                    "| " + r.Scenario,
                    r.Cells.ToString("N0", CultureInfo.InvariantCulture),
                    PerfResult.Megabytes(r.FileBytes).ToString("N2", CultureInfo.InvariantCulture) + " MB",
                    r.ElapsedMilliseconds.ToString("N0", CultureInfo.InvariantCulture) + " ms",
                    r.CellsPerSecond.ToString("N0", CultureInfo.InvariantCulture),
                    PerfResult.Megabytes(r.AllocatedBytes).ToString("N1", CultureInfo.InvariantCulture) + " MB",
                    r.BytesPerCellAllocated.ToString("N0", CultureInfo.InvariantCulture),
                    PerfResult.Megabytes(r.RetainedBytes).ToString("N1", CultureInfo.InvariantCulture) + " MB",
                    r.BytesPerCellRetained.ToString("N0", CultureInfo.InvariantCulture),
                    r.RetainedToFileRatio.ToString("N0", CultureInfo.InvariantCulture) + "x",
                    delta + " |"
                }));
            }

            report.AppendLine();
            report.AppendLine("`ret:file` is retained memory as a multiple of the file being read. " +
                              "`Retained` is what a caller still holds once the worksheet is materialised — " +
                              "the figure #70 is about — and is measured after a forced collection, so it is " +
                              "reproducible to the byte. Time is reported but never gated: it varies too much " +
                              "between runners to threshold meaningfully.");

            return report.ToString();
        }
    }
}
