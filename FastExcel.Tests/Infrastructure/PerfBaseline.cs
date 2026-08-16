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

    /// <summary>The committed baseline: the recorded numbers plus where they came from.</summary>
    public sealed class PerfBaselineFile
    {
        /// <summary>
        /// The processor architecture the numbers were recorded on, e.g. "X64".
        /// <para>
        /// This is load-bearing rather than informational. Allocated bytes are portable —
        /// measured byte-identical on linux-x64, win-x64 and osx-arm64 — but retained bytes
        /// are not, and differ by about a factor of two between x64 and arm64. Recording the
        /// architecture lets the gate check each metric only where it means something.
        /// </para>
        /// </summary>
        public string Architecture { get; set; }

        public List<PerfBaselineEntry> Scenarios { get; set; } = new List<PerfBaselineEntry>();

        [JsonIgnore]
        public IReadOnlyDictionary<string, PerfBaselineEntry> ByScenario =>
            Scenarios.ToDictionary(e => e.Scenario, StringComparer.Ordinal);

        [JsonIgnore]
        public bool IsEmpty => Scenarios.Count == 0;

        /// <summary>Whether retained memory recorded here is comparable on the current machine.</summary>
        public bool MatchesCurrentArchitecture =>
            string.Equals(Architecture, PerfBaseline.CurrentArchitecture, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>Which recorded metrics a comparison should check.</summary>
    [Flags]
    public enum PerfMetrics
    {
        None = 0,

        /// <summary>Total bytes allocated. Portable across operating systems and architectures.</summary>
        Allocated = 1,

        /// <summary>Bytes still live after the read. Comparable only within one architecture.</summary>
        Retained = 2,

        All = Allocated | Retained
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

        /// <summary>The architecture this process runs on, e.g. "X64" or "Arm64".</summary>
        public static string CurrentArchitecture =>
            System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString();

        public static PerfBaselineFile Load(string path = null)
        {
            path ??= DefaultPath;
            if (!File.Exists(path)) return new PerfBaselineFile();

            var json = File.ReadAllText(path).TrimStart();

            // The first version of this file was a bare array with no architecture recorded.
            // Read it rather than throw, so a stale checkout still runs. With no architecture
            // to compare against, only allocated memory gets gated, which is the safe default.
            if (json.StartsWith("[", StringComparison.Ordinal))
            {
                return new PerfBaselineFile
                {
                    Architecture = null,
                    Scenarios = JsonSerializer.Deserialize<List<PerfBaselineEntry>>(json, JsonOptions)
                                ?? new List<PerfBaselineEntry>()
                };
            }

            return JsonSerializer.Deserialize<PerfBaselineFile>(json, JsonOptions)
                   ?? new PerfBaselineFile();
        }

        public static void Save(IEnumerable<PerfResult> results, string path = null)
        {
            path ??= DefaultPath;

            var file = new PerfBaselineFile
            {
                Architecture = CurrentArchitecture,
                Scenarios = results
                    .OrderBy(r => r.Scenario, StringComparer.Ordinal)
                    .Select(r => new PerfBaselineEntry
                    {
                        Scenario = r.Scenario,
                        Cells = r.Cells,
                        FileBytes = r.FileBytes,
                        AllocatedBytes = r.AllocatedBytes,
                        RetainedBytes = r.RetainedBytes
                    })
                    .ToList()
            };

            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, JsonSerializer.Serialize(file, JsonOptions) + Environment.NewLine);
        }

        /// <summary>
        /// Memory metrics that grew beyond the tolerance. Scenarios absent from the baseline are
        /// not regressions — a newly added shape has nothing to compare against yet.
        /// </summary>
        public static IReadOnlyList<PerfRegression> Compare(
            IEnumerable<PerfResult> results,
            IReadOnlyDictionary<string, PerfBaselineEntry> baseline,
            PerfMetrics metrics = PerfMetrics.All,
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

                if (metrics.HasFlag(PerfMetrics.Retained))
                {
                    Check("retained", recorded.RetainedBytes, result.RetainedBytes);
                }

                if (metrics.HasFlag(PerfMetrics.Allocated))
                {
                    Check("allocated", recorded.AllocatedBytes, result.AllocatedBytes);
                }
            }

            return regressions;
        }

        /// <summary>Renders the results as a markdown table, with movement against the baseline.</summary>
        public static string ToMarkdown(
            IReadOnlyList<PerfResult> results,
            IReadOnlyDictionary<string, PerfBaselineEntry> baseline,
            PerfMetrics gated = PerfMetrics.All,
            string baselineArchitecture = null)
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
            report.AppendLine("`Retained` is what a caller still holds once the worksheet is materialised, " +
                              "which is the figure #70 is about. `ret:file` expresses it as a multiple of " +
                              "the file being read. Time is reported but never gated, because runner speed " +
                              "varies too much to threshold.");

            if (!gated.HasFlag(PerfMetrics.Retained))
            {
                report.AppendLine();
                report.AppendLine($"> Retained memory is reported here but **not gated**. The baseline was " +
                                  $"recorded on {baselineArchitecture ?? "another architecture"} and this " +
                                  $"machine is {CurrentArchitecture}. Retained memory reproduces to the byte " +
                                  $"within one architecture but differs by roughly a factor of two between " +
                                  $"x64 and arm64. Allocated memory is portable and is still gated here.");
            }

            return report.ToString();
        }
    }
}
