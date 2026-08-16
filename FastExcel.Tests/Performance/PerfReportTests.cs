using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using FastExcel.Tests.Infrastructure;
using Xunit;
using Xunit.Abstractions;

namespace FastExcel.Tests.Performance
{
    /// <summary>
    /// Runs the scenario catalogue, publishes a report, and fails the build if memory has
    /// regressed against the committed baseline.
    /// <para>
    /// Performance is treated as testing here: this runs on every build, on the same triggers
    /// as everything else, and a regression is a failure rather than a note in an artifact
    /// nobody opens. Only memory is compared. Elapsed time is reported for context but never
    /// thresholded, because a shared runner's speed varies enough that any gate on it would
    /// either flake or be too loose to catch a real problem.
    /// </para>
    /// </summary>
    public class PerfReportTests
    {
        private readonly ITestOutputHelper _output;

        public PerfReportTests(ITestOutputHelper output) => _output = output;

        /// <summary>Set to 1 to run the large scenarios as well as the smoke set.</summary>
        private static bool IncludeFullTier =>
            Environment.GetEnvironmentVariable("FASTEXCEL_PERF") == "1";

        /// <summary>
        /// Set to 1 to rewrite the committed baseline from this run. Do this deliberately —
        /// after a fix that legitimately changes the numbers — and review the diff, which is
        /// the record of what the change actually bought.
        /// </summary>
        private static bool UpdateBaseline =>
            Environment.GetEnvironmentVariable("FASTEXCEL_PERF_UPDATE_BASELINE") == "1";

        [Fact]
        public void MemoryHasNotRegressedAgainstTheBaseline()
        {
            var scenarios = (IncludeFullTier ? PerfScenario.All : PerfScenario.Smoke).ToList();

            var results = scenarios
                .Select(scenario => PerfMeasurement.MeasureRead(scenario))
                .ToList();

            var baseline = PerfBaseline.Load();
            var report = PerfBaseline.ToMarkdown(results, baseline);

            _output.WriteLine(report);
            PublishReport(report, results);

            if (UpdateBaseline)
            {
                // Only the tiers that actually ran may be rewritten, or a smoke-only run would
                // silently drop the large scenarios from the baseline.
                var merged = MergeWithExisting(results, baseline);
                PerfBaseline.Save(merged);
                _output.WriteLine($"Baseline rewritten at {PerfBaseline.DefaultPath}");
                return;
            }

            Assert.True(baseline.Count > 0,
                $"No performance baseline found at {PerfBaseline.DefaultPath}. Create one with " +
                "FASTEXCEL_PERF_UPDATE_BASELINE=1 dotnet test and commit the result.");

            var regressions = PerfBaseline.Compare(results, baseline);

            Assert.True(regressions.Count == 0,
                "Memory regressed against the committed baseline:" + Environment.NewLine +
                string.Join(Environment.NewLine, regressions.Select(r => "  " + r)) +
                Environment.NewLine + Environment.NewLine +
                $"Tolerance is {PerfBaseline.RegressionTolerance:P0}. If the increase is intended, " +
                "rerun with FASTEXCEL_PERF_UPDATE_BASELINE=1 and commit the updated baseline." +
                Environment.NewLine + Environment.NewLine + report);
        }

        [Fact]
        public void EveryScenarioInTheBaselineStillExists()
        {
            var baseline = PerfBaseline.Load();
            if (baseline.Count == 0) return;

            var known = PerfScenario.All.Select(s => s.Name).ToHashSet(StringComparer.Ordinal);
            var orphaned = baseline.Keys.Where(name => !known.Contains(name)).ToList();

            // A renamed or deleted scenario leaves a stale entry that silently stops being
            // checked, which is how a coverage gap hides in plain sight.
            Assert.True(orphaned.Count == 0,
                "The baseline records scenarios that no longer exist: " +
                string.Join(", ", orphaned) +
                ". Remove them with FASTEXCEL_PERF_UPDATE_BASELINE=1 dotnet test.");
        }

        [Fact]
        public void RetainedMemoryIsMeasuredDeterministically()
        {
            // The regression gate is only trustworthy if repeated measurement of the same
            // workbook agrees. If this ever becomes flaky the tolerance is hiding it.
            var scenario = PerfScenario.ByName("tall-small");

            var first = PerfMeasurement.MeasureRead(scenario);
            var second = PerfMeasurement.MeasureRead(scenario);

            _output.WriteLine($"run 1: {first.RetainedBytes:N0} bytes retained");
            _output.WriteLine($"run 2: {second.RetainedBytes:N0} bytes retained");

            var drift = Math.Abs(second.RetainedBytes - first.RetainedBytes) / (double)first.RetainedBytes;

            Assert.True(drift < 0.02,
                $"Retained memory moved {drift:P1} between two identical reads, so the " +
                "regression gate cannot distinguish a real change from measurement noise.");
        }

        private static IEnumerable<PerfResult> MergeWithExisting(
            IReadOnlyList<PerfResult> results,
            IReadOnlyDictionary<string, PerfBaselineEntry> baseline)
        {
            var measured = results.Select(r => r.Scenario).ToHashSet(StringComparer.Ordinal);
            var known = PerfScenario.All.Select(s => s.Name).ToHashSet(StringComparer.Ordinal);

            var carriedOver = baseline.Values
                .Where(e => !measured.Contains(e.Scenario) && known.Contains(e.Scenario))
                .Select(e => new PerfResult
                {
                    Scenario = e.Scenario,
                    Cells = e.Cells,
                    FileBytes = e.FileBytes,
                    AllocatedBytes = e.AllocatedBytes,
                    RetainedBytes = e.RetainedBytes
                });

            return results.Concat(carriedOver);
        }

        /// <summary>
        /// Writes the report where CI can pick it up as an artifact and add it to the job summary.
        /// </summary>
        private static void PublishReport(string report, IReadOnlyList<PerfResult> results)
        {
            var path = Environment.GetEnvironmentVariable("FASTEXCEL_PERF_REPORT");
            if (string.IsNullOrEmpty(path)) return;

            var header =
                $"### Read performance ({results.Count} scenarios, " +
                $"{RuntimeLabel()}){Environment.NewLine}{Environment.NewLine}";

            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)));
            File.WriteAllText(path, header + report);
        }

        private static string RuntimeLabel() =>
            $"{System.Runtime.InteropServices.RuntimeInformation.OSDescription.Trim()}, " +
            $".NET {Environment.Version}";
    }
}
