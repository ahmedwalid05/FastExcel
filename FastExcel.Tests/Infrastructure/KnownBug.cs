using System;
using Xunit.Sdk;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// Marks an assertion that describes <b>correct</b> behaviour which the library does
    /// not yet exhibit — a defect that is known, reproduced and tracked, but not fixed.
    /// <para>
    /// This is the "expected failure" pattern (pytest's <c>xfail</c> with strict XPASS).
    /// The test body asserts what the library <i>should</i> do, and the assertion is
    /// wrapped here. While the bug exists the test passes; the moment the bug is fixed
    /// the test <b>fails loudly</b> and tells you to unwrap it.
    /// </para>
    /// <para>
    /// The point is that CI stays green and meaningful — a red build always means a new
    /// regression, never a backlog item — while every known defect is still committed as
    /// executable, self-promoting documentation rather than a comment or a skipped test.
    /// A skipped test would rot silently; this one cannot.
    /// </para>
    /// </summary>
    public static class KnownBug
    {
        /// <summary>
        /// Asserts that <paramref name="correctBehaviour"/> currently does NOT hold.
        /// </summary>
        /// <param name="issue">Issue reference, e.g. "#87".</param>
        /// <param name="expectation">
        /// Plain-English statement of what the library should do once fixed. Surfaced in the
        /// failure message when the bug is fixed, so write it for whoever lands the fix.
        /// </param>
        /// <param name="correctBehaviour">Assertion describing the correct behaviour.</param>
        public static void StillBroken(string issue, string expectation, Action correctBehaviour)
        {
            if (correctBehaviour == null) throw new ArgumentNullException(nameof(correctBehaviour));

            try
            {
                correctBehaviour();
            }
            catch (Exception exception)
            {
                // Expected: the defect is still present. Any failure mode counts — some of
                // these bugs surface as a wrong value, others as an unhandled exception.
                Explain(issue, exception);
                return;
            }

            Explain(issue, null);

            throw new XunitException(
                $"KNOWN BUG {issue} APPEARS TO BE FIXED — promote this test.{Environment.NewLine}" +
                $"Expected behaviour: {expectation}{Environment.NewLine}" +
                $"The library now behaves correctly, so this assertion no longer belongs " +
                $"inside KnownBug.StillBroken(). Unwrap it into a plain assertion and, if " +
                $"nothing else covers {issue}, close the issue.");
        }

        /// <summary>
        /// Set FASTEXCEL_KNOWNBUG_LOG to a file path to record how each known bug actually
        /// fails right now.
        /// <para>
        /// This exists because the weakness of the expected-failure pattern is that a test
        /// which fails for the WRONG reason — a typo, a bad fixture — looks identical to one
        /// failing for the right reason. Dumping the real exception makes that auditable, and
        /// doubles as a quick "what is broken today" report while working through the fixes.
        /// </para>
        /// </summary>
        private static void Explain(string issue, Exception exception)
        {
            var path = Environment.GetEnvironmentVariable("FASTEXCEL_KNOWNBUG_LOG");
            if (string.IsNullOrEmpty(path)) return;

            var caller = new System.Diagnostics.StackTrace(2, false).GetFrame(0)?.GetMethod();
            var detail = exception == null
                ? "NOW PASSING"
                : $"{exception.GetType().Name}: {Flatten(exception.Message)}";

            lock (typeof(KnownBug))
            {
                System.IO.File.AppendAllText(path,
                    $"{issue}\t{caller?.DeclaringType?.Name}.{caller?.Name}\t{detail}{Environment.NewLine}");
            }
        }

        private static string Flatten(string message)
        {
            var single = message.Replace("\r", " ").Replace("\n", " ");
            while (single.Contains("  ")) single = single.Replace("  ", " ");
            return single.Length > 220 ? single.Substring(0, 220) + "…" : single;
        }
    }
}
