using System;
using System.Threading;

namespace FastExcel.Tests.Infrastructure
{
    /// <summary>
    /// Runs work under a time limit and reports whether it finished.
    ///
    /// Needed because this library contains at least one loop with no exit condition
    /// (<c>Worksheet.ReadHeadersAndFooters</c>, see #80), and .NET Core cannot abort a running
    /// thread. A test that simply called the method would wedge the run forever.
    ///
    /// The work therefore runs on a <b>background</b> thread: a hung call keeps burning a core
    /// for the rest of the run, but the process can still exit, so CI fails in minutes rather
    /// than hanging until the job times out. Keep the limits short — the defects this guards
    /// against fail to advance immediately, so waiting longer buys nothing.
    /// </summary>
    internal static class Timebox
    {
        /// <summary>
        /// Returns true if <paramref name="work"/> finished within <paramref name="limit"/>.
        /// An exception thrown by the work is rethrown, because "threw" and "never returned" are
        /// different defects and must not be reported as the same thing.
        /// </summary>
        internal static bool Completes(TimeSpan limit, Action work, string threadName = "timeboxed")
        {
            Exception failure = null;
            var finished = new ManualResetEventSlim(false);

            var thread = new Thread(() =>
            {
                try { work(); }
                catch (Exception ex) { failure = ex; }
                finally { finished.Set(); }
            })
            {
                IsBackground = true,
                Name = threadName
            };

            thread.Start();

            if (!finished.Wait(limit))
            {
                return false;
            }

            if (failure != null)
            {
                throw new InvalidOperationException(
                    "the work threw rather than hanging: " + failure.GetType().Name + ": " + failure.Message,
                    failure);
            }

            return true;
        }
    }
}
