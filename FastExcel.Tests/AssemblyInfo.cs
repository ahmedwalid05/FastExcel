using Xunit;

// Several tests set CultureInfo.CurrentCulture to prove that output is culture-invariant.
// Culture is per-thread, and xUnit runs test collections in parallel across threads, so a
// parallel run could observe another test's culture and fail intermittently. The whole suite
// takes well under a second, so serialising it is a free way to remove that entire class of
// flakiness.
[assembly: CollectionBehavior(DisableTestParallelization = true)]
