# FastExcel test suite

Run everything:

```bash
dotnet test
```

## How known bugs are represented

Some tests describe behaviour the library does **not** have yet. Rather than committing them
red — which would leave CI permanently failing and make a real regression indistinguishable
from a backlog item — they assert the *correct* behaviour and wrap it in `KnownBug.StillBroken`:

```csharp
KnownBug.StillBroken("#55",
    "under ru-RU the value 1.2 is stored as \"1.2\"",
    () => Assert.Equal("1.2", FirstCellValue(WriteRowAndGetSheetXml(workspace, 1.2d))));
```

This is the expected-failure pattern (pytest's `xfail` with strict XPASS):

- while the bug exists, the wrapped assertion fails and **the test passes**;
- the moment the bug is fixed, the assertion succeeds and **the test fails**, telling whoever
  landed the fix to unwrap it into a plain assertion and close the issue.

So a green build always means "nothing new is broken", every known defect is committed as
executable documentation instead of a comment, and no fix can quietly land without the suite
noticing. Unlike a `Skip`, these cannot rot — they are executed on every run.

**When you fix a bug, you will see a failure like this. That is the suite working:**

```
KNOWN BUG #55 APPEARS TO BE FIXED — promote this test.
Expected behaviour: under ru-RU the value 1.2 is stored as "1.2"
```

Delete the `KnownBug.StillBroken` wrapper, keep the assertion, and the test becomes a
permanent regression guard.

### Auditing what is actually broken

The weakness of this pattern is that a test failing for the *wrong* reason (a typo, a bad
fixture) looks the same as one failing for the right reason. To check, set
`FASTEXCEL_KNOWNBUG_LOG` and inspect how each one really fails:

```bash
FASTEXCEL_KNOWNBUG_LOG=/tmp/knownbugs.tsv dotnet test
column -t -s $'\t' /tmp/knownbugs.tsv
```

CI writes this file on every run and uploads it as the `known-bugs` artifact, so the current
"what is still broken" report is always one click away from a build.

## Layout

| Path | Covers |
| --- | --- |
| `SharedStringsTests.cs` | Positional shared-string resolution, rich text, text fidelity, escaping (#87, #81, #10, #76) |
| `CellWriteTests.cs` | Value-to-XML conversion: types, cultures, cell ordering (#55, #77, #72, #61) |
| `RowReadTests.cs` | Rows to cells, gaps, non-cell content, lazy-enumeration lifetime (#88, #83, #78, #22) |
| `DefinedNameTests.cs` | Defined names, sheet scoping, column aliases (#84, #49, #86, #89) |
| `StreamLifecycleTests.cs` | Constructors, read-only streams, update semantics, disposal (#75, #69, #74, #71) |
| `WorksheetResolutionTests.cs` | Sheet name/index to package part resolution (#82, PR #62) |
| `ColumnNameTests.cs` | Column letter/number conversion across Excel's full range |
| `Performance/AllocationBudgetTests.cs` | Per-cell memory budgets and scaling (#70) |
| `Performance/PerfReportTests.cs` | Scenario catalogue, report, regression gate vs the committed baseline |
| `Performance/ThroughputTests.cs` | Wall-clock throughput and scaling (opt-in) |
| `FastExcelTests.cs` | The original end-to-end tests |

## Infrastructure

- **`XlsxBuilder`** builds a minimal valid `.xlsx` in memory with exactly the XML a test needs.
  Most defects here are about how one specific piece of spreadsheet XML is interpreted, so the
  tests spell that XML out literally instead of relying on opaque binary fixtures.
- **`CultureScope`** runs a block under a given culture. The xlsx format requires an invariant
  decimal point, so culture bugs are invisible unless tests actually run under `ru-RU`, `de-DE`
  and friends.
- **`TempWorkspace`** gives each test its own temp directory, so tests never share output files.
- **`KnownBug`** is described above.

Test parallelisation is disabled assembly-wide (`AssemblyInfo.cs`): culture is thread-local, and
the whole suite runs in well under a second, so serialising it removes a class of flakiness for
free.

## Performance tests

Performance is treated as testing: `Performance/` runs on every build, on the same triggers as
everything else, and **fails the build when memory regresses**.

- **`PerfReportTests` is the gate.** It measures the scenario catalogue and compares against
  `perf-baseline.json`, failing if allocated or retained memory grows by more than 10%.
- **`AllocationBudgetTests`** holds standalone per-cell ceilings.
- **`ThroughputTests` are skipped unless `FASTEXCEL_PERF=1`.** Elapsed time on a shared runner
  varies too much to threshold; only memory is gated.

```bash
dotnet test                                    # smoke scenarios + gate
FASTEXCEL_PERF=1 dotnet test                   # adds 1M-4M cell scenarios and timing tests
FASTEXCEL_PERF_REPORT=perf.md dotnet test      # also write the markdown report
```

Two memory numbers are recorded per scenario because they move independently: **allocated**
(total bytes, including what the GC reclaims) and **retained** (still live once the worksheet is
materialised, which decides whether a file fits in RAM). An optimisation can improve one and not
the other, so gating on allocated memory alone would miss #70 entirely.

### What is portable, and what is not

The two metrics do not travel equally well, so the gate checks each one only where it means
something:

| Metric | Portability | Gated |
| --- | --- | --- |
| Allocated | Measured byte-identical on linux-x64, win-x64 and osx-arm64 | Everywhere |
| Retained | Reproduces to the byte within one architecture, but arm64 reports roughly 2x the x64 figure for the same object graph | Only where the architecture matches the baseline |

The baseline records the architecture it was measured on. On a machine that matches, both
metrics are gated. On one that does not, retained memory is still measured and reported, and the
report says plainly that it is not being checked. `RetainedMemoryIsMeasuredDeterministically`
guards reproducibility, and `TheCommittedBaselineRecordsItsArchitecture` stops a baseline
without an architecture from silently weakening the gate.

`PerfMeasurement.Settle` forces a blocking, compacting collection rather than calling
`GC.Collect()`, which may answer a gen2 request with a background, non-compacting pass. Reading
a workbook allocates about 1,300 bytes per cell and retains about 470, so most of the heap is
garbage when the measurement happens and any survivor would inflate the result.

### Updating the baseline

After a change that legitimately moves the numbers:

```bash
FASTEXCEL_PERF=1 FASTEXCEL_PERF_UPDATE_BASELINE=1 dotnet test
```

Commit the diff — it records what the change bought. Without `FASTEXCEL_PERF` only the smoke-tier
entries are rewritten and the large scenarios are carried over untouched.

Scenarios live in `Infrastructure/PerfScenario.cs` and are shared with the benchmark project, so
adding a shape adds it to both. For statistical detail and the comparison against MiniExcel,
ClosedXML and Sylvan, see `FastExcel.Benchmarks`.
