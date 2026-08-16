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

`Performance/` holds assertions about speed and memory. They are split by what can be trusted on
a shared CI runner:

- **`AllocationBudgetTests` run everywhere.** They assert allocated *bytes per cell*, which is
  essentially deterministic for the same input on any machine — so a real threshold is possible.
  This is what guards #70.
- **`ThroughputTests` are skipped unless `FASTEXCEL_PERF=1`.** Elapsed time on a shared runner
  varies enough that gating merges on it means either flaky builds or thresholds too loose to
  catch anything.

```bash
dotnet test                                   # allocation budgets included
FASTEXCEL_PERF=1 dotnet test                  # timing tests as well
```

Current baseline is about **1,280 bytes allocated per cell**, roughly 400x the size of the file
being read. The budget is set at 1,800 to catch regressions with room for runtime variation; the
target is under 256, tracked as a `KnownBug`. **Lower the budget when the read path gets
cheaper** — it is meant to ratchet.

For real numbers rather than pass/fail thresholds, see `FastExcel.Benchmarks`.
