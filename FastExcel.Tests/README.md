# FastExcel test suite

Run everything:

```bash
dotnet test
```

## Test strategy

The package has ~780,000 downloads across ~124 dependent repositories, and several of those
vendored the source rather than referencing the package — so they never receive a fix. The suite
exists to make it hard for a contributor to break one of them without noticing.

Two failure modes drive the design, because neither is caught by ordinary behavioural tests:

- **A compiled consumer stops starting.** Adding a parameter to a public constructor keeps every
  behavioural test green and raises `MissingMethodException` in code that was compiled against
  the previous release. `PublicApiTests` exists for this.
- **A written workbook is unreadable to other software.** This library escapes and unescapes
  symmetrically, and parses with the same assumptions it writes with, so a file can round-trip
  perfectly through FastExcel while every other reader chokes on it. `XlsxAssert` checks the
  package is valid; `InteropTests` goes further and reads the output with ClosedXML, MiniExcel
  and Sylvan. Three independent implementations dropping the same cell is the strongest evidence
  available without launching Excel.

### The layers

| Layer | Purpose | Roughly | Where |
| --- | --- | --- | --- |
| **Unit** | One decision at a time — column-letter arithmetic, cell-to-XML conversion, a single parse rule. Fast, no file I/O beyond an in-memory package. | ~55% | `CellWriteTests`, `ColumnNameTests`, `RowReadTests`, `SharedStringsTests` |
| **Regression** | One test per reported defect, named with its issue number, asserting the behaviour the reporter expected. Never deleted after a fix — that is the point. | ~25% | every `KnownBug.StillBroken` site; see the table below |
| **Integration** | A real package through the real API: open, read, write, update, dispose. Catches interactions between the archive, the shared-string table and the write path that no unit test sees. | ~10% | `StreamLifecycleTests`, `WorksheetResolutionTests`, `FastExcelTests` |
| **Contract** | The promises made to consumers rather than to a caller in this repo: the shape of the public API, and whether other software can read what we write. | ~5% | `PublicApiTests`, `InteropTests`, `Infrastructure/XlsxAssert` |
| **Performance** | Memory and throughput against a committed baseline. Treated as correctness because #70 (12 GB for a 50 MB file) is a defect, not a preference. | ~5% | `Performance/` |

The ratios are a description of where effort belongs, not a quota to enforce. Unit tests dominate
because they are the ones that stay fast and readable; regression tests are the second-largest
group because this library's defect history is its best specification.

### Coverage

The build **fails** below these minimums, measured on the Linux job:

| Metric | Minimum | Currently |
| --- | --- | --- |
| Line | 80% | 83.7% |
| Branch | 74% | 77.9% |

They are a ratchet. Raise them when a change lifts the real number; never lower them to make a
build pass. They currently sit a few points below the measured figure on purpose: the numbers above
were measured on Windows and the gate runs on Linux, so the headroom absorbs any small difference
until a real CI run tells us what Linux actually reports. Tighten them once it has. A coverage number is a floor on what is *executed*, not evidence that anything is
asserted — several tests in this suite execute code and assert nothing, which is why the number
alone is not the goal.

**There is a ceiling, and it is not laziness.** `FastExcel` holds two private fields,
`AddWorksheets` and `DeleteWorksheets`, that are read in nine places and assigned in none. Every
branch guarded by them is unreachable — roughly 180 lines across `UpdateRelations`,
`UpdateWorkbook`, `RenameAndRebildWorksheetProperties` and `UpdateContentTypes`, plus
`Worksheet.ValidateNewWorksheet`, `Worksheet.AddSettings` and all of `WorksheetAddSettings`. No
test can cover it because nothing can execute it. `PublicApiTests.ThereIsNoPublicWayToAddOrRemoveAWorksheet`
pins that, and fails the day someone wires the feature up. Excluding it, the reachable library is
covered in the mid-nineties.

### Running each layer

```bash
dotnet test                                                  # everything except the timing tests
dotnet test --filter "FullyQualifiedName~PublicApiTests"      # contract: public API surface
dotnet test --filter "FullyQualifiedName~Performance"         # memory budgets and the perf gate
FASTEXCEL_PERF=1 dotnet test                                  # adds 1M-4M cell scenarios and timing
dotnet test --collect:"XPlat Code Coverage"                   # produces coverage.cobertura.xml
FASTEXCEL_KNOWNBUG_LOG=/tmp/knownbugs.tsv dotnet test         # records how each known defect fails
```

### Writing a new test

1. **Reuse the infrastructure.** `XlsxBuilder` builds a package from the exact XML a test needs,
   `CultureScope` runs a block under a given culture, `TempWorkspace` gives each test its own
   directory. Do not check in a binary fixture for something `XlsxBuilder` can express.
2. **Assert on values, not on absence.** `Assert.DoesNotThrow` and `DoesNotContain` both pass when
   the library returns nothing at all.
3. **Call `XlsxAssert.IsValidPackage` from anything that writes a file**, and for anything about
   values reaching a user, add a case to `InteropTests`. A FastExcel round trip is not evidence
   the file is valid — it is barely evidence of anything, because both halves of the library share
   the same wrong assumptions.
4. **Parse, then assert on the string.** For raw-XML expectations, load the part with `XDocument`
   first and assert on the text after — well-formedness comes free and the exact bytes are still
   pinned.
5. **Put a time limit on anything that parses.** `WorksheetXmlLayoutTests` documents a real
   non-terminating loop; an unbounded parser test turns a CI run into a billing incident.

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
| `PublicApiTests.cs` | The public API surface, against a committed baseline; unreachable exported types (#94) |
| `WorksheetXmlLayoutTests.cs` | Line breaks and namespace prefixes in the worksheet part (#80, #22) |
| `WorksheetAuthoringTests.cs` | `PopulateRows*`, `AddRow`, `AddValue`, `GetCellsInRange` — building a sheet in memory |
| `WorkbookNavigationTests.cs` | `Worksheets`, sheet lookup by name and number, and all eight `Write` overloads |
| `SmallSurfaceTests.cs` | `ExcelColumnAttribute`, `WorksheetProperties`, cell identity properties, single-cell references |
| `MalformedInputTests.cs` | Corrupt zips, missing parts, bad row and cell references, bad shared-string indexes |
| `CellTypeTests.cs` | Every `t` value on read — `s`, `n`, `b`, `e`, `str`, `inlineStr` — plus formulas, dates and text fidelity |
| `RoundTripTests.cs` | Write-then-read type fidelity, and Excel's row, column and cell-length limits |
| `TemplateHeaderTests.cs` | Writing into a template that already has heading rows, and what the template keeps |
| `InteropTests.cs` | FastExcel's output read by ClosedXML/MiniExcel/Sylvan, and their output read by FastExcel |
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
- **`XlsxAssert`** opens a produced file through the OPC layer and parses every XML part. This is
  the check a FastExcel round trip cannot give you: the library escapes and unescapes
  symmetrically, so a file can be perfectly readable to itself and refused by Excel. Call
  `XlsxAssert.IsValidPackage` from any test that writes.
- **`Timebox`** runs work under a time limit on a background thread. Needed because
  `Worksheet.ReadHeadersAndFooters` contains a loop with no exit condition (#80) and .NET Core
  cannot abort a thread — without it, one test would wedge the whole run. Use it for anything that
  parses.
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
