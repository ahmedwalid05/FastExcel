# FastExcel benchmarks

Performance is treated as testing here. There are two layers:

| Layer | Where | Runs | Gates the build? |
| --- | --- | --- | --- |
| Regression gate | `FastExcel.Tests/Performance` | Every build, all three platforms | **Yes** — memory only |
| Measurement | `FastExcel.Benchmarks` (this project) | Same triggers, plus on demand | No — it reports |

```bash
# the gate, as CI runs it
dotnet test

# include the large scenarios (1M-4M cells)
FASTEXCEL_PERF=1 dotnet test

# full statistical numbers, including the comparison against other libraries
dotnet run -c Release --project FastExcel.Benchmarks -- --filter *

# quick numbers while iterating on a fix
dotnet run -c Release --project FastExcel.Benchmarks -- --filter *ReadBenchmarks* --job Short
```

Workbooks come from `LargeWorkbook`, shared with the test suite so both layers measure identical
files, written directly rather than through FastExcel so generation never lands inside a
measurement, and cached on disk by shape.

## Reading the output

Three memory numbers are reported, because they answer different questions and **move
independently**:

| Column | Means | Use it for |
| --- | --- | --- |
| `Allocated` | Total bytes allocated, including everything the GC reclaimed | CPU and GC pressure |
| `Retained` | Bytes still live once the worksheet is materialised | **Whether a file fits in RAM** |
| `ret:file` | Retained as a multiple of the file being read | The headline footprint number |

This distinction is the reason the gate is trustworthy. An optimisation can cut `Allocated`
several-fold while leaving `Retained` untouched — and `Retained` is what #70 is about. Reporting
only allocations, as BenchmarkDotNet's `MemoryDiagnoser` does, cannot tell those apart. The
defined-name finding below is exactly that case.

Elapsed time is reported but **never gated**: a shared CI runner varies enough that any
threshold on it would either flake or be too loose to catch a real regression. `Retained` is
measured after a forced collection and reproduced to the byte across repeated runs (0.00%
variation over five runs), which is what makes a 10% tolerance meaningful.

## The regression gate

`FastExcel.Tests/perf-baseline.json` holds reviewed numbers per scenario. Every build re-measures
and fails if allocated or retained memory grows more than 10%.

```
Memory regressed against the committed baseline:
  tall-small: retained 2.2 MB -> 4.5 MB (+100.0%)
Tolerance is 10%. If the increase is intended, rerun with
FASTEXCEL_PERF_UPDATE_BASELINE=1 and commit the updated baseline.
```

When a fix legitimately changes the numbers:

```bash
FASTEXCEL_PERF=1 FASTEXCEL_PERF_UPDATE_BASELINE=1 dotnet test
```

Then commit the diff — it is the record of what the change actually bought. Omit `FASTEXCEL_PERF`
and only the smoke-tier scenarios are rewritten; the large ones are carried over untouched.

## Baseline, August 2026

Pre-fix numbers, so Phase 1 and Phase 3 have something to beat. 4-core Linux container, .NET 8.

| Scenario | Cells | File | Time | Cells/sec | Allocated | alloc B/cell | Retained | ret B/cell | ret:file |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| tall-small | 10,000 | 0.03 MB | 112 ms | 89,399 | 12.7 MB | 1,328 | 4.5 MB | 470 | 149x |
| tall-medium | 200,000 | 0.52 MB | 1,241 ms | 161,148 | 252.5 MB | 1,324 | 90.4 MB | 474 | 173x |
| tall-large | 1,000,000 | 2.59 MB | 5,495 ms | 181,971 | 1,247.4 MB | 1,308 | 446.0 MB | 468 | 172x |
| wide-medium | 250,000 | 0.76 MB | 1,124 ms | 222,413 | 307.4 MB | 1,289 | 108.8 MB | 456 | 143x |
| wide-large | 4,000,000 | 17.93 MB | 23,691 ms | 168,838 | 4,959.4 MB | 1,300 | 1,764.2 MB | 462 | 98x |
| numeric-medium | 200,000 | 0.70 MB | 1,060 ms | 188,618 | 257.1 MB | 1,348 | 76.7 MB | 402 | 110x |
| numeric-large | 1,000,000 | 3.35 MB | 5,587 ms | 178,986 | 1,262.6 MB | 1,324 | 373.5 MB | 392 | 112x |
| names-none | 100,000 | 0.26 MB | 536 ms | 186,486 | 125.9 MB | 1,320 | 44.8 MB | 470 | 170x |
| names-10 | 100,000 | 0.26 MB | 876 ms | 114,096 | 386.7 MB | 4,055 | 44.8 MB | 470 | 170x |
| names-100 | 100,000 | 0.26 MB | 2,809 ms | 35,601 | 2,637.6 MB | 27,658 | 44.8 MB | 470 | 170x |

**A cell costs about 470 bytes for as long as you hold the worksheet** — remarkably constant
across every size and shape, and roughly **100-170x the size of the file on disk**. Numeric cells
are slightly cheaper (~400 bytes) because they carry no shared-string reference. An 18 MB
workbook retains 1.76 GB; extrapolating to the 50 MB file in #70 gives roughly 5 GB retained,
which is the same order as the 12 GB reported there (working set exceeds live heap, and .NET does
not promptly return freed pages to the OS).

Throughput is roughly 170,000-220,000 cells/sec, scaling linearly with cell count. Width does not
change the per-cell cost meaningfully.

### Defined names: churn, not footprint

Identical 100,000 cells, varying only how many defined names the workbook declares:

| Defined names | Time | Allocated | Retained |
| ---: | ---: | ---: | ---: |
| 0 | 536 ms | 125.9 MB | 44.8 MB |
| 10 | 876 ms | 386.7 MB | 44.8 MB |
| 100 | 2,809 ms | **2,637.6 MB** | **44.8 MB** |

100 defined names make the same read **5x slower and allocate 21x more, while retaining exactly
the same memory**. It is a garbage-churn and CPU problem, not a footprint problem — a genuinely
different defect from the 470 bytes/cell above, and one the old allocations-only output could
not have distinguished.

`Cell`'s constructor calls `DefinedNamesExtensions.FindColumnName` **and** `FindCellNames` for
every cell. Each runs a LINQ scan across the whole defined-name dictionary and allocates a
`ToUpper()` string, a concatenated reference string and (for `FindCellNames`) a `List<string>`.
The read is therefore O(cells x names) instead of O(cells), and the per-cell cost is paid even
when a workbook declares **no** defined names at all. Excel creates defined names by itself for
print areas and tables, so this is an ordinary workbook rather than a pathological one.

## How FastExcel compares

Reading the same file, every library asked to visit every cell and materialise its value:

| Library | 200k cells | vs FastExcel | Allocated |
| --- | ---: | ---: | ---: |
| **FastExcel** | 842 ms | 1.00x | 251.7 MB |
| MiniExcel | 541 ms | **0.64x** | 412.8 MB |
| ClosedXML | 967 ms | 1.15x | 359.2 MB |
| Sylvan.Data.Excel | **89 ms** | **0.11x** | **0.33 MB** |

Read these fairly. Sylvan is a forward-only reader and does not write, which is precisely why it
is cheap — it never builds an object model. ClosedXML builds a much richer one, including styles
and formulas, and is only 15% slower than FastExcel.

The uncomfortable comparison is MiniExcel, which occupies the same niche with the same broad
capabilities and reads the file **1.6x faster**. On the evidence here the package's central claim
— that it is a *fast* way to read xlsx files — no longer holds: it is the second-slowest of four,
and the fastest option uses 0.1% of its memory.

This matters for more than tuning. It is the evidence the revive-versus-deprecate decision turns
on, which is why it lives in the repository rather than in someone's head.

## Where the per-cell cost goes

Per cell, before any defined-name work: two `Regex.Replace` calls to split the reference, a
retained `XElement` (which keeps the whole parsed `XDocument` alive), a boxed value, and a
`List<string>` of cell names. Cheap wins in rough order of payoff:

1. Short-circuit both defined-name lookups when the dictionary is empty — removes the 21x churn
   penalty for the common case, and it is a few lines.
2. Index defined names by sheet/column/row instead of scanning them per cell.
3. Parse cell references by hand rather than with `Regex`.
4. Stop retaining `XElement` per cell — the single biggest lever on the 470 bytes.
5. Offer a streaming read path for callers who do not need random access, which is the design
   that lets Sylvan run in constant memory.

## Suites

| Suite | Measures |
| --- | --- |
| `ReadBenchmarks` | Reading every cell across sizes and payload types |
| `DefinedNameScalingBenchmarks` | How read cost scales with defined-name count |
| `WriteBenchmarks` | Writing rows, numeric vs string |
| `ReadComparisonBenchmarks` | FastExcel vs MiniExcel, ClosedXML, Sylvan |
| `WriteComparisonBenchmarks` | FastExcel vs MiniExcel, ClosedXML (Sylvan is read-only) |

Comparison packages are benchmark-only and never reach the shipped FastExcel package. EPPlus is
deliberately absent: since v5 it is Polyform Noncommercial licensed, which would be a licensing
consideration even for a benchmark reference.

Automated runs skip `WriteComparisonBenchmarks`, where ClosedXML writing 50k rows dominates the
wall time. Run everything with `--filter *` or via the workflow's manual dispatch.
