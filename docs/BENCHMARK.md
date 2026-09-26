# Benchmark Results: xlsxrb-adapters vs. rubyXL & xlsxrb

This document details the performance and memory footprint measurements of `xlsxrb-adapters` (`Xlsxrb::Adapters::RubyXL`) compared to official `rubyXL` (3.4.38) and native `xlsxrb`.

## Methodology

The benchmark suite follows the isolated subprocess approach established in [`xlsxrb`](https://github.com/niku/xlsxrb):
- **Process Isolation**: Each library run executes in a dedicated subprocess (`Bundler.with_unbundled_env`) with initial GC compaction to eliminate cross-library memory pollution and retain pure execution measurements.
- **Time Measurement**: Monotonic clock via `Process.clock_gettime(Process::CLOCK_MONOTONIC)`.
- **Peak Memory**: Linux kernel High Water Mark (`VmHWM`) via `/proc/self/status`, tracking maximum physical memory consumption.
- **GC Overhead**: Tracking garbage collection cycles with `GC.stat[:count]`.
- **Statistical Aggregation**: 3 isolated runs per benchmark, reporting Median as primary and Mean as secondary.

---

## 1. Large-Scale Benchmark (1,000,000 cells: 100,000 rows × 10 cols)

### Write Performance (Generation & Export)

| Library / Mode | Model | Time (Median) | Time (Mean) | Peak Memory (VmHWM) | GC Count | vs. rubyXL Speed | vs. rubyXL Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`rubyXL` (3.4.38)** | In-Memory | 35.010 s | 35.023 s | 2,198.8 MB | 94.0 | 1.0x (baseline) | 100% (baseline) |
| **`xlsxrb-adapters`** | In-Memory (RubyXL API) | **4.696 s** | **4.716 s** | **489.0 MB** | **16.0** | **~7.5x faster** | **77.8% reduced** |
| `xlsxrb (In-Memory)` | In-Memory (Native) | 2.558 s | 2.667 s | 240.6 MB | 12.0 | ~13.7x faster | 89.1% reduced |
| `xlsxrb (Streaming)` | Streaming (Native) | 0.948 s | 0.959 s | 74.9 MB | 27.0 | ~36.9x faster | 96.6% reduced |

### Read Performance (Parsing & Cell Iteration)

| Library / Mode | Model | Time (Median) | Time (Mean) | Peak Memory (VmHWM) | GC Count | vs. rubyXL Speed | vs. rubyXL Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`rubyXL` (3.4.38)** | In-Memory | 35.718 s | 37.097 s | 2,469.5 MB | 140.0 | 1.0x (baseline) | 100% (baseline) |
| **`xlsxrb-adapters`** | In-Memory (RubyXL API) | **4.504 s** | **4.505 s** | **556.1 MB** | **35.0** | **~7.9x faster** | **77.5% reduced** |
| `xlsxrb (In-Memory)` | In-Memory (Native) | 2.307 s | 2.337 s | 392.4 MB | 28.0 | ~15.5x faster | 84.1% reduced |
| `xlsxrb (Streaming)` | Streaming (Native) | 1.972 s | 1.982 s | 87.6 MB | 51.0 | ~18.1x faster | 96.5% reduced |

---

## 2. Scaling Profile across Dataset Sizes

Measurements across varying workloads (Median values):

| Dataset Size | Metric | Official `rubyXL` | `xlsxrb-adapters` | Improvement |
| :--- | :--- | :--- | :--- | :--- |
| **10,000 cells**<br>(1,000 rows × 10 cols) | Write Time / Memory<br>Read Time / Memory | 0.218 s / 57.1 MB<br>0.224 s / 64.6 MB | **0.042 s / 53.9 MB**<br>**0.040 s / 60.1 MB** | **5.2x faster** / 6% less memory<br>**5.6x faster** / 7% less memory |
| **100,000 cells**<br>(10,000 rows × 10 cols) | Write Time / Memory<br>Read Time / Memory | 2.605 s / 256.6 MB<br>2.702 s / 274.3 MB | **0.454 s / 97.3 MB**<br>**0.405 s / 108.6 MB** | **5.7x faster** / 62% less memory<br>**6.7x faster** / 60% less memory |
| **1,000,000 cells**<br>(100,000 rows × 10 cols) | Write Time / Memory<br>Read Time / Memory | 35.010 s / 2,198.8 MB<br>35.718 s / 2,469.5 MB | **4.696 s / 489.0 MB**<br>**4.504 s / 556.1 MB** | **7.5x faster** / 78% less memory<br>**7.9x faster** / 77% less memory |

---

## 3. Architectural Design & Trade-offs

### Two Complementary Approaches to XLSX Processing

The performance and memory characteristics reflect different design goals, each serving distinct application needs:

1. **`rubyXL`'s Design Philosophy — Full Document Object Model (DOM) & Fidelity**:
   - `rubyXL` is a mature, feature-rich library designed to parse and preserve the complete OpenXML tree structure, supporting in-depth template editing, complex styles, drawings, and arbitrary XML nodes.
   - To provide this level of manipulation, `rubyXL` builds comprehensive XML node objects for each workbook component. For typical spreadsheets (thousands of cells), this provides unbeatable convenience and fidelity. For very large datasets (100,000+ rows), memory scales with the node graph size.

2. **`xlsxrb-adapters`'s Approach — High-Throughput Batch Processing with rubyXL API Familiarity**:
   - `xlsxrb-adapters` bridges the familiar, intuitive API established by `rubyXL` with `xlsxrb`'s low-overhead streaming pull parser and immutable data pipeline.
   - **Read**: XML parsing delegates directly to `xlsxrb`'s tokenized stream parser before creating lightweight adapter row/cell wrappers.
   - **Write**: The mutable adapter structure compiles into frozen data elements (`Data.define`) via `to_xlsxrb`, which are written out through `xlsxrb`'s streaming shared strings table (SST) writer.
   - This delivers a drop-in option for data-intensive pipelines (e.g. large ETL, bulk export/import) where memory footprint and processing speed are critical, while preserving 100% of existing `rubyXL` code patterns.

---

## 4. Acknowledgements

We express our deep respect and gratitude to the creators and maintainers of **`rubyXL`** (Vivek Bhagwat, Wesha, and all contributors). `rubyXL` has set the standard for Excel manipulation in the Ruby ecosystem for over a decade. The clean, intuitive API design of `xlsxrb-adapters`—and the very test suite that guarantees its 100% compatibility—stands upon their pioneering work.

---

## 5. Reproducing Locally

Run the benchmark suite with:

```bash
# Run standard 10,000 rows x 10 cols benchmark
bundle exec rake benchmark

# Run 100,000 rows x 10 cols benchmark with 3 runs
RUNS=3 bundle exec rake benchmark rows=100000 cols=10
```
