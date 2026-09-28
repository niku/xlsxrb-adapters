# Benchmark Results: xlsxrb-adapters vs. caxlsx, rubyXL & xlsxrb

This document provides dedicated benchmarks comparing migration paths from peer XLSX libraries to `xlsxrb`:
1. **Caxlsx Migration**: Comparing `caxlsx` (4.5.0), `xlsxrb-adapters (Caxlsx)` (both In-Memory and Streaming), and native `xlsxrb`.
2. **RubyXL Migration**: Comparing `rubyXL` (3.4.38), `xlsxrb-adapters (RubyXL)`, and native `xlsxrb`.

## Methodology

The benchmark suite follows the isolated subprocess approach established in [`xlsxrb`](https://github.com/niku/xlsxrb):
- **Process Isolation**: Each library run executes in a dedicated subprocess (`Bundler.with_unbundled_env`) with initial GC compaction to eliminate cross-library memory pollution and retain pure execution measurements.
- **Time Measurement**: Monotonic clock via `Process.clock_gettime(Process::CLOCK_MONOTONIC)`.
- **Peak Memory**: Linux kernel High Water Mark (`VmHWM`) via `/proc/self/status`, tracking maximum physical memory consumption.
- **GC Overhead**: Tracking garbage collection cycles with `GC.stat[:count]`.
- **Statistical Aggregation**: Isolated runs per benchmark, reporting Median and Mean.

---

## 1. Caxlsx Migration Benchmarks (caxlsx vs. xlsxrb-adapters vs. xlsxrb)

*Note: Official `caxlsx` is a write-only library designed for building and exporting spreadsheets. Accordingly, benchmarks evaluate generation and serialization throughput and memory efficiency. To read XLSX files, use `xlsxrb` directly (`Xlsxrb.read` / `Xlsxrb.stream`).*

### 1.1 Large-Scale (1,000,000 cells: 100,000 rows × 10 cols)

#### Write Performance (Generation & Export)

| Library / Mode | Adapter / Engine | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. caxlsx Speed |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **`caxlsx` (4.5.0)** | Native Generative Tree | 5.732 s | 186.9 MB | 20.0 | 1.0x (baseline) |
| **`xlsxrb-adapters (Caxlsx)`** | **Caxlsx API + xlsxrb core** | **5.710 s** | **385.1 MB** | **19.0** | **~1.0x (matches)** |
| **`xlsxrb-adapters (Caxlsx Streaming)`** | **Caxlsx DSL + Streaming Writer** | **0.998 s** | **78.2 MB** | **18.0** | **~5.7x faster** |
| `xlsxrb (In-Memory)` | Native `Xlsxrb::Elements` | 2.723 s | 240.7 MB | 12.0 | ~2.1x faster |
| `xlsxrb (Streaming)` | Native Streaming Writer | 0.971 s | 75.3 MB | 27.0 | ~5.9x faster |

---

### 1.2 Medium-Scale (100,000 cells: 10,000 rows × 10 cols)

#### Write Performance

| Library | Time (Median) | Peak Memory | GC Count | vs. caxlsx Speed |
| :--- | :--- | :--- | :--- | :--- |
| **`caxlsx` (Original 4.5.0)** | 0.539 s | 49.1 MB | 7.0 | 1.0x (baseline) |
| **`xlsxrb-adapters (Caxlsx)`** | **0.493 s** | **82.9 MB** | **6.0** | **~1.1x faster** |
| **`xlsxrb-adapters (Caxlsx Streaming)`** | **0.099 s** | **60.2 MB** | **1.0** | **~5.4x faster** |
| `xlsxrb (In-Memory)` | 0.227 s | 66.8 MB | 4.0 | ~2.4x faster |
| `xlsxrb (Streaming)` | 0.087 s | 49.3 MB | 6.0 | ~6.2x faster |

---

### 1.3 Small-Scale (10,000 cells: 1,000 rows × 10 cols)

#### Write Performance

| Library | Time (Median) | Peak Memory | GC Count | vs. caxlsx Speed |
| :--- | :--- | :--- | :--- | :--- |
| **`caxlsx` (Original 4.5.0)** | 0.081 s | 36.6 MB | 0.0 | 1.0x (baseline) |
| **`xlsxrb-adapters (Caxlsx)`** | **0.057 s** | **55.4 MB** | **0.0** | **~1.4x faster** |
| **`xlsxrb-adapters (Caxlsx Streaming)`** | **0.011 s** | **50.2 MB** | **0.0** | **~7.3x faster** |
| `xlsxrb (In-Memory)` | 0.025 s | 48.1 MB | 1.0 | ~3.2x faster |
| `xlsxrb (Streaming)` | 0.011 s | 47.1 MB | 0.0 | ~7.3x faster |

---

## 2. RubyXL Migration Benchmarks (rubyXL vs. xlsxrb-adapters vs. xlsxrb)

### 2.1 Large-Scale (1,000,000 cells: 100,000 rows × 10 cols)

#### Write Performance (Generation & Export)

| Library / Mode | Adapter / Engine | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. rubyXL Speed | vs. rubyXL Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`rubyXL` (3.4.38)** | Native In-Memory DOM | 48.733 s | 2,198.5 MB | 94.0 | 1.0x (baseline) | 100% (baseline) |
| **`xlsxrb-adapters (RubyXL)`** | **RubyXL API + xlsxrb core** | **5.178 s** | **488.9 MB** | **16.0** | **~9.4x faster** | **77.8% reduced** |
| `xlsxrb (In-Memory)` | Native `Xlsxrb::Elements` | 2.723 s | 240.7 MB | 12.0 | ~17.9x faster | 89.1% reduced |
| `xlsxrb (Streaming)` | Native Streaming Writer | 0.971 s | 75.3 MB | 27.0 | ~50.2x faster | 96.6% reduced |

#### Read Performance (Parsing & Cell Iteration)

| Library / Mode | Adapter / Engine | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. rubyXL Speed | vs. rubyXL Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **`rubyXL` (3.4.38)** | Native In-Memory DOM | 57.216 s | 2,462.4 MB | 140.0 | 1.0x (baseline) | 100% (baseline) |
| **`xlsxrb-adapters (RubyXL)`** | **RubyXL API + xlsxrb core** | **5.075 s** | **556.5 MB** | **35.0** | **~11.3x faster** | **77.4% reduced** |
| `xlsxrb (In-Memory)` | Native `Xlsxrb::Elements` | 2.779 s | 394.9 MB | 28.0 | ~20.6x faster | 84.0% reduced |
| `xlsxrb (Streaming)` | Native Streaming Reader | 1.811 s | 87.7 MB | 51.0 | ~31.6x faster | 96.4% reduced |

---

### 2.2 Medium-Scale (100,000 cells: 10,000 rows × 10 cols)

#### Write Performance

| Library | Time (Median) | Peak Memory | GC Count | vs. rubyXL Speed | vs. rubyXL Memory |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **`rubyXL` (Original 3.4.38)** | 2.668 s | 256.0 MB | 42.0 | 1.0x (baseline) | 100% (baseline) |
| **`xlsxrb-adapters (RubyXL)`** | **0.460 s** | **94.5 MB** | **10.0** | **~5.8x faster** | **63.1% reduced** |
| `xlsxrb (In-Memory)` | 0.227 s | 66.8 MB | 4.0 | ~11.7x faster | 73.9% reduced |
| `xlsxrb (Streaming)` | 0.087 s | 49.3 MB | 6.0 | ~30.6x faster | 80.7% reduced |

#### Read Performance

| Library | Time (Median) | Peak Memory | GC Count | vs. rubyXL Speed | vs. rubyXL Memory |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **`rubyXL` (Original 3.4.38)** | 2.705 s | 272.0 MB | 50.0 | 1.0x (baseline) | 100% (baseline) |
| **`xlsxrb-adapters (RubyXL)`** | **0.400 s** | **108.0 MB** | **8.0** | **~6.8x faster** | **60.3% reduced** |
| `xlsxrb (In-Memory)` | 0.202 s | 84.1 MB | 5.0 | ~13.4x faster | 69.1% reduced |
| `xlsxrb (Streaming)` | 0.182 s | 83.1 MB | 11.0 | ~14.9x faster | 69.4% reduced |

---

### 2.3 Small-Scale (10,000 cells: 1,000 rows × 10 cols)

#### Write Performance

| Library | Time (Median) | Peak Memory | GC Count | vs. rubyXL Speed |
| :--- | :--- | :--- | :--- | :--- |
| **`rubyXL` (Original 3.4.38)** | 0.233 s | 57.3 MB | 13.0 | 1.0x (baseline) |
| **`xlsxrb-adapters (RubyXL)`** | **0.045 s** | **53.9 MB** | **0.0** | **~5.2x faster** |
| `xlsxrb (In-Memory)` | 0.025 s | 48.1 MB | 1.0 | ~9.3x faster |
| `xlsxrb (Streaming)` | 0.011 s | 47.1 MB | 0.0 | ~21.2x faster |

#### Read Performance

| Library | Time (Median) | Peak Memory | GC Count | vs. rubyXL Speed |
| :--- | :--- | :--- | :--- | :--- |
| **`rubyXL` (Original 3.4.38)** | 0.219 s | 64.5 MB | 15.0 | 1.0x (baseline) |
| **`xlsxrb-adapters (RubyXL)`** | **0.040 s** | **60.2 MB** | **0.0** | **~5.5x faster** |
| `xlsxrb (In-Memory)` | 0.023 s | 52.1 MB | 1.0 | ~9.5x faster |
| `xlsxrb (Streaming)` | 0.031 s | 54.0 MB | 1.0 | ~7.1x faster |

---

## 3. Key Takeaways

1. **Caxlsx Migration**:
   - `xlsxrb-adapters (Caxlsx)` provides 100% drop-in API compatibility for existing caxlsx generation code with zero application changes.
   - `xlsxrb-adapters (Caxlsx Streaming)` offers **5.4x–7.3x faster writes** with constant $O(1)$ memory, enabling massive report generation without RAM exhaustion.
2. **RubyXL Migration**:
   - `xlsxrb-adapters (RubyXL)` executes **~9.4x–11.3x faster** on 1,000,000 cells while reducing RAM consumption by **~78%** (~489 MB vs. ~2.2 GB).
   - Full 100% compatibility with official rubyXL spec suite without any code changes in user applications.

---

## 4. Reproducing Locally

Run the benchmark suite with:

```bash
# Run all migration benchmarks (default: 10,000 rows x 10 cols)
bundle exec ruby benchmark.rb 10000 10 all

# Run only Caxlsx migration benchmarks
bundle exec ruby benchmark.rb 10000 10 caxlsx

# Run only RubyXL migration benchmarks
bundle exec ruby benchmark.rb 10000 10 rubyxl
```
