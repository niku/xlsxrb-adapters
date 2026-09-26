# xlsxrb-adapters

Compatibility adapters for migrating from peer XLSX libraries (e.g. `rubyXL`) to [`xlsxrb`](https://github.com/niku/xlsxrb) with zero downtime and drop-in safety.

## Overview

[`rubyXL`](https://github.com/weshatheleopard/rubyXL) is a mature, full-featured library for comprehensive OpenXML spreadsheet creation and manipulation in Ruby, maintained for more than a decade. Its intuitive object model (`sheet.add_cell`, row/cell indexing, and in-memory mutability) is familiar to many Ruby developers.

[`xlsxrb`](https://github.com/niku/xlsxrb) is a pure Ruby, zero-dependency, streaming-capable, low-memory XLSX engine designed for high throughput.

`xlsxrb-adapters` bridges the best of both worlds by providing drop-in compatible adapter layers to:
1. Support gradual migration (Strangler Fig pattern) from peer XLSX libraries to `xlsxrb` without rewriting application logic.
2. Enable high-throughput, low-memory execution in batch processing and resource-constrained environments while retaining `rubyXL`'s familiar and battle-tested API.
3. Construct interoperability test harnesses against peer libraries and real-world fixtures.
4. Keep `xlsxrb` core strictly zero-dependency, mutant-tested, and type-safe while providing rich compatibility layers.

## Design Principles

1. **No Global Hijacking**:
   Does not reopen or hijack top-level constants like `::RubyXL`. Instead, exposes namespaces like `Xlsxrb::Adapters::RubyXL` so you can use both in the same application during migration or run side-by-side comparison tests.
2. **Mutable-to-Immutable Boundary**:
   Maintains a mutable in-memory wrapper structure compatible with legacy workflows, translating into `xlsxrb`'s immutable data models (`Data.define`, frozen) upon save/export.
3. **Native Bridge Conversion**:
   Supports seamless conversions between adapter structures and native `xlsxrb` objects via `adapter_wb.to_xlsxrb` and `Xlsxrb::Adapters::RubyXL.from_xlsxrb(xlsxrb_wb)`.

## Development & Dev Container

The Dev Container configuration mounts the sibling `../xlsxrb` repository at `/workspaces/xlsxrb` inside the container:

```json
"mounts": [
  "source=${localWorkspaceFolder}/../xlsxrb,target=/workspaces/xlsxrb,type=bind"
]
```

Run test suite with:

```bash
bundle install
bundle exec rake test
```

## License

MIT
