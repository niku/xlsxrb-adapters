# Bootstrap Specification: `xlsxrb-adapters`

This document serves as the briefing and blueprint for the `xlsxrb-adapters` project.

---

## 1. Project Overview & Mission

- **Repository**: `xlsxrb-adapters` (RubyGem: `xlsxrb-adapters`)
- **Core Dependency**: [`xlsxrb`](https://github.com/niku/xlsxrb) (Pure Ruby, zero external dependencies, streaming $O(1)$ memory & immutable DOM).
- **Mission**:
  Provide lightweight, non-intrusive adapter proxies for popular Ruby XLSX libraries ([docs/PEER_LIBRARIES.md](file:///workspaces/xlsxrb/docs/PEER_LIBRARIES.md)) to:
  1. Enable zero-friction, drop-in performance upgrades for existing codebases via the Strangler Fig pattern.
  2. Borrow test suites and fixtures from peer libraries to build an exhaustive interoperability test harness for `xlsxrb`.
  3. Keep `xlsxrb` core strictly pure (zero dependencies, 100% mutation kill rate, Steep 0 errors) while absorbing third-party API quirks in this adapter repository.

---

## 2. Core Architectural Principles

1. **Explicit Namespacing (No Global Hijacking)**:
   - Do NOT monkey-patch or overwrite global constants (e.g., do not redefine `::RubyXL`).
   - Place adapters under explicit namespaces: `Xlsxrb::Adapters::RubyXL`, `Xlsxrb::Adapters::Roo`, `Xlsxrb::Adapters::Xlsxtream`.
   - **Rationale**: Allows original gems to coexist in the same `Gemfile`, enabling side-by-side differential testing (`assert_equal original.value, adapter.value`) and progressive safe migration.
2. **Mutable-to-Immutable Boundary**:
   - `xlsxrb` core is strictly immutable (`Data.define`, frozen rows/cells, `update_cell`).
   - The adapter layer holds lightweight mutable wrapper state (`RubyXL::Cell#change_contents`, `Worksheet#add_cell`) and materializes immutable `xlsxrb` structures upon saving (`write`).
3. **Bi-directional Bridge**:
   - Provide conversion helpers (e.g., `adapter_workbook.to_xlsxrb` and `Adapter.from_xlsxrb(xlsxrb_wb)`) so users can transition from adapters to native `xlsxrb` APIs smoothly.

---

## 3. Dev Container & Environment Setup

The repository is designed to be developed alongside `xlsxrb` in sibling directories on the host.

### `.devcontainer/devcontainer.json`
```json
{
  "name": "xlsxrb-adapters Development",
  "build": {
    "dockerfile": "Dockerfile"
  },
  "workspaceFolder": "/workspaces/xlsxrb-adapters",
  "mounts": [
    "source=${localWorkspaceFolder}/../xlsxrb,target=/workspaces/xlsxrb,type=bind"
  ],
  "remoteUser": "vscode",
  "customizations": {
    "vscode": {
      "extensions": [
        "Shopify.ruby-lsp"
      ]
    }
  },
  "postCreateCommand": "bundle install"
}
```

### `Gemfile` (Initial Development)
```ruby
source "https://rubygems.org"

gemspec

gem "xlsxrb", path: "../xlsxrb"

group :test do
  gem "minitest", "~> 5.0"
  gem "rubyXL" # For side-by-side differential testing
end
```

---

## 4. Phase 1 Target: `Xlsxrb::Adapters::RubyXL` Specification

### Scope for Phase 1
- **Focus**: Values (strings, numbers, booleans, dates, times, nil) and formulas (`Formula`).
- **Deferred**: Advanced dynamic style creation/mutation (fonts, fills, borders) — styles will be preserved or handled in later phases.

### Class Architecture

```
Xlsxrb::Adapters::RubyXL::Parser
  ├── .parse(filepath) -> Workbook
  └── .parse_buffer(io_or_string) -> Workbook

Xlsxrb::Adapters::RubyXL::Workbook
  ├── #[index_or_name] -> Worksheet?
  ├── #worksheets -> Array<Worksheet>
  ├── #write(filepath) -> void
  ├── #stream -> StringIO
  └── #to_xlsxrb -> Xlsxrb::Elements::Workbook

Xlsxrb::Adapters::RubyXL::Worksheet
  ├── #[row_index] -> Row
  ├── #add_cell(row_index, col_index, value, formula = nil) -> Cell
  ├── #each { |row| ... }
  └── #to_xlsxrb -> Xlsxrb::Elements::Worksheet

Xlsxrb::Adapters::RubyXL::Row
  ├── #[col_index] -> Cell?
  └── #cells -> Array<Cell?> (dense 0-based array, nil for empty cells)

Xlsxrb::Adapters::RubyXL::Cell
  ├── #value -> Object?
  ├── #raw_value -> Object?
  ├── #formula -> String?
  ├── #change_contents(new_value, formula = nil) -> Object
  └── #style_index -> Integer?
```

### Data Flow
1. **Reading**: `Parser.parse(file)` calls `Xlsxrb.read(file).load` to get `Elements::Workbook`, wrapping sheets/rows/cells into `RubyXL` wrapper objects.
2. **Editing**: `worksheet[row][col].change_contents(...)` and `worksheet.add_cell(...)` mutate the wrapper objects.
3. **Saving**: `workbook.write(dest)` serializes via `Xlsxrb.write(dest, wb)` or `Xlsxrb.modify` for high-throughput in-place template edits.

---

## 5. Peer Libraries Roadmap (Future Phases)

- **Phase 1 (Read-Write DOM)**: `rubyXL`
- **Phase 2 (Streaming Writer)**: `xlsxtream` (`write_row`), `fast_excel`, `write_xlsx`, `caxlsx`
- **Phase 3 (Streaming Reader)**: `roo` (`Roo::Excelx`), `simple_xlsx_reader`, `creek`, `xsv`
