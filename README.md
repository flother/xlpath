xlpath
======

`xlpath` is a CLI to query the XML within XLSX files using XPath. You can run an XPath expression
against that XML. The output is grep-like (`file:part: value`) by default, so it composes with the
usual shell tools for surveying feature usage across folders of workbooks.

User documentation lives in [`docs/`](docs/) and can be read in the repo or at
<https://code.flother.is/xlpath>.

Building and testing
--------------------

Usual Rust commands apply:

```sh
cargo build                     # Debug build
cargo build --release           # Release build
cargo test                      # All unit and integration tests
cargo test --test integration   # Integration tests only
cargo clippy --all-targets      # Lint with Clippy
cargo fmt                       # Format source code
cargo run -- <XPATH> <PATH>...  # Run the CLI locally
```

Architecture
------------

The crate is both a library (`src/lib.rs`) and a binary (`src/main.rs`). The codebase has eight
modules.

- `main`: compiles the query and entry filter, resolves paths, runs one Rayon task per input file.
  Each worker builds its file's output as a single `String` and emits it in one call. Exit codes are
  `0` (matches), `1` (no matches), or `2` (at least one file errored).
- `cli`: clap `Parser` struct and `OutputMode` enum. Flag conflicts via clap attributes rather than
  runtime checks.
- `walk`: resolves positional paths into a concrete file list. Recursive directory walk, `-` meaning
  stdin, extension filter (excluding `.xlsb`), and `~$` Excel-lock-file skipping.
- `xlsx`: opens each workbook, detects OLE2-wrapped (encrypted) files by magic-number peek before
  handing to the zip parser, and streams internal XML and `.rels` parts through a `PartFilter`.
- `xpath`: namespace registry seeded from OOXML defaults, compiled query, and per-document
  evaluation. The XPath expression is re-parsed per document because `sxd-xpath`'s compiled tree
  isn't `Send`/`Sync`. Storing the expression string lets a single query be shared across Rayon
  workers without `unsafe`.
- `output`: formatters for each `OutputMode`
- `error`: error types

Testing conventions
-------------------

- Unit tests live in `#[cfg(test)] mod tests` blocks at the bottom of each module.
- Integration tests are in `tests/integration.rs` and use `assert_cmd` to invoke the built binary.
  `tests/common/mod.rs` provides helpers and canonical XML fixtures. Tests construct synthetic
  zips on the fly rather than checking in `.xlsx` binaries.

Licence
-------

MIT OR Apache-2.0.
