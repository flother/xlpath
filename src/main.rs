use std::io::{self, BufReader};
use std::path::PathBuf;
use std::process::ExitCode;
use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};

use anyhow::{anyhow, Context, Result};
use clap::Parser;
use rayon::prelude::*;

use xlpath::cli::Cli;
use xlpath::error::{FileWarning, SkipReason};
use xlpath::output::{self, Writer};
use xlpath::walk;
use xlpath::xlsx::{self, PartFilter};
use xlpath::xpath::{EvalOptions, Query};

fn main() -> ExitCode {
    reset_sigpipe();
    match run() {
        Ok(code) => code,
        Err(e) => {
            eprintln!("xlpath: {e:#}");
            ExitCode::from(2)
        }
    }
}

/// Restore the default `SIGPIPE` handler so that writing to a closed pipe (e.g.
/// `xlpath ... | head -1`) terminates the process immediately, matching the
/// convention used by grep, ripgrep, fd, and other grep-like tools. Without
/// this, Rust's default behaviour turns the signal into an
/// `ErrorKind::BrokenPipe` that our worker loop would otherwise swallow,
/// leaving the process churning through remaining workbooks for no effect.
#[cfg(unix)]
fn reset_sigpipe() {
    // Safety: setting SIGPIPE to SIG_DFL before spawning any threads is a
    // well-established idiom; it affects process-global signal state and so
    // must run before any rayon worker or output writer comes up.
    unsafe {
        libc::signal(libc::SIGPIPE, libc::SIG_DFL);
    }
}

#[cfg(not(unix))]
fn reset_sigpipe() {}

fn run() -> Result<ExitCode> {
    let mut cli = Cli::parse();

    let ns_pairs = parse_namespace_args(&cli.namespaces)?;
    let query = Query::compile(&cli.xpath, &ns_pairs)
        .map_err(|e| anyhow!("{}", e))
        .context("failed to compile XPath expression")?;

    let filter = PartFilter::new(&cli.includes, &cli.excludes)
        .map_err(|e| anyhow!("{}", e))
        .context("failed to compile include/exclude globs")?;

    // Resolve the positional path arguments into a concrete file list. With no
    // PATH given, default to the current working directory so that `xlpath
    // <XPATH>` surveys the whole tree from where the user is standing.
    let inputs: Vec<PathBuf> = if cli.paths.is_empty() {
        vec![PathBuf::from(".")]
    } else {
        std::mem::take(&mut cli.paths)
    };
    let stdin = io::stdin();
    let mut walk_warnings = String::new();
    let paths = walk::collect(&inputs, BufReader::new(stdin.lock()), cli.follow, |e| {
        let msg = match e.path() {
            Some(p) => format!(
                "xlpath: {}: {}\n",
                p.display(),
                e.io_error()
                    .map_or_else(|| e.to_string(), |ie| ie.to_string())
            ),
            None => format!("xlpath: {e}\n"),
        };
        walk_warnings.push_str(&msg);
    })
    .context("failed to resolve input paths")?;

    // Explicit thread count if given, otherwise rayon's default (logical CPUs).
    if let Some(n) = cli.threads {
        // `build_global` fails if the global pool was already initialised. The
        // only caller is this one-shot CLI entry point, so that can only happen
        // in tests that exercise `run()` twice in-process; silently keeping the
        // existing pool is the right behaviour there.
        rayon::ThreadPoolBuilder::new()
            .num_threads(n.max(1))
            .build_global()
            .ok();
    }

    let writer = Writer::new();
    let match_count = AtomicUsize::new(0);
    let had_error = AtomicBool::new(false);

    if !walk_warnings.is_empty() {
        had_error.store(true, Ordering::Relaxed);
        writer.emit_err(&walk_warnings);
    }

    let mode = cli.output_mode();
    let no_filename = cli.no_filename;
    let no_path_flag = cli.no_part;
    let eval_opts = EvalOptions { as_tag: cli.tag };

    paths.par_iter().for_each(|path| {
        let mut per_part: Vec<(String, Vec<xlpath::xpath::Match>)> = Vec::new();
        let mut part_warnings = String::new();

        let result = xlsx::process_parts(path, &filter, |part_name, data| {
            let Ok(xml) = std::str::from_utf8(data) else {
                part_warnings.push_str(
                    &FileWarning {
                        path: path.clone(),
                        reason: SkipReason::MalformedXml {
                            part: part_name.to_string(),
                            message: "not valid UTF-8".to_string(),
                        },
                    }
                    .format(),
                );
                had_error.store(true, Ordering::Relaxed);
                return;
            };

            match query.evaluate_xml_with(xml, eval_opts) {
                Ok(matches) => {
                    per_part.push((part_name.to_string(), matches));
                }
                Err(e) => {
                    part_warnings.push_str(
                        &FileWarning {
                            path: path.clone(),
                            reason: SkipReason::MalformedXml {
                                part: part_name.to_string(),
                                message: e.to_string(),
                            },
                        }
                        .format(),
                    );
                    had_error.store(true, Ordering::Relaxed);
                }
            }
        });

        match result {
            Ok(()) => {
                let total: usize = per_part.iter().map(|(_, ms)| ms.len()).sum();
                if total > 0 {
                    match_count.fetch_add(total, Ordering::Relaxed);
                }
                let rendered =
                    output::format_file(mode, no_filename, no_path_flag, path, &per_part);
                if !rendered.is_empty() {
                    writer.emit_out(&rendered);
                }
                // Per-part warnings surface on stderr after any matches from
                // the same file have been printed.
                if !part_warnings.is_empty() {
                    writer.emit_err(&part_warnings);
                }
            }
            Err(reason) => {
                had_error.store(true, Ordering::Relaxed);
                let warning = FileWarning {
                    path: path.clone(),
                    reason,
                };
                writer.emit_err(&warning.format());
            }
        }
    });

    writer.flush();

    if had_error.load(Ordering::Relaxed) {
        Ok(ExitCode::from(2))
    } else if match_count.load(Ordering::Relaxed) > 0 {
        Ok(ExitCode::from(0))
    } else {
        Ok(ExitCode::from(1))
    }
}

fn parse_namespace_args(args: &[String]) -> Result<Vec<(String, String)>> {
    let mut out = Vec::with_capacity(args.len());
    for a in args {
        let (prefix, uri) = a
            .split_once('=')
            .ok_or_else(|| anyhow!("--ns expects `prefix=uri`, got `{a}`"))?;
        if prefix.is_empty() {
            return Err(anyhow!("--ns prefix may not be empty: `{a}`"));
        }
        out.push((prefix.to_string(), uri.to_string()));
    }
    Ok(out)
}
