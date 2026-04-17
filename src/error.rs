use std::io;
use std::path::PathBuf;

use thiserror::Error;

/// Why `xlpath` declined to process a particular input file. Emitted to stderr
/// as a warning; the run continues with the remaining files.
#[derive(Debug, Error)]
pub enum SkipReason {
    #[error("encrypted workbook")]
    Encrypted,

    #[error("corrupt zip: {0}")]
    CorruptZip(String),

    #[error("malformed XML in `{part}`: {message}")]
    MalformedXml { part: String, message: String },

    #[error("i/o error: {0}")]
    Io(#[from] io::Error),
}

/// A per-file warning bundled with the path it applies to.
#[derive(Debug)]
pub struct FileWarning {
    pub path: PathBuf,
    pub reason: SkipReason,
}

impl FileWarning {
    pub fn format(&self) -> String {
        format!("xlpath: {}: {}\n", self.path.display(), self.reason)
    }
}
