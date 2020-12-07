#![forbid(unsafe_code)]

use std::fmt;
use std::io;
use std::{error, string};

/// A type alias for `Result<T, ovba::Error>`.
pub type Result<T> = std::result::Result<T, Error>;

/// Public error type.
#[derive(Debug)]
pub enum Error {
    /// I/O Error.
    Io(io::Error),
    /// Error originating from the cfb implementation.
    Cfb(io::Error),
    // TODO: Add details to make the diagnostic more meaningful to clients.
    /// Error originating from the `CompressedContainer` decompressor.
    Decompressor,
    // TODO: Add details to make the diagnostic more meaningful to clients.
    /// Generic parsing error.
    Parser,
    /// Requested module cannot be found.
    ModuleNotFound(string::String),
}

impl From<io::Error> for Error {
    // This provides automatic conversion from `io::Error` to `Error::Io`. The cfb crate doesn't provide a
    // custom error type and repurposes `io::Error` instead. Library code that handles cfb failures thus
    // needs to manually `map_err` to produce the respective `Error::Cfb` variant.
    fn from(e: io::Error) -> Self {
        Error::Io(e)
    }
}

impl error::Error for Error {
    fn source(&self) -> Option<&(dyn error::Error + 'static)> {
        match self {
            Error::Io(e) => Some(e),
            Error::Cfb(e) => Some(e),
            Error::Decompressor => None,
            Error::Parser => None,
            Error::ModuleNotFound(_) => None,
        }
    }
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Error::Io(e) => write!(f, "I/O error: {}", e),
            Error::Cfb(e) => write!(f, "CFB error: {}", e),
            Error::Decompressor => write!(f, "Decompressor error"),
            Error::Parser => write!(f, "Parse error"),
            Error::ModuleNotFound(name) => write!(f, r#"Module "{}" not found"#, name),
        }
    }
}
