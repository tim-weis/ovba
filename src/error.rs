#![forbid(unsafe_code)]
#![warn(rust_2018_idioms)]

use thiserror::Error;

#[derive(Debug, Error)]
pub enum Error {
    #[error("Not a valid Office Open XML document.\n\t{0}")]
    InvalidDocument(Box<dyn std::error::Error>),
    #[error("Failed to perform I/O.\n\t{0}")]
    Io(Box<dyn std::error::Error>),
    // TODO: Implement proper error handling. The module ovba should probably get its own error type.
    #[error("Unknown error.")]
    Unknown,
}
