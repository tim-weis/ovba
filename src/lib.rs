//! A crate for parsing Office VBA projects and extracting compressed content.
//!
//! This is an implementation of the [\[MS-OVBA\]: Office VBA File Format Structure][MS-OVBA] protocol
//! (Revision 9.1, published 2020-02-19).
//!
//! [MS-OVBA]: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc

#![forbid(unsafe_code)]
#![warn(rust_2018_idioms)]

pub mod error;
pub type Result<T> = std::result::Result<T, error::Error>;

// TODO: Implement better error handling.
#[doc(inline)]
pub use crate::error::Error;

use cfb::CompoundFile;

use std::io::{Cursor, Read};

/// Represents a VBA project.
pub struct Project {
    // TODO: Figure out how to make this generic (attempts have failed with trait bound violations)
    #[doc(hidden)]
    container: CompoundFile<Cursor<Vec<u8>>>,
}

#[derive(Debug)]
pub enum SysKind {
    Win16,
    Win32,
    MacOs,
    Win64,
}

/// Version Independent Project Information
#[derive(Debug)]
pub struct ProjectInformation {
    pub information: Information,
    pub references: Vec<Reference>,
    pub modules: Modules,
}

#[derive(Debug)]
pub struct ReferenceControl {
    /// (Optional) Name and NameUnicode entries
    name: Option<(String, String)>,
    libid_original: Option<String>,
    libid_twiddled: String,
    name_extended: Option<(String, String)>,
    libid_extended: String,
    guid: Vec<u8>, // Should be an `[u8; 16]`, though I'm not sure how to convert &[u8] returned by the parser into an array.
    /// Unique for each `ReferenceControl`
    cookie: u32,
}

#[derive(Debug)]
pub struct ReferenceOriginal {
    /// (Optional) Name and NameUnicode entries
    name: Option<(String, String)>,
    libid_original: String,
}

#[derive(Debug)]
pub struct ReferenceRegistered {
    name: Option<(String, String)>,
    libid: String,
}

#[derive(Debug)]
pub struct ReferenceProject {
    name: Option<(String, String)>,
    libid_absolute: String,
    libid_relative: String,
    major_version: u32,
    minor_version: u16,
}

#[derive(Debug)]
pub enum Reference {
    Control(ReferenceControl),
    Original(ReferenceOriginal),
    Registered(ReferenceRegistered),
    Project(ReferenceProject),
}

#[derive(Debug)]
pub struct Information {
    /// System kind.
    pub sys_kind: SysKind,
    lcid: u32,
    lcid_invoke: u32,
    code_page: u16,
    name: String,
    doc_string: String,
    doc_string_unicode: String,
    help_file_1: String,
    help_file_2: String,
    help_context: u32,
    lib_flags: u32,
    version_major: u32,
    version_minor: u16,
    constants: String,
    constants_unicode: String,
}

#[derive(Debug)]
pub struct Modules {
    pub count: u16,
    pub cookie: u16,
    pub modules: Vec<Module>,
}

#[derive(Debug)]
pub enum ModuleType {
    Procedural,
    DocClsDesigner,
}

#[derive(Debug)]
pub struct Module {
    pub name: String,
    pub name_unicode: Option<String>,
    pub stream_name: String,
    pub stream_name_unicode: String,
    pub doc_string: String,
    pub doc_string_unicode: String,
    pub text_offset: u32,
    pub help_context: u32,
    pub cookie: u16,
    pub module_type: ModuleType,
    pub read_only: bool,
    pub private: bool,
}

impl Project {
    pub fn list(&self) -> Result<Vec<(String, String)>> {
        let mut result = Vec::new();
        for entry in self.container.walk_storage("/").map_err(Error::Cfb)? {
            result.push((
                entry.name().to_owned(),
                entry.path().to_str().unwrap_or_default().to_owned(),
            ));
        }
        Ok(result)
    }

    pub fn read_stream(&mut self, stream_name: &str) -> Result<Vec<u8>> {
        let mut stream = self
            .container
            .open_stream(stream_name)
            .map_err(Error::Cfb)?;
        let mut buffer = Vec::new();
        stream.read_to_end(&mut buffer).map_err(Error::Cfb)?;

        Ok(buffer)
    }

    pub fn decompress_stream_from(&mut self, stream_name: &str, offset: usize) -> Result<Vec<u8>> {
        let data = self.read_stream(stream_name)?;
        let data = parser::decompress(&data[offset..])
            .map_err(|_| Error::Decompressor)?
            .1;
        Ok(data)
    }

    /// Returns version independent project information.
    pub fn information(&mut self) -> Result<ProjectInformation> {
        const DIR_STREAM_PATH: &str = r#"/VBA\dir"#;

        // Read *dir* stream
        let mut stream = self
            .container
            .open_stream(DIR_STREAM_PATH)
            .map_err(Error::Cfb)?;
        let mut buffer = Vec::new();
        stream.read_to_end(&mut buffer).map_err(Error::Cfb)?;

        // Decompress stream
        let (remainder, buffer) = parser::decompress(&buffer).map_err(|_| Error::Decompressor)?;
        debug_assert!(remainder.is_empty());

        // Parse binary data
        let (remainder, information) =
            parser::parse_project_information(&buffer).map_err(|_| Error::Parser)?;
        debug_assert_eq!(remainder.len(), 0, "Stream not fully consumed");

        // Return structured information
        Ok(information)
    }
}

pub fn open_project(raw: Vec<u8>) -> Result<Project> {
    let cursor = Cursor::new(raw);
    let container = CompoundFile::open(cursor).map_err(Error::Cfb)?;
    let proj = Project { container };

    Ok(proj)
}
