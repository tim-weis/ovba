//! An Office VBA project parser written in 100% safe Rust.
//!
//! This is a (partial) implementation of the [\[MS-OVBA\]: Office VBA File Format
//! Structure][MS-OVBA] protocol (Revision 9.1, published 2020-02-19).
//!
//! The main entry point into the API is the [`Project`] type, returned by the
//! [`open_project`] function.
//!
//! # Usage
//!
//! Opening a project:
//!
//! ```rust,no_run
//! use std::fs::read;
//! use ovba::open_project;
//!
//! let data = read("vbaProject.bin")?;
//! let project = open_project(data)?;
//! # Ok::<(), ovba::Error>(())
//! ```
//!
//! A more complete example that dumps an entire VBA project's source code:
//!
//! ```rust,no_run
//! use std::fs::{read, write};
//! use ovba::open_project;
//!
//! let data = read("vbaProject.bin")?;
//! let project = open_project(data)?;
//!
//! for module in &project.modules {
//!     let src_code = project.module_source_raw(&module.name)?;
//!     write("./out/".to_string() + &module.name, src_code)?;
//! }
//! # Ok::<(), ovba::Error>(())
//! ```
//!
//! The API also supports low-level access to the [\[MS-CFB\]: Compound File Binary File
//! Format][MS-CFB] data. The following example lists all CFB entries:
//!
//! ```rust,no_run
//! use std::fs::read;
//! use ovba::open_project;
//!
//! let data = read("vbaProject.bin")?;
//! let project = open_project(data)?;
//! for (name, path) in &project.list()? {
//!     println!(r#"Name: "{}"; Path: "{}""#, name, path);
//! }
//! # Ok::<(), ovba::Error>(())
//! ```
//!
//! [MS-OVBA]: https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc
//! [MS-CFB]: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b

#![forbid(unsafe_code)]
#![warn(rust_2018_idioms, missing_docs)]

mod error;
pub use crate::error::{Error, Result};

mod parser;

use cfb::CompoundFile;
use parser::cp_to_string;

use std::{
    cell::RefCell,
    io::{Cursor, Read},
    path::Path,
};

/// Represents a VBA project.
///
/// This type serves as the entry point into this crate's functionality and exposes the
/// public API surface.
pub struct Project {
    /// Specifies version-independent information for the VBA project.
    pub information: Information,
    /// Specifies the external references of the VBA project.
    pub references: Vec<Reference>,
    /// Specifies the modules in the project.
    pub modules: Vec<Module>,
    // TODO: Figure out how to make this generic (attempts have failed with
    //       trait bound violations). This would allow [`open_project`] to
    //       accept a wider range of input types.
    container: RefCell<CompoundFile<Cursor<Vec<u8>>>>,
}

/// Specifies the platform for which the VBA project is created.
#[derive(Debug)]
pub enum SysKind {
    /// For 16-bit Windows Platforms.
    Win16,
    /// For 32-bit Windows Platforms.
    Win32,
    /// For Macintosh Platforms.
    MacOs,
    /// For 64-bit Windows Platforms.
    Win64,
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]

/// Specifies a reference to a twiddled type library and its extended type library.
#[derive(Debug)]
pub struct ReferenceControl {
    /// (Optional) Name entry
    name: Option<String>,
    libid_original: Option<String>,
    libid_twiddled: String,
    name_extended: Option<String>,
    libid_extended: String,
    guid: Vec<u8>, // Should be an `[u8; 16]`, though I'm not sure how to convert &[u8] returned by the parser into an array.
    /// MUST be Unique for each `ReferenceControl` in the VBA projectwith the same
    /// libid_original.
    cookie: u32,
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]

/// Specifies the identifier of the Automation type library the containing
/// [`ReferenceControl`]'s twiddled type library was generated from.
#[derive(Debug)]
pub struct ReferenceOriginal {
    /// (Optional) Name entry
    name: Option<String>,
    libid_original: String,
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]

/// Specifies a reference to an Automation type library.
#[derive(Debug)]
pub struct ReferenceRegistered {
    name: Option<String>,
    libid: String,
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]

/// Specifies a reference to an external VBA project.
#[derive(Debug)]
pub struct ReferenceProject {
    name: Option<String>,
    libid_absolute: String,
    libid_relative: String,
    major_version: u32,
    minor_version: u16,
}

/// Specifies a reference to an Automation type library or VBA project.
#[derive(Debug)]
pub enum Reference {
    /// The `Reference` is a [`ReferenceControl`].
    Control(ReferenceControl),
    /// The `Reference` is a [`ReferenceOriginal`].
    Original(ReferenceOriginal),
    /// The `Reference` is a [`ReferenceRegistered`].
    Registered(ReferenceRegistered),
    /// The `Reference` is a [`ReferenceProject`].
    Project(ReferenceProject),
}

// TODO: Remove exemption once the implementation is complete.
#[allow(dead_code)]

/// Specifies version-independent information for the VBA project.
#[derive(Debug)]
pub struct Information {
    /// Specifies the platform for which the VBA project is created.
    pub sys_kind: SysKind,
    lcid: u32,
    lcid_invoke: u32,
    /// Specifies the code page for the VBA project.
    ///
    pub code_page: u16,
    name: String,
    doc_string: String,
    help_file_1: String,
    help_context: u32,
    lib_flags: u32,
    version_major: u32,
    version_minor: u16,
    constants: String,
}

/// Specifies the containing module's type.
#[derive(Debug)]
pub enum ModuleType {
    /// Specifies a procedural module.
    ///
    /// A procedural module is a collection of subroutines and functions.
    Procedural,
    /// Specifies a document module, class module, or designer module.
    ///
    /// A document module is a type of VBA project item that specifies a module for
    /// embedded macros and programmatic access operations that are associated with a
    /// document.
    ///
    /// A class module is a module that contains the definition for a new object. Each
    /// instance of a class creates a new object, and procedures that are defined in the
    /// module become properties and methods of the object.
    ///
    /// A designer module is a VBA module that extends the methods and properties of an
    /// ActiveX control that has been registered with the project.
    ///
    /// The file format specification doesn't distinguish between these three module
    /// types and encodes them using a single umbrella type ID.
    DocClsDesigner,
}

/// Specifies data for a module.
#[derive(Debug)]
pub struct Module {
    /// Specifies a VBA identifier as the name of the containing `Module`.
    pub name: String,
    /// Specifies the stream name in the VBA storage corresponding to the containing
    /// `Module`.
    pub stream_name: String,
    /// Specifies the description for the containing `Module`.
    pub doc_string: String,
    /// Specifies the location of the source code within the stream that corresponds to
    /// the containing `Module`.
    pub text_offset: usize,
    /// Specifies the Help topic identifier for the containing `Module`.
    pub help_context: u32,
    /// Specifies whether the containing `Module` is a procedural module, document
    /// module, class module, or designer module.
    pub module_type: ModuleType,
    /// Specifies that the containing `Module` is read-only.
    pub read_only: bool,
    /// Specifies that the containing `Module` is only usable from within the current VBA
    /// project.
    pub private: bool,
}

impl Project {
    /// Returns a stream's decompressed data.
    ///
    /// This function reads a stream referenced by `stream_path` and passes the data
    /// starting at `offset` into the RLE decompressor.
    ///
    /// The primary use case for this function is to extract source code from VBA
    /// [`Module`]s. The respective `offset` is reported by [`Module::text_offset`].
    ///
    /// This is a low-level function that is useful for very specific use cases only.
    /// Client code that needs to read source code should use [`Project::module_source`]
    /// or [`Project::module_source_raw`] instead.
    // TODO: Code example
    pub fn decompress_stream_from<P>(&self, stream_path: P, offset: usize) -> Result<Vec<u8>>
    where
        P: AsRef<Path>,
    {
        let data = self.read_stream(stream_path)?;
        let data = parser::decompress(&data[offset..])
            .map_err(|_| Error::Decompressor)?
            .1;
        Ok(data)
    }

    // TODO: This should probably live someplace else. It exposes information internal to
    //       the CFB implementation, that's not *immediately* useful or related to this
    //       library's primary responsibility.

    /// Returns a list of entries (storages and streams) in the raw binary data. Each
    /// entry is represented as a tuple of two `String`s, where the first element
    /// contains the entry's name and the second element the entry's path inside the
    /// CFB.
    ///
    /// The raw binary data is encoded as a [Compound File Binary][MS-CFB]
    ///
    /// [MS-CFB]: https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b
    pub fn list(&self) -> Result<Vec<(String, String)>> {
        let mut result = Vec::new();
        for entry in self
            .container
            .borrow()
            .walk_storage("/")
            .map_err(Error::Cfb)?
        {
            result.push((
                entry.name().to_owned(),
                entry.path().to_str().unwrap_or_default().to_owned(),
            ));
        }
        Ok(result)
    }

    /// Returns a module's source code.
    ///
    /// Similar to [`Project::module_source_raw`] this function returns the source code
    /// of a project's module. After the raw source code has been decoded it is then
    /// converted to a `String` using the project's code page.
    pub fn module_source(&self, name: &str) -> Result<String> {
        let source_raw = self.module_source_raw(name)?;
        let source = cp_to_string(&source_raw, self.information.code_page);

        Ok(source)
    }

    /// Returns the raw source code from a module.
    ///
    /// The result contains a module's source code as is. No character encoding conversion
    /// is done. The data is encoded using the project's code page available through
    /// [`Information::code_page`].
    pub fn module_source_raw(&self, name: &str) -> Result<Vec<u8>> {
        let module = self
            .modules
            .iter()
            .find(|&module| module.name == name)
            .ok_or_else(|| Error::ModuleNotFound(name.to_owned()))?;

        let path = format!("/VBA\\{}", &module.stream_name);
        let offset = module.text_offset;
        let src_code = self.decompress_stream_from(path, offset)?;

        Ok(src_code)
    }

    /// Returns a stream's contents.
    ///
    /// This is a low-level function operating on the CFB data. The CFB is the storage
    /// container of the raw binary VBA project.
    pub fn read_stream<P>(&self, stream_path: P) -> Result<Vec<u8>>
    where
        P: AsRef<Path>,
    {
        let mut buffer = Vec::new();
        self.container
            .borrow_mut()
            .open_stream(stream_path)
            .map_err(Error::Cfb)?
            .read_to_end(&mut buffer)
            .map_err(Error::Cfb)?;

        Ok(buffer)
    }
}

/// Opens a VBA project.
///
/// This function consumes `raw` and returns a [`Project`] struct on success, populated
/// with data from the parsed binary input.
pub fn open_project(raw: Vec<u8>) -> Result<Project> {
    let cursor = Cursor::new(raw);
    let mut container = CompoundFile::open(cursor).map_err(Error::Cfb)?;

    // Read *dir* stream
    const DIR_STREAM_PATH: &str = r#"/VBA\dir"#;

    let mut buffer = Vec::new();
    container
        .open_stream(DIR_STREAM_PATH)
        .map_err(Error::Cfb)?
        .read_to_end(&mut buffer)
        .map_err(Error::Cfb)?;

    // Decompress stream
    let (remainder, buffer) = parser::decompress(&buffer).map_err(|_| Error::Decompressor)?;
    debug_assert!(remainder.is_empty());

    // Parse binary data
    let (remainder, information) =
        parser::parse_project_information(&buffer).map_err(|_| Error::Parser)?;
    debug_assert_eq!(remainder.len(), 0, "Stream not fully consumed");

    Ok(Project {
        information: information.information,
        references: information.references,
        modules: information.modules,
        container: RefCell::new(container),
    })
}

#[cfg(test)]
mod tests;
