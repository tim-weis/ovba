#![forbid(unsafe_code)]
#![warn(rust_2018_idioms)]

use crate::error::Error;

use cfb::CompoundFile;

use std::io::{Cursor, Read};

pub(crate) struct Project {
    // TODO: Figure out how to make this generic (attempts have failed with trait bound violations)
    container: CompoundFile<Cursor<Vec<u8>>>,
}

#[derive(Debug)]
pub(crate) enum SysKind {
    Win16,
    Win32,
    MacOs,
    Win64,
}

/// Version Independent Project Information
#[derive(Debug)]
pub(crate) struct ProjectInformation {
    information: Information,
    references: Vec<Reference>,
}

#[derive(Debug)]
pub(crate) struct ReferenceControl {
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
pub(crate) struct ReferenceOriginal {
    /// (Optional) Name and NameUnicode entries
    name: Option<(String, String)>,
    libid_original: String,
}

#[derive(Debug)]
pub(crate) struct ReferenceRegistered {
    name: Option<(String, String)>,
    libid: String,
}

#[derive(Debug)]
pub(crate) struct ReferenceProject {
    name: Option<(String, String)>,
    libid_absolute: String,
    libid_relative: String,
    major_version: u32,
    minor_version: u16,
}

#[derive(Debug)]
pub(crate) enum Reference {
    Control(ReferenceControl),
    Original(ReferenceOriginal),
    Registered(ReferenceRegistered),
    Project(ReferenceProject),
}

#[derive(Debug)]
pub(crate) struct Information {
    sys_kind: SysKind,
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

impl Project {
    pub(crate) fn list(&self) -> Vec<(String, String)> {
        let mut result = Vec::new();
        for entry in self.container.walk_storage("/").unwrap() {
            result.push((
                entry.name().to_owned(),
                entry.path().to_str().unwrap_or_default().to_owned(),
            ));
        }
        result
    }

    pub(crate) fn read_stream(&mut self, stream_name: &str) -> Result<Vec<u8>, Error> {
        let mut stream = self
            .container
            .open_stream(stream_name)
            .map_err(|e| Error::Io(e.into()))?;
        let mut buffer = Vec::new();
        stream
            .read_to_end(&mut buffer)
            .map_err(|e| Error::Io(e.into()))?;

        Ok(buffer)
    }

    /// Returns version independent project information.
    pub(crate) fn information(&mut self) -> Result<ProjectInformation, Error> {
        const DIR_STREAM_PATH: &str = r#"/VBA\dir"#;

        // Read *dir* stream
        let mut stream = self
            .container
            .open_stream(DIR_STREAM_PATH)
            .map_err(|e| Error::Io(e.into()))?;
        let mut buffer = Vec::new();
        stream
            .read_to_end(&mut buffer)
            .map_err(|e| Error::Io(e.into()))?;

        // Decompress stream
        let (remainder, buffer) = parser::decompress(&buffer).map_err(|_| Error::Unknown)?;
        debug_assert!(remainder.is_empty());
        println!("Buffer length: {}", buffer.len());

        // Parse binary data
        let (remainder, information) =
            parser::parse_project_information(&buffer).map_err(|_| Error::Unknown)?;

        // Return structured information
        Ok(information)
    }
}

pub(crate) fn open_project(raw: Vec<u8>) -> Result<Project, Error> {
    let cursor = Cursor::new(raw);
    let container = CompoundFile::open(cursor).map_err(|e| Error::InvalidDocument(e.into()))?;
    let proj = Project { container };

    Ok(proj)
}

#[doc(hidden)]
/// Internal parser implementations
mod parser {
    use super::{
        Information, ProjectInformation, Reference, ReferenceControl, ReferenceOriginal,
        ReferenceProject, ReferenceRegistered, SysKind,
    };
    use codepage::to_encoding;
    use encoding_rs::{CoderResult, UTF_16LE};
    use nom::{
        bytes::complete::{tag, take},
        error::{ErrorKind, ParseError},
        multi::length_data,
        number::complete::{le_u16, le_u32, le_u8},
        Err::Error,
        IResult,
    };

    // TODO: Make this error private by translating to a crate-level error type
    //       at the public parser interface.
    #[derive(Debug, PartialEq)]
    pub(crate) enum FormatError<I> {
        UnexpectedValue,
        Nom(I, ErrorKind),
    }

    impl<I> ParseError<I> for FormatError<I> {
        fn from_error_kind(input: I, kind: ErrorKind) -> Self {
            FormatError::Nom(input, kind)
        }
        fn append(_: I, _: ErrorKind, other: Self) -> Self {
            other
        }
    }

    fn uncompressed_chunk_parser(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        Ok((&[], i.to_vec()))
    }

    fn compressed_chunk_parser(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        // Initialize output storage; Chunks are at most 4096 decompressed bytes
        let mut result = Vec::<u8>::with_capacity(4096);
        // Loop until `i` is depleted
        let mut input = i;
        while !input.is_empty() {
            // Read FlagByte
            let (i, flag_byte) = le_u8(input)?;
            input = i;
            // Loop over bits
            for flag_bit_index in 0..=7 {
                // Return, if we have reached the end of this chunk
                if input.is_empty() {
                    return Ok((input, result));
                }
                // Determine token type (0b0 == LiteralToken; 0b1 == CopyToken)
                let is_copy_token = (flag_byte & (1 << flag_bit_index)) != 0;
                // Delegate work based on TokenType
                if is_copy_token {
                    // TODO: Move the CopyToken decoder into its own, dedicated parser.
                    let (i, copy_token_raw) = le_u16(input)?;
                    input = i;
                    // Calculate length/offset masks
                    let diff = result.len();
                    let mut bit_count = 4_usize;
                    while 1 << bit_count < diff {
                        bit_count += 1;
                    }
                    let length_mask = 0xffff_u16 >> bit_count;
                    let offset_mask = !length_mask;
                    // Calculate length/offset
                    let length = ((copy_token_raw & length_mask) + 3) as usize;
                    let offset =
                        (((copy_token_raw & offset_mask) >> (16 - bit_count)) + 1) as usize;
                    // Copy `length` bytes starting at index `offset`
                    for index in result.len() - offset..result.len() - offset + length {
                        result.push(result[index]);
                    }
                } else {
                    // LiteralToken -> Copy token from input stream
                    let (i, byte) = le_u8(input)?;
                    input = i;
                    result.push(byte);
                }
            }
        }

        Ok((input, result))
    }

    fn chunk_parser(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        // CompressedChunkHeader (12 bits: size minus 3; 3 bits: 0b110; 1 bit: flag)
        // Delegate to specific parser (compressed/uncompressed) depending on the `flag`
        let (i, header_raw) = le_u16(i)?;
        // Check header magic (0b110) in bit positions 12..=14
        if (header_raw >> 12) & 0b111 != 0b011 {
            return Err(Error(FormatError::UnexpectedValue));
        }
        // Extract compressed/uncompressed flag
        let flag = ((header_raw >> 15) & 0b1) != 0;
        // Extract length
        let length = (header_raw & 0xfff) as usize + 1;

        let i = &i[..length];
        if flag {
            Ok(compressed_chunk_parser(i)?)
        } else {
            Ok(uncompressed_chunk_parser(i)?)
        }
    }

    /// Decompress a CompressedContainer.
    pub(crate) fn decompress(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        const COMPRESSED_CONTAINER_SIGNATURE: &[u8] = &[0x01];
        let (i, _) = tag(COMPRESSED_CONTAINER_SIGNATURE)(i)?;

        // This is the main `Chunk` parser:
        // * It parses 1 or more chunks, returning a `Vec<u8>` with decoded content.
        // * It appends the contents of the most recent `Chunk` to the existing decoded stream.
        // * If all data has been consumed, return an `Ok()` value.
        nom::combinator::all_consuming(nom::multi::fold_many1(
            chunk_parser,
            Vec::new(),
            |mut acc: Vec<_>, data| {
                acc.extend(data);
                acc
            },
        ))(i)
    }

    // -------------------------------------------------------------------------
    // -------------------------------------------------------------------------

    // Several size fields in the binary format have fixed values.
    const U32_FIXED_SIZE_4: &[u8] = &[0x04, 0x00, 0x00, 0x00];
    const U32_FIXED_SIZE_2: &[u8] = &[0x02, 0x00, 0x00, 0x00];

    fn parse_syskind(i: &[u8]) -> IResult<&[u8], SysKind, FormatError<&[u8]>> {
        const SYS_KIND_SIGNATURE: &[u8] = &[0x01, 0x00];
        let (i, _) = tag(SYS_KIND_SIGNATURE)(i)?;
        let (i, _) = tag(U32_FIXED_SIZE_4)(i)?;

        let (i, value) = le_u32(i)?;
        match value {
            0x0000_0000 => Ok((i, SysKind::Win16)),
            0x0000_0001 => Ok((i, SysKind::Win32)),
            0x0000_0002 => Ok((i, SysKind::MacOs)),
            0x0000_0003 => Ok((i, SysKind::Win64)),
            _ => Err(Error(FormatError::UnexpectedValue)),
        }
    }

    fn parse_lcid(i: &[u8]) -> IResult<&[u8], u32, FormatError<&[u8]>> {
        const LCID_SIGNATURE: &[u8] = &[0x02, 0x00];
        let (i, _) = tag(LCID_SIGNATURE)(i)?;
        let (i, _) = tag(U32_FIXED_SIZE_4)(i)?;
        let (i, value) = le_u32(i)?;
        Ok((i, value))
    }

    fn parse_lcid_invoke(i: &[u8]) -> IResult<&[u8], u32, FormatError<&[u8]>> {
        const LCID_INVOKE_SIGNATURE: &[u8] = &[0x14, 0x00];
        let (i, _) = tag(LCID_INVOKE_SIGNATURE)(i)?;
        let (i, _) = tag(U32_FIXED_SIZE_4)(i)?;
        let (i, value) = le_u32(i)?;
        Ok((i, value))
    }

    fn parse_code_page(i: &[u8]) -> IResult<&[u8], u16, FormatError<&[u8]>> {
        const CODE_PAGE_SIGNATURE: &[u8] = &[0x03, 0x00];
        let (i, _) = tag(CODE_PAGE_SIGNATURE)(i)?;
        let (i, _) = tag(U32_FIXED_SIZE_2)(i)?;
        let (i, value) = le_u16(i)?;
        Ok((i, value))
    }

    fn parse_name(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        const NAME_SIGNATURE: &[u8] = &[0x04, 0x00];
        let (i, _) = tag(NAME_SIGNATURE)(i)?;
        let (i, value) = length_data(le_u32)(i)?;
        Ok((i, value.to_vec()))
    }

    fn parse_doc_string(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        const DOC_STRING_SIGNATURE: &[u8] = &[0x05, 0x00];
        let (i, _) = tag(DOC_STRING_SIGNATURE)(i)?;
        let (i, value) = length_data(le_u32)(i)?;
        Ok((i, value.to_vec()))
    }

    fn parse_doc_string_unicode(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        const DOC_STRING_UNICODE_SIGNATURE: &[u8] = &[0x40, 0x00];
        let (i, _) = tag(DOC_STRING_UNICODE_SIGNATURE)(i)?;
        let (i, value) = length_data(le_u32)(i)?;
        // `value` represents a sequence of UTF-16 code units. If its length is uneven, the
        // input is malformed.
        if (value.len() & 1_usize) != 0 {
            Err(Error(FormatError::UnexpectedValue))
        } else {
            Ok((i, value.to_vec()))
        }
    }

    fn parse_help_file_1(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        const HELP_FILE_1_SIGNATURE: &[u8] = &[0x06, 0x00];
        let (i, _) = tag(HELP_FILE_1_SIGNATURE)(i)?;
        let (i, value) = length_data(le_u32)(i)?;
        Ok((i, value.to_vec()))
    }

    fn parse_help_file_2(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        const HELP_FILE_2_SIGNATURE: &[u8] = &[0x3d, 0x00];
        let (i, _) = tag(HELP_FILE_2_SIGNATURE)(i)?;
        let (i, value) = length_data(le_u32)(i)?;
        Ok((i, value.to_vec()))
    }

    fn parse_help_context(i: &[u8]) -> IResult<&[u8], u32, FormatError<&[u8]>> {
        const HELP_CONTEXT_SIGNATURE: &[u8] = &[0x07, 0x00];
        let (i, _) = tag(HELP_CONTEXT_SIGNATURE)(i)?;
        let (i, _) = tag(U32_FIXED_SIZE_4)(i)?;
        Ok(le_u32(i)?)
    }

    fn parse_lib_flags(i: &[u8]) -> IResult<&[u8], u32, FormatError<&[u8]>> {
        const LIB_FLAGS_SIGNATURE: &[u8] = &[0x08, 0x00];
        let (i, _) = tag(LIB_FLAGS_SIGNATURE)(i)?;
        let (i, _) = tag(U32_FIXED_SIZE_4)(i)?;
        Ok(le_u32(i)?)
    }

    fn parse_version(i: &[u8]) -> IResult<&[u8], (u32, u16), FormatError<&[u8]>> {
        const VERSION_SIGNATURE: &[u8] = &[0x09, 0x00];
        let (i, _) = tag(VERSION_SIGNATURE)(i)?;
        let (i, _) = tag(U32_FIXED_SIZE_4)(i)?;
        let (i, version_major) = le_u32(i)?;
        let (i, version_minor) = le_u16(i)?;
        Ok((i, (version_major, version_minor)))
    }

    fn parse_constants(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        const CONSTANTS_SIGNATURE: &[u8] = &[0x0c, 0x00];
        let (i, _) = tag(CONSTANTS_SIGNATURE)(i)?;
        let (i, value) = length_data(le_u32)(i)?;
        Ok((i, value.to_vec()))
    }

    fn parse_constants_unicode(i: &[u8]) -> IResult<&[u8], Vec<u8>, FormatError<&[u8]>> {
        const CONSTANTS_UNICODE_SIGNATURE: &[u8] = &[0x3c, 0x00];
        let (i, _) = tag(CONSTANTS_UNICODE_SIGNATURE)(i)?;
        let (i, value) = length_data(le_u32)(i)?;
        Ok((i, value.to_vec()))
    }

    // -------------------------------------------------------------------------
    // -------------------------------------------------------------------------

    #[allow(clippy::type_complexity)]
    fn parse_reference_name(
        i: &[u8],
        code_page: u16,
    ) -> IResult<&[u8], Option<(String, String)>, FormatError<&[u8]>> {
        const NAME_SIGNATURE: u16 = 0x0016_u16;
        let (i_next, id) = le_u16(i)?;

        // If this is not a REFERENCENAME Record, make sure to return to original slice
        if id != NAME_SIGNATURE {
            return Ok((i, None));
        }

        // Update remaining slice since we have a REFERENCENAME Record
        let i = i_next;
        let (i, value) = length_data(le_u32)(i)?;
        let name = cp_to_string(value, code_page);

        const NAME_UNICODE_SIGNATURE: &[u8] = &[0x3e, 0x00];
        let (i, _) = tag(NAME_UNICODE_SIGNATURE)(i)?;
        let (i, value) = length_data(le_u32)(i)?;
        let name_unicode = utf16_to_string(value);

        Ok((i, Some((name, name_unicode))))
    }

    fn parse_reference_original(
        i: &[u8],
        code_page: u16,
    ) -> IResult<&[u8], String, FormatError<&[u8]>> {
        const ORIGINAL_SIGNATURE: &[u8] = &[0x33, 0x00];
        let (i, _) = tag(ORIGINAL_SIGNATURE)(i)?;
        let (i, libid_original) = length_data(le_u32)(i)?;
        let libid_original = cp_to_string(libid_original, code_page);
        Ok((i, libid_original))
    }

    fn parse_reference_control(
        i: &[u8],
        code_page: u16,
    ) -> IResult<&[u8], ReferenceControl, FormatError<&[u8]>> {
        // REFERENCEORIGINAL Record is optional here
        let (_, id) = le_u16(i)?;
        let (i, libid_original) = match id {
            0x0033_u16 => {
                let (i, libid_original) = parse_reference_original(i, code_page)?;
                (i, Some(libid_original))
            }
            _ => (i, None),
        };

        const CONTROL_SIGNATURE: &[u8] = &[0x2f, 0x00];
        let (i, _) = tag(CONTROL_SIGNATURE)(i)?;
        let (i, _combined_size) = le_u32(i)?;
        let (i, libid_twiddled) = length_data(le_u32)(i)?;
        let libid_twiddled = cp_to_string(libid_twiddled, code_page);

        const RESERVED_1: &[u8] = &[0x00, 0x00, 0x00, 0x00];
        const RESERVED_2: &[u8] = &[0x00, 0x00];
        let (i, _) = tag(RESERVED_1)(i)?;
        let (i, _) = tag(RESERVED_2)(i)?;

        let (i, name_extended) = parse_reference_name(i, code_page)?;

        const RESERVED_3: &[u8] = &[0x30, 0x00];
        let (i, _) = tag(RESERVED_3)(i)?;
        let (i, _combined_size) = le_u32(i)?;
        let (i, libid_extended) = length_data(le_u32)(i)?;
        let libid_extended = cp_to_string(libid_extended, code_page);

        const RESERVED_4: &[u8] = &[0x00, 0x00, 0x00, 0x00];
        const RESERVED_5: &[u8] = &[0x00, 0x00];
        let (i, _) = tag(RESERVED_4)(i)?;
        let (i, _) = tag(RESERVED_5)(i)?;

        let (i, guid) = take(16_usize)(i)?;
        let guid = guid.to_vec();

        let (i, cookie) = le_u32(i)?;

        Ok((
            i,
            ReferenceControl {
                name: None,
                libid_original,
                libid_twiddled,
                name_extended,
                libid_extended,
                guid,
                cookie,
            },
        ))
    }

    fn parse_reference_registered(
        i: &[u8],
        code_page: u16,
    ) -> IResult<&[u8], ReferenceRegistered, FormatError<&[u8]>> {
        const REGISTERED_SIGNATURE: &[u8] = &[0x0d, 0x00];
        let (i, _) = tag(REGISTERED_SIGNATURE)(i)?;
        let (i, _combined_size) = le_u32(i)?;
        let (i, libid) = length_data(le_u32)(i)?;
        let libid = cp_to_string(libid, code_page);

        const RESERVED_1: &[u8] = &[0x00, 0x00, 0x00, 0x00];
        const RESERVED_2: &[u8] = &[0x00, 0x00];
        let (i, _) = tag(RESERVED_1)(i)?;
        let (i, _) = tag(RESERVED_2)(i)?;

        Ok((i, ReferenceRegistered { name: None, libid }))
    }

    fn parse_reference_project(
        i: &[u8],
        code_page: u16,
    ) -> IResult<&[u8], ReferenceProject, FormatError<&[u8]>> {
        let (i, _) = tag(&[0x0e, 0x00])(i)?;
        let (i, _combined_size) = le_u32(i)?;

        let (i, libid_absolute) = length_data(le_u32)(i)?;
        let libid_absolute = cp_to_string(libid_absolute, code_page);

        let (i, libid_relative) = length_data(le_u32)(i)?;
        let libid_relative = cp_to_string(libid_relative, code_page);

        let (i, major_version) = le_u32(i)?;
        let (i, minor_version) = le_u16(i)?;

        Ok((
            i,
            ReferenceProject {
                name: None,
                libid_absolute,
                libid_relative,
                major_version,
                minor_version,
            },
        ))
    }

    /// Parses a single REFERENCE Record.
    ///
    /// There are several tricky bits to this:
    /// * The first entry (NameRecord) is optional.
    /// * The REFERENCE Record can be one of 4 variants.
    /// * The length is implied through a terminator (0x000F) that starts a PROJECTMODULES Record.
    ///
    /// Returns `Some(reference)` if a variant was found, `None` if the end of the array was
    /// reached, or an error.
    fn parse_reference(
        i: &[u8],
        code_page: u16,
    ) -> IResult<&[u8], Option<Reference>, FormatError<&[u8]>> {
        let (i, name) = parse_reference_name(i, code_page)?;
        // Determine REFERENCE Record variant (or end of array)
        let (_, id) = le_u16(i)?;
        match id {
            0x002f_u16 => {
                let (i, mut value) = parse_reference_control(i, code_page)?;
                value.name = name;
                Ok((i, Some(Reference::Control(value))))
            }
            0x0033_u16 => {
                let (i, libid_original) = parse_reference_original(i, code_page)?;
                let original = ReferenceOriginal {
                    name,
                    libid_original,
                };
                Ok((i, Some(Reference::Original(original))))
            }
            0x000d_u16 => {
                let (i, mut value) = parse_reference_registered(i, code_page)?;
                value.name = name;
                Ok((i, Some(Reference::Registered(value))))
            }
            0x000e_u16 => {
                let (i, mut value) = parse_reference_project(i, code_page)?;
                value.name = name;
                Ok((i, Some(Reference::Project(value))))
            }
            0x000f_u16 => Ok((i, None)),
            _ => Err(Error(FormatError::UnexpectedValue)),
        }
    }

    fn parse_references(
        i: &[u8],
        code_page: u16,
    ) -> IResult<&[u8], Vec<Reference>, FormatError<&[u8]>> {
        let mut result = Vec::new();
        let mut i = i;
        loop {
            // TODO: Verify whether `i` stays alive at the end of the loop.
            let (remainder, value) = parse_reference(i, code_page)?;
            i = remainder;
            if let Some(reference) = value {
                result.push(reference);
            } else {
                return Ok((i, result));
            }
        }
    }

    // -------------------------------------------------------------------------
    // -------------------------------------------------------------------------

    // -------------------------------------------------------------------------
    // -------------------------------------------------------------------------

    /// *dir* stream parser.
    pub(crate) fn parse_project_information(
        i: &[u8],
    ) -> IResult<&[u8], ProjectInformation, FormatError<&[u8]>> {
        let (i, sys_kind) = parse_syskind(i)?;
        let (i, lcid) = parse_lcid(i)?;
        let (i, lcid_invoke) = parse_lcid_invoke(i)?;
        let (i, code_page) = parse_code_page(i)?;

        let (i, name) = parse_name(i)?;
        let name = cp_to_string(&name, code_page);

        let (i, doc_string) = parse_doc_string(i)?;
        let doc_string = cp_to_string(&doc_string, code_page);

        let (i, doc_string_unicode) = parse_doc_string_unicode(i)?;
        let doc_string_unicode = utf16_to_string(&doc_string_unicode);

        let (i, help_file_1) = parse_help_file_1(i)?;
        let help_file_1 = cp_to_string(&help_file_1, code_page);

        let (i, help_file_2) = parse_help_file_2(i)?;
        let help_file_2 = cp_to_string(&help_file_2, code_page);

        let (i, help_context) = parse_help_context(i)?;
        let (i, lib_flags) = parse_lib_flags(i)?;
        let (i, (version_major, version_minor)) = parse_version(i)?;

        let (i, constants) = parse_constants(i)?;
        let constants = cp_to_string(&constants, code_page);

        let (i, constants_unicode) = parse_constants_unicode(i)?;
        let constants_unicode = utf16_to_string(&constants_unicode);

        let (i, references) = parse_references(i, code_page)?;

        Ok((
            i,
            ProjectInformation {
                information: Information {
                    sys_kind,
                    lcid,
                    lcid_invoke,
                    code_page,
                    name,
                    doc_string,
                    doc_string_unicode,
                    help_file_1,
                    help_file_2,
                    help_context,
                    lib_flags,
                    version_major,
                    version_minor,
                    constants,
                    constants_unicode,
                },
                references,
            },
        ))
    }

    // -------------------------------------------------------------------------
    // -------------------------------------------------------------------------

    /// # Panics
    ///
    /// This function panics, if:
    /// * the passed in code page cannot be mapped to an encoding.
    /// * the maximum length of the output would overflow a `usize`.
    /// * part of the input could not be decoded into the allocated output `String`.
    ///
    /// This is a temporary solution that allows me to postpone implementing error reporting
    /// to a later time, when the set of expected errors and the overall error handling strategy
    /// are better understood.
    fn cp_to_string(data: &[u8], code_page: u16) -> String {
        let encoding = to_encoding(code_page).expect("Failed to map code page to an encoding.");
        let mut decoder = encoding.new_decoder_without_bom_handling();
        // The following returns `None` on overflow. That case is only expected with malformed document
        // input, so let's just panic in this case.
        let max_length = decoder.max_utf8_buffer_length(data.len()).unwrap();
        let mut result = String::with_capacity(max_length);
        let (decoder_result, _, _) = decoder.decode_to_string(data, &mut result, true);
        assert_eq!(
            decoder_result,
            CoderResult::InputEmpty,
            "Failed to decode full MBCS sequence."
        );

        result
    }

    fn utf16_to_string(data: &[u8]) -> String {
        let mut decoder = UTF_16LE.new_decoder_without_bom_handling();
        let max_length = decoder.max_utf8_buffer_length(data.len()).unwrap();
        let mut result = String::with_capacity(max_length);
        let (decoder_result, _, _) = decoder.decode_to_string(data, &mut result, true);
        assert_eq!(
            decoder_result,
            CoderResult::InputEmpty,
            "Failed to decode full UTF-16 sequence."
        );

        result
    }
}
