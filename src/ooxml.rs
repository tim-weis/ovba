#![forbid(unsafe_code)]
#![warn(rust_2018_idioms)]

//! Contains structures and functions to inspect Office Open XML documents and extract data.

use crate::error::Error;

use sxd_document::parser;
use sxd_xpath::{nodeset::Node, Context, Factory, Value};
use zip::ZipArchive;

use std::{
    fs::read,
    io::{stdin, Cursor, Read},
    path::PathBuf,
};

/// Opaque data type that represents an Office Open XML file.
pub(crate) struct Document {
    data: Vec<u8>,
}

impl Document {
    /// Creates a new instance holding the entire document contents.
    ///
    /// The document is read from a file if `source` is `Some`, otherwise from standard input.
    pub(crate) fn new(source: &Option<PathBuf>) -> Result<Self, Error> {
        match source {
            Some(path_name) => Ok(Self {
                data: read(path_name).map_err(|e| Error::Io(e.into()))?,
            }),
            None => {
                let mut buffer = Vec::<u8>::new();
                stdin()
                    .read_to_end(&mut buffer)
                    .map_err(|e| Error::Io(e.into()))?;
                Ok(Document { data: buffer })
            }
        }
    }

    /// Returns the name of the contained VBA project, if present.
    pub(crate) fn vba_project_name(&self) -> Result<Option<String>, Error> {
        let factory = Factory::new();
        let xpath = factory
            .build(
                "/ns:Types/ns:Override[@ContentType='application/vnd.ms-office.vbaProject']/@PartName",
            )
            .map_err(|e| Error::InvalidDocument(e.into()))?
            .unwrap();

        let mut context = Context::new();
        context.set_namespace(
            "ns",
            "http://schemas.openxmlformats.org/package/2006/content-types",
        );

        let content_types = self.content_types()?;
        let package =
            parser::parse(&content_types).map_err(|e| Error::InvalidDocument(e.into()))?;

        let value = xpath
            .evaluate(&context, package.as_document().root())
            .map_err(|e| Error::InvalidDocument(e.into()))?;
        if let Value::Nodeset(nodeset) = &value {
            if let Some(node) = nodeset.document_order_first() {
                if let Node::Attribute(attribute) = &node {
                    return Ok(Some(attribute.value().trim_start_matches('/').to_owned()));
                }
            }
        }
        Ok(None)
    }

    /// Extracts a part with a given `part_name` from the document.
    pub(crate) fn part(&self, part_name: &str) -> Result<Vec<u8>, Error> {
        let mut cursor = Cursor::new(&self.data);
        let mut archive =
            ZipArchive::new(&mut cursor).map_err(|e| Error::InvalidDocument(e.into()))?;
        let mut part = archive
            .by_name(&part_name)
            .map_err(|e| Error::InvalidDocument(e.into()))?;
        let mut data = Vec::<u8>::new();
        part.read_to_end(&mut data)
            .map_err(|e| Error::InvalidDocument(e.into()))?;
        Ok(data)
    }

    /// Returns the root content types XML document.
    fn content_types(&self) -> Result<String, Error> {
        const CONTENT_TYPES_NAME: &str = "[Content_Types].xml";

        let mut cursor = Cursor::new(&self.data);
        let mut archive =
            ZipArchive::new(&mut cursor).map_err(|e| Error::InvalidDocument(e.into()))?;

        let mut content = archive
            .by_name(CONTENT_TYPES_NAME)
            .map_err(|e| Error::InvalidDocument(e.into()))?;
        let mut xml_text = String::new();
        content
            .read_to_string(&mut xml_text)
            .map_err(|e| Error::InvalidDocument(e.into()))?;
        Ok(xml_text)
    }
}
