#![forbid(unsafe_code)]
#![warn(rust_2018_idioms)]

use crate::error::Error;

use cfb::CompoundFile;

use std::io::{Cursor, Read};

pub(crate) struct Project {
    // TODO: Figure out how to make this generic (attempts have failed with trait bound violations)
    container: CompoundFile<Cursor<Vec<u8>>>,
}

impl Project {
    pub(crate) fn list(&self) -> Vec<String> {
        let mut result = Vec::new();
        for entry in self.container.walk_storage("/").unwrap() {
            result.push(entry.name().to_owned());
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
}

pub(crate) fn open_project(raw: Vec<u8>) -> Result<Project, Error> {
    let cursor = Cursor::new(raw);
    let container = CompoundFile::open(cursor).map_err(|e| Error::InvalidDocument(e.into()))?;
    let proj = Project { container };

    Ok(proj)
}
