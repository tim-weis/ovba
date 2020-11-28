An Office VBA project parser written in 100% safe Rust. This is an implementation of the [\[MS-OVBA\]: Office VBA File Format Structure](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc) protocol (Revision 9.1, published 2020-02-19).

## Motivation

Binary file format parsers have historically been an attractive target for attackers. A combination of complex code logic with frequently unchecked memory accesses have produced uncounted successful remote code execution vulnerability exploits.

Rust is a perfect tool in addressing these security concerns, empowering this crate to deliver a safe parser implementation.

## Features

This library provides read-only access to VBA projects' metadata and source code. Notable features include:

* Extract source code.
* Inspect metadata, like contained modules, references, etc.

This library does not provide a way to extract the raw binary VBA project data from an Office document. This is the responsibility of client code. The companion [ovba-cli](https://github.com/tim-weis/ovba-cli) tool illustrates how this can be done.

## Usage

List all CFB entries contained in a VBA project:

```rust
use ovba::{open_project, Result};
use std::fs::read;

fn main() -> Result<()> {
    // Read raw data
    let data = read("vbaProject.bin")?;
    // Open project
    let project = open_project(data)?;
    // Iterate over CFB entries
    for (name, path) in project.list()? {
        println!(r#"Name: "{}"; Path: "{}""#, name, path);
    }

    Ok(())
}
```

## Backwards compatibility

This is a preview release. There will be breaking changes before reaching a 1.0 release. This crate has been published to allow others to use it, and solicit feedback to help drive decisions on the future direction.

## Future work

All future work is tracked [here](https://github.com/tim-weis/ovba/issues). Notable future work includes:

* [Streamline API](https://github.com/tim-weis/ovba/issues/8). This is intended to remove some noise and redundancy from the API surface.

If you are missing a feature, found a bug, have a question, or want to provide feedback, make sure to [file an issue](https://github.com/tim-weis/ovba/issues/new).
