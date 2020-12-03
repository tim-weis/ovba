[![crates.io](https://img.shields.io/crates/v/ovba.svg)](https://crates.io/crates/ovba)
[![docs.rs](https://docs.rs/ovba/badge.svg)](https://docs.rs/ovba)
[![tests](https://github.com/tim-weis/ovba/workflows/tests/badge.svg?event=push)](https://github.com/tim-weis/ovba/actions)

An Office VBA project parser written in 100% safe Rust. This is an implementation of the [\[MS-OVBA\]: Office VBA File Format Structure](https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-ovba/575462ba-bf67-4190-9fac-c275523c75fc) protocol (Revision 9.1, published 2020-02-19).

## Motivation

Binary file format parsers have historically been an attractive target for attackers. A combination of complex code logic with frequently unchecked memory accesses have had them fall victim to remote code execution exploits many times over.

Rust is a reliable ally in addressing many of these security concerns, empowering this crate to deliver a safe parser implementation.

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
    // Read raw project container
    let data = read("vbaProject.bin")?;
    let project = open_project(data)?;
    // Iterate over CFB entries
    for (name, path) in project.list()? {
        println!(r#"Name: "{}"; Path: "{}""#, name, path);
    }

    Ok(())
}
```

Write out all modules' source code:

```rust
use ovba::{open_project, Result};
use std::fs::{read, write};

fn main() -> Result<()> {
    let data = read("vbaProject.bin")?;
    let project = open_project(data)?;

    for module in &project.information()?.modules {
        let path = format!("/VBA\\{}", &module.stream_name);
        let offset = module.text_offset;
        let src_code = project.decompress_stream_from(&path, offset)?;
        write("./out/".to_string() + &module.name, src_code)?;
    }

    Ok(())
}
```

## Backwards compatibility

At this time, both API and implementation are under development. It is expected to see breaking changes before reaching a 1.0 release. With 0.X.Y releases, breaking changes are signified by a bump in the 0.X version number, leaving non-breaking changes to a bump in the Y version number.

This is a preview release. It has been published to allow others to use it, and solicit feedback to help drive future decision.

## Future work

All future work is tracked [here](https://github.com/tim-weis/ovba/issues). Notable issues include:

* [Improve error reporting](https://github.com/tim-weis/issues/10).

If you are missing a feature, found a bug, have a question, or want to provide feedback, make sure to [file an issue](https://github.com/tim-weis/ovba/issues/new).
