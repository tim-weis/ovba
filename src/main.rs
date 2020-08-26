#![forbid(unsafe_code)]
#![warn(rust_2018_idioms)]

mod error;
mod ooxml;
mod ovba;

use error::Error;
use ooxml::Document;

use clap::Clap;

use std::{
    fs::write,
    io::{stdout, Write},
    path::PathBuf,
};

#[derive(Clap, Debug)]
#[clap(author, about, version)]
struct Opts {
    #[clap(subcommand)]
    subcmd: SubCommand,
}

#[derive(Clap, Debug)]
enum SubCommand {
    /// Dump contents of the VBA project file
    Dump(DumpArgs),
    /// Display a list of storages and streams
    List(ListArgs),
    /// Display VBA project information
    Info(InfoArgs),
}

#[derive(Clap, Debug)]
struct DumpArgs {
    /// Input file. Reads from STDIN if omitted.
    #[clap(short, long, parse(from_os_str))]
    input: Option<PathBuf>,
    /// Output file. Writes to STDOUT if omitted.
    #[clap(short, long, parse(from_os_str))]
    output: Option<PathBuf>,
    /// Module to output.
    #[clap(short, long)]
    module: Option<String>,
}

#[derive(Clap, Debug)]
struct ListArgs {
    /// Input file. Reads from STDIN if omitted.
    #[clap(short, long, parse(from_os_str))]
    input: Option<PathBuf>,
}

#[derive(Clap, Debug)]
struct InfoArgs {
    /// Input file. Reads from STDIN if omitted.
    #[clap(short, long, parse(from_os_str))]
    input: Option<PathBuf>,
}

fn write_output(to: &Option<PathBuf>, data: &[u8]) -> Result<(), Error> {
    match to {
        Some(path_name) => write(path_name, data).map_err(|e| Error::Io(e.into())),
        _ => stdout().write_all(data).map_err(|e| Error::Io(e.into())),
    }
}

fn main() -> Result<(), Error> {
    let opts = Opts::parse();

    match opts.subcmd {
        SubCommand::Dump(dump_opts) => {
            let doc = Document::new(&dump_opts.input)?;
            let part_name = doc.vba_project_name()?;
            match &part_name {
                Some(part_name) => {
                    let data = doc.part(part_name)?;
                    match dump_opts.module {
                        Some(module_name) => {
                            let mut project = ovba::open_project(data)?;
                            let info = project.information()?;
                            let module_record = info
                                .modules
                                .modules
                                .iter()
                                .find(|module| module.name == module_name);
                            if let Some(module_record) = module_record {
                                let stream_name = format!("/VBA\\{}", module_record.stream_name);
                                let stream_data = project.decompress_stream_from(
                                    &stream_name,
                                    module_record.text_offset as _,
                                )?;
                                write_output(&dump_opts.output, &stream_data)?;
                            }
                        }

                        None => {
                            // Dump full binary project if no module to extract is specified
                            write_output(&dump_opts.output, &data)?;
                        }
                    }
                }
                None => eprintln!("Document doesn't contain a VBA project."),
            }
        }
        SubCommand::List(list_opts) => {
            let doc = Document::new(&list_opts.input)?;
            let part_name = doc.vba_project_name()?;
            if let Some(part_name) = part_name {
                let part = doc.part(&part_name)?;
                let project = ovba::open_project(part)?;
                let entries = project.list();
                for entry in &entries {
                    println!("Entry: {} ({})", entry.0, entry.1);
                }
            }
        }
        SubCommand::Info(info_opts) => {
            // TODO: Implementation
            let doc = Document::new(&info_opts.input)?;
            let part_name = doc.vba_project_name()?;
            if let Some(part_name) = part_name {
                let part = doc.part(&part_name)?;
                let mut project = ovba::open_project(part)?;
                let info = project.information()?;
                println!("Version Independent Project Information:\n{:#?}", info);
            }
        }
    }

    Ok(())
}
