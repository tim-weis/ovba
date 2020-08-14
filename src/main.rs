mod error;
mod oox;

use error::Error;
use oox::Document;

use clap::Clap;

use std::{
    fs::write,
    io::{stdout, Write},
    path::PathBuf,
};

#[derive(Clap, Debug)]
#[clap(author, about, version)]
struct Opts {
    /// Input file. Reads from STDIN if omitted.
    #[clap(short, long, parse(from_os_str))]
    input: Option<PathBuf>,
    #[clap(subcommand)]
    subcmd: SubCommand,
}

#[derive(Clap, Debug)]
enum SubCommand {
    Dump(Dump),
}

/// Dump binary VBA project file
#[derive(Clap, Debug)]
struct Dump {
    /// Output file. Writes to STDOUT if omitted.
    #[clap(short, long, parse(from_os_str))]
    output: Option<PathBuf>,
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
            let doc = Document::new(&opts.input)?;
            let part_name = doc.vba_project_name()?;
            match &part_name {
                Some(part_name) => {
                    let data = doc.part(part_name)?;
                    write_output(&dump_opts.output, &data)?;
                }
                None => eprintln!("Document doesn't contain a VBA project."),
            }
        }
    }

    Ok(())
}
