use clap::Clap;
use std::{
    fs::{read, write},
    io::{stdin, stdout, Cursor, Seek, Write},
    io::{Error, Read},
    path::PathBuf,
};
use sxd_document::parser;
use sxd_xpath::{nodeset::Node, Context, Factory, Value};
use zip::ZipArchive;

/// Inspect and extract VBA projects from Office Open XML documents.
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

fn read_input(from: &Option<PathBuf>) -> Result<Vec<u8>, Error> {
    match from {
        Some(path_name) => read(path_name),
        None => {
            let mut buffer = Vec::<u8>::new();
            stdin().read_to_end(&mut buffer)?;
            Ok(buffer)
        }
    }
}

fn get_content_types<T: Read + Seek>(archive: &mut ZipArchive<T>) -> Result<String, Error> {
    let mut content = archive.by_name("[Content_Types].xml")?;
    let mut xml_text = String::new();
    content.read_to_string(&mut xml_text)?;
    Ok(xml_text)
}

fn get_project_name(xml_text: &str) -> Result<Option<String>, Error> {
    // TODO: Map errors
    let package = parser::parse(&xml_text).unwrap();
    let document = package.as_document();

    let factory = Factory::new();
    let xpath = factory
        .build(
            "/ns:Types/ns:Override[@ContentType='application/vnd.ms-office.vbaProject']/@PartName",
        )
        // TODO: Map errors
        .unwrap()
        .unwrap();

    let mut context = Context::new();
    context.set_namespace(
        "ns",
        "http://schemas.openxmlformats.org/package/2006/content-types",
    );

    let value = xpath.evaluate(&context, document.root()).unwrap();
    if let Value::Nodeset(nodeset) = &value {
        if let Some(node) = nodeset.document_order_first() {
            if let Node::Attribute(attribute) = &node {
                return Ok(Some(attribute.value().trim_start_matches('/').to_owned()));
            }
        }
    }
    Ok(None)
}

fn write_output(to: &Option<PathBuf>, data: &[u8]) -> Result<(), Error> {
    match to {
        Some(path_name) => write(path_name, data),
        _ => stdout().write_all(data),
    }
}

fn main() -> Result<(), Error> {
    let opts = Opts::parse();

    // `ZipArchive` operates on `Reader`s, and while re-reading files works, this isn't true for
    // STDIN. So we need to keep the entire document in memory.
    let input = read_input(&opts.input)?;
    let mut cursor = Cursor::new(&input);
    let mut zip_archive = ZipArchive::new(&mut cursor)?;

    match opts.subcmd {
        SubCommand::Dump(dump_opts) => {
            let xml_text = get_content_types(&mut zip_archive)?;
            let part_name = get_project_name(&xml_text)?;

            if let Some(part_name) = part_name {
                let mut zip = ZipArchive::new(&mut cursor)?;
                let mut content = zip.by_name(&part_name)?;
                let mut vba_project = Vec::<u8>::new();
                content.read_to_end(&mut vba_project)?;

                write_output(&dump_opts.output, &vba_project)?;
            }
        }
    }

    Ok(())
}
