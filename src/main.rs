use std::{fs::File, io::Read, io::Write};
use sxd_document::parser;
use sxd_xpath::{nodeset::Node, Context, Factory, Value};
use zip::ZipArchive;

fn main() {
    let input = File::open(r#"c:\Users\Tim\AppData\Roaming\Microsoft\Excel\XLSTART\PERSONAL.XLSB"#)
        .unwrap();
    let mut zip = ZipArchive::new(&input).unwrap();
    for name in zip.file_names() {
        println!("File: {}", name);
    }

    let mut content = zip.by_name("[Content_Types].xml").unwrap();
    let mut xml_text = String::new();
    content.read_to_string(&mut xml_text).unwrap();

    let package = parser::parse(&xml_text).unwrap();
    let document = package.as_document();

    let factory = Factory::new();
    let xpath = factory
        .build(
            "/ns:Types/ns:Override[@ContentType='application/vnd.ms-office.vbaProject']/@PartName",
        )
        .unwrap()
        .unwrap();

    let mut context = Context::new();
    context.set_namespace(
        "ns",
        "http://schemas.openxmlformats.org/package/2006/content-types",
    );

    let mut part_name = None;
    let value = xpath.evaluate(&context, document.root()).unwrap();
    if let Value::Nodeset(nodeset) = &value {
        if let Some(node) = nodeset.document_order_first() {
            if let Node::Attribute(attribute) = &node {
                part_name = Some(attribute.value());
            }
        }
    }

    if let Some(part_name) = part_name {
        let mut zip = ZipArchive::new(&input).unwrap();
        let mut content = zip.by_name(part_name.trim_start_matches('/')).unwrap();
        let mut vba_project = Vec::<u8>::new();
        content.read_to_end(&mut vba_project).unwrap();

        let mut output = File::create("vbaProject.bin").unwrap();
        output.write_all(&vba_project).unwrap();
    }
}
