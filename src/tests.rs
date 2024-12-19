use super::parser::{decompress, parse_project_information};

#[test]
fn copy_token_decoder() {
    // CopyTokens store offset and length information in a single 16-bit value. The bit
    // range used for either one changes with the current position in the output stream.
    //
    // This test verifies the implementation by running the decompressor against input
    // where the first CopyToken is encountered at positions 31, 32, and 33 in the
    // output stream.
    //
    // The input was generated using Excel 2013, by adding non-repeating character
    // sequences to a module, until the full size reached the desired length.
    // The prefix `Attribute VB_Name = "a"\r\n` gets added by Excel for every code
    // module, where `"a"` is the respective module name.
    // The resulting *vbaProject.bin* files were then extracted from the Excel documents,
    // opened in a hex editor, and the byte sequences corresponding to the respective
    // compressed containers copied here.

    // CompressedContainer with first CopyToken at position 31:
    // 01 27 B0 00 41 74 74 72 69 62 75 74 00 65 20 56 42 5F 4E 61 6D 00 65 20 3D 20 22 61 22 0D 80 0A 61 62 63 64 65 66 06 F0 00 0D 0A
    const CONTAINER_1: &[u8] = b"\x01\x27\xB0\x00\x41\x74\x74\x72\x69\x62\x75\x74\x00\x65\x20\x56\x42\x5F\x4E\x61\x6D\x00\x65\x20\x3D\x20\x22\x61\x22\x0D\x80\x0A\x61\x62\x63\x64\x65\x66\x06\xF0\x00\x0D\x0A";
    const CONTENTS_1: &[u8] = b"Attribute VB_Name = \"a\"\x0D\x0AabcdefAttribute\x0D\x0A";
    let contents = decompress(CONTAINER_1).unwrap().1;
    assert_eq!(contents, CONTENTS_1);

    // CompressedContainer with first CopyToken at position 32:
    // 01 28 B0 00 41 74 74 72 69 62 75 74 00 65 20 56 42 5F 4E 61 6D 00 65 20 3D 20 22 61 22 0D 00 0A 61 62 63 64 65 66 67 01 06 F8 0D 0A
    const CONTAINER_2: &[u8] = b"\x01\x28\xB0\x00\x41\x74\x74\x72\x69\x62\x75\x74\x00\x65\x20\x56\x42\x5F\x4E\x61\x6D\x00\x65\x20\x3D\x20\x22\x61\x22\x0D\x00\x0A\x61\x62\x63\x64\x65\x66\x67\x01\x06\xF8\x0D\x0A";
    const CONTENTS_2: &[u8] = b"Attribute VB_Name = \"a\"\x0D\x0AabcdefgAttribute\x0D\x0A";
    let contents = decompress(CONTAINER_2).unwrap().1;
    assert_eq!(contents, CONTENTS_2);

    // CompressedContainer with first CopyToken at position 33:
    // 01 29 B0 00 41 74 74 72 69 62 75 74 00 65 20 56 42 5F 4E 61 6D 00 65 20 3D 20 22 61 22 0D 00 0A 61 62 63 64 65 66 67 02 68 06 80 0D 0A
    const CONTAINER_3: &[u8] = b"\x01\x29\xB0\x00\x41\x74\x74\x72\x69\x62\x75\x74\x00\x65\x20\x56\x42\x5F\x4E\x61\x6D\x00\x65\x20\x3D\x20\x22\x61\x22\x0D\x00\x0A\x61\x62\x63\x64\x65\x66\x67\x02\x68\x06\x80\x0D\x0A";
    const CONTENTS_3: &[u8] = b"Attribute VB_Name = \"a\"\x0D\x0AabcdefghAttribute\x0D\x0A";
    let contents = decompress(CONTAINER_3).unwrap().1;
    assert_eq!(contents, CONTENTS_3);
}

#[test]
fn proj_info_opt_records() {
    // Version 11 of the `[MS-OVBA]` specification introduced an optional
    // `PROJECTCOMPATVERSION` record following the `PROJECTSYSKIND` record. This test
    // verifies that this addition is properly handled by the parser.
    //
    // In addition, this test verifies that the final `PROJECTCONSTANTS` is treated as
    // optional (which it should have been all along).
    //
    // The four test inputs represent the `2x2` matrix of combinations of optional
    // records.

    const INPUT_NONE_NONE: &[u8] = b"\x01\x00\x04\x00\x00\x00\x02\x00\x00\x00\
        \x02\x00\x04\x00\x00\x00\x09\x04\x00\x00\
        \x14\x00\x04\x00\x00\x00\x09\x04\x00\x00\
        \x03\x00\x02\x00\x00\x00\xE4\x04\
        \x04\x00\x01\x00\x00\x00\x41\
        \x05\x00\x01\x00\x00\x00\x41\x40\x00\x02\x00\x00\x00\x41\x00\
        \x06\x00\x00\x00\x00\x00\x3D\x00\x00\x00\x00\x00\
        \x07\x00\x04\x00\x00\x00\x00\x00\x00\x00\
        \x08\x00\x04\x00\x00\x00\x00\x00\x00\x00\
        \x09\x00\x04\x00\x00\x00\x00\x00\x00\x00\x00\x00\
        \x0F\x00\x02\x00\x00\x00\x00\x00\
        \x13\x00\x02\x00\x00\x00\xFF\xFF\
        \x10\x00\
        \x00\x00\x00\x00";
    let res = parse_project_information(INPUT_NONE_NONE);
    assert!(res.is_ok());

    const INPUT_NONE_SOME: &[u8] = b"\x01\x00\x04\x00\x00\x00\x02\x00\x00\x00\
        \x02\x00\x04\x00\x00\x00\x09\x04\x00\x00\
        \x14\x00\x04\x00\x00\x00\x09\x04\x00\x00\
        \x03\x00\x02\x00\x00\x00\xE4\x04\
        \x04\x00\x01\x00\x00\x00\x41\
        \x05\x00\x01\x00\x00\x00\x41\x40\x00\x02\x00\x00\x00\x41\x00\
        \x06\x00\x00\x00\x00\x00\x3D\x00\x00\x00\x00\x00\
        \x07\x00\x04\x00\x00\x00\x00\x00\x00\x00\
        \x08\x00\x04\x00\x00\x00\x00\x00\x00\x00\
        \x09\x00\x04\x00\x00\x00\x00\x00\x00\x00\x00\x00\
        \x0C\x00\x00\x00\x00\x00\x3C\x00\x00\x00\x00\x00\
        \x0F\x00\x02\x00\x00\x00\x00\x00\
        \x13\x00\x02\x00\x00\x00\xFF\xFF\
        \x10\x00\
        \x00\x00\x00\x00";
    let res = parse_project_information(INPUT_NONE_SOME);
    assert!(res.is_ok());
}
